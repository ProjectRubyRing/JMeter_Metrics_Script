<#
.SYNOPSIS
    perf-metrics.sh (性能試験用メトリクス収集) が出力した perf/ ディレクトリを、
    モノトーンの Excel レポート (.xlsx) に変換します (Windows 11 / PowerShell 5.1・7)。

.DESCRIPTION
    ① jboss-cli (DB 接続プール・HTTP・IO ワーカー・スレッドプール・トランザクション・JVM)、
    ② jstat / jcmd (ヒープ・GC)、③ top / ps / cgroup (CPU・メモリ・プロセス・スレッド) の
    CSV を読み込み、次のシートを持つブックを作成します。

      サマリー           収集情報・所見・主要指標 (最大 / 平均 / 95% 値 / 最小)・データソース別・収集プロセスの負荷
      グラフ             HTTP・スレッド / DB 接続プール / ヒープ・GC / CPU・メモリ の推移 (共通の時間軸)
      DB接続プール       データソース別の時系列 (使用率・区間の待ち / 取得失敗など)
      HTTP・スレッド     スループット・エラー・処理時間・IO ワーカー・スレッドプール・セッション
      JVM・GC（CLI）     ヒープ・非ヒープ・メモリプール・スレッド・クラス・GC (platform-mbean)
      ヒープ・GC（jstat） 各領域の使用量・容量、区間の GC 回数・時間・GC 時間率
      ヒープ詳細（jcmd）  GC.heap_info (リージョン・メタスペース)
      CPU・メモリ（top）  ロードアベレージ・CPU 内訳・メモリ・JBoss と収集プロセスの %CPU / RES
      スレッド別CPU       top -H の上位スレッドをスレッド名ごとに集計
      プロセス（ps）      JBoss のスレッド数・RSS・CPU と収集プロセスの合計
      プロセス一覧（ps）  ps の全行 (オートフィルター付き)
      コンテナ（cgroup）  CPU 使用率・スロットリング・メモリ使用量・OOM
      収集情報 / 凡例

    ・フォント Meiryo UI、黒・グレーのモノトーン。グラフは線の種類で系列を区別します。
    ・Excel のインストールは不要です (Office Open XML を直接生成します)。
    ・時刻は既定で「この PC のタイムゾーン」に換算して表示します (JMeter の結果と突き合わせやすくするため)。
      コンテナで記録した時刻のまま表示する場合は -UseRecordedTime を指定します。

    ※ 保存時の文字コードは UTF-8 (BOM 付き・なしのどちらでも可) または ANSI (Shift_JIS) で動作します。

.PARAMETER InputPath
    次のいずれか
      ・perf/ ディレクトリ (perf_meta.csv や cli_*.csv を含む)
      ・その親ディレクトリ (mid/ など)。下にある perf/ をすべて探し、1 つずつブックを作成します
      ・perf-metrics.sh pack で作った .tar.gz / .tar、または perf/ を固めた .zip

.PARAMETER OutputPath
    出力する .xlsx (perf/ が 1 つのとき)、または出力先ディレクトリ (複数のとき)。
    省略時はカレントディレクトリに perf_report_<LOG_ID>.xlsx を作成します。

.PARAMETER Title
    レポートのタイトル (既定: 性能試験メトリクス レポート)。

.PARAMETER From
    この時刻以降のデータだけを使います (例: "2026-10-20 13:00")。JMeter の実行時間帯に絞るときに指定します。

.PARAMETER To
    この時刻以前のデータだけを使います。

.PARAMETER UseRecordedTime
    時刻を PC のタイムゾーンに換算せず、コンテナで記録した時刻 (perf_meta.csv の utc_offset) のまま表示します。

.PARAMETER MaxRows
    1 シートの最大行数 (既定: 200000)。超えた分は CSV を参照してください。

.PARAMETER Open
    作成後に Excel で開きます。

.EXAMPLE
    .\New-PerfMetricsReport.ps1 -InputPath .\perf

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\New-PerfMetricsReport.ps1 .\perf_20261020-abc.tar.gz -From "2026-10-20 13:00" -To "2026-10-20 14:00" -Open
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputPath,
    [Parameter(Position = 1)][string]$OutputPath = '',
    [string]$Title = '',
    [string]$From = '',
    [string]$To = '',
    [switch]$UseRecordedTime,
    [int]$MaxRows = 200000,
    [switch]$Open
)

#------------------------------------------------------------------------------
# Bootstrap (ASCII only).
# Windows PowerShell 5.1 reads a .ps1 saved without a BOM in the ANSI code page
# (CP932 on Japanese Windows) and the Japanese text in the program would break it.
# The program is therefore kept in the block comment below and is read back from
# this file with the right encoding: UTF-8 (with or without BOM), otherwise ANSI.
#------------------------------------------------------------------------------
if (-not $PSCommandPath) { throw 'Run this script from a saved .ps1 file (e.g. powershell -ExecutionPolicy Bypass -File .\New-PerfMetricsReport.ps1 ...).' }
try {
    $BootText = [IO.File]::ReadAllText($PSCommandPath, (New-Object System.Text.UTF8Encoding($false, $true)))
} catch {
    $BootText = [IO.File]::ReadAllText($PSCommandPath, [System.Text.Encoding]::Default)
}
$BootMain = [regex]::Match($BootText, '(?s)<#PERFRPT-MAIN\r?\n(.*?)\r?\nPERFRPT-MAIN#>')
if (-not $BootMain.Success) { throw 'The program block (<#PERFRPT-MAIN ... PERFRPT-MAIN#>) was not found. Copy the whole script again.' }
& ([ScriptBlock]::Create($BootMain.Groups[1].Value)) @PSBoundParameters
return

<#PERFRPT-MAIN
[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$OutputPath = '',
    [string]$Title = '',
    [string]$From = '',
    [string]$To = '',
    [switch]$UseRecordedTime,
    [int]$MaxRows = 200000,
    [switch]$Open
)
if (-not $Title) { $Title = '性能試験メトリクス レポート' }
$ErrorActionPreference = 'Stop'
$Version = '1.0.0'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Inv = [Globalization.CultureInfo]::InvariantCulture
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$NumStyle = [Globalization.NumberStyles]::Float
$XmlNeedRe = [regex]'[&<>"\x01-\x08\x0B\x0C\x0E-\x1F]'
$CtlRe = [regex]'[\x01-\x08\x0B\x0C\x0E-\x1F]'
$NumRe = [regex]'^-?[0-9]+(\.[0-9]+)?([eE][-+]?[0-9]+)?$'

function Write-Info([string]$Message) { Write-Host ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message) }

#------------------------------------------------------------------------------
# スタイル (styles.xml の cellXfs の並び)
#------------------------------------------------------------------------------
$S_TITLE = 1; $S_NOTE = 2; $S_SEC = 3; $S_HDR = 4; $S_TEXT = 5; $S_TEXTC = 6; $S_INT = 7; $S_DEC1 = 8; $S_DEC2 = 9
$S_PCT = 10; $S_DT = 11; $S_KEY = 12; $S_SEVH = 13; $S_SEVM = 14; $S_SEVL = 15; $S_WRAP = 16; $S_BAND = 17; $S_MUTED = 18
$S_DEC3 = 19; $S_TEXTB = 20; $S_NOTEW = 21; $S_DEC1B = 22; $S_INTB = 23; $S_COLKEY = 24; $S_HEAD = 25; $S_PCTB = 26
$FMTS = @{ text = $S_TEXT; textc = $S_TEXTC; int = $S_INT; dec1 = $S_DEC1; dec2 = $S_DEC2; dec3 = $S_DEC3; pct = $S_PCT; dt = $S_DT }
$NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$FONT = '<a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/><a:cs typeface="Meiryo UI"/>'

$StylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="6">
<numFmt numFmtId="164" formatCode="yyyy/mm/dd hh:mm:ss"/>
<numFmt numFmtId="165" formatCode="yyyy/mm/dd hh:mm"/>
<numFmt numFmtId="166" formatCode="#,##0.0"/>
<numFmt numFmtId="167" formatCode="0.0%"/>
<numFmt numFmtId="170" formatCode="#,##0.00"/>
<numFmt numFmtId="171" formatCode="#,##0.000"/>
</numFmts>
<fonts count="9">
<font><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="16"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><sz val="9"/><color rgb="FF595959"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><i/><sz val="9"/><color rgb="FF808080"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><sz val="8"/><color rgb="FF7F7F7F"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="12"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="3"/><charset val="128"/></font>
</fonts>
<fills count="8">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF404040"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF262626"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFBFBFBF"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE7E7E7"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="4">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right><top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>
<border><left/><right/><top/><bottom style="medium"><color rgb="FF000000"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="27">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="5" fillId="0" borderId="3" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="170" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="171" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="166" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="3" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="7" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="167" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="3">
<dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="FF404040"/><bgColor rgb="FF404040"/></patternFill></fill></dxf>
<dxf><font><b/><color rgb="FF000000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFBFBFBF"/><bgColor rgb="FFBFBFBF"/></patternFill></fill></dxf>
<dxf><font><color rgb="FF000000"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFE7E7E7"/><bgColor rgb="FFE7E7E7"/></patternFill></fill></dxf>
</dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@

$COLN = New-Object 'string[]' 729
for ($i = 1; $i -le 728; $i++) {
    $n = $i; $s = ''
    while ($n -gt 0) { $rem = ($n - 1) % 26; $s = [string][char](65 + $rem) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
    $COLN[$i] = $s
}

#------------------------------------------------------------------------------
# 共通関数
#------------------------------------------------------------------------------
function X([string]$s) {
    if ([string]::IsNullOrEmpty($s) -or -not $XmlNeedRe.IsMatch($s)) { return $s }
    $s = $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
    return $CtlRe.Replace($s, '')
}
function Split-Csv([string]$line) {
    if ($line.IndexOf([char]34) -lt 0) { return , $line.Split(',') }
    $list = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inq = $false; $n = $line.Length
    for ($i = 0; $i -lt $n; $i++) {
        $ch = $line[$i]
        if ($inq) {
            if ($ch -eq [char]34) {
                if ($i + 1 -lt $n -and $line[$i + 1] -eq [char]34) { [void]$sb.Append([char]34); $i++ } else { $inq = $false }
            } else { [void]$sb.Append($ch) }
        } elseif ($ch -eq [char]34) { $inq = $true }
        elseif ($ch -eq [char]44) { $list.Add($sb.ToString()); [void]$sb.Clear() }
        else { [void]$sb.Append($ch) }
    }
    $list.Add($sb.ToString())
    return , $list.ToArray()
}
function NStr($v) {
    $d = [double]$v
    if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return '' }
    return $d.ToString('R', $Inv)
}

function SerialText([double]$x) {
    $dt = [DateTime]::FromOADate($x)
    return $dt.ToString('yyyy/MM/dd HH:mm:ss', $Inv)
}
function Units([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return 0.0 }
    $nb = $Utf8.GetByteCount($s); $asc = 0
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -lt 128) { $asc++ } }
    $na = $nb - $asc
    return ($nb - $na) + $na / 3 * 2
}
function WrapHt([string]$s, [double]$width) { $l = [Math]::Truncate((Units $s) / ($width * 0.95)) + 1; return $l * 15 + 3 }

# セル
function StrC([string]$ref, [string]$s, [int]$st) { return '<c r="' + $ref + '" s="' + $st + '" t="inlineStr"><is><t xml:space="preserve">' + (X $s) + '</t></is></c>' }
function NumC([string]$ref, $v, [int]$st) { if ($null -eq $v) { return '<c r="' + $ref + '" s="' + $st + '"/>' }; $t = NStr $v; if ($t -eq '') { return '<c r="' + $ref + '" s="' + $st + '"/>' }; return '<c r="' + $ref + '" s="' + $st + '"><v>' + $t + '</v></c>' }
function EmptyC([string]$ref, [int]$st) { return '<c r="' + $ref + '" s="' + $st + '"/>' }
function ValC([string]$ref, $v, [string]$fmt) {
    $st = $FMTS[$fmt]; if ($null -eq $st) { $st = $S_TEXT }
    if ($null -eq $v) { return EmptyC $ref $st }
    if ($v -is [string]) {
        if ($fmt -eq 'text' -or $fmt -eq 'textc') { return StrC $ref $v $st }
        $d = $PerfRpt::D($v); if ($null -eq $d) { return StrC $ref $v $S_TEXT }; return NumC $ref $d $st
    }
    if ($fmt -eq 'text' -or $fmt -eq 'textc') { return StrC $ref (NStr $v) $st }
    return NumC $ref $v $st
}
function Row([int]$r, [string]$cells, [double]$ht) {
    if ($ht -gt 0) { return '<row r="' + $r + '" ht="' + $ht.ToString($Inv) + '" customHeight="1">' + $cells + '</row>' }
    return '<row r="' + $r + '">' + $cells + '</row>'
}
function Ws-Begin([string]$tab, [bool]$fit, [int]$xs, [int]$ys, [bool]$selected, [int]$zoom) {
    $s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<worksheet ' + $NS + '>'
    $s += '<sheetPr><tabColor rgb="' + $tab + '"/>' + $(if ($fit) { '<pageSetUpPr fitToPage="1"/>' } else { '' }) + '</sheetPr>'
    $s += '<sheetViews><sheetView' + $(if ($selected) { ' tabSelected="1"' } else { '' }) + ' showGridLines="0" zoomScale="' + $zoom + '" zoomScaleNormal="' + $zoom + '" workbookViewId="0">'
    if ($xs -gt 0 -or $ys -gt 0) {
        $tl = $COLN[$xs + 1] + ($ys + 1)
        $pane = $(if ($xs -gt 0 -and $ys -gt 0) { 'bottomRight' } elseif ($ys -gt 0) { 'bottomLeft' } else { 'topRight' })
        $s += '<pane' + $(if ($xs -gt 0) { ' xSplit="' + $xs + '"' } else { '' }) + $(if ($ys -gt 0) { ' ySplit="' + $ys + '"' } else { '' }) + ' topLeftCell="' + $tl + '" activePane="' + $pane + '" state="frozen"/>'
        $s += '<selection pane="' + $pane + '" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
    } else { $s += '<selection activeCell="A1" sqref="A1"/>' }
    $s += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
    return $s
}
function Ws-End([string]$orient, [bool]$fit, [bool]$drawing, [string]$repTitle, [string]$extra) {
    $s = $extra
    $s += '<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'
    $s += '<pageSetup paperSize="9" orientation="' + $orient + '"' + $(if ($fit) { ' fitToWidth="1" fitToHeight="0"' } else { '' }) + '/>'
    $hf = '&L&"Meiryo UI,Regular"&9' + $repTitle + '&R&"Meiryo UI,Regular"&9&A'
    $s += '<headerFooter><oddHeader>' + (X $hf) + '</oddHeader><oddFooter>' + (X '&C&"Meiryo UI,Regular"&9&P / &N') + '</oddFooter></headerFooter>'
    if ($drawing) { $s += '<drawing r:id="rId1"/>' }
    return $s + '</worksheet>'
}
function Stat($list) {
    if ($null -eq $list -or $list.Count -eq 0) { return $null }
    $arr = $list.ToArray(); [Array]::Sort($arr)
    $n = $arr.Length; $sum = 0.0
    foreach ($v in $arr) { $sum += $v }
    $k = [int][Math]::Ceiling($n * 0.95) - 1; if ($k -lt 0) { $k = 0 }
    return [pscustomobject]@{ Max = $arr[$n - 1]; Min = $arr[0]; Avg = $sum / $n; P95 = $arr[$k]; N = $n; Sum = $sum }
}
function NewList() { return , (New-Object 'System.Collections.Generic.List[double]') }

#------------------------------------------------------------------------------
# 高速化: 行ごとに何十回も呼ぶ処理 (CSV の分解・数値化・増分・丸め・時刻・行の XML 出力) は
# .NET の静的メソッドにまとめる。PowerShell の関数は 1 回 100µs 以上かかり、24 時間分
# (数十万行) では数十分になるため。C# (Add-Type) を優先し、使えない環境 (csc の実行が
# 禁止されている等) では同じ API の PowerShell クラスへ自動で切り替える (結果は同じ)。
#------------------------------------------------------------------------------
$PerfRptCsSource = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

public sealed class PerfRptTableV1 { public string Name; public string[] H; public Hashtable I; public List<string[]> Rows; }

public static class PerfRptCsV1 {
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly Regex XmlNeed = new Regex("[&<>\"\\x01-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    static readonly Regex Ctl = new Regex("[\\x01-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    static readonly Dictionary<long, double> OffCache = new Dictionary<long, double>();
    static PropertyInfo BaseProp;
    public static bool UseRecorded = false;
    public static double RecOffset = 0;
    public static void ResetTz() { OffCache.Clear(); }

    static object Unwrap(object o) {
        if (o == null) return null;
        Type t = o.GetType();
        if (t.FullName == "System.Management.Automation.PSObject") {
            if (BaseProp == null) BaseProp = t.GetProperty("BaseObject");
            return BaseProp.GetValue(o, null);
        }
        return o;
    }
    static bool Num(object o, out double v) {
        v = 0; o = Unwrap(o);
        if (o == null) return false;
        if (o is double) { v = (double)o; return true; }
        if (o is int) { v = (int)o; return true; }
        if (o is long) { v = (long)o; return true; }
        if (o is decimal) { v = (double)(decimal)o; return true; }
        if (o is float) { v = (float)o; return true; }
        string s = o as string;
        if (s == null) { try { v = Convert.ToDouble(o, Inv); return true; } catch { return false; } }
        if (s.Length == 0 || s == "-") return false;
        return double.TryParse(s, NumberStyles.Float, Inv, out v);
    }
    public static object D(object s) { double v; if (Num(s, out v)) return v; return null; }
    public static string[] SplitCsv(string line) {
        if (line.IndexOf('"') < 0) return line.Split(',');
        List<string> list = new List<string>(); StringBuilder sb = new StringBuilder(); bool inq = false;
        for (int i = 0; i < line.Length; i++) {
            char c = line[i];
            if (inq) {
                if (c == '"') { if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; } else inq = false; }
                else sb.Append(c);
            } else if (c == '"') inq = true;
            else if (c == ',') { list.Add(sb.ToString()); sb.Length = 0; }
            else sb.Append(c);
        }
        list.Add(sb.ToString());
        return list.ToArray();
    }
    public static PerfRptTableV1 ReadCsv(string path, string name, double from, double to) {
        if (!File.Exists(path)) return null;
        PerfRptTableV1 t = new PerfRptTableV1(); t.Name = name; t.I = new Hashtable(); t.Rows = new List<string[]>();
        using (StreamReader r = new StreamReader(path, new UTF8Encoding(false))) {
            string line = r.ReadLine();
            if (line == null) return null;
            t.H = SplitCsv(line);
            for (int i = 0; i < t.H.Length; i++) t.I[t.H[i]] = i;
            int ei = t.I.ContainsKey("epoch") ? (int)t.I["epoch"] : -1;
            int nh = t.H.Length;
            while ((line = r.ReadLine()) != null) {
                if (line.Length == 0) continue;
                string[] f = SplitCsv(line);
                if (f.Length < nh) continue;
                if (ei >= 0) {
                    double e;
                    if (!double.TryParse(f[ei], NumberStyles.Float, Inv, out e)) continue;
                    if (e < from || e > to) continue;
                }
                t.Rows.Add(f);
            }
        }
        return t;
    }
    public static string G(object t, object row, string key) {
        PerfRptTableV1 tb = (PerfRptTableV1)Unwrap(t); string[] r = (string[])Unwrap(row);
        object o = tb.I[key]; if (o == null) return null;
        return r[(int)o];
    }
    public static object GD(object t, object row, string key) { return D(G(t, row, key)); }
    public static object Dl(object cur, object prev) {
        double c, p; if (!Num(cur, out c) || !Num(prev, out p)) return null;
        double d = c - p; return d < 0 ? c : d;
    }
    public static object Div(object a, object b) { double x, y; if (!Num(a, out x) || !Num(b, out y) || y == 0) return null; return x / y; }
    public static object Rnd(object v, int n) { double x; if (!Num(v, out x)) return null; return Math.Round(x, n); }
    public static object MB(object b) { double x; if (!Num(b, out x)) return null; return Math.Round(x / 1048576.0, 1); }
    public static object KBtoMB(object kb) { double x; if (!Num(kb, out x)) return null; return Math.Round(x / 1024.0, 1); }
    public static void AddV(List<double> list, object v) { double x; if (Num(v, out x)) list.Add(x); }
    public static double XT(double ep) {
        if (UseRecorded) return (ep + RecOffset) / 86400.0 + 25569.0;
        long h = (long)Math.Floor(ep / 3600);
        double off;
        if (!OffCache.TryGetValue(h, out off)) {
            off = TimeZoneInfo.Local.GetUtcOffset(DateTimeOffset.FromUnixTimeSeconds((long)ep).UtcDateTime).TotalSeconds;
            OffCache[h] = off;
        }
        return (ep + off) / 86400.0 + 25569.0;
    }
    public static string X(string s) {
        if (string.IsNullOrEmpty(s) || !XmlNeed.IsMatch(s)) return s;
        s = s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        return Ctl.Replace(s, "");
    }
    public static string NStr(object v) {
        double d; if (!Num(v, out d)) return "";
        if (double.IsNaN(d) || double.IsInfinity(d)) return "";
        return d.ToString("R", Inv);
    }
    // 表の行を <row> の XML にして書き出す (fm / st は 0 始まりの列ごとの書式・スタイル番号)
    public static void WriteRows(TextWriter w, IList rows, string[] fm, int[] st, int textStyle, int firstRow, string[] coln) {
        StringBuilder sb = new StringBuilder(8192);
        int nc = fm.Length;
        for (int i = 0; i < rows.Count; i++) {
            object[] vals = (object[])Unwrap(rows[i]);
            string rs = (firstRow + i).ToString(Inv);
            sb.Length = 0;
            sb.Append("<row r=\"").Append(rs).Append("\">");
            for (int c = 0; c < nc; c++) {
                object v = c < vals.Length ? Unwrap(vals[c]) : null;
                string rf = coln[c + 1] + rs;
                if (v == null) { sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(st[c]).Append("\"/>"); continue; }
                string f = fm[c];
                if (f == "text" || f == "textc") {
                    string sv = v as string; if (sv == null) sv = Convert.ToString(v, Inv);
                    sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(st[c]).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(X(sv)).Append("</t></is></c>");
                } else if (v is string) {
                    double d;
                    if (!Num(v, out d)) sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(textStyle).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(X((string)v)).Append("</t></is></c>");
                    else sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(st[c]).Append("\"><v>").Append(NStr(d)).Append("</v></c>");
                } else {
                    string t = NStr(v);
                    if (t.Length == 0) sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(st[c]).Append("\"/>");
                    else sb.Append("<c r=\"").Append(rf).Append("\" s=\"").Append(st[c]).Append("\"><v>").Append(t).Append("</v></c>");
                }
            }
            sb.Append("</row>");
            w.Write(sb.ToString());
        }
    }
}
'@

# 同じ API の PowerShell 版 (Add-Type が使えないとき)
class PerfRptTablePsV1 { [string]$Name; [string[]]$H; [hashtable]$I; [System.Collections.Generic.List[string[]]]$Rows }
class PerfRptPsV1 {
    static [Globalization.CultureInfo]$Inv = [Globalization.CultureInfo]::InvariantCulture
    static [regex]$CsvRe = [regex]'(?:^|,)(?:"((?:[^"]|"")*)"|([^,]*))'
    static [regex]$XmlNeed = [regex]'[&<>"\x01-\x08\x0B\x0C\x0E-\x1F]'
    static [regex]$Ctl = [regex]'[\x01-\x08\x0B\x0C\x0E-\x1F]'
    static [hashtable]$Off = @{}
    static [bool]$UseRecorded = $false
    static [double]$RecOffset = 0
    static [void] ResetTz() { [PerfRptPsV1]::Off = @{} }
    static [object] D([object]$s) {
        if ($null -eq $s) { return $null }
        if ($s -is [double]) { return $s }
        if ($s -is [int] -or $s -is [long] -or $s -is [decimal] -or $s -is [single]) { return [double]$s }
        $t = [string]$s
        if ($t.Length -eq 0 -or $t -eq '-') { return $null }
        $v = 0.0
        if ([double]::TryParse($t, [Globalization.NumberStyles]::Float, [PerfRptPsV1]::Inv, [ref]$v)) { return $v }
        return $null
    }
    static [string[]] SplitCsv([string]$line) {
        if ($line.IndexOf([char]34) -lt 0) { return $line.Split(',') }
        $list = New-Object System.Collections.Generic.List[string]
        foreach ($m in [PerfRptPsV1]::CsvRe.Matches($line)) {
            if ($m.Groups[1].Success) { $list.Add($m.Groups[1].Value.Replace('""', '"')) } else { $list.Add($m.Groups[2].Value) }
        }
        return $list.ToArray()
    }
    static [object] ReadCsv([string]$path, [string]$name, [double]$from, [double]$to) {
        if (-not [IO.File]::Exists($path)) { return $null }
        $t = [PerfRptTablePsV1]::new(); $t.Name = $name; $t.I = @{}; $t.Rows = New-Object 'System.Collections.Generic.List[string[]]'
        $first = $true; $ei = -1; $nh = 0
        foreach ($line in [IO.File]::ReadLines($path, (New-Object System.Text.UTF8Encoding($false)))) {
            if ($first) {
                $first = $false; $t.H = [PerfRptPsV1]::SplitCsv($line)
                for ($i = 0; $i -lt $t.H.Length; $i++) { $t.I[$t.H[$i]] = $i }
                if ($t.I.ContainsKey('epoch')) { $ei = [int]$t.I['epoch'] }
                $nh = $t.H.Length
                continue
            }
            if ($line.Length -eq 0) { continue }
            $f = [PerfRptPsV1]::SplitCsv($line)
            if ($f.Length -lt $nh) { continue }
            if ($ei -ge 0) {
                $e = 0.0
                if (-not [double]::TryParse($f[$ei], [Globalization.NumberStyles]::Float, [PerfRptPsV1]::Inv, [ref]$e)) { continue }
                if ($e -lt $from -or $e -gt $to) { continue }
            }
            $t.Rows.Add($f)
        }
        if ($first) { return $null }
        return $t
    }
    static [string] G([object]$t, [object]$row, [string]$key) { $o = $t.I[$key]; if ($null -eq $o) { return $null }; return $row[[int]$o] }
    static [object] GD([object]$t, [object]$row, [string]$key) { return [PerfRptPsV1]::D([PerfRptPsV1]::G($t, $row, $key)) }
    static [object] Dl([object]$cur, [object]$prev) {
        $c = [PerfRptPsV1]::D($cur); $p = [PerfRptPsV1]::D($prev)
        if ($null -eq $c -or $null -eq $p) { return $null }
        $d = $c - $p; if ($d -lt 0) { return $c }; return $d
    }
    static [object] Div([object]$a, [object]$b) {
        $x = [PerfRptPsV1]::D($a); $y = [PerfRptPsV1]::D($b)
        if ($null -eq $x -or $null -eq $y -or $y -eq 0) { return $null }; return $x / $y
    }
    static [object] Rnd([object]$v, [int]$n) { $x = [PerfRptPsV1]::D($v); if ($null -eq $x) { return $null }; return [Math]::Round([double]$x, $n) }
    static [object] MB([object]$b) { $x = [PerfRptPsV1]::D($b); if ($null -eq $x) { return $null }; return [Math]::Round([double]$x / 1048576.0, 1) }
    static [object] KBtoMB([object]$kb) { $x = [PerfRptPsV1]::D($kb); if ($null -eq $x) { return $null }; return [Math]::Round([double]$x / 1024.0, 1) }
    static [void] AddV([System.Collections.Generic.List[double]]$list, [object]$v) { $x = [PerfRptPsV1]::D($v); if ($null -ne $x) { $list.Add([double]$x) } }
    static [double] XT([double]$ep) {
        if ([PerfRptPsV1]::UseRecorded) { return ($ep + [PerfRptPsV1]::RecOffset) / 86400.0 + 25569.0 }
        $h = [long][Math]::Floor($ep / 3600)
        if (-not [PerfRptPsV1]::Off.ContainsKey($h)) {
            [PerfRptPsV1]::Off[$h] = [TimeZoneInfo]::Local.GetUtcOffset([DateTimeOffset]::FromUnixTimeSeconds([long]$ep).UtcDateTime).TotalSeconds
        }
        return ($ep + [double][PerfRptPsV1]::Off[$h]) / 86400.0 + 25569.0
    }
    static [string] X([string]$s) {
        if ([string]::IsNullOrEmpty($s) -or -not [PerfRptPsV1]::XmlNeed.IsMatch($s)) { return $s }
        $s = $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
        return [PerfRptPsV1]::Ctl.Replace($s, '')
    }
    static [string] NStr([object]$v) {
        $x = [PerfRptPsV1]::D($v); if ($null -eq $x) { return '' }
        $d = [double]$x
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return '' }
        return $d.ToString('R', [PerfRptPsV1]::Inv)
    }
    static [void] WriteRows([IO.TextWriter]$w, [System.Collections.IList]$rows, [string[]]$fm, [int[]]$st, [int]$textStyle, [int]$firstRow, [string[]]$coln) {
        $sb = New-Object System.Text.StringBuilder 8192
        $nc = $fm.Length
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $vals = $rows[$i]
            $rs = [string]($firstRow + $i)
            [void]$sb.Clear(); [void]$sb.Append('<row r="').Append($rs).Append('">')
            for ($c = 0; $c -lt $nc; $c++) {
                $v = $null; if ($c -lt $vals.Length) { $v = $vals[$c] }
                $rf = $coln[$c + 1] + $rs
                if ($null -eq $v) { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"/>'); continue }
                $f = $fm[$c]
                if ($f -eq 'text' -or $f -eq 'textc') {
                    [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('" t="inlineStr"><is><t xml:space="preserve">').Append([PerfRptPsV1]::X([string]$v)).Append('</t></is></c>')
                } elseif ($v -is [string]) {
                    $d = [PerfRptPsV1]::D($v)
                    if ($null -eq $d) { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($textStyle).Append('" t="inlineStr"><is><t xml:space="preserve">').Append([PerfRptPsV1]::X($v)).Append('</t></is></c>') }
                    else { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"><v>').Append([PerfRptPsV1]::NStr($d)).Append('</v></c>') }
                } else {
                    $t = [PerfRptPsV1]::NStr($v)
                    if ($t.Length -eq 0) { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"/>') }
                    else { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"><v>').Append($t).Append('</v></c>') }
                }
            }
            [void]$sb.Append('</row>')
            $w.Write($sb.ToString())
        }
    }
}

$PerfRpt = $null
if ($env:PERFRPT_NO_ADDTYPE -ne '1') {
    try {
        if (-not ('PerfRptCsV1' -as [type])) { Add-Type -TypeDefinition $PerfRptCsSource -Language CSharp -ErrorAction Stop }
        $PerfRpt = [PerfRptCsV1]
    } catch {
        Write-Warning ('C# のコンパイル (Add-Type) が使えないため、PowerShell だけで処理します (大きなデータでは時間がかかります): ' + $_.Exception.Message)
    }
}
if ($null -eq $PerfRpt) { $PerfRpt = [PerfRptPsV1] }

#------------------------------------------------------------------------------
# 入力 (ディレクトリ / 親ディレクトリ / .zip / .tar.gz)
#------------------------------------------------------------------------------
$TempDirs = New-Object System.Collections.Generic.List[string]
function Test-PerfDir([string]$d) {
    if (Test-Path -LiteralPath (Join-Path $d 'perf_meta.csv')) { return $true }
    return (@(Get-ChildItem -LiteralPath $d -Filter 'cli_*.csv' -File -ErrorAction SilentlyContinue).Count -gt 0) -or
           (@(Get-ChildItem -LiteralPath $d -Filter 'os_*.csv' -File -ErrorAction SilentlyContinue).Count -gt 0)
}
function Resolve-PerfDirs([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { throw "入力が見つかりません: $path" }
    $item = Get-Item -LiteralPath $path
    $root = $item.FullName
    if (-not $item.PSIsContainer) {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ('perfrpt_' + [Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tmp | Out-Null
        $TempDirs.Add($tmp)
        $nm = $item.Name.ToLowerInvariant()
        if ($nm.EndsWith('.zip')) {
            [IO.Compression.ZipFile]::ExtractToDirectory($item.FullName, $tmp)
        } elseif ($nm.EndsWith('.tar.gz') -or $nm.EndsWith('.tgz') -or $nm.EndsWith('.tar')) {
            $tarExe = Join-Path $env:SystemRoot 'System32\tar.exe'
            if (-not (Test-Path -LiteralPath $tarExe)) { $tarExe = 'tar.exe' }
            & $tarExe -xf $item.FullName -C $tmp
            if ($LASTEXITCODE -ne 0) { throw "展開に失敗しました: $($item.FullName)" }
        } elseif ($nm.EndsWith('.csv')) {
            return , @($item.DirectoryName)
        } else {
            throw "対応していない入力です (ディレクトリ / .zip / .tar.gz / .tar): $path"
        }
        $root = $tmp
    }
    if (Test-PerfDir $root) { return , @($root) }
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($f in (Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'perf_meta.csv' -ErrorAction SilentlyContinue)) { $found.Add($f.DirectoryName) }
    if ($found.Count -eq 0) {
        foreach ($f in (Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'cli_jvm.csv' -ErrorAction SilentlyContinue)) { $found.Add($f.DirectoryName) }
        foreach ($f in (Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'os_ps.csv' -ErrorAction SilentlyContinue)) { if (-not $found.Contains($f.DirectoryName)) { $found.Add($f.DirectoryName) } }
    }
    if ($found.Count -eq 0) { throw "perf-metrics の出力 (perf_meta.csv / cli_*.csv / os_*.csv) が見つかりません: $path" }
    return , ($found.ToArray() | Sort-Object)
}

# CSV を読む。epoch で時間帯を絞り、途中で切れた行 (コンテナ停止時) は捨てる。
function Read-PerfCsv([string]$dir, [string]$name) {
    return $PerfRpt::ReadCsv((Join-Path $dir $name), $name, $script:FromEpoch, $script:ToEpoch)
}
# キーごとに行をまとめる (初出順)
function Group-By($t, [string[]]$keys) {
    $od = New-Object System.Collections.Specialized.OrderedDictionary
    if ($null -eq $t) { return $od }
    $ix = @(); foreach ($k in $keys) { $ix += $t.I[$k] }
    foreach ($r in $t.Rows) {
        $parts = @(); foreach ($i in $ix) { $parts += $r[$i] }
        $k = $parts -join "`t"
        if (-not $od.Contains($k)) { $od[$k] = New-Object System.Collections.Generic.List[object] }
        $od[$k].Add($r)
    }
    return $od
}

#------------------------------------------------------------------------------
# シート (表) のモデル
#------------------------------------------------------------------------------
function New-Sheet([string]$name, [string]$title, [string]$note, [int]$freezeCols) {
    return [pscustomobject]@{
        Name = $name; Title = $title; Note = $note; FreezeCols = $freezeCols
        Cols = (New-Object System.Collections.Generic.List[object])
        Rows = (New-Object System.Collections.Generic.List[object])
        Groups = (New-Object System.Collections.Generic.List[object])
        CF = (New-Object System.Collections.Generic.List[object])
        Truncated = $false; Index = 0
    }
}
# 列を足して列番号 (1 始まり) を返す。 cat=分類 label=見出し key=元の CSV 列名 fmt=書式 w=列幅
function Add-Col($sh, [string]$cat, [string]$label, [string]$key, [string]$fmt, [double]$w) {
    $sh.Cols.Add([pscustomobject]@{ Cat = $cat; Label = $label; Key = $key; Fmt = $fmt; W = $w })
    return $sh.Cols.Count
}
function Add-Row($sh, [object[]]$vals, [string]$group) {
    if ($sh.Rows.Count -ge $MaxRows) { $sh.Truncated = $true; return }
    $n = $sh.Rows.Count
    if ($sh.Groups.Count -eq 0 -or $sh.Groups[$sh.Groups.Count - 1].Key -ne $group) {
        $sh.Groups.Add([pscustomobject]@{ Key = $group; Start = $n; End = $n; R1 = 0; R2 = 0 })
    } else { $sh.Groups[$sh.Groups.Count - 1].End = $n }
    $sh.Rows.Add($vals)
}
# 条件付き書式 (列番号, 種類, 値, dxf)  種類: ge / gt / bar
function Add-CF($sh, [int]$col, [string]$kind, [string]$val, [int]$dxf) { $sh.CF.Add(@($col, $kind, $val, $dxf)) }
$DATA_ROW = 7
function Fix-Ranges($sh) { foreach ($g in $sh.Groups) { $g.R1 = $DATA_ROW + $g.Start; $g.R2 = $DATA_ROW + $g.End } }

function Write-TableSheet([System.IO.Compression.ZipArchive]$zip, $sh, [int]$sheetIdx) {
    $e = $zip.CreateEntry("xl/worksheets/sheet$sheetIdx.xml", [IO.Compression.CompressionLevel]::Optimal)
    $w = New-Object IO.StreamWriter($e.Open(), $Utf8)
    $nc = $sh.Cols.Count; $nr = $sh.Rows.Count
    $fit = $true
    $w.Write((Ws-Begin 'FF404040' $fit $(if ($nr -gt 0) { $sh.FreezeCols } else { 0 }) $(if ($nr -gt 0) { 6 } else { 0 }) $false 90))
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<cols>')
    for ($c = 1; $c -le $nc; $c++) { [void]$sb.Append('<col min="' + $c + '" max="' + $c + '" width="' + ([double]$sh.Cols[$c - 1].W).ToString($Inv) + '" customWidth="1"/>') }
    [void]$sb.Append('</cols><sheetData>')
    [void]$sb.Append((Row 1 (StrC 'A1' $sh.Title $S_TITLE) 24))
    $note = $sh.Note
    if ($sh.Truncated) { $note += '（行数が多いため先頭 ' + $MaxRows + ' 行だけを載せています。残りは CSV を参照してください）' }
    [void]$sb.Append((Row 2 (StrC 'A2' $note $S_NOTE) 0))
    $merges = New-Object System.Collections.Generic.List[string]
    # 4 行目: 分類 (同じ分類が続く列は結合)、5 行目: 見出し、6 行目: 元の CSV 列名
    $out = New-Object System.Text.StringBuilder
    $c = 1
    while ($c -le $nc) {
        $j = $c; while ($j -lt $nc -and $sh.Cols[$j].Cat -ceq $sh.Cols[$c - 1].Cat) { $j++ }
        [void]$out.Append((StrC ($COLN[$c] + '4') $sh.Cols[$c - 1].Cat $S_BAND))
        for ($i = $c + 1; $i -le $j; $i++) { [void]$out.Append((EmptyC ($COLN[$i] + '4') $S_BAND)) }
        if ($j -gt $c) { $merges.Add($COLN[$c] + '4:' + $COLN[$j] + '4') }
        $c = $j + 1
    }
    [void]$sb.Append((Row 4 $out.ToString() 18))
    [void]$out.Clear(); for ($c = 1; $c -le $nc; $c++) { [void]$out.Append((StrC ($COLN[$c] + '5') $sh.Cols[$c - 1].Label $S_HDR)) }
    [void]$sb.Append((Row 5 $out.ToString() 32))
    [void]$out.Clear(); for ($c = 1; $c -le $nc; $c++) { [void]$out.Append((StrC ($COLN[$c] + '6') $sh.Cols[$c - 1].Key $S_COLKEY)) }
    [void]$sb.Append((Row 6 $out.ToString() 13.5))
    $w.Write($sb.ToString())
    if ($nr -eq 0) {
        $w.Write((Row 7 (StrC 'A7' '（この期間のデータがありません）' $S_NOTE) 0))
    }
    $fm = New-Object 'string[]' $nc; $st = New-Object 'int[]' $nc
    for ($c = 0; $c -lt $nc; $c++) { $fm[$c] = $sh.Cols[$c].Fmt; $s2 = $FMTS[$fm[$c]]; if ($null -eq $s2) { $s2 = $S_TEXT }; $st[$c] = $s2 }
    $PerfRpt::WriteRows($w, $sh.Rows, $fm, $st, $S_TEXT, $DATA_ROW, $COLN)
    $w.Write('</sheetData>')
    $extra = ''
    $last = $DATA_ROW + [Math]::Max($nr, 1) - 1
    if ($nr -gt 0) {
        $extra += '<autoFilter ref="A6:' + $COLN[$nc] + $last + '"/>'
        Add-DefName $sheetIdx '_xlnm._FilterDatabase' ("'" + $sh.Name + "'!`$A`$6:`$" + $COLN[$nc] + '$' + $last)
    }
    if ($merges.Count -gt 0) {
        $extra += '<mergeCells count="' + $merges.Count + '">'
        foreach ($m in $merges) { $extra += '<mergeCell ref="' + $m + '"/>' }
        $extra += '</mergeCells>'
    }
    if ($nr -gt 0 -and $sh.CF.Count -gt 0) {
        $prio = 1
        foreach ($cf in $sh.CF) {
            $col = $COLN[[int]$cf[0]]; $sq = $col + $DATA_ROW + ':' + $col + $last
            $extra += '<conditionalFormatting sqref="' + $sq + '">'
            switch ($cf[1]) {
                'ge' { $extra += '<cfRule type="cellIs" dxfId="' + $cf[3] + '" priority="' + $prio + '" operator="greaterThanOrEqual"><formula>' + $cf[2] + '</formula></cfRule>' }
                'gt' { $extra += '<cfRule type="cellIs" dxfId="' + $cf[3] + '" priority="' + $prio + '" operator="greaterThan"><formula>' + $cf[2] + '</formula></cfRule>' }
                'bar' { $extra += '<cfRule type="dataBar" priority="' + $prio + '"><dataBar><cfvo type="num" val="0"/><cfvo type="num" val="' + $cf[2] + '"/><color rgb="FFA6A6A6"/></dataBar></cfRule>' }
            }
            $extra += '</conditionalFormatting>'
            $prio++
        }
    }
    Add-DefName $sheetIdx '_xlnm.Print_Titles' ("'" + $sh.Name + "'!`$4:`$6")
    $w.Write((Ws-End 'landscape' $fit $false $sh.Title $extra))
    $w.Dispose()
}

$DefNames = New-Object System.Collections.Generic.List[object]
function Add-DefName([int]$sheetIdx, [string]$name, [string]$ref) { $DefNames.Add(@(($sheetIdx - 1), $name, $ref)) }

#------------------------------------------------------------------------------
# グラフ (DrawingML 散布図。時間軸は全グラフ共通)
#------------------------------------------------------------------------------
$LineStyles = @(
    @('000000', 22225, 'solid'), @('595959', 15875, 'dash'), @('8C8C8C', 15875, 'sysDot'),
    @('262626', 12700, 'lgDash'), @('A6A6A6', 15875, 'dashDot'), @('404040', 12700, 'sysDash'), @('7F7F7F', 12700, 'lgDashDot')
)
function C-Title([string]$t, [int]$sz) {
    return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + $sz + '" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' + $FONT + '</a:defRPr></a:pPr><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="' + $sz + '" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>' + $FONT + '</a:rPr><a:t>' + (X $t) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
}
function Ax-Title([string]$t) {
    return '<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="0"><a:solidFill><a:srgbClr val="404040"/></a:solidFill>' + $FONT + '</a:defRPr></a:pPr><a:r><a:rPr lang="ja-JP" altLang="en-US" sz="800" b="0"><a:solidFill><a:srgbClr val="404040"/></a:solidFill>' + $FONT + '</a:rPr><a:t>' + (X $t) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
}
function Tx-Pr([int]$sz, [string]$rot) {
    return '<c:txPr><a:bodyPr' + $(if ($rot -ne '') { ' rot="' + $rot + '" vert="horz"' } else { '' }) + '/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + $sz + '"><a:solidFill><a:srgbClr val="404040"/></a:solidFill>' + $FONT + '</a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>'
}
$AXLN = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>'
$GRID = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
function Rng([string]$sheet, [int]$col, [int]$r1, [int]$r2) { return "'" + $sheet + "'!`$" + $COLN[$col] + '$' + $r1 + ':$' + $COLN[$col] + '$' + $r2 }
function Ser-Sc([int]$i, $s) {
    $ln = $LineStyles[$i % $LineStyles.Count]
    $sp = '<c:spPr><a:ln w="' + $ln[1] + '" cap="rnd"><a:solidFill><a:srgbClr val="' + $ln[0] + '"/></a:solidFill><a:prstDash val="' + $ln[2] + '"/><a:round/></a:ln></c:spPr>'
    return '<c:ser><c:idx val="' + $i + '"/><c:order val="' + $i + '"/><c:tx><c:v>' + (X $s.N) + '</c:v></c:tx>' + $sp + '<c:marker><c:symbol val="none"/></c:marker><c:xVal><c:numRef><c:f>' + (X (Rng $s.S $s.X $s.R1 $s.R2)) + '</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>' + (X (Rng $s.S $s.Y $s.R1 $s.R2)) + '</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>'
}
function X-Scale() {
    $script:XSCL = ''; $script:XMAJ = ''; $script:XFMT = 'hh:mm'
    if ($null -eq $script:TMin -or $null -eq $script:TMax) { return }
    $x1 = $script:TMin; $x2 = $script:TMax; $span = $x2 - $x1
    if ($span -le 0) { $span = 1.0 / 1440 }
    $units = @((1 / 1440), (2 / 1440), (5 / 1440), (10 / 1440), (15 / 1440), (30 / 1440), (1 / 24), (2 / 24), (3 / 24), (6 / 24), (12 / 24), 1, 2, 7)
    $u = $units[$units.Count - 1]
    foreach ($cand in $units) { if ($span / $cand -le 12) { $u = $cand; break } }
    $lo = [Math]::Floor($x1 / $u + 1e-9) * $u
    $hi = [Math]::Ceiling($x2 / $u - 1e-9) * $u; if ($hi -le $lo) { $hi = $lo + $u }
    $script:XSCL = '<c:max val="' + $hi.ToString('R', $Inv) + '"/><c:min val="' + $lo.ToString('R', $Inv) + '"/>'
    $script:XMAJ = $u.ToString('R', $Inv)
    if ($span -ge 1.5) { $script:XFMT = 'm/d hh:mm' }
    if ($u -ge 1) { $script:XFMT = 'm/d' }
}
function Chart-Xml($ch, [int]$n) {
    $ax1 = 500000 + $n * 10 + 1; $ax2 = $ax1 + 1
    $plot = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
    $i = 0; foreach ($s in $ch.Series) { $plot += Ser-Sc $i $s; $i++ }
    $plot += '<c:axId val="' + $ax1 + '"/><c:axId val="' + $ax2 + '"/></c:scatterChart>'
    # X (時刻)
    $plot += '<c:valAx><c:axId val="' + $ax1 + '"/><c:scaling><c:orientation val="minMax"/>' + $script:XSCL + '</c:scaling><c:delete val="0"/><c:axPos val="b"/>'
    $plot += '<c:numFmt formatCode="' + (X $script:XFMT) + '" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>' + $AXLN + (Tx-Pr 800 '-2700000')
    $plot += '<c:crossAx val="' + $ax2 + '"/><c:crosses val="min"/><c:crossBetween val="midCat"/>' + $(if ($script:XMAJ -ne '') { '<c:majorUnit val="' + $script:XMAJ + '"/>' } else { '' }) + '</c:valAx>'
    # Y
    $ysc = '<c:min val="0"/>'; if ($ch.YMax -ne '') { $ysc += '<c:max val="' + $ch.YMax + '"/>' }
    $plot += '<c:valAx><c:axId val="' + $ax2 + '"/><c:scaling><c:orientation val="minMax"/>' + $ysc + '</c:scaling><c:delete val="0"/><c:axPos val="l"/>' + $GRID
    if ($ch.YTitle -ne '') { $plot += Ax-Title $ch.YTitle }
    $plot += '<c:numFmt formatCode="' + (X $ch.YFmt) + '" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' + $AXLN + (Tx-Pr 800 '')
    $plot += '<c:crossAx val="' + $ax1 + '"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><c:chart>' + (C-Title $ch.Title 1000) + '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' + $plot + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + (Tx-Pr 800 '') + '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900">' + $FONT + '</a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr></c:chartSpace>'
}
# 系列: 名前・シート・列 (X は日時列)。グループ (データソース別など) の行範囲を使う
function Sr([string]$name, $sh, [int]$ycol, $grp) {
    if ($null -eq $sh -or $sh.Rows.Count -eq 0) { return $null }
    $g = $grp
    if ($null -eq $g) { $r1 = $DATA_ROW; $r2 = $DATA_ROW + $sh.Rows.Count - 1 } else { $r1 = $g.R1; $r2 = $g.R2 }
    return [pscustomobject]@{ N = $name; S = $sh.Name; X = 1; Y = $ycol; R1 = $r1; R2 = $r2 }
}
$ChartSecs = New-Object System.Collections.Generic.List[object]
function Add-ChartSection([string]$title) { $ChartSecs.Add([pscustomobject]@{ Title = $title; Charts = (New-Object System.Collections.Generic.List[object]) }) }
function Add-Chart([string]$title, [string]$ytitle, [string]$yfmt, [object[]]$series, [string]$ymax) {
    $ss = @(); foreach ($s in $series) { if ($null -ne $s) { $ss += $s } }
    if ($ss.Count -eq 0) { return }
    if ($ChartSecs.Count -eq 0) { Add-ChartSection '' }
    $ChartSecs[$ChartSecs.Count - 1].Charts.Add([pscustomobject]@{ Title = $title; YTitle = $ytitle; YFmt = $yfmt; Series = $ss; YMax = $ymax })
}

function Write-ChartsSheet([System.IO.Compression.ZipArchive]$zip, [int]$sheetIdx, [string]$sheetName, [string]$title, [string]$note) {
    X-Scale
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((Ws-Begin 'FF262626' $true 0 0 $false 85))
    [void]$sb.Append('<cols><col min="1" max="18" width="9.5" customWidth="1"/></cols><sheetData>')
    [void]$sb.Append((Row 1 (StrC 'A1' $title $S_TITLE) 24))
    [void]$sb.Append((Row 2 (StrC 'A2' $note $S_NOTE) 0))
    $anch = New-Object System.Text.StringBuilder; $rels = New-Object System.Text.StringBuilder
    $n = 0; $r = 3; $brk = ''; $nbrk = 0
    foreach ($sec in $ChartSecs) {
        if ($sec.Charts.Count -eq 0) { continue }
        if ($r -gt 3) { $brk += '<brk id="' + ($r) + '" max="16383" man="1"/>'; $nbrk++ }
        $r++
        [void]$sb.Append((Row $r (StrC ('A' + $r) ('■ ' + $sec.Title) $S_HEAD) 21))
        $k = 0
        foreach ($ch in $sec.Charts) {
            $n++
            $col = $(if ($k % 2 -eq 0) { 0 } else { 9 })
            $top = $r + [int][Math]::Floor($k / 2) * 19
            # 印刷: 1 ページに 2 段 (4 グラフ) まで。グラフがページをまたがないよう改ページを入れる
            if ($k -gt 0 -and $k % 4 -eq 0) { $brk += '<brk id="' + $top + '" max="16383" man="1"/>'; $nbrk++ }
            $e = $zip.CreateEntry("xl/charts/chart$n.xml", [IO.Compression.CompressionLevel]::Optimal)
            $cw = New-Object IO.StreamWriter($e.Open(), $Utf8); $cw.Write((Chart-Xml $ch $n)); $cw.Dispose()
            [void]$anch.Append('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + $col + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $top + '</xdr:row><xdr:rowOff>60000</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + ($col + 9) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ($top + 18) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + ($n + 1) + '" name="Chart ' + $n + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + $n + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
            [void]$rels.Append('<Relationship Id="rId' + $n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + $n + '.xml"/>')
            $k++
        }
        $r += [int][Math]::Ceiling($sec.Charts.Count / 2.0) * 19 + 1
    }
    if ($n -eq 0) { [void]$sb.Append((Row 4 (StrC 'A4' '（グラフに使える時系列データがありません）' $S_NOTE) 0)) }
    [void]$sb.Append('</sheetData>')
    $extra = ''
    if ($nbrk -gt 0) { $extra = '<rowBreaks count="' + $nbrk + '" manualBreakCount="' + $nbrk + '">' + $brk + '</rowBreaks>' }
    # rowBreaks は pageSetup の後ろに置く必要がある
    $end = Ws-End 'landscape' $true ($n -gt 0) $title ''
    if ($extra -ne '') { $end = $end.Replace('</headerFooter>', '</headerFooter>' + $extra) }
    [void]$sb.Append($end)
    if ($n -gt 0) {
        $e = $zip.CreateEntry('xl/drawings/drawing1.xml', [IO.Compression.CompressionLevel]::Optimal)
        $dw = New-Object IO.StreamWriter($e.Open(), $Utf8)
        $dw.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + $anch.ToString() + '</xdr:wsDr>'); $dw.Dispose()
        $e = $zip.CreateEntry('xl/drawings/_rels/drawing1.xml.rels', [IO.Compression.CompressionLevel]::Optimal)
        $dw = New-Object IO.StreamWriter($e.Open(), $Utf8)
        $dw.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels.ToString() + '</Relationships>'); $dw.Dispose()
    }
    Add-DefName $sheetIdx '_xlnm.Print_Area' ("'" + $sheetName + "'!`$A`$1:`$R`$" + [Math]::Max($r, 4))
    $script:NCH = $n
    return $sb.ToString()
}

#------------------------------------------------------------------------------
# 項目名 (収集情報シート)
#------------------------------------------------------------------------------
$MetaLabel = @{
    'perf.version' = 'perf-metrics.sh の版'; 'perf.start_time' = '収集開始 (コンテナの時刻)'; 'perf.utc_offset' = 'コンテナの UTC オフセット'
    'perf.output_dir' = '出力先'; 'perf.work_dir' = '作業ディレクトリ'; 'perf.collectors' = '動かした収集'
    'perf.cli_interval' = '① jboss-cli 間隔 (秒)'; 'perf.jstat_interval' = '② jstat 間隔 (秒)'; 'perf.jcmd_interval' = '② jcmd 間隔 (秒)'
    'perf.top_interval' = '③ top 間隔 (秒)'; 'perf.ps_interval' = '③ ps / cgroup 間隔 (秒)'; 'perf.duration' = '収集時間 (秒。0 = 停止まで)'
    'perf.jstat_options' = 'jstat のオプション'; 'perf.jcmd_commands' = 'jcmd のコマンド'; 'perf.statistics_opts' = '統計の有効化 (起動引数)'
    'container.hostname' = 'ホスト名'; 'container.service_name' = 'サービス名'; 'container.component_name' = 'コンポーネント名'
    'container.component_role' = '役割'; 'container.log_id' = 'LOG_ID (タスク専用ディレクトリ)'; 'container.jboss_home' = 'JBOSS_HOME'
    'container.cpus_online' = 'CPU 数 (OS から見える数)'; 'container.mem_total_kb' = 'メモリ (OS から見える量, KB)'; 'container.os' = 'OS'
    'container.command' = '起動コマンド'; 'ecs.TaskARN' = 'ECS タスク ARN'; 'ecs.Cluster' = 'ECS クラスター'; 'ecs.Family' = 'タスク定義'
    'ecs.Revision' = 'リビジョン'; 'ecs.AvailabilityZone' = 'AZ'; 'ecs.LaunchType' = '起動タイプ'; 'ecs.TaskLimits' = 'タスクの上限 (CPU / Memory)'
    'jboss.product-name' = '製品'; 'jboss.product-version' = '製品の版'; 'jboss.release-version' = 'コアの版'; 'jboss.server-name' = 'サーバー名'
    'jboss.launch-type' = '起動形態'; 'jboss.running-mode' = '実行モード'; 'jboss.config-file' = '設定ファイル'; 'jboss.node-name' = 'ノード名'
    'jboss.vm-name' = 'JVM'; 'jboss.vm-vendor' = 'JVM ベンダー'; 'jboss.vm-version' = 'JVM の版'; 'jboss.jvm-start-time' = 'JVM 起動時刻 (epoch ms)'
    'jboss.available-processors' = 'JVM が使える CPU 数'; 'jvm.version' = 'JVM (jcmd VM.version)'; 'jvm.jstat_command' = 'jstat の起動方法'
    'jvm.jcmd_command' = 'jcmd の起動方法'; 'jvm_flag.MaxHeapSize' = '最大ヒープ (バイト)'; 'jvm_flag.InitialHeapSize' = '初期ヒープ (バイト)'
    'jvm_flag.UseG1GC' = 'G1 GC'; 'jvm_flag.UseParallelGC' = 'Parallel GC'; 'jvm_flag.UseSerialGC' = 'Serial GC'; 'jvm_flag.UseZGC' = 'Z GC'
    'jvm_flag.MaxGCPauseMillis' = 'MaxGCPauseMillis'; 'jvm_flag.ActiveProcessorCount' = 'ActiveProcessorCount'; 'jvm_flag.MaxRAMPercentage' = 'MaxRAMPercentage'
    'cli.controller' = 'jboss-cli の接続先'; 'cli.mode' = 'jboss-cli の方式'; 'cli.ds_stats_enabled_at_runtime' = 'プール統計を実行時に有効化した時刻'
    'os.top_version' = 'top の版'
}
$MetaCat = @{ perf = '収集条件'; container = 'コンテナ'; ecs = 'ECS'; jboss = 'JBoss EAP'; jvm = 'JVM'; jvm_flag = 'JVM フラグ'; cli = 'jboss-cli'; os = 'OS' }

#------------------------------------------------------------------------------
# 1 つの perf/ ディレクトリからブックを作る
#------------------------------------------------------------------------------
function Build-Report([string]$dir, [string]$outPath) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Write-Info "入力: $dir"
    $DefNames.Clear(); $ChartSecs.Clear()

    # ---- 収集情報 ----
    $meta = [ordered]@{}
    $mp = Join-Path $dir 'perf_meta.csv'
    if (Test-Path -LiteralPath $mp) {
        foreach ($l in [IO.File]::ReadAllLines($mp, $Utf8)) {
            if ($l.Length -eq 0) { continue }
            $f = Split-Csv $l
            if ($f.Length -lt 3) { continue }
            $meta[$f[0] + '.' + $f[1]] = $f[2]
        }
    }
    function M([string]$k) { if ($meta.Contains($k)) { return [string]$meta[$k] }; return '' }
    $script:RecOffset = 0
    $uo = M 'perf.utc_offset'
    if ($uo -match '^([+-])([0-9]{2})([0-9]{2})$') { $sg = $(if ($Matches[1] -eq '-') { -1 } else { 1 }); $script:RecOffset = $sg * ([int]$Matches[2] * 3600 + [int]$Matches[3] * 60) }
    $PerfRpt::UseRecorded = [bool]$UseRecordedTime; $PerfRpt::RecOffset = [double]$script:RecOffset; $PerfRpt::ResetTz()
    $script:FromEpoch = [double]::MinValue; $script:ToEpoch = [double]::MaxValue
    foreach ($pair in @(@($From, 'FromEpoch'), @($To, 'ToEpoch'))) {
        if ($pair[0] -eq '') { continue }
        $dt = [DateTime]::Parse($pair[0], $Inv)
        if ($UseRecordedTime) { $ep = ($dt - [DateTime]'1970-01-01').TotalSeconds - $script:RecOffset }
        else { $ep = ([DateTimeOffset]::new($dt, [TimeZoneInfo]::Local.GetUtcOffset($dt))).ToUnixTimeSeconds() }
        Set-Variable -Scope Script -Name $pair[1] -Value ([double]$ep)
    }

    # ---- CSV ----
    $TB = @{}
    foreach ($n in @('cli_datasource', 'cli_undertow', 'cli_session', 'cli_io_worker', 'cli_threadpool', 'cli_transactions', 'cli_jvm', 'cli_gc', 'cli_mempool',
                     'jvm_jstat_gc', 'jvm_jcmd_heap', 'os_top_summary', 'os_top_process', 'os_top_threads', 'os_ps', 'os_cgroup')) {
        $t = Read-PerfCsv $dir ($n + '.csv')
        if ($null -ne $t -and $t.Rows.Count -gt 0) { $TB[$n] = $t; Write-Info ("読み込み: {0}.csv ({1:N0} 行)" -f $n, $t.Rows.Count) }
    }
    $extraJstat = @()
    foreach ($f in (Get-ChildItem -LiteralPath $dir -Filter 'jvm_jstat_*.csv' -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -eq 'jvm_jstat_gc.csv') { continue }
        $t = Read-PerfCsv $dir $f.Name; if ($null -ne $t -and $t.Rows.Count -gt 0) { $extraJstat += $t }
    }
    if ($TB.Count -eq 0) { Write-Warning "この期間のデータがありません: $dir"; return }

    # 全体の時間範囲 (グラフの時間軸)
    $script:TMin = $null; $script:TMax = $null
    foreach ($t in $TB.Values) {
        $ei = $t.I['epoch']; if ($null -eq $ei) { continue }
        $a = $PerfRpt::D($t.Rows[0][$ei]); $b = $PerfRpt::D($t.Rows[$t.Rows.Count - 1][$ei])
        if ($null -ne $a -and ($null -eq $script:TMin -or ($PerfRpt::XT($a)) -lt $script:TMin)) { $script:TMin = $PerfRpt::XT($a)}
        if ($null -ne $b -and ($null -eq $script:TMax -or ($PerfRpt::XT($b)) -gt $script:TMax)) { $script:TMax = $PerfRpt::XT($b)}
    }

    $KPI = @{}          # 主要指標用の値の列
    $Find = New-Object System.Collections.Generic.List[object]   # 所見 (重さ, 分類, 内容)
    $Sheets = New-Object System.Collections.Generic.List[object]

    # ==== DB 接続プール ====
    Write-Info '集計: DB 接続プール'
    $dsSum = New-Object System.Collections.Generic.List[object]
    $shDS = $null
    if ($TB.ContainsKey('cli_datasource')) {
        $t = $TB['cli_datasource']
        $sh = New-Sheet 'DB接続プール' 'DB 接続プール（データソース別の時系列）' ('① jboss-cli の /subsystem=datasources/data-source=*/statistics=pool・jdbc。「区間」は前回の取得からの増分です。最大プールが未設定のデータソースは IronJacamar の既定値 20 で使用率を計算しています。使用率 80% 以上・待ち・取得失敗は濃淡で表示します。') 2
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cN = Add-Col $sh '基本' 'データソース' 'ds_name' 'text' 18
        $cIn = Add-Col $sh '接続数' '使用中' 'InUseCount' 'int' 8
        $cAc = Add-Col $sh '接続数' 'アクティブ' 'ActiveCount' 'int' 9
        $cId = Add-Col $sh '接続数' 'アイドル' 'IdleCount' 'int' 8
        $cAv = Add-Col $sh '接続数' '取得可能' 'AvailableCount' 'int' 8
        $cMx = Add-Col $sh '接続数' '最大プール' 'max_pool_size' 'int' 9
        $cUs = Add-Col $sh '接続数' '使用率' '(InUse÷最大)' 'pct' 9
        $cWd = Add-Col $sh '待ち・失敗' '待ち（区間）' 'Δ WaitCount' 'int' 9
        $cFd = Add-Col $sh '待ち・失敗' '取得失敗（区間）' 'Δ BlockingFailureCount' 'int' 10
        $cWc = Add-Col $sh '待ち・失敗' '待ち（累計）' 'WaitCount' 'int' 10
        $cMw = Add-Col $sh '待ち・失敗' '最大同時待ち' 'MaxWaitCount' 'int' 9
        $cMu = Add-Col $sh '待ち・失敗' '最大使用数' 'MaxUsedCount' 'int' 9
        $cAg = Add-Col $sh '時間（ミリ秒）' '平均取得' 'AverageGetTime' 'int' 9
        $cMg = Add-Col $sh '時間（ミリ秒）' '最大取得' 'MaxGetTime' 'int' 9
        $cAb = Add-Col $sh '時間（ミリ秒）' '平均ブロック' 'AverageBlockingTime' 'int' 9
        $cMwt = Add-Col $sh '時間（ミリ秒）' '最大待ち' 'MaxWaitTime' 'int' 9
        $cAu = Add-Col $sh '時間（ミリ秒）' '平均使用' 'AverageUsageTime' 'int' 9
        $cMuT = Add-Col $sh '時間（ミリ秒）' '最大使用' 'MaxUsageTime' 'int' 9
        $cMc = Add-Col $sh '時間（ミリ秒）' '最大作成' 'MaxCreationTime' 'int' 9
        $cCr = Add-Col $sh '物理接続' '作成（区間）' 'Δ CreatedCount' 'int' 9
        $cDe = Add-Col $sh '物理接続' '破棄（区間）' 'Δ DestroyedCount' 'int' 9
        $cTo = Add-Col $sh '物理接続' 'アイドル破棄（区間）' 'Δ TimedOut' 'int' 10
        $cPh = Add-Col $sh '文キャッシュ' 'ヒット率（区間）' 'Δ PSCacheHit÷Access' 'pct' 10
        $cPs = Add-Col $sh '文キャッシュ' 'サイズ' 'PreparedStatementCacheCurrentSize' 'int' 8
        $cMn = Add-Col $sh '設定' '最小プール' 'min_pool_size' 'int' 8
        $cBt = Add-Col $sh '設定' '取得タイムアウト(ms)' 'blocking_timeout_ms' 'int' 11
        $cJn = Add-Col $sh '設定' 'JNDI 名' 'jndi_name' 'text' 30
        $cSe = Add-Col $sh '設定' '統計' 'statistics_enabled' 'textc' 7
        Add-CF $sh $cUs 'bar' '1' 0; Add-CF $sh $cUs 'ge' '0.8' 1; Add-CF $sh $cWd 'gt' '0' 1; Add-CF $sh $cFd 'gt' '0' 0
        $grp = Group-By $t @('ds_type', 'ds_name')
        foreach ($key in $grp.Keys) {
            $rows = $grp[$key]
            $name = $PerfRpt::G($t, $rows[0], 'ds_name')
            $p = $null
            $lIn = NewList; $lUs = NewList; $wsum = 0.0; $fsum = 0.0; $maxWaitT = $null; $atMax = 0; $mxp = $null; $lastAg = $null; $maxGet = $null
            foreach ($r in $rows) {
                $mx = $PerfRpt::GD($t, $r, 'max_pool_size'); if ($null -eq $mx) { $mx = 20.0 }
                $in = $PerfRpt::GD($t, $r, 'InUseCount')
                $wd = $null; $fd = $null; $cr = $null; $de = $null; $tmo = $null; $ph = $null
                if ($null -ne $p) {
                    $wd = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'WaitCount')), ($PerfRpt::GD($t, $p, 'WaitCount')))
                    $fd = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'BlockingFailureCount')), ($PerfRpt::GD($t, $p, 'BlockingFailureCount')))
                    $cr = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'CreatedCount')), ($PerfRpt::GD($t, $p, 'CreatedCount')))
                    $de = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'DestroyedCount')), ($PerfRpt::GD($t, $p, 'DestroyedCount')))
                    $tmo = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'TimedOut')), ($PerfRpt::GD($t, $p, 'TimedOut')))
                    $ph = $PerfRpt::Div(($PerfRpt::Dl(($PerfRpt::GD($t, $r, 'PreparedStatementCacheHitCount')), ($PerfRpt::GD($t, $p, 'PreparedStatementCacheHitCount')))), ($PerfRpt::Dl(($PerfRpt::GD($t, $r, 'PreparedStatementCacheAccessCount')), ($PerfRpt::GD($t, $p, 'PreparedStatementCacheAccessCount')))))
                }
                $us = $PerfRpt::Div($in, $mx)
                $PerfRpt::AddV($lIn, $in); $PerfRpt::AddV($lUs, $us)
                if ($null -ne $wd) { $wsum += $wd }; if ($null -ne $fd) { $fsum += $fd }
                if ($null -ne $in -and $in -ge $mx) { $atMax++ }
                $v = $PerfRpt::GD($t, $r, 'MaxWaitTime'); if ($null -ne $v -and ($null -eq $maxWaitT -or $v -gt $maxWaitT)) { $maxWaitT = $v }
                $v = $PerfRpt::GD($t, $r, 'MaxGetTime'); if ($null -ne $v -and ($null -eq $maxGet -or $v -gt $maxGet)) { $maxGet = $v }
                $lastAg = $PerfRpt::GD($t, $r, 'AverageGetTime'); $mxp = $mx
                $vals = New-Object 'object[]' $sh.Cols.Count
                $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::GD($t, $r, 'epoch'))); $vals[$cN - 1] = $name
                $vals[$cIn - 1] = $in; $vals[$cAc - 1] = $PerfRpt::GD($t, $r, 'ActiveCount'); $vals[$cId - 1] = $PerfRpt::GD($t, $r, 'IdleCount'); $vals[$cAv - 1] = $PerfRpt::GD($t, $r, 'AvailableCount')
                $vals[$cMx - 1] = $mx; $vals[$cUs - 1] = $PerfRpt::Rnd($us, 4); $vals[$cWd - 1] = $wd; $vals[$cFd - 1] = $fd
                $vals[$cWc - 1] = $PerfRpt::GD($t, $r, 'WaitCount'); $vals[$cMw - 1] = $PerfRpt::GD($t, $r, 'MaxWaitCount'); $vals[$cMu - 1] = $PerfRpt::GD($t, $r, 'MaxUsedCount')
                $vals[$cAg - 1] = $PerfRpt::GD($t, $r, 'AverageGetTime'); $vals[$cMg - 1] = $PerfRpt::GD($t, $r, 'MaxGetTime'); $vals[$cAb - 1] = $PerfRpt::GD($t, $r, 'AverageBlockingTime')
                $vals[$cMwt - 1] = $PerfRpt::GD($t, $r, 'MaxWaitTime'); $vals[$cAu - 1] = $PerfRpt::GD($t, $r, 'AverageUsageTime'); $vals[$cMuT - 1] = $PerfRpt::GD($t, $r, 'MaxUsageTime')
                $vals[$cMc - 1] = $PerfRpt::GD($t, $r, 'MaxCreationTime'); $vals[$cCr - 1] = $cr; $vals[$cDe - 1] = $de; $vals[$cTo - 1] = $tmo
                $vals[$cPh - 1] = $PerfRpt::Rnd($ph, 4); $vals[$cPs - 1] = $PerfRpt::GD($t, $r, 'PreparedStatementCacheCurrentSize')
                $vals[$cMn - 1] = $PerfRpt::GD($t, $r, 'min_pool_size'); $vals[$cBt - 1] = $PerfRpt::GD($t, $r, 'blocking_timeout_ms'); $vals[$cJn - 1] = $PerfRpt::G($t, $r, 'jndi_name'); $vals[$cSe - 1] = $PerfRpt::G($t, $r, 'statistics_enabled')
                Add-Row $sh $vals $key
                $p = $r
            }
            $dsSum.Add([pscustomobject]@{ Name = $name; Max = $mxp; In = (Stat $lIn); Us = (Stat $lUs); Wait = $wsum; Fail = $fsum; MaxWait = $maxWaitT; AtMax = $atMax; N = $rows.Count; AvgGet = $lastAg; MaxGet = $maxGet; StatsOff = (($PerfRpt::G($t, $rows[$rows.Count - 1], 'statistics_enabled')) -eq 'false') })
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shDS = $sh
        $script:DSC = @{ In = $cIn; Ac = $cAc; Mx = $cMx; Wd = $cWd; Fd = $cFd; Mwt = $cMwt; Ag = $cAg }
    }

    # ==== HTTP・スレッド (jboss-cli の取得 1 回 = 1 行) ====
    Write-Info '集計: HTTP・スレッド'
    $shHT = $null
    if ($TB.ContainsKey('cli_jvm')) {
        $tj = $TB['cli_jvm']
        $byEp = @{}
        foreach ($nm in @('cli_undertow', 'cli_io_worker', 'cli_threadpool', 'cli_session', 'cli_transactions')) {
            $byEp[$nm] = @{}
            if (-not $TB.ContainsKey($nm)) { continue }
            $t = $TB[$nm]; $ei = $t.I['epoch']
            foreach ($r in $t.Rows) { $k = $r[$ei]; if (-not $byEp[$nm].ContainsKey($k)) { $byEp[$nm][$k] = New-Object System.Collections.Generic.List[object] }; $byEp[$nm][$k].Add($r) }
        }
        $sh = New-Sheet 'HTTP・スレッド' 'HTTP・リクエスト処理スレッド・セッション・トランザクション' ('① jboss-cli の取得 1 回を 1 行にまとめています。HTTP は全リスナーの合計、「/秒」「区間」は前回の取得からの増分です。平均処理時間は Undertow の processing-time から計算するため、リスナーの record-request-start-time=true のときだけ値が入ります。IO ワーカーが最大に達した行・キュー待ちは濃淡で表示します。') 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cRps = Add-Col $sh 'HTTP（全リスナー）' 'リクエスト/秒' 'Δ request-count÷秒' 'dec1' 10
        $cEps = Add-Col $sh 'HTTP（全リスナー）' 'エラー/秒' 'Δ error-count÷秒' 'dec2' 9
        $cAms = Add-Col $sh 'HTTP（全リスナー）' '平均処理時間(ms)' 'Δ processing-time÷Δ request' 'dec1' 11
        $cMms = Add-Col $sh 'HTTP（全リスナー）' '最大処理時間(ms)' 'max-processing-time' 'dec1' 11
        $cRx = Add-Col $sh 'HTTP（全リスナー）' '受信 KB/秒' 'Δ bytes-received' 'dec1' 9
        $cTx = Add-Col $sh 'HTTP（全リスナー）' '送信 KB/秒' 'Δ bytes-sent' 'dec1' 9
        $cRq = Add-Col $sh 'HTTP（全リスナー）' 'リクエスト（区間）' 'Δ request-count' 'int' 10
        $cEr = Add-Col $sh 'HTTP（全リスナー）' 'エラー（区間）' 'Δ error-count' 'int' 9
        $cIb = Add-Col $sh 'リクエスト処理スレッド（IO ワーカー）' '処理中' 'busy-task-thread-count' 'int' 8
        $cIm = Add-Col $sh 'リクエスト処理スレッド（IO ワーカー）' '最大' 'max-pool-size' 'int' 7
        $cIu = Add-Col $sh 'リクエスト処理スレッド（IO ワーカー）' '使用率' '(処理中÷最大)' 'pct' 8
        $cIq = Add-Col $sh 'リクエスト処理スレッド（IO ワーカー）' 'キュー' 'queue-size' 'int' 7
        $cIo = Add-Col $sh 'リクエスト処理スレッド（IO ワーカー）' 'IO スレッド' 'io-thread-count' 'int' 8
        $cEa = Add-Col $sh 'アプリのスレッドプール' 'EJB 使用中' 'ejb3 active' 'int' 8
        $cEq = Add-Col $sh 'アプリのスレッドプール' 'EJB キュー' 'ejb3 queue' 'int' 8
        $cXa = Add-Col $sh 'アプリのスレッドプール' 'EE 使用中' 'ee-executor active' 'int' 8
        $cXq = Add-Col $sh 'アプリのスレッドプール' 'EE キュー' 'ee-executor queue' 'int' 8
        $cXh = Add-Col $sh 'アプリのスレッドプール' 'EE ハング' 'hung-thread-count' 'int' 8
        $cSa = Add-Col $sh 'HTTP セッション' 'アクティブ' 'active-sessions' 'int' 9
        $cSc = Add-Col $sh 'HTTP セッション' '作成（区間）' 'Δ sessions-created' 'int' 9
        $cSr = Add-Col $sh 'HTTP セッション' '拒否（区間）' 'Δ rejected-sessions' 'int' 9
        $cTc = Add-Col $sh 'トランザクション' 'コミット（区間）' 'Δ committed' 'int' 9
        $cTa = Add-Col $sh 'トランザクション' 'ロールバック（区間）' 'Δ aborted' 'int' 10
        $cTt = Add-Col $sh 'トランザクション' 'タイムアウト（区間）' 'Δ timed-out' 'int' 10
        $cTi = Add-Col $sh 'トランザクション' '実行中' 'inflight' 'int' 7
        $cAr = Add-Col $sh 'サーバー' 'アクティブ要求' 'active_requests' 'int' 9
        $cJt = Add-Col $sh 'サーバー' 'JVM スレッド数' 'thread_count' 'int' 9
        $cSs = Add-Col $sh 'サーバー' '状態' 'server_state' 'textc' 10
        $cCm = Add-Col $sh 'サーバー' 'CLI 応答(ms)' 'cli_ms' 'int' 9
        Add-CF $sh $cIu 'bar' '1' 0; Add-CF $sh $cIu 'ge' '1' 1; Add-CF $sh $cIq 'gt' '0' 2; Add-CF $sh $cEr 'gt' '0' 1; Add-CF $sh $cTt 'gt' '0' 1
        $prevL = @{}; $prevS = @{}; $prevTx = $null; $pe = $null
        $lRps = NewList; $lAms = NewList; $lIb = NewList; $lIq = NewList; $lCm = NewList; $lJt = NewList; $lSa = NewList; $reqTot = 0.0; $errTot = 0.0
        $ioMax = $null; $ioFullN = 0; $lsn = @{}
        foreach ($r in $tj.Rows) {
            $epS = $PerfRpt::G($tj, $r, 'epoch'); $ep = $PerfRpt::D($epS)
            $dt = $null; if ($null -ne $pe) { $dt = $ep - $pe }
            # HTTP: リスナーごとの増分を合計
            $dq = $null; $de = $null; $dp = $null; $drx = $null; $dtx = $null; $mms = $null
            $tu = $TB['cli_undertow']
            if ($byEp['cli_undertow'].ContainsKey($epS)) {
                $dq = 0.0; $de = 0.0; $dp = 0.0; $drx = 0.0; $dtx = 0.0; $any = $false
                foreach ($u in $byEp['cli_undertow'][$epS]) {
                    $lk = ($PerfRpt::G($tu, $u, 'server')) + '/' + ($PerfRpt::G($tu, $u, 'listener_type')) + '/' + ($PerfRpt::G($tu, $u, 'listener'))
                    $m = $PerfRpt::GD($tu, $u, 'max-processing-time'); if ($null -ne $m -and ($null -eq $mms -or $m -gt $mms)) { $mms = $m }
                    if ($prevL.ContainsKey($lk)) {
                        $q0 = $prevL[$lk]; $any = $true
                        $x = $PerfRpt::Dl(($PerfRpt::GD($tu, $u, 'request-count')), ($PerfRpt::GD($tu, $q0, 'request-count'))); if ($null -ne $x) { $dq += $x; if (-not $lsn.ContainsKey($lk)) { $lsn[$lk] = 0.0 }; $lsn[$lk] += $x }
                        $x = $PerfRpt::Dl(($PerfRpt::GD($tu, $u, 'error-count')), ($PerfRpt::GD($tu, $q0, 'error-count'))); if ($null -ne $x) { $de += $x }
                        $x = $PerfRpt::Dl(($PerfRpt::GD($tu, $u, 'processing-time')), ($PerfRpt::GD($tu, $q0, 'processing-time'))); if ($null -ne $x) { $dp += $x }
                        $x = $PerfRpt::Dl(($PerfRpt::GD($tu, $u, 'bytes-received')), ($PerfRpt::GD($tu, $q0, 'bytes-received'))); if ($null -ne $x) { $drx += $x }
                        $x = $PerfRpt::Dl(($PerfRpt::GD($tu, $u, 'bytes-sent')), ($PerfRpt::GD($tu, $q0, 'bytes-sent'))); if ($null -ne $x) { $dtx += $x }
                    }
                    $prevL[$lk] = $u
                }
                if (-not $any) { $dq = $null; $de = $null; $dp = $null; $drx = $null; $dtx = $null }
            }
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cT - 1] = $PerfRpt::XT($ep)
            if ($null -ne $dq -and $null -ne $dt -and $dt -gt 0) {
                $vals[$cRps - 1] = $PerfRpt::Rnd(($dq / $dt), 2); $vals[$cEps - 1] = $PerfRpt::Rnd(($de / $dt), 3)
                $vals[$cRx - 1] = $PerfRpt::Rnd(($drx / $dt / 1024), 1); $vals[$cTx - 1] = $PerfRpt::Rnd(($dtx / $dt / 1024), 1)
                $vals[$cRq - 1] = $dq; $vals[$cEr - 1] = $de
                $PerfRpt::AddV($lRps, ($dq / $dt)); $reqTot += $dq; $errTot += $de
                if ($dq -gt 0 -and $dp -gt 0) { $a = $dp / $dq / 1000000.0; $vals[$cAms - 1] = $PerfRpt::Rnd($a, 2); $PerfRpt::AddV($lAms, $a)}
            }
            if ($null -ne $mms -and $mms -gt 0) { $vals[$cMms - 1] = $PerfRpt::Rnd(($mms / 1000000.0), 1)}
            # IO ワーカー (全ワーカーの合計)
            if ($byEp['cli_io_worker'].ContainsKey($epS)) {
                $ti = $TB['cli_io_worker']; $b = 0.0; $mx = 0.0; $q = 0.0; $io = 0.0
                foreach ($u in $byEp['cli_io_worker'][$epS]) {
                    $v = $PerfRpt::GD($ti, $u, 'busy-task-thread-count'); if ($null -ne $v) { $b += $v }
                    $v = $PerfRpt::GD($ti, $u, 'max-pool-size'); if ($null -eq $v) { $v = $PerfRpt::GD($ti, $u, 'task-max-threads')}; if ($null -ne $v) { $mx += $v }
                    $v = $PerfRpt::GD($ti, $u, 'queue-size'); if ($null -ne $v) { $q += $v }
                    $v = $PerfRpt::GD($ti, $u, 'io-thread-count'); if ($null -ne $v) { $io += $v }
                }
                $vals[$cIb - 1] = $b; $vals[$cIm - 1] = $mx; $vals[$cIu - 1] = $PerfRpt::Rnd(($PerfRpt::Div($b, $mx)), 4); $vals[$cIq - 1] = $q; $vals[$cIo - 1] = $io
                $PerfRpt::AddV($lIb, $b); $PerfRpt::AddV($lIq, $q); $ioMax = $mx; if ($mx -gt 0 -and $b -ge $mx) { $ioFullN++ }
            }
            if ($byEp['cli_threadpool'].ContainsKey($epS)) {
                $tp = $TB['cli_threadpool']; $ea = $null; $eq = $null; $xa = $null; $xq = $null; $xh = $null
                foreach ($u in $byEp['cli_threadpool'][$epS]) {
                    $ty = $PerfRpt::G($tp, $u, 'pool_type')
                    if ($ty -eq 'ejb3') { $ea = [double]$ea + [double]($PerfRpt::GD($tp, $u, 'active')); $eq = [double]$eq + [double]($PerfRpt::GD($tp, $u, 'queue')) }
                    elseif ($ty -eq 'ee-executor') { $xa = [double]$xa + [double]($PerfRpt::GD($tp, $u, 'active')); $xq = [double]$xq + [double]($PerfRpt::GD($tp, $u, 'queue')); $xh = [double]$xh + [double]($PerfRpt::GD($tp, $u, 'hung')) }
                }
                $vals[$cEa - 1] = $ea; $vals[$cEq - 1] = $eq; $vals[$cXa - 1] = $xa; $vals[$cXq - 1] = $xq; $vals[$cXh - 1] = $xh
            }
            if ($byEp['cli_session'].ContainsKey($epS)) {
                $ts = $TB['cli_session']; $sa = 0.0; $sc = $null; $sr = $null
                foreach ($u in $byEp['cli_session'][$epS]) {
                    $dk = $PerfRpt::G($ts, $u, 'deployment')
                    $v = $PerfRpt::GD($ts, $u, 'active-sessions'); if ($null -ne $v) { $sa += $v }
                    if ($prevS.ContainsKey($dk)) {
                        $x = $PerfRpt::Dl(($PerfRpt::GD($ts, $u, 'sessions-created')), ($PerfRpt::GD($ts, $prevS[$dk], 'sessions-created'))); if ($null -ne $x) { $sc = [double]$sc + $x }
                        $x = $PerfRpt::Dl(($PerfRpt::GD($ts, $u, 'rejected-sessions')), ($PerfRpt::GD($ts, $prevS[$dk], 'rejected-sessions'))); if ($null -ne $x) { $sr = [double]$sr + $x }
                    }
                    $prevS[$dk] = $u
                }
                $vals[$cSa - 1] = $sa; $vals[$cSc - 1] = $sc; $vals[$cSr - 1] = $sr; $PerfRpt::AddV($lSa, $sa)
            }
            if ($byEp['cli_transactions'].ContainsKey($epS)) {
                $tt = $TB['cli_transactions']; $u = $byEp['cli_transactions'][$epS][0]
                if ($null -ne $prevTx) {
                    $vals[$cTc - 1] = $PerfRpt::Dl(($PerfRpt::GD($tt, $u, 'number-of-committed-transactions')), ($PerfRpt::GD($tt, $prevTx, 'number-of-committed-transactions')))
                    $vals[$cTa - 1] = $PerfRpt::Dl(($PerfRpt::GD($tt, $u, 'number-of-aborted-transactions')), ($PerfRpt::GD($tt, $prevTx, 'number-of-aborted-transactions')))
                    $vals[$cTt - 1] = $PerfRpt::Dl(($PerfRpt::GD($tt, $u, 'number-of-timed-out-transactions')), ($PerfRpt::GD($tt, $prevTx, 'number-of-timed-out-transactions')))
                }
                $vals[$cTi - 1] = $PerfRpt::GD($tt, $u, 'number-of-inflight-transactions')
                $prevTx = $u
            }
            $vals[$cAr - 1] = $PerfRpt::GD($tj, $r, 'active_requests'); $vals[$cJt - 1] = $PerfRpt::GD($tj, $r, 'thread_count')
            $vals[$cSs - 1] = $PerfRpt::G($tj, $r, 'server_state'); $vals[$cCm - 1] = $PerfRpt::GD($tj, $r, 'cli_ms')
            $PerfRpt::AddV($lCm, ($PerfRpt::GD($tj, $r, 'cli_ms'))); $PerfRpt::AddV($lJt, ($PerfRpt::GD($tj, $r, 'thread_count')))
            Add-Row $sh $vals ''
            $pe = $ep
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shHT = $sh
        $KPI['rps'] = Stat $lRps; $KPI['ams'] = Stat $lAms; $KPI['iob'] = Stat $lIb; $KPI['ioq'] = Stat $lIq; $KPI['clims'] = Stat $lCm; $KPI['jthr'] = Stat $lJt; $KPI['ses'] = Stat $lSa
        $KPI['reqTot'] = $reqTot; $KPI['errTot'] = $errTot; $KPI['ioMax'] = $ioMax; $KPI['ioFullN'] = $ioFullN; $KPI['lsn'] = $lsn
        $script:HTC = @{ Rps = $cRps; Eps = $cEps; Ams = $cAms; Ib = $cIb; Im = $cIm; Iq = $cIq; Sa = $cSa; Ar = $cAr; Jt = $cJt }
    }

    # ==== JVM・GC (CLI / platform-mbean) ====
    Write-Info '集計: JVM・GC（CLI）'
    $shJV = $null
    if ($TB.ContainsKey('cli_jvm')) {
        $tj = $TB['cli_jvm']
        $gcNames = New-Object System.Collections.Generic.List[string]; $gcBy = @{}
        if ($TB.ContainsKey('cli_gc')) { $tg = $TB['cli_gc']; foreach ($r in $tg.Rows) { $nm = $PerfRpt::G($tg, $r, 'collector'); if (-not $gcNames.Contains($nm)) { $gcNames.Add($nm) }; $gcBy[($PerfRpt::G($tg, $r, 'epoch')) + "`t" + $nm] = $r } }
        $mpBy = @{}
        if ($TB.ContainsKey('cli_mempool')) {
            $tm = $TB['cli_mempool']
            foreach ($r in $tm.Rows) {
                $nm = $PerfRpt::G($tm, $r, 'pool'); $g2 = 'Other'
                if ($nm -match 'Eden') { $g2 = 'Eden' } elseif ($nm -match 'Survivor') { $g2 = 'Survivor' } elseif ($nm -match 'Old|Tenured') { $g2 = 'Old' }
                elseif ($nm -match '^Metaspace') { $g2 = 'Metaspace' } elseif ($nm -match 'Compressed') { $g2 = 'CCS' } elseif ($nm -match 'CodeHeap|Code_Cache|Code Cache') { $g2 = 'Code' }
                $k2 = ($PerfRpt::G($tm, $r, 'epoch')) + "`t" + $g2
                $v = $PerfRpt::GD($tm, $r, 'used'); if ($null -eq $v) { continue }
                if ($mpBy.ContainsKey($k2)) { $mpBy[$k2] += $v } else { $mpBy[$k2] = $v }
            }
        }
        $sh = New-Sheet 'JVM・GC（CLI）' 'JVM（platform-mbean）：ヒープ・メモリプール・スレッド・クラス・GC' '① jboss-cli の /core-service=platform-mbean。jstat / jcmd が使えない環境 (/tmp が書けない等) でも取得できます。GC の「区間」は前回の取得からの増分です。' 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cHu = Add-Col $sh 'ヒープ（MB）' '使用' 'heap_used' 'dec1' 9
        $cHc = Add-Col $sh 'ヒープ（MB）' 'コミット' 'heap_committed' 'dec1' 9
        $cHm = Add-Col $sh 'ヒープ（MB）' '最大' 'heap_max' 'dec1' 9
        $cHp = Add-Col $sh 'ヒープ（MB）' '使用率' '(使用÷最大)' 'pct' 8
        $cNu = Add-Col $sh '非ヒープ（MB）' '使用' 'nonheap_used' 'dec1' 9
        $cNc = Add-Col $sh '非ヒープ（MB）' 'コミット' 'nonheap_committed' 'dec1' 9
        $mpCols = @{}
        foreach ($g2 in @('Eden', 'Survivor', 'Old', 'Metaspace', 'CCS', 'Code')) { $mpCols[$g2] = Add-Col $sh 'メモリプール使用量（MB）' $(switch ($g2) { 'CCS' { 'Compressed Class' } 'Code' { 'CodeHeap' } default { $g2 } }) $g2 'dec1' 9 }
        $cTh = Add-Col $sh 'スレッド' 'スレッド数' 'thread_count' 'int' 8
        $cTp = Add-Col $sh 'スレッド' 'ピーク' 'peak_thread_count' 'int' 7
        $cTd = Add-Col $sh 'スレッド' 'デーモン' 'daemon_thread_count' 'int' 7
        $cCl = Add-Col $sh 'クラス' 'ロード中' 'loaded_class_count' 'int' 9
        $cCu = Add-Col $sh 'クラス' 'アンロード累計' 'unloaded_class_count' 'int' 9
        $gcCols = @{}
        foreach ($nm in $gcNames) {
            $lab = $nm.Replace('_', ' ')
            $gcCols[$nm] = @((Add-Col $sh ('GC: ' + $lab) '回数（区間）' 'Δ collection-count' 'int' 9), (Add-Col $sh ('GC: ' + $lab) '時間 ms（区間）' 'Δ collection-time' 'int' 10))
        }
        $cLa = Add-Col $sh 'OS' 'ロードアベレージ' 'system_load_average' 'dec2' 9
        Add-CF $sh $cHp 'bar' '1' 0; Add-CF $sh $cHp 'ge' '0.85' 1
        $prevGc = @{}
        $lHu = NewList; $lNu = NewList; $hmx = $null; $gcDelta = @{}
        foreach ($r in $tj.Rows) {
            $epS = $PerfRpt::G($tj, $r, 'epoch')
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::D($epS)))
            $hu = $PerfRpt::MB(($PerfRpt::GD($tj, $r, 'heap_used'))); $hm = $PerfRpt::MB(($PerfRpt::GD($tj, $r, 'heap_max')))
            $vals[$cHu - 1] = $hu; $vals[$cHc - 1] = $PerfRpt::MB(($PerfRpt::GD($tj, $r, 'heap_committed'))); $vals[$cHm - 1] = $hm; $vals[$cHp - 1] = $PerfRpt::Rnd(($PerfRpt::Div($hu, $hm)), 4)
            $vals[$cNu - 1] = $PerfRpt::MB(($PerfRpt::GD($tj, $r, 'nonheap_used'))); $vals[$cNc - 1] = $PerfRpt::MB(($PerfRpt::GD($tj, $r, 'nonheap_committed')))
            foreach ($g2 in $mpCols.Keys) { $k2 = $epS + "`t" + $g2; if ($mpBy.ContainsKey($k2)) { $vals[$mpCols[$g2] - 1] = $PerfRpt::MB($mpBy[$k2])} }
            $vals[$cTh - 1] = $PerfRpt::GD($tj, $r, 'thread_count'); $vals[$cTp - 1] = $PerfRpt::GD($tj, $r, 'peak_thread_count'); $vals[$cTd - 1] = $PerfRpt::GD($tj, $r, 'daemon_thread_count')
            $vals[$cCl - 1] = $PerfRpt::GD($tj, $r, 'loaded_class_count'); $vals[$cCu - 1] = $PerfRpt::GD($tj, $r, 'unloaded_class_count'); $vals[$cLa - 1] = $PerfRpt::GD($tj, $r, 'system_load_average')
            foreach ($nm in $gcNames) {
                $k2 = $epS + "`t" + $nm
                if (-not $gcBy.ContainsKey($k2)) { continue }
                $g = $gcBy[$k2]
                if ($prevGc.ContainsKey($nm)) {
                    $dc = $PerfRpt::Dl(($PerfRpt::GD($tg, $g, 'collection_count')), ($PerfRpt::GD($tg, $prevGc[$nm], 'collection_count'))); $dtm = $PerfRpt::Dl(($PerfRpt::GD($tg, $g, 'collection_time_ms')), ($PerfRpt::GD($tg, $prevGc[$nm], 'collection_time_ms')))
                    $vals[$gcCols[$nm][0] - 1] = $dc; $vals[$gcCols[$nm][1] - 1] = $dtm
                    if (-not $gcDelta.ContainsKey($nm)) { $gcDelta[$nm] = @(0.0, 0.0) }
                    if ($null -ne $dc) { $gcDelta[$nm][0] += $dc }; if ($null -ne $dtm) { $gcDelta[$nm][1] += $dtm }
                }
                $prevGc[$nm] = $g
            }
            $PerfRpt::AddV($lHu, $hu); $PerfRpt::AddV($lNu, $vals[$cNu - 1]); if ($null -ne $hm) { $hmx = $hm }
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shJV = $sh
        $KPI['cliHeap'] = Stat $lHu; $KPI['cliHeapMax'] = $hmx; $KPI['cliGc'] = $gcDelta; $KPI['nonheap'] = Stat $lNu
        $script:JVC = @{ Hu = $cHu; Hc = $cHc; Hm = $cHm; Th = $cTh; Mp = $mpCols; Gc = $gcCols }
    }

    # ==== ヒープ・GC (jstat -gc) ====
    Write-Info '集計: ヒープ・GC（jstat）'
    $shJS = $null
    if ($TB.ContainsKey('jvm_jstat_gc')) {
        $t = $TB['jvm_jstat_gc']
        $sh = New-Sheet 'ヒープ・GC（jstat）' 'ヒープ・GC（jstat -gc）' '② jstat -gc。容量・使用量は MB、GC の「区間」は前回の行からの増分、GC 時間率は区間の GC 時間 ÷ 経過時間です。Full GC・Old 使用率 85% 以上・GC 時間率 5% 以上は濃淡で表示します。' 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cUp = Add-Col $sh '基本' 'JVM 稼働(秒)' 'uptime_s' 'dec1' 10
        $cHu = Add-Col $sh 'ヒープ（MB）' '使用（合計）' 'EU+S0U+S1U+OU' 'dec1' 10
        $cHc = Add-Col $sh 'ヒープ（MB）' '容量（合計）' 'EC+S0C+S1C+OC' 'dec1' 10
        $cEu = Add-Col $sh 'ヒープ（MB）' 'Eden 使用' 'EU' 'dec1' 9
        $cEc = Add-Col $sh 'ヒープ（MB）' 'Eden 容量' 'EC' 'dec1' 9
        $cSu = Add-Col $sh 'ヒープ（MB）' 'Survivor 使用' 'S0U+S1U' 'dec1' 9
        $cSc = Add-Col $sh 'ヒープ（MB）' 'Survivor 容量' 'S0C+S1C' 'dec1' 9
        $cOu = Add-Col $sh 'ヒープ（MB）' 'Old 使用' 'OU' 'dec1' 9
        $cOc = Add-Col $sh 'ヒープ（MB）' 'Old 容量' 'OC' 'dec1' 9
        $cOp = Add-Col $sh 'ヒープ（MB）' 'Old 使用率' '(OU÷OC)' 'pct' 9
        $cMu = Add-Col $sh 'メタスペース（MB）' '使用' 'MU' 'dec1' 8
        $cMc = Add-Col $sh 'メタスペース（MB）' '容量' 'MC' 'dec1' 8
        $cCu = Add-Col $sh 'メタスペース（MB）' 'CCS 使用' 'CCSU' 'dec1' 8
        $cYn = Add-Col $sh 'GC（区間）' 'Young 回数' 'Δ YGC' 'int' 8
        $cYt = Add-Col $sh 'GC（区間）' 'Young 時間 ms' 'Δ YGCT' 'dec1' 9
        $cYa = Add-Col $sh 'GC（区間）' 'Young 平均停止 ms' 'Δ YGCT÷Δ YGC' 'dec1' 10
        $cFn = Add-Col $sh 'GC（区間）' 'Full 回数' 'Δ FGC' 'int' 8
        $cFt = Add-Col $sh 'GC（区間）' 'Full 時間 ms' 'Δ FGCT' 'dec1' 9
        $cCn = Add-Col $sh 'GC（区間）' 'Concurrent 回数' 'Δ CGC' 'int' 10
        $cCt = Add-Col $sh 'GC（区間）' 'Concurrent 時間 ms' 'Δ CGCT' 'dec1' 10
        $cGt = Add-Col $sh 'GC（区間）' 'GC 時間 ms' 'Δ GCT' 'dec1' 9
        $cGr = Add-Col $sh 'GC（区間）' 'GC 時間率' 'Δ GCT÷経過' 'pct' 9
        $cYc = Add-Col $sh 'GC（累計）' 'YGC' 'YGC' 'int' 7
        $cFc = Add-Col $sh 'GC（累計）' 'FGC' 'FGC' 'int' 7
        $cCc = Add-Col $sh 'GC（累計）' 'CGC' 'CGC' 'int' 7
        $cGc = Add-Col $sh 'GC（累計）' 'GCT 秒' 'GCT' 'dec3' 8
        Add-CF $sh $cOp 'bar' '1' 0; Add-CF $sh $cOp 'ge' '0.85' 1; Add-CF $sh $cFn 'gt' '0' 0; Add-CF $sh $cGr 'ge' '0.05' 1
        $p = $null
        $lHu = NewList; $lOu = NewList; $lOp = NewList; $lMu = NewList; $lGr = NewList; $lYa = NewList; $ygS = 0.0; $ytS = 0.0; $fgS = 0.0; $ftS = 0.0; $cS = 0.0; $gtS = 0.0; $hcMax = $null
        $first = $null; $lastR = $null
        foreach ($r in $t.Rows) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            $ep = $PerfRpt::GD($t, $r, 'epoch')
            $eu = $PerfRpt::GD($t, $r, 'EU'); $s0 = $PerfRpt::GD($t, $r, 'S0U'); $s1 = $PerfRpt::GD($t, $r, 'S1U'); $ou = $PerfRpt::GD($t, $r, 'OU')
            $ec = $PerfRpt::GD($t, $r, 'EC'); $s0c = $PerfRpt::GD($t, $r, 'S0C'); $s1c = $PerfRpt::GD($t, $r, 'S1C'); $oc = $PerfRpt::GD($t, $r, 'OC')
            $hu = $null; if ($null -ne $eu -and $null -ne $ou) { $hu = $eu + [double]$s0 + [double]$s1 + $ou }
            $hc = $null; if ($null -ne $ec -and $null -ne $oc) { $hc = $ec + [double]$s0c + [double]$s1c + $oc }
            $vals[$cT - 1] = $PerfRpt::XT($ep); $vals[$cUp - 1] = $PerfRpt::GD($t, $r, 'uptime_s')
            $vals[$cHu - 1] = $PerfRpt::KBtoMB($hu); $vals[$cHc - 1] = $PerfRpt::KBtoMB($hc); $vals[$cEu - 1] = $PerfRpt::KBtoMB($eu); $vals[$cEc - 1] = $PerfRpt::KBtoMB($ec)
            $vals[$cSu - 1] = $PerfRpt::KBtoMB(([double]$s0 + [double]$s1)); $vals[$cSc - 1] = $PerfRpt::KBtoMB(([double]$s0c + [double]$s1c))
            $vals[$cOu - 1] = $PerfRpt::KBtoMB($ou); $vals[$cOc - 1] = $PerfRpt::KBtoMB($oc); $op = $PerfRpt::Div($ou, $oc); $vals[$cOp - 1] = $PerfRpt::Rnd($op, 4)
            $vals[$cMu - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'MU'))); $vals[$cMc - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'MC'))); $vals[$cCu - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'CCSU')))
            $vals[$cYc - 1] = $PerfRpt::GD($t, $r, 'YGC'); $vals[$cFc - 1] = $PerfRpt::GD($t, $r, 'FGC'); $vals[$cCc - 1] = $PerfRpt::GD($t, $r, 'CGC'); $vals[$cGc - 1] = $PerfRpt::GD($t, $r, 'GCT')
            if ($null -ne $p) {
                $el = $ep - ($PerfRpt::GD($t, $p, 'epoch'))
                $yn = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'YGC')), ($PerfRpt::GD($t, $p, 'YGC'))); $yt = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'YGCT')), ($PerfRpt::GD($t, $p, 'YGCT')))
                $fn = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'FGC')), ($PerfRpt::GD($t, $p, 'FGC'))); $ft = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'FGCT')), ($PerfRpt::GD($t, $p, 'FGCT')))
                $cgcn = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'CGC')), ($PerfRpt::GD($t, $p, 'CGC'))); $cgct = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'CGCT')), ($PerfRpt::GD($t, $p, 'CGCT')))
                $gt = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'GCT')), ($PerfRpt::GD($t, $p, 'GCT')))
                $vals[$cYn - 1] = $yn; $vals[$cYt - 1] = $PerfRpt::Rnd(([double]$yt * 1000), 1); $vals[$cFn - 1] = $fn; $vals[$cFt - 1] = $PerfRpt::Rnd(([double]$ft * 1000), 1)
                $vals[$cCn - 1] = $cgcn; $vals[$cCt - 1] = $PerfRpt::Rnd(([double]$cgct * 1000), 1); $vals[$cGt - 1] = $PerfRpt::Rnd(([double]$gt * 1000), 1)
                if ($null -ne $yn -and $yn -gt 0 -and $null -ne $yt) { $a = $yt * 1000 / $yn; $vals[$cYa - 1] = $PerfRpt::Rnd($a, 1); $PerfRpt::AddV($lYa, $a)}
                if ($el -gt 0 -and $null -ne $gt) { $gr = $gt / $el; $vals[$cGr - 1] = $PerfRpt::Rnd($gr, 4); $PerfRpt::AddV($lGr, $gr)}
                if ($null -ne $yn) { $ygS += $yn }; if ($null -ne $yt) { $ytS += $yt }; if ($null -ne $fn) { $fgS += $fn }; if ($null -ne $ft) { $ftS += $ft }; if ($null -ne $cgcn) { $cS += $cgcn }; if ($null -ne $gt) { $gtS += $gt }
            } else { $first = $r }
            $PerfRpt::AddV($lHu, ($PerfRpt::KBtoMB($hu))); $PerfRpt::AddV($lOu, ($PerfRpt::KBtoMB($ou))); $PerfRpt::AddV($lOp, $op); $PerfRpt::AddV($lMu, ($PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'MU')))))
            if ($null -ne $hc -and ($null -eq $hcMax -or $hc -gt $hcMax)) { $hcMax = $hc }
            Add-Row $sh $vals ''
            $p = $r; $lastR = $r
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shJS = $sh
        $elapsed = 0.0; if ($null -ne $first -and $null -ne $lastR) { $elapsed = ($PerfRpt::GD($t, $lastR, 'epoch')) - ($PerfRpt::GD($t, $first, 'epoch')) }
        $KPI['jsHeap'] = Stat $lHu; $KPI['jsOld'] = Stat $lOu; $KPI['jsOldP'] = Stat $lOp; $KPI['jsMeta'] = Stat $lMu; $KPI['jsGr'] = Stat $lGr; $KPI['jsYa'] = Stat $lYa
        $KPI['jsY'] = $ygS; $KPI['jsYt'] = $ytS; $KPI['jsF'] = $fgS; $KPI['jsFt'] = $ftS; $KPI['jsC'] = $cS; $KPI['jsGt'] = $gtS; $KPI['jsEl'] = $elapsed; $KPI['jsHcMax'] = $PerfRpt::KBtoMB($hcMax)
        $script:JSC = @{ Hu = $cHu; Hc = $cHc; Eu = $cEu; Su = $cSu; Ou = $cOu; Oc = $cOc; Mu = $cMu; Mc = $cMc; Cu = $cCu; Yn = $cYn; Fn = $cFn; Cn = $cCn; Yt = $cYt; Ft = $cFt; Ct = $cCt; Gt = $cGt; Gr = $cGr }
    }

    # ==== ヒープ詳細 (jcmd GC.heap_info) ====
    if ($TB.ContainsKey('jvm_jcmd_heap')) {
        $t = $TB['jvm_jcmd_heap']
        $sh = New-Sheet 'ヒープ詳細（jcmd）' 'ヒープ詳細（jcmd GC.heap_info）' '② jcmd <pid> GC.heap_info を数値化したもの。G1 はリージョン数、Serial / Parallel は世代ごとの容量・使用量が入ります。元の出力は jvm_jcmd.log にあります。' 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cG = Add-Col $sh '基本' 'GC' 'gc' 'textc' 7
        $cHt = Add-Col $sh 'ヒープ（MB）' '容量' 'heap_total_kb' 'dec1' 9
        $cHu = Add-Col $sh 'ヒープ（MB）' '使用' 'heap_used_kb' 'dec1' 9
        $cHp = Add-Col $sh 'ヒープ（MB）' '使用率' '(使用÷容量)' 'pct' 8
        $cRs = Add-Col $sh 'G1 リージョン' 'サイズ KB' 'region_size_kb' 'int' 8
        $cYr = Add-Col $sh 'G1 リージョン' 'Young 数' 'young_regions' 'int' 8
        $cYk = Add-Col $sh 'G1 リージョン' 'Young MB' 'young_kb' 'dec1' 8
        $cSr = Add-Col $sh 'G1 リージョン' 'Survivor 数' 'survivor_regions' 'int' 9
        $cSk = Add-Col $sh 'G1 リージョン' 'Survivor MB' 'survivor_kb' 'dec1' 9
        $cYg = Add-Col $sh '世代（Serial / Parallel）' 'Young 使用 MB' 'young_gen_used_kb' 'dec1' 10
        $cOg = Add-Col $sh '世代（Serial / Parallel）' 'Old 使用 MB' 'old_gen_used_kb' 'dec1' 10
        $cMu = Add-Col $sh 'メタスペース（MB）' '使用' 'metaspace_used_kb' 'dec1' 8
        $cMc = Add-Col $sh 'メタスペース（MB）' 'コミット' 'metaspace_committed_kb' 'dec1' 9
        $cMr = Add-Col $sh 'メタスペース（MB）' '予約' 'metaspace_reserved_kb' 'dec1' 9
        $cCu = Add-Col $sh 'クラス空間（MB）' '使用' 'class_used_kb' 'dec1' 8
        $cCc = Add-Col $sh 'クラス空間（MB）' 'コミット' 'class_committed_kb' 'dec1' 9
        $cMs = Add-Col $sh '基本' 'jcmd 所要 ms' 'jcmd_ms' 'int' 9
        foreach ($r in $t.Rows) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            $ht = $PerfRpt::GD($t, $r, 'heap_total_kb'); $hu = $PerfRpt::GD($t, $r, 'heap_used_kb')
            $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::GD($t, $r, 'epoch'))); $vals[$cG - 1] = $PerfRpt::G($t, $r, 'gc')
            $vals[$cHt - 1] = $PerfRpt::KBtoMB($ht); $vals[$cHu - 1] = $PerfRpt::KBtoMB($hu); $vals[$cHp - 1] = $PerfRpt::Rnd(($PerfRpt::Div($hu, $ht)), 4)
            $vals[$cRs - 1] = $PerfRpt::GD($t, $r, 'region_size_kb'); $vals[$cYr - 1] = $PerfRpt::GD($t, $r, 'young_regions'); $vals[$cYk - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'young_kb')))
            $vals[$cSr - 1] = $PerfRpt::GD($t, $r, 'survivor_regions'); $vals[$cSk - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'survivor_kb')))
            $vals[$cYg - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'young_gen_used_kb'))); $vals[$cOg - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'old_gen_used_kb')))
            $vals[$cMu - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'metaspace_used_kb'))); $vals[$cMc - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'metaspace_committed_kb'))); $vals[$cMr - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'metaspace_reserved_kb')))
            $vals[$cCu - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'class_used_kb'))); $vals[$cCc - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'class_committed_kb'))); $vals[$cMs - 1] = $PerfRpt::GD($t, $r, 'jcmd_ms')
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh)
    }
    foreach ($t in $extraJstat) {
        $opt = $t.Name -replace '^jvm_jstat_', '' -replace '\.csv$', ''
        $sh = New-Sheet ('jstat -' + $opt) ('jstat -' + $opt) ('② jstat -' + $opt + ' の出力そのものです (列名は jstat の見出し)。') 1
        $cols = @()
        foreach ($h in $t.H) {
            if ($h -eq 'epoch') { $cols += 0; continue }
            if ($h -eq 'timestamp') { $cols += (Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19); continue }
            $f = 'dec1'; if ($h -match 'GCC|LGCC') { $f = 'text' }
            $cols += (Add-Col $sh 'jstat' $h $h $f $(if ($f -eq 'text') { 22 } else { 9 }))
        }
        foreach ($r in $t.Rows) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            for ($i = 0; $i -lt $t.H.Length; $i++) {
                $c = $cols[$i]; if ($c -eq 0) { continue }
                if ($t.H[$i] -eq 'timestamp') { $vals[$c - 1] = $PerfRpt::XT(($PerfRpt::GD($t, $r, 'epoch')))} else { $vals[$c - 1] = $r[$i] }
            }
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh)
    }

    # ==== CPU・メモリ (top) ====
    Write-Info '集計: CPU・メモリ（top）'
    $shTop = $null
    $ovh = @{}
    if ($TB.ContainsKey('os_top_summary')) {
        $ts = $TB['os_top_summary']
        $pBy = @{}
        if ($TB.ContainsKey('os_top_process')) {
            $tp = $TB['os_top_process']
            foreach ($r in $tp.Rows) {
                $k2 = $PerfRpt::G($tp, $r, 'epoch')
                if (-not $pBy.ContainsKey($k2)) { $pBy[$k2] = @{} }
                $role = $PerfRpt::G($tp, $r, 'role')
                $c = $PerfRpt::GD($tp, $r, 'cpu_pct'); $m = $PerfRpt::GD($tp, $r, 'res_kb')
                if (-not $pBy[$k2].ContainsKey($role)) { $pBy[$k2][$role] = @(0.0, 0.0, 0) }
                $pBy[$k2][$role][0] += [double]$c; $pBy[$k2][$role][1] += [double]$m; $pBy[$k2][$role][2] += 1
            }
        }
        $sh = New-Sheet 'CPU・メモリ（top）' 'CPU・メモリ（top）' ('③ top -b の要約行とプロセス行。プロセスの %CPU は「1 コア = 100%」(複数コアを使うと 100% を超えます)。CPU 内訳は全 CPU に対する割合です。収集プロセス (jboss-cli・jstat・jcmd・スクリプト) 自身の負荷も確認できます。') 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cL1 = Add-Col $sh 'ロードアベレージ' '1 分' 'load1' 'dec2' 7
        $cL5 = Add-Col $sh 'ロードアベレージ' '5 分' 'load5' 'dec2' 7
        $cL15 = Add-Col $sh 'ロードアベレージ' '15 分' 'load15' 'dec2' 7
        $cCu = Add-Col $sh 'CPU（全体に対する %）' '使用率' '100 - id' 'dec1' 8
        $cUs = Add-Col $sh 'CPU（全体に対する %）' 'us' 'cpu_us' 'dec1' 6
        $cSy = Add-Col $sh 'CPU（全体に対する %）' 'sy' 'cpu_sy' 'dec1' 6
        $cWa = Add-Col $sh 'CPU（全体に対する %）' 'wa' 'cpu_wa' 'dec1' 6
        $cSt = Add-Col $sh 'CPU（全体に対する %）' 'st' 'cpu_st' 'dec1' 6
        $cMt = Add-Col $sh 'メモリ（MiB）' '合計' 'mem_total_mib' 'dec1' 8
        $cMu = Add-Col $sh 'メモリ（MiB）' '使用' 'mem_used_mib' 'dec1' 8
        $cMa = Add-Col $sh 'メモリ（MiB）' '利用可能' 'mem_avail_mib' 'dec1' 8
        $cMb = Add-Col $sh 'メモリ（MiB）' 'buff/cache' 'mem_buffcache_mib' 'dec1' 9
        $cJc = Add-Col $sh 'JBoss EAP' '%CPU' 'role=jboss cpu_pct' 'dec1' 8
        $cJr = Add-Col $sh 'JBoss EAP' 'RES MB' 'role=jboss res_kb' 'dec1' 8
        $cKc = Add-Col $sh '収集プロセス（性能試験用）' 'jboss-cli %CPU' 'role=jboss-cli' 'dec1' 10
        $cKr = Add-Col $sh '収集プロセス（性能試験用）' 'jboss-cli RES MB' 'role=jboss-cli' 'dec1' 10
        $cSc = Add-Col $sh '収集プロセス（性能試験用）' 'jstat %CPU' 'role=jstat' 'dec1' 9
        $cSr = Add-Col $sh '収集プロセス（性能試験用）' 'jstat RES MB' 'role=jstat' 'dec1' 9
        $cCc = Add-Col $sh '収集プロセス（性能試験用）' 'jcmd %CPU' 'role=jcmd' 'dec1' 9
        $cPc = Add-Col $sh '収集プロセス（性能試験用）' 'スクリプト %CPU' 'perf-metrics/top/ps' 'dec1' 10
        $cOt = Add-Col $sh '収集プロセス（性能試験用）' '収集の合計 %CPU' '(上の合計)' 'dec1' 10
        $cTk = Add-Col $sh 'タスク' 'プロセス数' 'tasks_total' 'int' 8
        $lJc = NewList; $lJr = NewList; $lCu = NewList; $lL1 = NewList; $lOt = NewList; $lKr = NewList; $lSr = NewList; $lKc = NewList
        foreach ($r in $ts.Rows) {
            $epS = $PerfRpt::G($ts, $r, 'epoch')
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::D($epS)))
            $vals[$cL1 - 1] = $PerfRpt::GD($ts, $r, 'load1'); $vals[$cL5 - 1] = $PerfRpt::GD($ts, $r, 'load5'); $vals[$cL15 - 1] = $PerfRpt::GD($ts, $r, 'load15')
            $id = $PerfRpt::GD($ts, $r, 'cpu_id'); if ($null -ne $id) { $vals[$cCu - 1] = $PerfRpt::Rnd((100 - $id), 1); $PerfRpt::AddV($lCu, (100 - $id))}
            $vals[$cUs - 1] = $PerfRpt::GD($ts, $r, 'cpu_us'); $vals[$cSy - 1] = $PerfRpt::GD($ts, $r, 'cpu_sy'); $vals[$cWa - 1] = $PerfRpt::GD($ts, $r, 'cpu_wa'); $vals[$cSt - 1] = $PerfRpt::GD($ts, $r, 'cpu_st')
            $vals[$cMt - 1] = $PerfRpt::Rnd(($PerfRpt::GD($ts, $r, 'mem_total_mib')), 1); $vals[$cMu - 1] = $PerfRpt::Rnd(($PerfRpt::GD($ts, $r, 'mem_used_mib')), 1); $vals[$cMa - 1] = $PerfRpt::Rnd(($PerfRpt::GD($ts, $r, 'mem_avail_mib')), 1); $vals[$cMb - 1] = $PerfRpt::Rnd(($PerfRpt::GD($ts, $r, 'mem_buffcache_mib')), 1)
            $vals[$cTk - 1] = $PerfRpt::GD($ts, $r, 'tasks_total'); $PerfRpt::AddV($lL1, ($PerfRpt::GD($ts, $r, 'load1')))
            if ($pBy.ContainsKey($epS)) {
                $pr = $pBy[$epS]
                if ($pr.ContainsKey('jboss')) { $vals[$cJc - 1] = $PerfRpt::Rnd($pr['jboss'][0], 1); $vals[$cJr - 1] = $PerfRpt::Rnd(($pr['jboss'][1] / 1024), 1); $PerfRpt::AddV($lJc, $pr['jboss'][0]); $PerfRpt::AddV($lJr, ($pr['jboss'][1] / 1024))}
                $kc = 0.0; $tot = 0.0
                if ($pr.ContainsKey('jboss-cli')) { $kc = $pr['jboss-cli'][0]; $vals[$cKc - 1] = $PerfRpt::Rnd($kc, 1); $vals[$cKr - 1] = $PerfRpt::Rnd(($pr['jboss-cli'][1] / 1024), 1); $PerfRpt::AddV($lKr, ($pr['jboss-cli'][1] / 1024))} else { $vals[$cKc - 1] = 0.0 }
                $tot += $kc; $PerfRpt::AddV($lKc, $kc)
                $sc = 0.0; if ($pr.ContainsKey('jstat')) { $sc = $pr['jstat'][0]; $vals[$cSr - 1] = $PerfRpt::Rnd(($pr['jstat'][1] / 1024), 1); $PerfRpt::AddV($lSr, ($pr['jstat'][1] / 1024))}; $vals[$cSc - 1] = $PerfRpt::Rnd($sc, 1); $tot += $sc
                $cc = 0.0; if ($pr.ContainsKey('jcmd')) { $cc = $pr['jcmd'][0] }; $vals[$cCc - 1] = $PerfRpt::Rnd($cc, 1); $tot += $cc
                $pc = 0.0; foreach ($rl in @('perf-metrics', 'top', 'ps')) { if ($pr.ContainsKey($rl)) { $pc += $pr[$rl][0] } }; $vals[$cPc - 1] = $PerfRpt::Rnd($pc, 1); $tot += $pc
                $vals[$cOt - 1] = $PerfRpt::Rnd($tot, 1); $PerfRpt::AddV($lOt, $tot)
                foreach ($rl in @('jboss-cli', 'jstat', 'jcmd', 'perf-metrics', 'top', 'ps')) {
                    if (-not $ovh.ContainsKey($rl)) { $ovh[$rl] = @((NewList), (NewList)) }
                    if ($pr.ContainsKey($rl)) { $PerfRpt::AddV($ovh[$rl][0], $pr[$rl][0]); $PerfRpt::AddV($ovh[$rl][1], ($pr[$rl][1] / 1024))} else { $PerfRpt::AddV($ovh[$rl][0], 0.0)}
                }
            }
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shTop = $sh
        $KPI['jbCpu'] = Stat $lJc; $KPI['jbRes'] = Stat $lJr; $KPI['sysCpu'] = Stat $lCu; $KPI['load1'] = Stat $lL1; $KPI['ovhCpu'] = Stat $lOt
        $script:TPC = @{ Jc = $cJc; Jr = $cJr; Kc = $cKc; Sc = $cSc; Cc = $cCc; Pc = $cPc; Ot = $cOt; L1 = $cL1; L5 = $cL5; L15 = $cL15; Cu = $cCu; Mu = $cMu; Kr = $cKr; Sr = $cSr }
    }

    # ==== スレッド別 CPU (top -H) ====
    Write-Info '集計: スレッド別 CPU'
    if ($TB.ContainsKey('os_top_threads')) {
        $t = $TB['os_top_threads']
        $frames = @{}; $agg = @{}
        foreach ($r in $t.Rows) {
            $ep = $PerfRpt::G($t, $r, 'epoch'); $frames[$ep] = 1
            $nm = [string]($PerfRpt::G($t, $r, 'thread'))
            $gname = [regex]::Replace($nm, '[0-9]+', '#')
            $c = $PerfRpt::D(($PerfRpt::G($t, $r, 'cpu_pct'))); if ($null -eq $c) { $c = 0.0 }
            if (-not $agg.ContainsKey($gname)) { $agg[$gname] = @{ Sum = @{}; MaxOne = 0.0; Seen = 0; Names = @{} } }
            $a = $agg[$gname]
            if ($a.Sum.ContainsKey($ep)) { $a.Sum[$ep] += $c } else { $a.Sum[$ep] = $c }
            if ($c -gt $a.MaxOne) { $a.MaxOne = $c }
            $a.Seen++; $a.Names[$nm] = 1
        }
        $nf = $frames.Count
        $sh = New-Sheet 'スレッド別CPU（top）' 'JBoss のスレッド別 CPU（top -H、スレッド名ごとに集計）' ('③ top -H -p <JBoss の PID> で各回の %CPU 上位 ' + $(if ($meta.Contains('perf.top_threads')) { [string]$meta['perf.top_threads'] + ' 件' } else { '' }) + 'のスレッドを記録し、スレッド名の数字を # にまとめて集計しています (例: default task-12 → default task-#)。上位に入らなかった回は 0 として平均しています。%CPU は 1 コア = 100%。') 1
        $cN = Add-Col $sh 'スレッド' 'スレッド名（まとめ）' 'thread' 'text' 34
        $cAv = Add-Col $sh '%CPU（グループ合計）' '平均' 'avg' 'dec2' 9
        $cMx = Add-Col $sh '%CPU（グループ合計）' '最大' 'max' 'dec1' 9
        $cMo = Add-Col $sh '%CPU（1 スレッド）' '最大' 'max(1 thread)' 'dec1' 9
        $cSe = Add-Col $sh '出現' '上位に入った回数' 'count' 'int' 10
        $cFr = Add-Col $sh '出現' '割合' '(回数÷取得回数)' 'pct' 8
        $cTn = Add-Col $sh '出現' 'スレッド数' '(名前の種類)' 'int' 8
        Add-CF $sh $cAv 'bar' ([string]([Math]::Max(1, [Math]::Ceiling((@($agg.Values | ForEach-Object { ($_.Sum.Values | Measure-Object -Sum).Sum }) | Measure-Object -Maximum).Maximum / [Math]::Max(1, $nf))))) 0
        $list = @()
        foreach ($gname in $agg.Keys) {
            $a = $agg[$gname]; $sum = 0.0; $mx = 0.0
            foreach ($v in $a.Sum.Values) { $sum += $v; if ($v -gt $mx) { $mx = $v } }
            $list += [pscustomobject]@{ N = $gname; Avg = $(if ($nf -gt 0) { $sum / $nf } else { 0 }); Max = $mx; MaxOne = $a.MaxOne; Seen = $a.Sum.Count; Names = $a.Names.Count }
        }
        foreach ($x in ($list | Sort-Object -Property @{ Expression = 'Avg'; Descending = $true }, @{ Expression = 'Max'; Descending = $true })) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cN - 1] = $x.N; $vals[$cAv - 1] = $PerfRpt::Rnd($x.Avg, 2); $vals[$cMx - 1] = $PerfRpt::Rnd($x.Max, 1); $vals[$cMo - 1] = $PerfRpt::Rnd($x.MaxOne, 1)
            $vals[$cSe - 1] = $x.Seen; $vals[$cFr - 1] = $PerfRpt::Rnd(($PerfRpt::Div($x.Seen, $nf)), 4); $vals[$cTn - 1] = $x.Names
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh)
        $KPI['threads'] = $list | Sort-Object -Property Avg -Descending | Select-Object -First 5
    }

    # ==== プロセス (ps) ====
    Write-Info '集計: プロセス（ps）'
    $shPs = $null
    if ($TB.ContainsKey('os_ps')) {
        $t = $TB['os_ps']
        $sh = New-Sheet 'プロセス（ps）' 'プロセス（ps）：JBoss のスレッド数・メモリ・CPU と収集プロセスの合計' '③ ps -e -o pid,ppid,user,stat,nlwp,rss,vsz,pcpu,pmem,etimes,times,args。CPU（区間）は累計 CPU 秒の増分 ÷ 経過秒で、1 コア = 100% です (ps の累計 CPU 秒は 1 秒単位のため、短い間隔では粗い値になります)。' 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cPid = Add-Col $sh 'JBoss EAP' 'PID' 'pid' 'int' 7
        $cNl = Add-Col $sh 'JBoss EAP' 'スレッド数' 'nlwp' 'int' 9
        $cRs = Add-Col $sh 'JBoss EAP' 'RSS MB' 'rss_kb' 'dec1' 9
        $cVs = Add-Col $sh 'JBoss EAP' 'VSZ MB' 'vsz_kb' 'dec1' 10
        $cCp = Add-Col $sh 'JBoss EAP' 'CPU（区間）%' 'Δ cpu_time_s÷経過' 'dec1' 10
        $cCs = Add-Col $sh 'JBoss EAP' '累計 CPU 秒' 'cpu_time_s' 'int' 9
        $cEt = Add-Col $sh 'JBoss EAP' '稼働秒' 'etimes' 'int' 9
        $cKr = Add-Col $sh '収集プロセス（性能試験用）' 'RSS 合計 MB' 'jboss-cli+jstat+jcmd+scripts' 'dec1' 11
        $cKc = Add-Col $sh '収集プロセス（性能試験用）' 'CPU（区間）%' 'Δ cpu_time_s÷経過' 'dec1' 10
        $cKn = Add-Col $sh '収集プロセス（性能試験用）' 'プロセス数' 'count' 'int' 9
        $cNp = Add-Col $sh '全体' 'プロセス数' 'count' 'int' 8
        $cAr = Add-Col $sh '全体' 'RSS 合計 MB' 'sum rss_kb' 'dec1' 10
        $byEp = New-Object System.Collections.Specialized.OrderedDictionary
        $ei = $t.I['epoch']
        foreach ($r in $t.Rows) { $k2 = $r[$ei]; if (-not $byEp.Contains($k2)) { $byEp[$k2] = New-Object System.Collections.Generic.List[object] }; $byEp[$k2].Add($r) }
        $pj = $null; $pk = $null; $pe = $null
        $lNl = NewList; $lRs = NewList; $lKr = NewList; $lCp = NewList
        $collect = @{ 'jboss-cli' = 1; 'jstat' = 1; 'jcmd' = 1; 'perf-metrics' = 1; 'top' = 1; 'ps' = 1 }
        foreach ($k2 in $byEp.Keys) {
            $ep = $PerfRpt::D($k2)
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cT - 1] = $PerfRpt::XT($ep)
            $jb = $null; $kr = 0.0; $kcpu = 0.0; $kn = 0; $np = 0; $ar = 0.0
            foreach ($r in $byEp[$k2]) {
                $np++; $rs = $PerfRpt::GD($t, $r, 'rss_kb'); if ($null -ne $rs) { $ar += $rs }
                $role = $PerfRpt::G($t, $r, 'role')
                if ($role -eq 'jboss') { $jb = $r }
                elseif ($collect.ContainsKey($role)) { $kn++; if ($null -ne $rs) { $kr += $rs }; $v = $PerfRpt::GD($t, $r, 'cpu_time_s'); if ($null -ne $v) { $kcpu += $v } }
            }
            if ($null -ne $jb) {
                $vals[$cPid - 1] = $PerfRpt::GD($t, $jb, 'pid'); $vals[$cNl - 1] = $PerfRpt::GD($t, $jb, 'nlwp'); $vals[$cRs - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $jb, 'rss_kb'))); $vals[$cVs - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $jb, 'vsz_kb')))
                $vals[$cCs - 1] = $PerfRpt::GD($t, $jb, 'cpu_time_s'); $vals[$cEt - 1] = $PerfRpt::GD($t, $jb, 'etimes')
                if ($null -ne $pj -and $null -ne $pe -and $ep -gt $pe -and ($PerfRpt::GD($t, $pj, 'pid')) -eq ($PerfRpt::GD($t, $jb, 'pid'))) {
                    $x = $PerfRpt::Dl(($PerfRpt::GD($t, $jb, 'cpu_time_s')), ($PerfRpt::GD($t, $pj, 'cpu_time_s'))); if ($null -ne $x) { $c = $x / ($ep - $pe) * 100; $vals[$cCp - 1] = $PerfRpt::Rnd($c, 1); $PerfRpt::AddV($lCp, $c)}
                }
                $PerfRpt::AddV($lNl, ($PerfRpt::GD($t, $jb, 'nlwp'))); $PerfRpt::AddV($lRs, ($PerfRpt::KBtoMB(($PerfRpt::GD($t, $jb, 'rss_kb'))))); $pj = $jb
            }
            if ($null -ne $pk -and $null -ne $pe -and $ep -gt $pe) { $x = $kcpu - $pk; if ($x -ge 0) { $vals[$cKc - 1] = $PerfRpt::Rnd(($x / ($ep - $pe) * 100), 1)} }
            $vals[$cKr - 1] = $PerfRpt::KBtoMB($kr); $vals[$cKn - 1] = $kn; $vals[$cNp - 1] = $np; $vals[$cAr - 1] = $PerfRpt::KBtoMB($ar)
            $PerfRpt::AddV($lKr, ($PerfRpt::KBtoMB($kr)))
            Add-Row $sh $vals ''
            $pk = $kcpu; $pe = $ep
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shPs = $sh
        $KPI['psNl'] = Stat $lNl; $KPI['psRs'] = Stat $lRs; $KPI['psKr'] = Stat $lKr; $KPI['psCp'] = Stat $lCp
        $script:PSC = @{ Nl = $cNl; Rs = $cRs; Kr = $cKr; Cp = $cCp }

        # 一覧 (全行)
        $sh = New-Sheet 'プロセス一覧（ps）' 'プロセス一覧（ps の全行）' '③ ps の出力を 1 プロセス 1 行で載せています。オートフィルターで「役割」「PID」などを絞り込めます。役割: jboss = JBoss EAP、jboss-cli / jstat / jcmd / perf-metrics / top / ps = 性能試験用の収集プロセス。パスワードらしき値は伏せています。' 2
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cPid = Add-Col $sh '基本' 'PID' 'pid' 'int' 7
        $cPp = Add-Col $sh '基本' 'PPID' 'ppid' 'int' 7
        $cRo = Add-Col $sh '基本' '役割' 'role' 'textc' 12
        $cUs = Add-Col $sh '基本' 'ユーザー' 'user' 'textc' 9
        $cSt = Add-Col $sh '状態' '状態' 'stat' 'textc' 6
        $cNl = Add-Col $sh '状態' 'スレッド数' 'nlwp' 'int' 8
        $cRs = Add-Col $sh 'メモリ' 'RSS MB' 'rss_kb' 'dec1' 8
        $cVs = Add-Col $sh 'メモリ' 'VSZ MB' 'vsz_kb' 'dec1' 9
        $cPm = Add-Col $sh 'メモリ' '%MEM' 'pmem' 'dec1' 7
        $cPc = Add-Col $sh 'CPU' '%CPU（起動からの平均）' 'pcpu' 'dec1' 11
        $cCs = Add-Col $sh 'CPU' '累計 CPU 秒' 'cpu_time_s' 'int' 9
        $cEt = Add-Col $sh 'CPU' '稼働秒' 'etimes' 'int' 8
        $cAg = Add-Col $sh 'コマンド' 'コマンドライン' 'args' 'text' 100
        foreach ($r in $t.Rows) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::GD($t, $r, 'epoch'))); $vals[$cPid - 1] = $PerfRpt::GD($t, $r, 'pid'); $vals[$cPp - 1] = $PerfRpt::GD($t, $r, 'ppid'); $vals[$cRo - 1] = $PerfRpt::G($t, $r, 'role')
            $vals[$cUs - 1] = $PerfRpt::G($t, $r, 'user'); $vals[$cSt - 1] = $PerfRpt::G($t, $r, 'stat'); $vals[$cNl - 1] = $PerfRpt::GD($t, $r, 'nlwp')
            $vals[$cRs - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'rss_kb'))); $vals[$cVs - 1] = $PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'vsz_kb'))); $vals[$cPm - 1] = $PerfRpt::GD($t, $r, 'pmem')
            $vals[$cPc - 1] = $PerfRpt::GD($t, $r, 'pcpu'); $vals[$cCs - 1] = $PerfRpt::GD($t, $r, 'cpu_time_s'); $vals[$cEt - 1] = $PerfRpt::GD($t, $r, 'etimes')
            $a = [string]($PerfRpt::G($t, $r, 'args')); if ($a.Length -gt 1000) { $a = $a.Substring(0, 1000) + ' …' }
            $vals[$cAg - 1] = $a
            Add-Row $sh $vals ''
        }
        Fix-Ranges $sh; $Sheets.Add($sh)
    }

    # ==== コンテナ (cgroup) ====
    Write-Info '集計: コンテナ（cgroup）'
    $shCg = $null
    if ($TB.ContainsKey('os_cgroup')) {
        $t = $TB['os_cgroup']
        $sh = New-Sheet 'コンテナ（cgroup）' 'コンテナ（cgroup）：CPU 使用率・スロットリング・メモリ' ('③ /sys/fs/cgroup。CPU 使用率は「1 vCPU = 100%」で、上限比は CPU 上限 (cpu.max / cfs_quota) に対する割合です。スロットリングは CPU 上限に当たって待たされた回数・時間 (区間)。メモリ使用率 90% 以上・OOM kill は濃淡で表示します。') 1
        $cT = Add-Col $sh '基本' '日時' 'timestamp' 'dt' 19
        $cCu = Add-Col $sh 'CPU' '使用率 %' 'Δ usage_usec÷経過' 'dec1' 9
        $cQ = Add-Col $sh 'CPU' '上限 vCPU' 'cpu_quota_cores' 'dec2' 9
        $cQr = Add-Col $sh 'CPU' '上限比' '(使用率÷上限)' 'pct' 8
        $cTn = Add-Col $sh 'CPU' 'スロットリング回数' 'Δ nr_throttled' 'int' 11
        $cTt = Add-Col $sh 'CPU' 'スロットリング ms' 'Δ throttled_usec' 'dec1' 11
        $cTr = Add-Col $sh 'CPU' 'スロットリング率' 'Δ nr_throttled÷Δ nr_periods' 'pct' 10
        $cMu = Add-Col $sh 'メモリ' '使用 MB' 'memory_usage_bytes' 'dec1' 9
        $cMl = Add-Col $sh 'メモリ' '上限 MB' 'memory_limit_bytes' 'dec1' 9
        $cMp = Add-Col $sh 'メモリ' '使用率' '(使用÷上限)' 'pct' 8
        $cMa = Add-Col $sh 'メモリ' 'anon MB' 'memory_anon_bytes' 'dec1' 9
        $cMf = Add-Col $sh 'メモリ' 'file MB' 'memory_file_bytes' 'dec1' 9
        $cOo = Add-Col $sh 'メモリ' 'OOM kill（区間）' 'Δ memory_oom_kill' 'int' 10
        $cPc = Add-Col $sh 'その他' 'プロセス数' 'pids_current' 'int' 9
        $cVv = Add-Col $sh 'その他' 'cgroup' 'cgroup_version' 'textc' 7
        Add-CF $sh $cQr 'bar' '1' 0; Add-CF $sh $cTn 'gt' '0' 2; Add-CF $sh $cMp 'bar' '1' 0; Add-CF $sh $cMp 'ge' '0.9' 1; Add-CF $sh $cOo 'gt' '0' 0
        $p = $null
        $lCu = NewList; $lQr = NewList; $lMu = NewList; $lMp = NewList; $thN = 0.0; $thT = 0.0; $oom = 0.0; $quota = $null; $mlim = $null
        foreach ($r in $t.Rows) {
            $vals = New-Object 'object[]' $sh.Cols.Count
            $ep = $PerfRpt::GD($t, $r, 'epoch')
            $vals[$cT - 1] = $PerfRpt::XT($ep)
            $q = $PerfRpt::GD($t, $r, 'cpu_quota_cores'); $vals[$cQ - 1] = $q; if ($null -ne $q) { $quota = $q }
            $mu = $PerfRpt::MB(($PerfRpt::GD($t, $r, 'memory_usage_bytes'))); $ml = $PerfRpt::MB(($PerfRpt::GD($t, $r, 'memory_limit_bytes'))); if ($null -ne $ml) { $mlim = $ml }
            $vals[$cMu - 1] = $mu; $vals[$cMl - 1] = $ml; $mp2 = $PerfRpt::Div($mu, $ml); $vals[$cMp - 1] = $PerfRpt::Rnd($mp2, 4)
            $vals[$cMa - 1] = $PerfRpt::MB(($PerfRpt::GD($t, $r, 'memory_anon_bytes'))); $vals[$cMf - 1] = $PerfRpt::MB(($PerfRpt::GD($t, $r, 'memory_file_bytes')))
            $vals[$cPc - 1] = $PerfRpt::GD($t, $r, 'pids_current'); $vals[$cVv - 1] = $PerfRpt::G($t, $r, 'cgroup_version')
            if ($null -ne $p) {
                $el = $ep - ($PerfRpt::GD($t, $p, 'epoch'))
                if ($el -gt 0) {
                    $du = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'cpu_usage_usec')), ($PerfRpt::GD($t, $p, 'cpu_usage_usec')))
                    if ($null -ne $du) { $cu = $du / ($el * 1000000.0) * 100; $vals[$cCu - 1] = $PerfRpt::Rnd($cu, 1); $PerfRpt::AddV($lCu, $cu); if ($null -ne $q -and $q -gt 0) { $qr = $cu / ($q * 100); $vals[$cQr - 1] = $PerfRpt::Rnd($qr, 4); $PerfRpt::AddV($lQr, $qr)} }
                }
                $tn = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'cpu_nr_throttled')), ($PerfRpt::GD($t, $p, 'cpu_nr_throttled'))); $vals[$cTn - 1] = $tn; if ($null -ne $tn) { $thN += $tn }
                $tt = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'cpu_throttled_usec')), ($PerfRpt::GD($t, $p, 'cpu_throttled_usec'))); if ($null -ne $tt) { $vals[$cTt - 1] = $PerfRpt::Rnd(($tt / 1000), 1); $thT += $tt / 1000 }
                $np = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'cpu_nr_periods')), ($PerfRpt::GD($t, $p, 'cpu_nr_periods'))); $vals[$cTr - 1] = $PerfRpt::Rnd(($PerfRpt::Div($tn, $np)), 4)
                $oo = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'memory_oom_kill')), ($PerfRpt::GD($t, $p, 'memory_oom_kill'))); $vals[$cOo - 1] = $oo; if ($null -ne $oo) { $oom += $oo }
            }
            $PerfRpt::AddV($lMu, $mu); $PerfRpt::AddV($lMp, $mp2)
            Add-Row $sh $vals ''
            $p = $r
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shCg = $sh
        $KPI['cgCpu'] = Stat $lCu; $KPI['cgQr'] = Stat $lQr; $KPI['cgMem'] = Stat $lMu; $KPI['cgMemP'] = Stat $lMp; $KPI['cgThN'] = $thN; $KPI['cgThT'] = $thT; $KPI['cgOom'] = $oom; $KPI['cgQuota'] = $quota; $KPI['cgLimit'] = $mlim
        $script:CGRC = @{ Cu = $cCu; Q = $cQ; Mu = $cMu; Ml = $cMl; Tn = $cTn }
    }

    # ==== 所見 ====
    function AddF([int]$lv, [string]$cat, [string]$msg) { $Find.Add(@($lv, $cat, $msg)) }
    foreach ($d in $dsSum) {
        if ($d.Fail -gt 0) { AddF 3 'DB 接続プール' ('「' + $d.Name + '」で接続の取得失敗 (タイムアウト) が ' + $d.Fail.ToString('#,0') + ' 件ありました。アプリでは IJ000453 等の例外になります。最大プール (' + $d.Max + ') と、接続を保持している処理 (スレッドダンプ) を確認してください。') }
        if ($d.AtMax -gt 0) { AddF 2 'DB 接続プール' ('「' + $d.Name + '」の使用中が最大プール (' + $d.Max + ') に達しました (' + $d.AtMax + ' / ' + $d.N + ' 回)。') }
        if ($d.Wait -gt 0) { AddF 2 'DB 接続プール' ('「' + $d.Name + '」で空き接続の待ちが ' + $d.Wait.ToString('#,0') + ' 件発生しました (最大待ち時間 ' + $(if ($null -ne $d.MaxWait) { $d.MaxWait.ToString('#,0') } else { '-' }) + ' ms)。') }
        elseif ($null -ne $d.Us -and $d.Us.Max -ge 0.8) { AddF 1 'DB 接続プール' ('「' + $d.Name + '」の使用率が最大 ' + ($d.Us.Max * 100).ToString('0') + '% まで上がりました (待ちは発生していません)。') }
        if ($d.StatsOff) { AddF 1 'DB 接続プール' ('「' + $d.Name + '」のプール統計が無効です。値は 0 のままになります。') }
    }
    if ($null -ne $KPI['ioMax'] -and $KPI['ioFullN'] -gt 0) {
        $qm = $(if ($null -ne $KPI['ioq']) { $KPI['ioq'].Max } else { 0 })
        AddF 2 'リクエスト処理' ('IO ワーカーの処理中スレッドが最大 (' + $KPI['ioMax'] + ') に達しました (' + $KPI['ioFullN'] + ' 回、キュー最大 ' + $qm + ')。リクエストが処理待ちになっています。task-max-threads と DB 接続待ちの有無を確認してください。')
    }
    if ($null -ne $KPI['errTot'] -and $KPI['errTot'] -gt 0) { AddF 1 'HTTP' ('HTTP エラー (error-count) が ' + $KPI['errTot'].ToString('#,0') + ' 件ありました (全リクエスト ' + $KPI['reqTot'].ToString('#,0') + ' 件)。') }
    if ($KPI.ContainsKey('jsF') -and $KPI['jsF'] -gt 0) { AddF 3 'GC' ('Full GC が ' + $KPI['jsF'] + ' 回発生しました (合計 ' + ($KPI['jsFt'] * 1000).ToString('#,0') + ' ms)。ヒープ不足・メタスペース不足・System.gc() を確認してください。') }
    if ($null -ne $KPI['jsOldP'] -and $KPI['jsOldP'].Max -ge 0.85) { AddF 2 'ヒープ' ('Old 領域の使用率が最大 ' + ($KPI['jsOldP'].Max * 100).ToString('0') + '% に達しました。') }
    if ($KPI.ContainsKey('jsEl') -and $KPI['jsEl'] -gt 0) {
        $ratio = $KPI['jsGt'] / $KPI['jsEl']
        if ($ratio -ge 0.05) { AddF 2 'GC' ('期間中の GC 時間率が ' + ($ratio * 100).ToString('0.0') + '% です (目安 5% 未満)。') }
    }
    if ($KPI.ContainsKey('cgThN') -and $KPI['cgThN'] -gt 0) { AddF 2 'CPU' ('CPU スロットリングが ' + $KPI['cgThN'].ToString('#,0') + ' 回 (合計 ' + ($KPI['cgThT'] / 1000).ToString('#,0.0') + ' 秒) 発生しました。タスクの CPU 上限 (' + $KPI['cgQuota'] + ' vCPU) に当たっています。') }
    if ($null -ne $KPI['cgQr'] -and $KPI['cgQr'].Avg -ge 0.8) { AddF 2 'CPU' ('コンテナの CPU 使用率が平均で上限の ' + ($KPI['cgQr'].Avg * 100).ToString('0') + '% です。') }
    if ($KPI.ContainsKey('cgOom') -and $KPI['cgOom'] -gt 0) { AddF 3 'メモリ' ('OOM kill が ' + $KPI['cgOom'] + ' 回発生しました。') }
    if ($null -ne $KPI['cgMemP'] -and $KPI['cgMemP'].Max -ge 0.9) { AddF 2 'メモリ' ('コンテナのメモリ使用率が最大 ' + ($KPI['cgMemP'].Max * 100).ToString('0') + '% に達しました (上限 ' + $KPI['cgLimit'] + ' MB)。') }
    if ($null -ne $KPI['ovhCpu'] -and $KPI['ovhCpu'].Avg -ge 5) { AddF 1 '収集の負荷' ('収集プロセスの CPU が平均 ' + $KPI['ovhCpu'].Avg.ToString('0.0') + '% (1 コア = 100%) です。間隔を長くすると負荷を下げられます。') }
    if ($Find.Count -eq 0) { AddF 0 '全般' '特記事項はありません (接続プールの待ち・取得失敗、Full GC、CPU スロットリング、メモリの逼迫は検出されませんでした)。' }

    # ==== グラフ ====
    if ($null -ne $shHT) {
        Add-ChartSection 'HTTP・リクエスト処理スレッド'
        $c = $script:HTC
        Add-Chart 'スループット（リクエスト/秒・エラー/秒）' '件/秒' '#,##0' @((Sr 'リクエスト/秒' $shHT $c.Rps $null), (Sr 'エラー/秒' $shHT $c.Eps $null)) ''
        Add-Chart 'リクエスト処理スレッド（IO ワーカー）' 'スレッド数' '#,##0' @((Sr '処理中' $shHT $c.Ib $null), (Sr '最大' $shHT $c.Im $null), (Sr 'キュー' $shHT $c.Iq $null), (Sr 'アクティブ要求' $shHT $c.Ar $null)) ''
        if ($null -ne $KPI['ams']) { Add-Chart '平均処理時間（ミリ秒）' 'ミリ秒' '#,##0.0' @((Sr '平均処理時間' $shHT $c.Ams $null)) '' }
        Add-Chart 'HTTP セッション（アクティブ）' 'セッション数' '#,##0' @((Sr 'アクティブセッション' $shHT $c.Sa $null)) ''
    }
    if ($null -ne $shDS) {
        Add-ChartSection 'DB 接続プール'
        $c = $script:DSC
        foreach ($g in $shDS.Groups) {
            $dn = ($g.Key -split "`t")[-1]
            Add-Chart ($dn + '：接続数') '接続数' '#,##0' @((Sr '使用中' $shDS $c.In $g), (Sr 'アクティブ' $shDS $c.Ac $g), (Sr '最大プール' $shDS $c.Mx $g)) ''
            Add-Chart ($dn + '：待ち・取得失敗（区間）') '件' '#,##0' @((Sr '待ち' $shDS $c.Wd $g), (Sr '取得失敗' $shDS $c.Fd $g)) ''
        }
    }
    if ($null -ne $shJS -or $null -ne $shJV) {
        Add-ChartSection 'ヒープ・GC'
        if ($null -ne $shJS) {
            $c = $script:JSC
            Add-Chart 'ヒープ使用量（jstat、MB）' 'MB' '#,##0' @((Sr '使用（合計）' $shJS $c.Hu $null), (Sr 'Old 使用' $shJS $c.Ou $null), (Sr 'Eden 使用' $shJS $c.Eu $null), (Sr '容量（合計）' $shJS $c.Hc $null)) ''
            Add-Chart 'GC 回数（区間）' '回' '#,##0' @((Sr 'Young' $shJS $c.Yn $null), (Sr 'Full' $shJS $c.Fn $null), (Sr 'Concurrent' $shJS $c.Cn $null)) ''
            Add-Chart 'GC 時間（区間、ミリ秒）' 'ミリ秒' '#,##0' @((Sr 'GC 合計' $shJS $c.Gt $null), (Sr 'Young' $shJS $c.Yt $null), (Sr 'Full' $shJS $c.Ft $null)) ''
            Add-Chart 'メタスペース（MB）' 'MB' '#,##0' @((Sr '使用' $shJS $c.Mu $null), (Sr '容量' $shJS $c.Mc $null), (Sr 'CCS 使用' $shJS $c.Cu $null)) ''
        }
        if ($null -ne $shJV) {
            $c = $script:JVC
            $ser = @((Sr 'ヒープ使用' $shJV $c.Hu $null), (Sr 'ヒープコミット' $shJV $c.Hc $null), (Sr 'ヒープ最大' $shJV $c.Hm $null))
            Add-Chart 'ヒープ（platform-mbean、MB）' 'MB' '#,##0' $ser ''
            if ($null -eq $shJS -and $c.Gc.Count -gt 0) {
                $ser = @(); foreach ($nm in $c.Gc.Keys) { $ser += Sr ($nm.Replace('_', ' ')) $shJV $c.Gc[$nm][1] $null }
                Add-Chart 'GC 時間（区間、ミリ秒、platform-mbean）' 'ミリ秒' '#,##0' $ser ''
            }
        }
    }
    if ($null -ne $shTop -or $null -ne $shCg -or $null -ne $shPs) {
        Add-ChartSection 'CPU・メモリ・プロセス'
        $cpuSer = @()
        if ($null -ne $shTop) { $cpuSer += Sr 'JBoss（top）' $shTop $script:TPC.Jc $null }
        if ($null -ne $shCg) { $cpuSer += Sr 'コンテナ全体（cgroup）' $shCg $script:CGRC.Cu $null }
        if ($null -ne $shPs) { $cpuSer += Sr 'JBoss（ps 区間）' $shPs $script:PSC.Cp $null }
        Add-Chart 'CPU 使用率（1 コア = 100%）' '%' '#,##0' $cpuSer ''
        $memSer = @()
        if ($null -ne $shTop) { $memSer += Sr 'JBoss RES（top）' $shTop $script:TPC.Jr $null }
        if ($null -ne $shCg) { $memSer += Sr 'コンテナ使用量（cgroup）'  $shCg $script:CGRC.Mu $null; $memSer += Sr 'コンテナ上限' $shCg $script:CGRC.Ml $null }
        Add-Chart 'メモリ（MB）' 'MB' '#,##0' $memSer ''
        $thrSer = @()
        if ($null -ne $shPs) { $thrSer += Sr 'OS から見たスレッド数（ps nlwp）' $shPs $script:PSC.Nl $null }
        if ($null -ne $shHT) { $thrSer += Sr 'JVM スレッド数（platform-mbean）' $shHT $script:HTC.Jt $null }
        Add-Chart 'JBoss のスレッド数' 'スレッド数' '#,##0' $thrSer ''
        if ($null -ne $shTop) {
            Add-Chart 'ロードアベレージ' '' '#,##0.0' @((Sr '1 分' $shTop $script:TPC.L1 $null), (Sr '5 分' $shTop $script:TPC.L5 $null), (Sr '15 分' $shTop $script:TPC.L15 $null)) ''
            Add-Chart '収集プロセスの CPU（性能試験用、1 コア = 100%）' '%' '#,##0.0' @((Sr 'jboss-cli' $shTop $script:TPC.Kc $null), (Sr 'jstat' $shTop $script:TPC.Sc $null), (Sr 'jcmd' $shTop $script:TPC.Cc $null), (Sr 'スクリプト' $shTop $script:TPC.Pc $null)) ''
        }
        if ($null -ne $shCg) { Add-Chart 'CPU スロットリング（区間の回数）' '回' '#,##0' @((Sr 'スロットリング' $shCg $script:CGRC.Tn $null)) '' }
    }

    # ==== ブックの書き出し ====
    $label = M 'container.log_id'; if ($label -eq '') { $label = Split-Path -Leaf (Split-Path -Parent $dir) }
    $svc = (@((M 'container.service_name'), (M 'container.component_name')) | Where-Object { $_ -ne '' }) -join ' / '
    $range = ''
    if ($null -ne $script:TMin) { $range = (SerialText $script:TMin) + ' ～ ' + (SerialText $script:TMax) + '（' + ([TimeSpan]::FromDays($script:TMax - $script:TMin)).ToString('hh\:mm\:ss') + '）' }
    $tzNote = $(if ($UseRecordedTime) { 'コンテナで記録した時刻 (UTC' + $uo + ')' } else { 'この PC のタイムゾーン (' + [TimeZoneInfo]::Local.DisplayName + ')' })
    $SheetNames = New-Object System.Collections.Generic.List[string]
    $SheetNames.Add('サマリー'); $SheetNames.Add('グラフ')
    foreach ($s in $Sheets) { $SheetNames.Add($s.Name) }
    $SheetNames.Add('収集情報'); $SheetNames.Add('凡例')
    $dir2 = Split-Path -Parent $outPath
    $tmp = Join-Path $dir2 ('.' + [IO.Path]::GetFileName($outPath) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $fs = [IO.File]::Open($tmp, [IO.FileMode]::Create)
    $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $idx = 3
        foreach ($s in $Sheets) {
            Write-Info ("シートを書いています: {0} ({1:N0} 行)" -f $s.Name, $s.Rows.Count)
            Write-TableSheet $zip $s $idx
            $idx++
        }
        Add-TextEntry $zip 'xl/worksheets/sheet1.xml' (Write-SummarySheet 1 $label $svc $range $tzNote $meta $KPI $dsSum $Find $ovh)
        Write-Info 'グラフを作成しています...'
        Add-TextEntry $zip 'xl/worksheets/sheet2.xml' (Write-ChartsSheet $zip 2 'グラフ' ('グラフ  ' + $svc + '  ' + $label) ('時間軸は全グラフ共通 (' + $range + '、' + $tzNote + ')。モノトーン表示のため線の種類で系列を区別しています (実線 → 破線 → 点線 → 長破線 の順)。元データは各シートにあります。'))
        $nChart = $script:NCH
        $mi = $Sheets.Count + 3
        Add-TextEntry $zip ('xl/worksheets/sheet' + $mi + '.xml') (Write-MetaSheet $mi $meta)
        Add-TextEntry $zip ('xl/worksheets/sheet' + ($mi + 1) + '.xml') (Write-LegendSheet ($mi + 1))
        Add-TextEntry $zip 'xl/styles.xml' ($StylesXml.Replace("`r`n", "`n").TrimEnd("`n") + "`n")
        $sorted = @($DefNames.ToArray())
        [Array]::Sort($sorted, [Comparison[object]] { param($p, $q) $c = [string]::CompareOrdinal([string]$p[1], [string]$q[1]); if ($c -ne 0) { return $c }; return ([int]$p[0]).CompareTo([int]$q[0]) })
        $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        $wb += '<workbookPr/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" tabRatio="900" activeTab="0"/></bookViews><sheets>'
        for ($i = 0; $i -lt $SheetNames.Count; $i++) { $wb += '<sheet name="' + (X $SheetNames[$i]) + '" sheetId="' + ($i + 1) + '" r:id="rId' + ($i + 1) + '"/>' }
        $wb += '</sheets>'
        if ($sorted.Count -gt 0) {
            $wb += '<definedNames>'
            foreach ($dn in $sorted) {
                if ($dn[1] -eq '_xlnm._FilterDatabase') { $wb += '<definedName name="' + $dn[1] + '" localSheetId="' + $dn[0] + '" hidden="1">' + (X $dn[2]) + '</definedName>' }
                else { $wb += '<definedName name="' + $dn[1] + '" localSheetId="' + $dn[0] + '">' + (X $dn[2]) + '</definedName>' }
            }
            $wb += '</definedNames>'
        }
        $wb += '<calcPr calcId="191029"/></workbook>'
        Add-TextEntry $zip 'xl/workbook.xml' $wb
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        for ($i = 0; $i -lt $SheetNames.Count; $i++) { $rels += '<Relationship Id="rId' + ($i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + ($i + 1) + '.xml"/>' }
        $rels += '<Relationship Id="rId' + ($SheetNames.Count + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        Add-TextEntry $zip 'xl/_rels/workbook.xml.rels' $rels
        if ($nChart -gt 0) {
            Add-TextEntry $zip 'xl/worksheets/_rels/sheet2.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
        }
        $cgct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        $cgct += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        $cgct += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        for ($i = 1; $i -le $SheetNames.Count; $i++) { $cgct += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' }
        $cgct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        if ($nChart -gt 0) {
            $cgct += '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            for ($i = 1; $i -le $nChart; $i++) { $cgct += '<Override PartName="/xl/charts/chart' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' }
        }
        $cgct += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        $cgct += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
        Add-TextEntry $zip '[Content_Types].xml' $cgct
        Add-TextEntry $zip '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
        $w3c = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', $Inv)
        Add-TextEntry $zip 'docProps/core.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + (X $Title) + '</dc:title><dc:subject>' + (X $label) + '</dc:subject><dc:creator>New-PerfMetricsReport.ps1</dc:creator><cp:lastModifiedBy>New-PerfMetricsReport.ps1</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">' + $w3c + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + $w3c + '</dcterms:modified></cp:coreProperties>')
        Add-TextEntry $zip 'docProps/app.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>New-PerfMetricsReport.ps1 ' + $Version + '</Application></Properties>')
    } catch {
        $zip.Dispose(); $fs.Dispose()
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw
    }
    $zip.Dispose(); $fs.Dispose()
    Move-Item -LiteralPath $tmp -Destination $outPath -Force
    Write-Info ("作成しました: {0} （シート {1}、グラフ {2}、{3:N1} 秒）" -f $outPath, $SheetNames.Count, $nChart, $sw.Elapsed.TotalSeconds)
    $script:Created.Add($outPath)
}

function Add-TextEntry([System.IO.Compression.ZipArchive]$zip, [string]$name, [string]$content) {
    $e = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal)
    $w = New-Object IO.StreamWriter($e.Open(), $Utf8); $w.Write($content); $w.Dispose()
}

#------------------------------------------------------------------------------
# サマリーシート
#------------------------------------------------------------------------------
function Fmt($v, [string]$f) { if ($null -eq $v) { return '－' }; return ([double]$v).ToString($f, $Inv) }
function Write-SummarySheet([int]$sheetIdx, [string]$label, [string]$svc, [string]$range, [string]$tzNote, $meta, $KPI, $dsSum, $Find, $ovh) {
    function M2([string]$k) { if ($meta.Contains($k)) { return [string]$meta[$k] }; return '' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((Ws-Begin 'FF000000' $true 0 0 $true 90))
    [void]$sb.Append('<cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="2" width="34" customWidth="1"/><col min="3" max="7" width="13" customWidth="1"/><col min="8" max="8" width="40" customWidth="1"/></cols><sheetData>')
    $merges = New-Object System.Collections.Generic.List[string]
    $r = 1
    [void]$sb.Append((Row $r (StrC 'A1' $Title $S_TITLE) 28)); $r++
    [void]$sb.Append((Row $r (StrC 'A2' ($svc + '　　タスク ' + $label + '　　' + $range) $S_HEAD) 20)); $r++
    [void]$sb.Append((Row $r (StrC 'A3' ('時刻は ' + $tzNote + ' で表示しています。作成: ' + (Get-Date -Format 'yyyy/MM/dd HH:mm') + '　New-PerfMetricsReport.ps1 ' + $Version) $S_NOTE) 0)); $r += 2

    # --- 所見 ---
    [void]$sb.Append((Row $r (StrC ('A' + $r) '■ 所見' $S_SEC) 20)); $r++
    [void]$sb.Append((Row $r ((StrC ('A' + $r) '重要度' $S_HDR) + (StrC ('B' + $r) '分類' $S_HDR) + (StrC ('C' + $r) '内容' $S_HDR) + (EmptyC ('D' + $r) $S_HDR) + (EmptyC ('E' + $r) $S_HDR) + (EmptyC ('F' + $r) $S_HDR) + (EmptyC ('G' + $r) $S_HDR) + (EmptyC ('H' + $r) $S_HDR)) 20))
    $merges.Add('C' + $r + ':H' + $r); $r++
    foreach ($f in ($Find | Sort-Object -Property @{ Expression = { $_[0] }; Descending = $true })) {
        $lv = [int]$f[0]
        $lab = switch ($lv) { 3 { '■ 要対応' } 2 { '▲ 注意' } 1 { '● 参考' } default { '○ 良好' } }
        $st = switch ($lv) { 3 { $S_SEVH } 2 { $S_SEVM } default { $S_SEVL } }
        $cells = (StrC ('A' + $r) $lab $st) + (StrC ('B' + $r) $f[1] $S_TEXTB) + (StrC ('C' + $r) $f[2] $S_WRAP)
        foreach ($cc in 'D', 'E', 'F', 'G', 'H') { $cells += EmptyC ($cc + $r) $S_WRAP }
        [void]$sb.Append((Row $r $cells (WrapHt $f[2] 105)))
        $merges.Add('C' + $r + ':H' + $r); $r++
    }
    $r++

    # --- 主要指標 ---
    [void]$sb.Append((Row $r (StrC ('A' + $r) '■ 主要指標' $S_SEC) 20)); $r++
    $hd = @('分類', '指標', '最大', '95% 値', '平均', '最小', '単位', '備考')
    $cells = ''; for ($i = 0; $i -lt 8; $i++) { $cells += StrC ($COLN[$i + 1] + $r) $hd[$i] $S_HDR }
    [void]$sb.Append((Row $r $cells 20)); $r++
    $kp = New-Object System.Collections.Generic.List[object]
    function KRow([string]$cat, [string]$name, $st, [string]$unit, [string]$note, [double]$scale, [string]$fmt) { $kp.Add(@($cat, $name, $st, $unit, $note, $scale, $fmt)) }
    function KOne([string]$cat, [string]$name, $v, [string]$unit, [string]$note, [string]$fmt) { $kp.Add(@($cat, $name, $v, $unit, $note, 0, $fmt)) }
    if ($null -ne $KPI['rps']) { KRow 'HTTP' 'スループット' $KPI['rps'] '件/秒' ('期間中のリクエスト ' + (Fmt $KPI['reqTot'] '#,0') + ' 件・エラー ' + (Fmt $KPI['errTot'] '#,0') + ' 件') 1 'dec1' }
    if ($null -ne $KPI['ams']) { KRow 'HTTP' '平均処理時間' $KPI['ams'] 'ミリ秒' '取得間隔ごとの平均 (Undertow processing-time)' 1 'dec1' }
    if ($null -ne $KPI['iob']) { KRow 'リクエスト処理' 'IO ワーカー 処理中スレッド' $KPI['iob'] 'スレッド' ('最大 ' + (Fmt $KPI['ioMax'] '#,0') + ' スレッド') 1 'int' }
    if ($null -ne $KPI['ioq']) { KRow 'リクエスト処理' 'IO ワーカー キュー' $KPI['ioq'] '件' '処理待ちのタスク' 1 'int' }
    if ($null -ne $KPI['ses']) { KRow 'HTTP' 'アクティブセッション' $KPI['ses'] '件' '' 1 'int' }
    foreach ($d in $dsSum) {
        if ($null -ne $d.In) { KRow 'DB 接続プール' ($d.Name + ' 使用中') $d.In '接続' ('最大プール ' + $d.Max + '・待ち ' + (Fmt $d.Wait '#,0') + ' 件・取得失敗 ' + (Fmt $d.Fail '#,0') + ' 件') 1 'int' }
    }
    if ($null -ne $KPI['jsHeap']) { KRow 'ヒープ（jstat）' 'ヒープ使用量' $KPI['jsHeap'] 'MB' ('容量 最大 ' + (Fmt $KPI['jsHcMax'] '#,0') + ' MB' + $(if ($meta.Contains('jvm_flag.MaxHeapSize')) { '・最大ヒープ ' + (Fmt ([double]$meta['jvm_flag.MaxHeapSize'] / 1048576) '#,0') + ' MB' } else { '' })) 1 'dec1' }
    if ($null -ne $KPI['jsOld']) { KRow 'ヒープ（jstat）' 'Old 使用量' $KPI['jsOld'] 'MB' ('Old 使用率 最大 ' + (Fmt ($KPI['jsOldP'].Max * 100) '0.0') + '%') 1 'dec1' }
    if ($null -ne $KPI['jsMeta']) { KRow 'ヒープ（jstat）' 'メタスペース使用量' $KPI['jsMeta'] 'MB' '' 1 'dec1' }
    if ($KPI.ContainsKey('jsY')) {
        $gr = $(if ($KPI['jsEl'] -gt 0) { $KPI['jsGt'] / $KPI['jsEl'] * 100 } else { $null })
        KOne 'GC（jstat）' 'Young GC 回数 / 合計時間' $KPI['jsY'] '回' ('合計 ' + (Fmt ($KPI['jsYt'] * 1000) '#,0') + ' ms・平均停止 ' + $(if ($null -ne $KPI['jsYa']) { (Fmt $KPI['jsYa'].Avg '0.0') + ' ms (最大 ' + (Fmt $KPI['jsYa'].Max '0.0') + ' ms)' } else { '－' })) 'int'
        KOne 'GC（jstat）' 'Full GC 回数 / 合計時間' $KPI['jsF'] '回' ('合計 ' + (Fmt ($KPI['jsFt'] * 1000) '#,0') + ' ms') 'int'
        KOne 'GC（jstat）' 'Concurrent GC 回数' $KPI['jsC'] '回' 'G1 の並行マーキング等' 'int'
        KOne 'GC（jstat）' 'GC 時間率（期間全体）' $gr '%' ('GC 時間 ' + (Fmt ($KPI['jsGt']) '#,0.00') + ' 秒 ÷ 経過 ' + (Fmt $KPI['jsEl'] '#,0') + ' 秒') 'dec2'
    } elseif ($null -ne $KPI['cliGc']) {
        foreach ($nm in $KPI['cliGc'].Keys) { KOne 'GC（CLI）' ($nm.Replace('_', ' ') + ' 回数') $KPI['cliGc'][$nm][0] '回' ('合計 ' + (Fmt $KPI['cliGc'][$nm][1] '#,0') + ' ms') 'int' }
    }
    if ($null -ne $KPI['cliHeap'] -and $null -eq $KPI['jsHeap']) { KRow 'ヒープ（CLI）' 'ヒープ使用量' $KPI['cliHeap'] 'MB' ('最大ヒープ ' + (Fmt $KPI['cliHeapMax'] '#,0') + ' MB') 1 'dec1' }
    if ($null -ne $KPI['jthr']) { KRow 'JVM' 'スレッド数' $KPI['jthr'] 'スレッド' 'platform-mbean' 1 'int' }
    if ($null -ne $KPI['jbCpu']) { KRow 'CPU（top）' 'JBoss の %CPU' $KPI['jbCpu'] '%' '1 コア = 100%' 1 'dec1' }
    if ($null -ne $KPI['cgCpu']) { KRow 'CPU（cgroup）' 'コンテナの CPU 使用率' $KPI['cgCpu'] '%' ('1 vCPU = 100%・上限 ' + (Fmt $KPI['cgQuota'] '0.00') + ' vCPU・スロットリング ' + (Fmt $KPI['cgThN'] '#,0') + ' 回') 1 'dec1' }
    if ($null -ne $KPI['sysCpu']) { KRow 'CPU（top）' '全体の CPU 使用率' $KPI['sysCpu'] '%' '全 CPU に対する割合 (100 - idle)' 1 'dec1' }
    if ($null -ne $KPI['load1']) { KRow 'CPU（top）' 'ロードアベレージ（1 分）' $KPI['load1'] '' '' 1 'dec2' }
    if ($null -ne $KPI['jbRes']) { KRow 'メモリ（top）' 'JBoss の RES' $KPI['jbRes'] 'MB' '' 1 'dec1' }
    if ($null -ne $KPI['cgMem']) { KRow 'メモリ（cgroup）' 'コンテナのメモリ使用量' $KPI['cgMem'] 'MB' ('上限 ' + (Fmt $KPI['cgLimit'] '#,0') + ' MB・OOM kill ' + (Fmt $KPI['cgOom'] '#,0') + ' 回') 1 'dec1' }
    if ($null -ne $KPI['psNl']) { KRow 'プロセス（ps）' 'JBoss のスレッド数' $KPI['psNl'] 'スレッド' '' 1 'int' }
    if ($null -ne $KPI['clims']) { KRow '収集' 'jboss-cli 1 回の応答時間' $KPI['clims'] 'ミリ秒' '① の全項目を取得する時間' 1 'int' }
    foreach ($k in $kp) {
        $cells = (StrC ('A' + $r) $k[0] $S_KEY) + (StrC ('B' + $r) $k[1] $S_TEXT)
        $st = $FMTS[$k[6]]
        if ($k[5] -eq 0) {
            $cells += NumC ('C' + $r) $k[2] $st; foreach ($cc in 'D', 'E', 'F') { $cells += StrC ($cc + $r) '' $S_MUTED }
        } elseif ($null -eq $k[2]) {
            foreach ($cc in 'C', 'D', 'E', 'F') { $cells += StrC ($cc + $r) '－' $S_MUTED }
        } else {
            $cells += NumC ('C' + $r) ($PerfRpt::Rnd($k[2].Max, 2)) $(if ($k[6] -eq 'int') { $S_INTB } else { $S_DEC1B })
            $cells += NumC ('D' + $r) ($PerfRpt::Rnd($k[2].P95, 2)) $st; $cells += NumC ('E' + $r) ($PerfRpt::Rnd($k[2].Avg, 2)) $st; $cells += NumC ('F' + $r) ($PerfRpt::Rnd($k[2].Min, 2)) $st
        }
        $cells += (StrC ('G' + $r) $k[3] $S_TEXTC) + (StrC ('H' + $r) $k[4] $S_WRAP)
        [void]$sb.Append((Row $r $cells (WrapHt $k[4] 40)))
        $r++
    }
    $r++

    # --- データソース別 ---
    if ($dsSum.Count -gt 0) {
        [void]$sb.Append((Row $r (StrC ('A' + $r) '■ DB 接続プール（データソース別）' $S_SEC) 20)); $r++
        $hd = @('データソース', '使用中 最大 / 最大プール', '使用率 最大', '使用率 平均', '待ち（件）', '取得失敗（件）', '最大待ち（ms）', '平均取得（ms、最終）')
        $cells = ''; for ($i = 0; $i -lt 8; $i++) { $cells += StrC ($COLN[$i + 1] + $r) $hd[$i] $S_HDR }
        [void]$sb.Append((Row $r $cells 30)); $r++
        foreach ($d in $dsSum) {
            $cells = (StrC ('A' + $r) $d.Name $S_KEY)
            $cells += StrC ('B' + $r) ((Fmt $(if ($null -ne $d.In) { $d.In.Max } else { $null }) '#,0') + ' / ' + (Fmt $d.Max '#,0')) $S_TEXTC
            $cells += NumC ('C' + $r) $(if ($null -ne $d.Us) { $PerfRpt::Rnd($d.Us.Max, 4)} else { $null }) $S_PCTB
            $cells += NumC ('D' + $r) $(if ($null -ne $d.Us) { $PerfRpt::Rnd($d.Us.Avg, 4)} else { $null }) $S_PCT
            $cells += NumC ('E' + $r) $d.Wait $(if ($d.Wait -gt 0) { $S_INTB } else { $S_INT })
            $cells += NumC ('F' + $r) $d.Fail $(if ($d.Fail -gt 0) { $S_INTB } else { $S_INT })
            $cells += NumC ('G' + $r) $d.MaxWait $S_INT
            $cells += NumC ('H' + $r) $d.AvgGet $S_INT
            [void]$sb.Append((Row $r $cells 0)); $r++
        }
        $r++
    }

    # --- HTTP リスナー別 ---
    if ($null -ne $KPI['lsn'] -and $KPI['lsn'].Count -gt 0) {
        [void]$sb.Append((Row $r (StrC ('A' + $r) '■ HTTP リスナー別のリクエスト数（期間中）' $S_SEC) 20)); $r++
        [void]$sb.Append((Row $r ((StrC ('A' + $r) 'リスナー' $S_HDR) + (StrC ('B' + $r) 'サーバー / 種別 / 名前' $S_HDR) + (StrC ('C' + $r) 'リクエスト' $S_HDR)) 20)); $r++
        foreach ($k in $KPI['lsn'].Keys) {
            $p = $k.Split('/')
            [void]$sb.Append((Row $r ((StrC ('A' + $r) $p[2] $S_KEY) + (StrC ('B' + $r) $k $S_TEXT) + (NumC ('C' + $r) $KPI['lsn'][$k] $S_INT)) 0)); $r++
        }
        $r++
    }

    # --- 収集プロセスの負荷 ---
    if ($ovh.Count -gt 0) {
        [void]$sb.Append((Row $r (StrC ('A' + $r) '■ 収集プロセスの負荷（性能試験用。結果を読むときの補正の目安）' $S_SEC) 20)); $r++
        $hd = @('プロセス', '内容', '平均 %CPU', '最大 %CPU', '最大 RES MB', '', '', '備考')
        $cells = ''; for ($i = 0; $i -lt 8; $i++) { $cells += StrC ($COLN[$i + 1] + $r) $hd[$i] $S_HDR }
        [void]$sb.Append((Row $r $cells 20)); $r++
        $desc = @{ 'jboss-cli' = '① jboss-cli 常駐セッション (JVM)'; 'jstat' = '② jstat 常駐 (JVM)'; 'jcmd' = '② jcmd (間隔ごとに起動)'; 'perf-metrics' = '収集スクリプト (sh / awk)'; 'top' = '③ top'; 'ps' = '③ ps' }
        foreach ($rl in @('jboss-cli', 'jstat', 'jcmd', 'perf-metrics', 'top', 'ps')) {
            if (-not $ovh.ContainsKey($rl)) { continue }
            $sc = Stat $ovh[$rl][0]; $sm = Stat $ovh[$rl][1]
            $cells = (StrC ('A' + $r) $rl $S_KEY) + (StrC ('B' + $r) $desc[$rl] $S_TEXT)
            $cells += NumC ('C' + $r) $(if ($null -ne $sc) { $PerfRpt::Rnd($sc.Avg, 2)} else { $null }) $S_DEC2
            $cells += NumC ('D' + $r) $(if ($null -ne $sc) { $PerfRpt::Rnd($sc.Max, 1)} else { $null }) $S_DEC1
            $cells += NumC ('E' + $r) $(if ($null -ne $sm) { $PerfRpt::Rnd($sm.Max, 1)} else { $null }) $S_DEC1
            $cells += (EmptyC ('F' + $r) $S_TEXT) + (EmptyC ('G' + $r) $S_TEXT) + (StrC ('H' + $r) '1 コア = 100%' $S_TEXT)
            [void]$sb.Append((Row $r $cells 0)); $r++
        }
        $r++
    }

    # --- 収集情報 (抜粋) ---
    [void]$sb.Append((Row $r (StrC ('A' + $r) '■ 収集情報（抜粋。全項目は「収集情報」シート）' $S_SEC) 20)); $r++
    $items = @(
        @('サービス', $svc), @('タスク (LOG_ID)', $label), @('ホスト名', (M2 'container.hostname')), @('ECS タスク', (M2 'ecs.TaskARN')),
        @('タスクの上限', (M2 'ecs.TaskLimits')), @('JBoss EAP', ((M2 'jboss.product-name') + ' ' + (M2 'jboss.product-version'))), @('JVM', ((M2 'jboss.vm-name') + ' ' + (M2 'jboss.vm-version'))),
        @('GC', $(if ((M2 'jvm_flag.UseG1GC') -eq 'true') { 'G1' } elseif ((M2 'jvm_flag.UseParallelGC') -eq 'true') { 'Parallel' } elseif ((M2 'jvm_flag.UseSerialGC') -eq 'true') { 'Serial' } elseif ((M2 'jvm_flag.UseZGC') -eq 'true') { 'Z' } else { '' })),
        @('最大ヒープ', $(if ((M2 'jvm_flag.MaxHeapSize') -ne '') { ([double](M2 'jvm_flag.MaxHeapSize') / 1048576).ToString('#,0', $Inv) + ' MB' } else { '' })),
        @('JVM が使える CPU', (M2 'jboss.available-processors')), @('収集間隔 (秒)', ('jboss-cli ' + (M2 'perf.cli_interval') + ' / jstat ' + (M2 'perf.jstat_interval') + ' / jcmd ' + (M2 'perf.jcmd_interval') + ' / top ' + (M2 'perf.top_interval') + ' / ps ' + (M2 'perf.ps_interval'))),
        @('出力先', (M2 'perf.output_dir')), @('収集開始', (M2 'perf.start_time') + ' (コンテナの時刻、UTC' + (M2 'perf.utc_offset') + ')')
    )
    foreach ($it in $items) {
        if ([string]::IsNullOrWhiteSpace($it[1]) -or $it[1].Trim() -eq '/') { continue }
        $cells = (StrC ('A' + $r) $it[0] $S_KEY) + (StrC ('B' + $r) $it[1] $S_TEXT)
        foreach ($cc in 'C', 'D', 'E', 'F', 'G', 'H') { $cells += EmptyC ($cc + $r) $S_TEXT }
        [void]$sb.Append((Row $r $cells 0)); $merges.Add('B' + $r + ':H' + $r); $r++
    }
    [void]$sb.Append('</sheetData>')
    $extra = '<mergeCells count="' + $merges.Count + '">'; foreach ($m in $merges) { $extra += '<mergeCell ref="' + $m + '"/>' }; $extra += '</mergeCells>'
    [void]$sb.Append((Ws-End 'portrait' $true $false $Title $extra))
    Add-DefName $sheetIdx '_xlnm.Print_Area' ("'サマリー'!`$A`$1:`$H`$" + $r)
    return $sb.ToString()
}

#------------------------------------------------------------------------------
# 収集情報シート
#------------------------------------------------------------------------------
function Write-MetaSheet([int]$sheetIdx, $meta) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((Ws-Begin 'FF7F7F7F' $true 0 4 $false 90))
    [void]$sb.Append('<cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="2" width="34" customWidth="1"/><col min="3" max="3" width="90" customWidth="1"/><col min="4" max="4" width="28" customWidth="1"/></cols><sheetData>')
    [void]$sb.Append((Row 1 (StrC 'A1' '収集情報' $S_TITLE) 24))
    [void]$sb.Append((Row 2 (StrC 'A2' 'perf_meta.csv の内容です (収集開始時の環境・条件と、jboss-cli / jcmd が取得したサーバー・JVM の情報)。' $S_NOTE) 0))
    [void]$sb.Append((Row 4 ((StrC 'A4' '分類' $S_HDR) + (StrC 'B4' '項目' $S_HDR) + (StrC 'C4' '値' $S_HDR) + (StrC 'D4' 'キー' $S_HDR)) 20))
    $r = 5
    foreach ($k in $meta.Keys) {
        $cat = $k.Split('.')[0]
        $cl = $MetaCat[$cat]; if ($null -eq $cl) { $cl = $cat }
        $lb = $MetaLabel[$k]; if ($null -eq $lb) { $lb = $k.Substring($cat.Length + 1) }
        $v = [string]$meta[$k]
        if ($k -eq 'jboss.jvm-start-time' -and $NumRe.IsMatch($v)) { $v = $v + '  (' + (SerialText ($PerfRpt::XT(([double]$v / 1000)))) + ')' }
        [void]$sb.Append((Row $r ((StrC ('A' + $r) $cl $S_KEY) + (StrC ('B' + $r) $lb $S_TEXT) + (StrC ('C' + $r) $v $S_WRAP) + (StrC ('D' + $r) $k $S_MUTED)) (WrapHt $v 88)))
        $r++
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append((Ws-End 'landscape' $true $false '収集情報' ''))
    return $sb.ToString()
}

#------------------------------------------------------------------------------
# 凡例シート
#------------------------------------------------------------------------------
$LegendText = @'
H|このブックの見方
P|・時刻は既定で、レポートを作成した PC のタイムゾーンで表示しています (-UseRecordedTime を付けるとコンテナで記録した時刻)。JMeter の結果 (JTL) と突き合わせるときは同じタイムゾーンで比べてください。
P|・各データシートは 4 行目が分類、5 行目が見出し、6 行目が元の CSV の列名です。6 行目にオートフィルターがあります。グラフは「グラフ」シートにまとめ、時間軸を全グラフで揃えています。
P|・「区間」は前回の取得からの増分です。累計カウンタが減った場合 (サーバー再起動・統計のクリア) は、その回の値をそのまま増分として扱います。
P|・95% 値は、期間中の取得値を小さい順に並べたときに 95% の位置にある値です。
B|
H|収集の方法 (perf-metrics.sh)
TH|区分|取得方法|主な出力
T|① jboss-cli|jboss-cli.sh を 1 つ常駐させ (FIFO で命令を送る)、管理モデルを間隔ごとに読む。読み取りだけで設定は変えない (プール統計だけは実行時属性で有効化)|cli_datasource / cli_undertow / cli_session / cli_io_worker / cli_threadpool / cli_transactions / cli_jvm / cli_gc / cli_mempool
T|② jstat|jstat -gc -t <pid> <間隔> を 1 つ常駐させ、各行に時刻を付ける (runtime イメージに jstat が無い場合は java -m jdk.jcmd で起動)|jvm_jstat_gc
T|② jcmd|jcmd <pid> GC.heap_info を間隔ごとに実行 (1 回ごとに JVM を起動するため既定 60 秒)。起動時に VM.version / VM.flags|jvm_jcmd_heap / jvm_jcmd.log / jvm_info.txt
T|③ top|top -b -c (全体) と top -b -H -p <JBoss> (スレッド別) を常駐させる。区間の %CPU が取れる|os_top_summary / os_top_process / os_top_threads / os_top.log
T|③ ps / cgroup|ps -e -o … と /sys/fs/cgroup を間隔ごとに読む|os_ps / os_cgroup
B|
H|主な項目
TH|項目|元の値|説明
T|使用中 / アクティブ / アイドル|InUseCount / ActiveCount / IdleCount|アプリが使用中の接続数 / プールに作成済みの物理接続数 / 未使用の接続数
T|使用率|InUseCount ÷ 最大プール|最大プール (max-pool-size) が未設定のときは既定値 20 で計算
T|待ち / 取得失敗|WaitCount / BlockingFailureCount|空き接続を待った取得要求 / 取得タイムアウト (blocking-timeout-wait-millis 超過。アプリでは IJ000453 等)
T|平均取得 / 最大待ち|AverageGetTime / MaxWaitTime|接続の取得時間 (待ちを含む) / 空きを待った最大時間 (ミリ秒。統計の有効化以降の累積値)
T|リクエスト/秒|request-count の増分 ÷ 秒|Undertow の全リスナー (http / https / ajp) の合計
T|平均処理時間|processing-time の増分 ÷ request-count の増分|リスナーの record-request-start-time=true のときだけ値が入る (false なら空欄)
T|IO ワーカー 処理中 / 最大 / キュー|busy-task-thread-count / max-pool-size / queue-size|リクエストを処理するスレッド (EAP 8 の Undertow は IO ワーカーのタスクスレッドで処理する)。最大に達するとキューで待つ
T|ヒープ (jstat)|EU + S0U + S1U + OU|Eden・Survivor・Old の使用量の合計。容量は EC + S0C + S1C + OC
T|GC 時間率|GCT の増分 ÷ 経過秒|アプリが GC で止まっていた時間の割合 (G1 の Concurrent は並行処理の停止部分)。目安 5% 未満
T|役割 (top / ps)|コマンドラインから判定|jboss = JBoss EAP、jboss-cli / jstat / jcmd = 性能試験用の収集 (目印 -Dperf.metrics.collector 付き)、perf-metrics / top / ps = 収集スクリプト、cli-other / jdk-tool = 手動で実行した jboss-cli / jstat / jcmd
T|%CPU (top / ps)|1 コア = 100%|複数のコアを使うと 100% を超える。コンテナ全体の CPU は cgroup の使用率 (1 vCPU = 100%) を参照
T|スロットリング|cpu.stat nr_throttled / throttled_usec|タスクの CPU 上限に当たり、実行を待たされた回数と時間
T|RES / RSS|top RES / ps rss|プロセスが使っている物理メモリ。JVM はヒープの外 (メタスペース・スレッドスタック・コード) も使う
B|
H|所見の判定
TH|重要度|内容|判定基準
T|■ 要対応|接続の取得失敗・Full GC・OOM kill|期間中の増分 > 0
T|▲ 注意|接続待ち・プールの上限到達・IO ワーカーの上限到達・Old 使用率 ≧ 85%・GC 時間率 ≧ 5%・CPU スロットリング・CPU 上限比の平均 ≧ 80%・メモリ使用率 ≧ 90%|期間中に 1 回以上
T|● 参考|接続プール使用率 ≧ 80% (待ちなし)・HTTP エラー・プール統計が無効・収集プロセスの CPU が平均 5% 以上|－
B|
H|注意事項
P|・収集プロセス (jboss-cli・jstat の JVM など) も同じコンテナの CPU・メモリを使います。サマリーの「収集プロセスの負荷」で量を確認できます。本番と比べる試験では、収集の有無で条件を揃えてください。
P|・-Dwildfly.statistics-enabled=true (統計の有効化) と、プール統計の実行時有効化によるわずかな負荷があります。
P|・jstat / jcmd は /tmp が書き込めないと JVM に接続できません。その場合も「JVM・GC（CLI）」シートでヒープ・GC を確認できます。
P|・グラフはモノトーンのため、線の種類 (実線・破線・点線・長破線) で系列を区別しています。
'@
function Write-LegendSheet([int]$sheetIdx) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((Ws-Begin 'FF7F7F7F' $true 0 0 $false 90))
    [void]$sb.Append('<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="44" customWidth="1"/><col min="3" max="3" width="70" customWidth="1"/></cols><sheetData>')
    [void]$sb.Append((Row 1 (StrC 'A1' '凡例・用語の説明' $S_TITLE) 24))
    $merges = New-Object System.Collections.Generic.List[string]
    $r = 3
    foreach ($ln in ($LegendText -split "`r?`n")) {
        if ($ln -eq '') { continue }
        $a = $ln.Split('|')
        switch ($a[0]) {
            'H' { [void]$sb.Append((Row $r (StrC ('A' + $r) ('■ ' + $a[1]) $S_SEC) 20)) }
            'TH' { [void]$sb.Append((Row $r ((StrC ('A' + $r) $a[1] $S_HDR) + (StrC ('B' + $r) $a[2] $S_HDR) + (StrC ('C' + $r) $a[3] $S_HDR)) 20)) }
            'T' { $h = [Math]::Max((WrapHt $a[2] 42), (WrapHt $a[3] 68)); [void]$sb.Append((Row $r ((StrC ('A' + $r) $a[1] $S_KEY) + (StrC ('B' + $r) $a[2] $S_WRAP) + (StrC ('C' + $r) $a[3] $S_WRAP)) $h)) }
            'P' { [void]$sb.Append((Row $r (StrC ('A' + $r) $a[1] $S_NOTEW) (WrapHt $a[1] 135))); $merges.Add('A' + $r + ':C' + $r) }
            'B' { }
        }
        $r++
    }
    [void]$sb.Append('</sheetData>')
    $extra = ''; if ($merges.Count -gt 0) { $extra = '<mergeCells count="' + $merges.Count + '">'; foreach ($m in $merges) { $extra += '<mergeCell ref="' + $m + '"/>' }; $extra += '</mergeCells>' }
    [void]$sb.Append((Ws-End 'landscape' $true $false '凡例' $extra))
    return $sb.ToString()
}

#------------------------------------------------------------------------------
# メイン
#------------------------------------------------------------------------------
$script:Created = New-Object System.Collections.Generic.List[string]
try {
    $dirs = Resolve-PerfDirs $InputPath
    # mid/ ごとコピーすると current (最後に起動したタスクへのリンク) 経由で同じ perf/ が
    # 2 回見つかるため、LOG_ID が同じものは 1 回だけ処理する (current を通らないパスを優先)
    $dirs = @($dirs | Sort-Object -Property @{ Expression = { $_ -match '[\\/]current[\\/]' } }, @{ Expression = { $_ } })
    $todo = New-Object System.Collections.Generic.List[object]; $seen = @{}
    foreach ($d in $dirs) {
        $lid = ''
        $mp = Join-Path $d 'perf_meta.csv'
        if (Test-Path -LiteralPath $mp) { foreach ($l in [IO.File]::ReadAllLines($mp, $Utf8)) { if ($l.StartsWith('container,log_id,')) { $lid = (Split-Csv $l)[2] } } }
        if ($lid -eq '') { $lid = Split-Path -Leaf (Split-Path -Parent $d) }
        $lid = [regex]::Replace($lid, '[^A-Za-z0-9._-]', '_')
        if ($seen.ContainsKey($lid)) { Write-Info "同じタスク ($lid) の perf/ が重複しているため省略します: $d"; continue }
        $seen[$lid] = 1; $todo.Add(@($d, $lid))
    }
    $multi = ($todo.Count -gt 1)
    if ($multi) { Write-Info ("perf/ ディレクトリが {0} 個見つかりました。1 つずつブックを作成します" -f $todo.Count) }
    foreach ($item in $todo) {
        $d = $item[0]; $lid = $item[1]
        $name = 'perf_report_' + $lid + '.xlsx'
        if ($OutputPath -eq '') { $out = Join-Path (Get-Location).ProviderPath $name }
        elseif ($multi -or (Test-Path -LiteralPath $OutputPath -PathType Container)) { $out = Join-Path $OutputPath $name }
        else { $out = $OutputPath; if (-not $out.EndsWith('.xlsx', [StringComparison]::OrdinalIgnoreCase)) { $out += '.xlsx' } }
        if (-not [IO.Path]::IsPathRooted($out)) { $out = Join-Path (Get-Location).ProviderPath $out }
        $od = Split-Path -Parent $out
        if (-not (Test-Path -LiteralPath $od -PathType Container)) { New-Item -ItemType Directory -Path $od -Force | Out-Null }
        Build-Report $d $out
    }
} finally {
    foreach ($t in $TempDirs) { Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction SilentlyContinue }
}
if ($Open) { foreach ($c in $script:Created) { Invoke-Item -LiteralPath $c } }
PERFRPT-MAIN#>
