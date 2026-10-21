<#
.SYNOPSIS
    perf-metrics.sh (性能試験用メトリクス収集) が出力した perf/ ディレクトリを、
    モノトーンの Excel レポート (.xlsx) に変換します (Windows 11 / PowerShell 5.1・7)。

.DESCRIPTION
    ① jboss-cli の DB 接続プール (cli_datasource.csv) と ② jstat -gc (jvm_jstat_gc.csv) を
    読み込み、次のシートを持つブックを作成します。

      サマリー             所見・主要指標 (最大 / 95% 値 / 平均 / 最小)・データソース別・収集情報の抜粋
      グラフ               DB 接続プール (データソース別) とヒープ・GC の推移 (共通の時間軸)
      DB接続プール         データソース別の時系列 (使用率・区間の待ち / 取得失敗など)
      ヒープ・GC（jstat）  各領域の使用量・容量、区間の GC 回数・時間・GC 時間率
      収集情報 / 凡例

    ・フォント Meiryo UI、黒・グレーのモノトーン。グラフは線の種類で系列を区別します。
    ・Excel のインストールは不要です (Office Open XML を直接生成します)。
    ・時刻は既定で「この PC のタイムゾーン」に換算して表示します (JMeter の結果と突き合わせやすくするため)。
      コンテナで記録した時刻のまま表示する場合は -UseRecordedTime を指定します。

    ※ 保存時の文字コードは UTF-8 (BOM 付き・なしのどちらでも可) または ANSI (Shift_JIS) で動作します。

.PARAMETER InputPath
    次のいずれか
      ・perf/ ディレクトリ (perf_meta.csv・cli_datasource.csv・jvm_jstat_gc.csv を含む)
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
$Version = '2.0.0'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Inv = [Globalization.CultureInfo]::InvariantCulture
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$XmlNeedRe = [regex]'[&<>"\x01-\x08\x0B\x0C\x0E-\x1F]'
$CtlRe = [regex]'[\x01-\x08\x0B\x0C\x0E-\x1F]'

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
function Split-Csv([string]$line) { return , [PerfRpt]::SplitCsv($line) }
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
function NumC([string]$ref, $v, [int]$st) { $t = [PerfRpt]::NStr($v); if ($t -eq '') { return '<c r="' + $ref + '" s="' + $st + '"/>' }; return '<c r="' + $ref + '" s="' + $st + '"><v>' + $t + '</v></c>' }
function EmptyC([string]$ref, [int]$st) { return '<c r="' + $ref + '" s="' + $st + '"/>' }
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
# 行ごとに何度も呼ぶ処理 (CSV の分解・数値化・増分・丸め・時刻・行の XML 出力)。
# PowerShell の関数は 1 回の呼び出しが重いため、クラスの静的メソッドにまとめる。
#------------------------------------------------------------------------------
class PerfRptTable { [string]$Name; [string[]]$H; [hashtable]$I; [System.Collections.Generic.List[string[]]]$Rows }
class PerfRpt {
    static [Globalization.CultureInfo]$Inv = [Globalization.CultureInfo]::InvariantCulture
    static [regex]$CsvRe = [regex]'(?:^|,)(?:"((?:[^"]|"")*)"|([^,]*))'
    static [regex]$XmlNeed = [regex]'[&<>"\x01-\x08\x0B\x0C\x0E-\x1F]'
    static [regex]$Ctl = [regex]'[\x01-\x08\x0B\x0C\x0E-\x1F]'
    static [hashtable]$Off = @{}
    static [bool]$UseRecorded = $false
    static [double]$RecOffset = 0
    static [void] ResetTz() { [PerfRpt]::Off = @{} }
    static [object] D([object]$s) {
        if ($null -eq $s) { return $null }
        if ($s -is [double]) { return $s }
        if ($s -is [int] -or $s -is [long] -or $s -is [decimal] -or $s -is [single]) { return [double]$s }
        $t = [string]$s
        if ($t.Length -eq 0 -or $t -eq '-') { return $null }
        $v = 0.0
        if ([double]::TryParse($t, [Globalization.NumberStyles]::Float, [PerfRpt]::Inv, [ref]$v)) { return $v }
        return $null
    }
    static [string[]] SplitCsv([string]$line) {
        if ($line.IndexOf([char]34) -lt 0) { return $line.Split(',') }
        $list = New-Object System.Collections.Generic.List[string]
        foreach ($m in [PerfRpt]::CsvRe.Matches($line)) {
            if ($m.Groups[1].Success) { $list.Add($m.Groups[1].Value.Replace('""', '"')) } else { $list.Add($m.Groups[2].Value) }
        }
        return $list.ToArray()
    }
    # CSV を読む。epoch で時間帯を絞り、途中で切れた行 (コンテナ停止時) は捨てる。
    static [object] ReadCsv([string]$path, [string]$name, [double]$from, [double]$to) {
        if (-not [IO.File]::Exists($path)) { return $null }
        $t = [PerfRptTable]::new(); $t.Name = $name; $t.I = @{}; $t.Rows = New-Object 'System.Collections.Generic.List[string[]]'
        $first = $true; $ei = -1; $nh = 0
        foreach ($line in [IO.File]::ReadLines($path, (New-Object System.Text.UTF8Encoding($false)))) {
            if ($first) {
                $first = $false; $t.H = [PerfRpt]::SplitCsv($line)
                for ($i = 0; $i -lt $t.H.Length; $i++) { $t.I[$t.H[$i]] = $i }
                if ($t.I.ContainsKey('epoch')) { $ei = [int]$t.I['epoch'] }
                $nh = $t.H.Length
                continue
            }
            if ($line.Length -eq 0) { continue }
            $f = [PerfRpt]::SplitCsv($line)
            if ($f.Length -lt $nh) { continue }
            if ($ei -ge 0) {
                $e = 0.0
                if (-not [double]::TryParse($f[$ei], [Globalization.NumberStyles]::Float, [PerfRpt]::Inv, [ref]$e)) { continue }
                if ($e -lt $from -or $e -gt $to) { continue }
            }
            $t.Rows.Add($f)
        }
        if ($first) { return $null }
        return $t
    }
    static [string] G([object]$t, [object]$row, [string]$key) { $o = $t.I[$key]; if ($null -eq $o) { return $null }; return $row[[int]$o] }
    static [object] GD([object]$t, [object]$row, [string]$key) { return [PerfRpt]::D([PerfRpt]::G($t, $row, $key)) }
    # 累計カウンタの増分 (減った = 再起動・統計のクリアのときは、その回の値をそのまま増分とする)
    static [object] Dl([object]$cur, [object]$prev) {
        $c = [PerfRpt]::D($cur); $p = [PerfRpt]::D($prev)
        if ($null -eq $c -or $null -eq $p) { return $null }
        $d = $c - $p; if ($d -lt 0) { return $c }; return $d
    }
    static [object] Div([object]$a, [object]$b) {
        $x = [PerfRpt]::D($a); $y = [PerfRpt]::D($b)
        if ($null -eq $x -or $null -eq $y -or $y -eq 0) { return $null }; return $x / $y
    }
    static [object] Rnd([object]$v, [int]$n) { $x = [PerfRpt]::D($v); if ($null -eq $x) { return $null }; return [Math]::Round([double]$x, $n) }
    static [object] KBtoMB([object]$kb) { $x = [PerfRpt]::D($kb); if ($null -eq $x) { return $null }; return [Math]::Round([double]$x / 1024.0, 1) }
    static [void] AddV([System.Collections.Generic.List[double]]$list, [object]$v) { $x = [PerfRpt]::D($v); if ($null -ne $x) { $list.Add([double]$x) } }
    # UNIX 秒 → Excel のシリアル値 (この PC のタイムゾーン、または記録時の UTC オフセット)
    static [double] XT([double]$ep) {
        if ([PerfRpt]::UseRecorded) { return ($ep + [PerfRpt]::RecOffset) / 86400.0 + 25569.0 }
        $h = [long][Math]::Floor($ep / 3600)
        if (-not [PerfRpt]::Off.ContainsKey($h)) {
            [PerfRpt]::Off[$h] = [TimeZoneInfo]::Local.GetUtcOffset([DateTimeOffset]::FromUnixTimeSeconds([long]$ep).UtcDateTime).TotalSeconds
        }
        return ($ep + [double][PerfRpt]::Off[$h]) / 86400.0 + 25569.0
    }
    static [string] X([string]$s) {
        if ([string]::IsNullOrEmpty($s) -or -not [PerfRpt]::XmlNeed.IsMatch($s)) { return $s }
        $s = $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
        return [PerfRpt]::Ctl.Replace($s, '')
    }
    static [string] NStr([object]$v) {
        $x = [PerfRpt]::D($v); if ($null -eq $x) { return '' }
        $d = [double]$x
        if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return '' }
        return $d.ToString('R', [PerfRpt]::Inv)
    }
    # 表の行を <row> の XML にして書き出す (fm / st は 0 始まりの列ごとの書式・スタイル番号)
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
                    [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('" t="inlineStr"><is><t xml:space="preserve">').Append([PerfRpt]::X([string]$v)).Append('</t></is></c>')
                } elseif ($v -is [string]) {
                    $d = [PerfRpt]::D($v)
                    if ($null -eq $d) { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($textStyle).Append('" t="inlineStr"><is><t xml:space="preserve">').Append([PerfRpt]::X($v)).Append('</t></is></c>') }
                    else { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"><v>').Append([PerfRpt]::NStr($d)).Append('</v></c>') }
                } else {
                    $t = [PerfRpt]::NStr($v)
                    if ($t.Length -eq 0) { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"/>') }
                    else { [void]$sb.Append('<c r="').Append($rf).Append('" s="').Append($st[$c]).Append('"><v>').Append($t).Append('</v></c>') }
                }
            }
            [void]$sb.Append('</row>')
            $w.Write($sb.ToString())
        }
    }
}
$PerfRpt = [PerfRpt]

#------------------------------------------------------------------------------
# 入力 (ディレクトリ / 親ディレクトリ / .zip / .tar.gz)
#------------------------------------------------------------------------------
$TempDirs = New-Object System.Collections.Generic.List[string]
$DataFiles = @('perf_meta.csv', 'cli_datasource.csv', 'jvm_jstat_gc.csv')
function Test-PerfDir([string]$d) {
    foreach ($f in $DataFiles) { if (Test-Path -LiteralPath (Join-Path $d $f)) { return $true } }
    return $false
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
            $tarExe = 'tar'
            if ($env:SystemRoot -and (Test-Path -LiteralPath (Join-Path $env:SystemRoot 'System32\tar.exe'))) { $tarExe = Join-Path $env:SystemRoot 'System32\tar.exe' }
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
    foreach ($fn in $DataFiles) {
        foreach ($f in (Get-ChildItem -LiteralPath $root -Recurse -File -Filter $fn -ErrorAction SilentlyContinue)) {
            if (-not $found.Contains($f.DirectoryName)) { $found.Add($f.DirectoryName) }
        }
    }
    if ($found.Count -eq 0) { throw "perf-metrics の出力 (perf_meta.csv / cli_datasource.csv / jvm_jstat_gc.csv) が見つかりません: $path" }
    return , ($found.ToArray() | Sort-Object)
}

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
        Truncated = $false
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
    $w.Write((Ws-Begin 'FF404040' $true $(if ($nr -gt 0) { $sh.FreezeCols } else { 0 }) $(if ($nr -gt 0) { 6 } else { 0 }) $false 90))
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
    $w.Write((Ws-End 'landscape' $true $false $sh.Title $extra))
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
    $plot += '<c:valAx><c:axId val="' + $ax2 + '"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>' + $GRID
    if ($ch.YTitle -ne '') { $plot += Ax-Title $ch.YTitle }
    $plot += '<c:numFmt formatCode="' + (X $ch.YFmt) + '" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' + $AXLN + (Tx-Pr 800 '')
    $plot += '<c:crossAx val="' + $ax1 + '"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><c:chart>' + (C-Title $ch.Title 1000) + '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' + $plot + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + (Tx-Pr 800 '') + '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900">' + $FONT + '</a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr></c:chartSpace>'
}
# 系列: 名前・シート・列 (X は日時列)。グループ (データソース別など) の行範囲を使う
function Sr([string]$name, $sh, [int]$ycol, $grp) {
    if ($null -eq $sh -or $sh.Rows.Count -eq 0) { return $null }
    if ($null -eq $grp) { $r1 = $DATA_ROW; $r2 = $DATA_ROW + $sh.Rows.Count - 1 } else { $r1 = $grp.R1; $r2 = $grp.R2 }
    return [pscustomobject]@{ N = $name; S = $sh.Name; X = 1; Y = $ycol; R1 = $r1; R2 = $r2 }
}
$ChartSecs = New-Object System.Collections.Generic.List[object]
function Add-ChartSection([string]$title) { $ChartSecs.Add([pscustomobject]@{ Title = $title; Charts = (New-Object System.Collections.Generic.List[object]) }) }
function Add-Chart([string]$title, [string]$ytitle, [string]$yfmt, [object[]]$series) {
    $ss = @(); foreach ($s in $series) { if ($null -ne $s) { $ss += $s } }
    if ($ss.Count -eq 0) { return }
    $ChartSecs[$ChartSecs.Count - 1].Charts.Add([pscustomobject]@{ Title = $title; YTitle = $ytitle; YFmt = $yfmt; Series = $ss })
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
        Add-TextEntry $zip 'xl/drawings/drawing1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + $anch.ToString() + '</xdr:wsDr>')
        Add-TextEntry $zip 'xl/drawings/_rels/drawing1.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels.ToString() + '</Relationships>')
    }
    Add-DefName $sheetIdx '_xlnm.Print_Area' ("'" + $sheetName + "'!`$A`$1:`$R`$" + [Math]::Max($r, 4))
    $script:NCH = $n
    return $sb.ToString()
}

#------------------------------------------------------------------------------
# 項目名 (収集情報シート)
#------------------------------------------------------------------------------
$MetaLabel = @{
    'perf.version' = 'perf-metrics.sh の版'; 'perf.start_time' = '収集開始 (コンテナの時刻)'; 'perf.start_epoch' = '収集開始 (UNIX 秒)'
    'perf.utc_offset' = 'コンテナの UTC オフセット'; 'perf.output_dir' = '出力先'; 'perf.collectors' = '動かした収集'
    'perf.cli_interval' = '① jboss-cli 間隔 (秒)'; 'perf.jstat_interval' = '② jstat 間隔 (秒)'; 'perf.statistics_opts' = '統計の有効化 (起動引数)'
    'container.hostname' = 'ホスト名'; 'container.service_name' = 'サービス名'; 'container.component_name' = 'コンポーネント名'
    'container.log_id' = 'LOG_ID (タスク専用ディレクトリ)'; 'container.command' = '起動コマンド'
    'cli.controller' = 'jboss-cli の接続先'; 'cli.ds_stats_enabled_at_runtime' = 'プール統計を実行時に有効化した時刻'
    'jvm.jstat_command' = 'jstat の起動方法'
}
$MetaCat = @{ perf = '収集条件'; container = 'コンテナ'; cli = 'jboss-cli'; jvm = 'jstat' }

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
    foreach ($n in @('cli_datasource', 'jvm_jstat_gc')) {
        $t = Read-PerfCsv $dir ($n + '.csv')
        if ($null -ne $t -and $t.Rows.Count -gt 0) { $TB[$n] = $t; Write-Info ("読み込み: {0}.csv ({1:N0} 行)" -f $n, $t.Rows.Count) }
    }
    if ($TB.Count -eq 0) { Write-Warning "この期間のデータがありません: $dir"; return }

    # 全体の時間範囲 (グラフの時間軸)
    $script:TMin = $null; $script:TMax = $null
    foreach ($t in $TB.Values) {
        $ei = $t.I['epoch']
        $a = $PerfRpt::XT([double]$t.Rows[0][$ei]); $b = $PerfRpt::XT([double]$t.Rows[$t.Rows.Count - 1][$ei])
        if ($null -eq $script:TMin -or $a -lt $script:TMin) { $script:TMin = $a }
        if ($null -eq $script:TMax -or $b -gt $script:TMax) { $script:TMax = $b }
    }

    $KPI = @{}          # 主要指標用の値の列
    $Find = New-Object System.Collections.Generic.List[object]   # 所見 (重さ, 分類, 内容)
    $Sheets = New-Object System.Collections.Generic.List[object]

    # ==== DB 接続プール ====
    $dsSum = New-Object System.Collections.Generic.List[object]
    $shDS = $null
    if ($TB.ContainsKey('cli_datasource')) {
        Write-Info '集計: DB 接続プール'
        $t = $TB['cli_datasource']
        $sh = New-Sheet 'DB接続プール' 'DB 接続プール（データソース別の時系列）' ('① jboss-cli の /subsystem=datasources/data-source=*/statistics=pool (XA データソースも同じ)。「区間」は前回の取得からの増分です。最大プールが未設定のデータソースは IronJacamar の既定値 20 で使用率を計算しています。使用率 80% 以上・待ち・取得失敗は濃淡で表示します。') 2
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
            $lIn = NewList; $lUs = NewList; $wsum = 0.0; $fsum = 0.0; $maxWaitT = $null; $atMax = 0; $mxp = $null; $lastAg = $null
            foreach ($r in $rows) {
                $mx = $PerfRpt::GD($t, $r, 'max_pool_size'); if ($null -eq $mx) { $mx = 20.0 }
                $in = $PerfRpt::GD($t, $r, 'InUseCount')
                $wd = $null; $fd = $null; $cr = $null; $de = $null; $tmo = $null
                if ($null -ne $p) {
                    $wd = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'WaitCount')), ($PerfRpt::GD($t, $p, 'WaitCount')))
                    $fd = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'BlockingFailureCount')), ($PerfRpt::GD($t, $p, 'BlockingFailureCount')))
                    $cr = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'CreatedCount')), ($PerfRpt::GD($t, $p, 'CreatedCount')))
                    $de = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'DestroyedCount')), ($PerfRpt::GD($t, $p, 'DestroyedCount')))
                    $tmo = $PerfRpt::Dl(($PerfRpt::GD($t, $r, 'TimedOut')), ($PerfRpt::GD($t, $p, 'TimedOut')))
                }
                $us = $PerfRpt::Div($in, $mx)
                $PerfRpt::AddV($lIn, $in); $PerfRpt::AddV($lUs, $us)
                if ($null -ne $wd) { $wsum += $wd }; if ($null -ne $fd) { $fsum += $fd }
                if ($null -ne $in -and $in -ge $mx) { $atMax++ }
                $v = $PerfRpt::GD($t, $r, 'MaxWaitTime'); if ($null -ne $v -and ($null -eq $maxWaitT -or $v -gt $maxWaitT)) { $maxWaitT = $v }
                $lastAg = $PerfRpt::GD($t, $r, 'AverageGetTime'); $mxp = $mx
                $vals = New-Object 'object[]' $sh.Cols.Count
                $vals[$cT - 1] = $PerfRpt::XT(($PerfRpt::GD($t, $r, 'epoch'))); $vals[$cN - 1] = $name
                $vals[$cIn - 1] = $in; $vals[$cAc - 1] = $PerfRpt::GD($t, $r, 'ActiveCount'); $vals[$cId - 1] = $PerfRpt::GD($t, $r, 'IdleCount'); $vals[$cAv - 1] = $PerfRpt::GD($t, $r, 'AvailableCount')
                $vals[$cMx - 1] = $mx; $vals[$cUs - 1] = $PerfRpt::Rnd($us, 4); $vals[$cWd - 1] = $wd; $vals[$cFd - 1] = $fd
                $vals[$cWc - 1] = $PerfRpt::GD($t, $r, 'WaitCount'); $vals[$cMw - 1] = $PerfRpt::GD($t, $r, 'MaxWaitCount'); $vals[$cMu - 1] = $PerfRpt::GD($t, $r, 'MaxUsedCount')
                $vals[$cAg - 1] = $lastAg; $vals[$cMg - 1] = $PerfRpt::GD($t, $r, 'MaxGetTime'); $vals[$cAb - 1] = $PerfRpt::GD($t, $r, 'AverageBlockingTime')
                $vals[$cMwt - 1] = $PerfRpt::GD($t, $r, 'MaxWaitTime'); $vals[$cAu - 1] = $PerfRpt::GD($t, $r, 'AverageUsageTime'); $vals[$cMuT - 1] = $PerfRpt::GD($t, $r, 'MaxUsageTime')
                $vals[$cMc - 1] = $PerfRpt::GD($t, $r, 'MaxCreationTime'); $vals[$cCr - 1] = $cr; $vals[$cDe - 1] = $de; $vals[$cTo - 1] = $tmo
                $vals[$cMn - 1] = $PerfRpt::GD($t, $r, 'min_pool_size'); $vals[$cBt - 1] = $PerfRpt::GD($t, $r, 'blocking_timeout_ms'); $vals[$cJn - 1] = $PerfRpt::G($t, $r, 'jndi_name'); $vals[$cSe - 1] = $PerfRpt::G($t, $r, 'statistics_enabled')
                Add-Row $sh $vals $key
                $p = $r
            }
            $dsSum.Add([pscustomobject]@{ Name = $name; Max = $mxp; In = (Stat $lIn); Us = (Stat $lUs); Wait = $wsum; Fail = $fsum; MaxWait = $maxWaitT; AtMax = $atMax; N = $rows.Count; AvgGet = $lastAg; StatsOff = (($PerfRpt::G($t, $rows[$rows.Count - 1], 'statistics_enabled')) -eq 'false') })
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shDS = $sh
        $script:DSC = @{ In = $cIn; Ac = $cAc; Mx = $cMx; Wd = $cWd; Fd = $cFd }
    }

    # ==== ヒープ・GC (jstat -gc) ====
    $shJS = $null
    if ($TB.ContainsKey('jvm_jstat_gc')) {
        Write-Info '集計: ヒープ・GC（jstat）'
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
        $lHu = NewList; $lOu = NewList; $lOp = NewList; $lMu = NewList; $lYa = NewList; $ygS = 0.0; $ytS = 0.0; $fgS = 0.0; $ftS = 0.0; $cS = 0.0; $gtS = 0.0; $hcMax = $null
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
                if ($null -ne $yn -and $yn -gt 0 -and $null -ne $yt) { $a = $yt * 1000 / $yn; $vals[$cYa - 1] = $PerfRpt::Rnd($a, 1); $PerfRpt::AddV($lYa, $a) }
                if ($el -gt 0 -and $null -ne $gt) { $vals[$cGr - 1] = $PerfRpt::Rnd(($gt / $el), 4) }
                if ($null -ne $yn) { $ygS += $yn }; if ($null -ne $yt) { $ytS += $yt }; if ($null -ne $fn) { $fgS += $fn }; if ($null -ne $ft) { $ftS += $ft }; if ($null -ne $cgcn) { $cS += $cgcn }; if ($null -ne $gt) { $gtS += $gt }
            }
            $PerfRpt::AddV($lHu, ($PerfRpt::KBtoMB($hu))); $PerfRpt::AddV($lOu, ($PerfRpt::KBtoMB($ou))); $PerfRpt::AddV($lOp, $op); $PerfRpt::AddV($lMu, ($PerfRpt::KBtoMB(($PerfRpt::GD($t, $r, 'MU')))))
            if ($null -ne $hc -and ($null -eq $hcMax -or $hc -gt $hcMax)) { $hcMax = $hc }
            Add-Row $sh $vals ''
            $p = $r
        }
        Fix-Ranges $sh; $Sheets.Add($sh); $shJS = $sh
        $elapsed = ($PerfRpt::GD($t, $t.Rows[$t.Rows.Count - 1], 'epoch')) - ($PerfRpt::GD($t, $t.Rows[0], 'epoch'))
        $KPI['jsHeap'] = Stat $lHu; $KPI['jsOld'] = Stat $lOu; $KPI['jsOldP'] = Stat $lOp; $KPI['jsMeta'] = Stat $lMu; $KPI['jsYa'] = Stat $lYa
        $KPI['jsY'] = $ygS; $KPI['jsYt'] = $ytS; $KPI['jsF'] = $fgS; $KPI['jsFt'] = $ftS; $KPI['jsC'] = $cS; $KPI['jsGt'] = $gtS; $KPI['jsEl'] = $elapsed; $KPI['jsHcMax'] = $PerfRpt::KBtoMB($hcMax)
        $script:JSC = @{ Hu = $cHu; Hc = $cHc; Eu = $cEu; Ou = $cOu; Mu = $cMu; Mc = $cMc; Cu = $cCu; Yn = $cYn; Fn = $cFn; Cn = $cCn; Yt = $cYt; Ft = $cFt; Gt = $cGt }
    }

    # ==== 所見 ====
    function AddF([int]$lv, [string]$cat, [string]$msg) { $Find.Add(@($lv, $cat, $msg)) }
    foreach ($d in $dsSum) {
        if ($d.Fail -gt 0) { AddF 3 'DB 接続プール' ('「' + $d.Name + '」で接続の取得失敗 (タイムアウト) が ' + $d.Fail.ToString('#,0') + ' 件ありました。アプリでは IJ000453 等の例外になります。最大プール (' + $d.Max + ') と、接続を長く保持している処理を確認してください。') }
        if ($d.AtMax -gt 0) { AddF 2 'DB 接続プール' ('「' + $d.Name + '」の使用中が最大プール (' + $d.Max + ') に達しました (' + $d.AtMax + ' / ' + $d.N + ' 回)。') }
        if ($d.Wait -gt 0) { AddF 2 'DB 接続プール' ('「' + $d.Name + '」で空き接続の待ちが ' + $d.Wait.ToString('#,0') + ' 件発生しました (最大待ち時間 ' + $(if ($null -ne $d.MaxWait) { $d.MaxWait.ToString('#,0') } else { '-' }) + ' ms)。') }
        elseif ($null -ne $d.Us -and $d.Us.Max -ge 0.8) { AddF 1 'DB 接続プール' ('「' + $d.Name + '」の使用率が最大 ' + ($d.Us.Max * 100).ToString('0') + '% まで上がりました (待ちは発生していません)。') }
        if ($d.StatsOff) { AddF 1 'DB 接続プール' ('「' + $d.Name + '」のプール統計が無効です。値は 0 のままになります。') }
    }
    if ($KPI.ContainsKey('jsF') -and $KPI['jsF'] -gt 0) { AddF 3 'GC' ('Full GC が ' + $KPI['jsF'] + ' 回発生しました (合計 ' + ($KPI['jsFt'] * 1000).ToString('#,0') + ' ms)。ヒープ不足・メタスペース不足・System.gc() を確認してください。') }
    if ($null -ne $KPI['jsOldP'] -and $KPI['jsOldP'].Max -ge 0.85) { AddF 2 'ヒープ' ('Old 領域の使用率が最大 ' + ($KPI['jsOldP'].Max * 100).ToString('0') + '% に達しました。') }
    if ($KPI.ContainsKey('jsEl') -and $KPI['jsEl'] -gt 0) {
        $ratio = $KPI['jsGt'] / $KPI['jsEl']
        if ($ratio -ge 0.05) { AddF 2 'GC' ('期間中の GC 時間率が ' + ($ratio * 100).ToString('0.0') + '% です (目安 5% 未満)。') }
    }
    if ($Find.Count -eq 0) { AddF 0 '全般' '特記事項はありません (接続プールの待ち・取得失敗・上限到達、Full GC、Old 領域・GC 時間の逼迫は検出されませんでした)。' }

    # ==== グラフ ====
    if ($null -ne $shDS) {
        Add-ChartSection 'DB 接続プール'
        $c = $script:DSC
        foreach ($g in $shDS.Groups) {
            $dn = ($g.Key -split "`t")[-1]
            Add-Chart ($dn + '：接続数') '接続数' '#,##0' @((Sr '使用中' $shDS $c.In $g), (Sr 'アクティブ' $shDS $c.Ac $g), (Sr '最大プール' $shDS $c.Mx $g))
            Add-Chart ($dn + '：待ち・取得失敗（区間）') '件' '#,##0' @((Sr '待ち' $shDS $c.Wd $g), (Sr '取得失敗' $shDS $c.Fd $g))
        }
    }
    if ($null -ne $shJS) {
        Add-ChartSection 'ヒープ・GC（jstat）'
        $c = $script:JSC
        Add-Chart 'ヒープ使用量（MB）' 'MB' '#,##0' @((Sr '使用（合計）' $shJS $c.Hu $null), (Sr 'Old 使用' $shJS $c.Ou $null), (Sr 'Eden 使用' $shJS $c.Eu $null), (Sr '容量（合計）' $shJS $c.Hc $null))
        Add-Chart 'GC 回数（区間）' '回' '#,##0' @((Sr 'Young' $shJS $c.Yn $null), (Sr 'Full' $shJS $c.Fn $null), (Sr 'Concurrent' $shJS $c.Cn $null))
        Add-Chart 'GC 時間（区間、ミリ秒）' 'ミリ秒' '#,##0' @((Sr 'GC 合計' $shJS $c.Gt $null), (Sr 'Young' $shJS $c.Yt $null), (Sr 'Full' $shJS $c.Ft $null))
        Add-Chart 'メタスペース（MB）' 'MB' '#,##0' @((Sr '使用' $shJS $c.Mu $null), (Sr '容量' $shJS $c.Mc $null), (Sr 'CCS 使用' $shJS $c.Cu $null))
    }

    # ==== ブックの書き出し ====
    $label = M 'container.log_id'; if ($label -eq '') { $label = Split-Path -Leaf (Split-Path -Parent $dir) }
    $svc = (@((M 'container.service_name'), (M 'container.component_name')) | Where-Object { $_ -ne '' }) -join ' / '
    $range = (SerialText $script:TMin) + ' ～ ' + (SerialText $script:TMax) + '（' + ([TimeSpan]::FromDays($script:TMax - $script:TMin)).ToString('hh\:mm\:ss') + '）'
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
        Add-TextEntry $zip 'xl/worksheets/sheet1.xml' (Write-SummarySheet 1 $label $svc $range $tzNote $meta $KPI $dsSum $Find)
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
function Write-SummarySheet([int]$sheetIdx, [string]$label, [string]$svc, [string]$range, [string]$tzNote, $meta, $KPI, $dsSum, $Find) {
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
    # KRow: 時系列の統計 (最大 / 95% 値 / 平均 / 最小)、KOne: 期間全体の値 1 つ
    function KRow([string]$cat, [string]$name, $st, [string]$unit, [string]$note, [string]$fmt) { $kp.Add(@($cat, $name, $st, $unit, $note, 1, $fmt)) }
    function KOne([string]$cat, [string]$name, $v, [string]$unit, [string]$note, [string]$fmt) { $kp.Add(@($cat, $name, $v, $unit, $note, 0, $fmt)) }
    foreach ($d in $dsSum) {
        KRow 'DB 接続プール' ($d.Name + ' 使用中') $d.In '接続' ('最大プール ' + $d.Max + '・待ち ' + (Fmt $d.Wait '#,0') + ' 件・取得失敗 ' + (Fmt $d.Fail '#,0') + ' 件') 'int'
    }
    if ($null -ne $KPI['jsHeap']) { KRow 'ヒープ（jstat）' 'ヒープ使用量' $KPI['jsHeap'] 'MB' ('容量 最大 ' + (Fmt $KPI['jsHcMax'] '#,0') + ' MB') 'dec1' }
    if ($null -ne $KPI['jsOld']) { KRow 'ヒープ（jstat）' 'Old 使用量' $KPI['jsOld'] 'MB' ('Old 使用率 最大 ' + (Fmt ($KPI['jsOldP'].Max * 100) '0.0') + '%') 'dec1' }
    if ($null -ne $KPI['jsMeta']) { KRow 'ヒープ（jstat）' 'メタスペース使用量' $KPI['jsMeta'] 'MB' '' 'dec1' }
    if ($KPI.ContainsKey('jsY')) {
        $gr = $(if ($KPI['jsEl'] -gt 0) { $KPI['jsGt'] / $KPI['jsEl'] * 100 } else { $null })
        KOne 'GC（jstat）' 'Young GC 回数 / 合計時間' $KPI['jsY'] '回' ('合計 ' + (Fmt ($KPI['jsYt'] * 1000) '#,0') + ' ms・平均停止 ' + $(if ($null -ne $KPI['jsYa']) { (Fmt $KPI['jsYa'].Avg '0.0') + ' ms (最大 ' + (Fmt $KPI['jsYa'].Max '0.0') + ' ms)' } else { '－' })) 'int'
        KOne 'GC（jstat）' 'Full GC 回数 / 合計時間' $KPI['jsF'] '回' ('合計 ' + (Fmt ($KPI['jsFt'] * 1000) '#,0') + ' ms') 'int'
        KOne 'GC（jstat）' 'Concurrent GC 回数' $KPI['jsC'] '回' 'G1 の並行マーキング等' 'int'
        KOne 'GC（jstat）' 'GC 時間率（期間全体）' $gr '%' ('GC 時間 ' + (Fmt ($KPI['jsGt']) '#,0.00') + ' 秒 ÷ 経過 ' + (Fmt $KPI['jsEl'] '#,0') + ' 秒') 'dec2'
    }
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
            $cells += NumC ('C' + $r) $(if ($null -ne $d.Us) { $PerfRpt::Rnd($d.Us.Max, 4) } else { $null }) $S_PCTB
            $cells += NumC ('D' + $r) $(if ($null -ne $d.Us) { $PerfRpt::Rnd($d.Us.Avg, 4) } else { $null }) $S_PCT
            $cells += NumC ('E' + $r) $d.Wait $(if ($d.Wait -gt 0) { $S_INTB } else { $S_INT })
            $cells += NumC ('F' + $r) $d.Fail $(if ($d.Fail -gt 0) { $S_INTB } else { $S_INT })
            $cells += NumC ('G' + $r) $d.MaxWait $S_INT
            $cells += NumC ('H' + $r) $d.AvgGet $S_INT
            [void]$sb.Append((Row $r $cells 0)); $r++
        }
        $r++
    }

    # --- 収集情報 (抜粋) ---
    [void]$sb.Append((Row $r (StrC ('A' + $r) '■ 収集情報（抜粋。全項目は「収集情報」シート）' $S_SEC) 20)); $r++
    $items = @(
        @('サービス', $svc), @('タスク (LOG_ID)', $label), @('ホスト名', (M2 'container.hostname')),
        @('収集間隔 (秒)', ('jboss-cli ' + (M2 'perf.cli_interval') + ' / jstat ' + (M2 'perf.jstat_interval'))),
        @('jboss-cli の接続先', (M2 'cli.controller')), @('jstat の起動方法', (M2 'jvm.jstat_command')),
        @('出力先', (M2 'perf.output_dir')), @('収集開始', ((M2 'perf.start_time') + ' (コンテナの時刻、UTC' + (M2 'perf.utc_offset') + ')'))
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
    [void]$sb.Append((Row 2 (StrC 'A2' 'perf_meta.csv の内容です (収集開始時の環境・条件と、jboss-cli / jstat の接続先・起動方法)。' $S_NOTE) 0))
    [void]$sb.Append((Row 4 ((StrC 'A4' '分類' $S_HDR) + (StrC 'B4' '項目' $S_HDR) + (StrC 'C4' '値' $S_HDR) + (StrC 'D4' 'キー' $S_HDR)) 20))
    $r = 5
    foreach ($k in $meta.Keys) {
        $cat = $k.Split('.')[0]
        $cl = $MetaCat[$cat]; if ($null -eq $cl) { $cl = $cat }
        $lb = $MetaLabel[$k]; if ($null -eq $lb) { $lb = $k.Substring($cat.Length + 1) }
        $v = [string]$meta[$k]
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
TH|区分|取得方法|出力
T|① jboss-cli|jboss-cli.sh を 1 つ常駐させ (FIFO で命令を送る)、データソース / XA データソースの statistics=pool と最小 / 最大プール等の設定を間隔ごとに読む。読み取りだけで設定は変えない (プール統計だけは実行時属性で有効化)|cli_datasource.csv
T|② jstat|jstat -gc -t <pid> <間隔> を 1 つ常駐させ、各行に時刻を付ける (runtime イメージに jstat が無い場合は java -m jdk.jcmd で起動)|jvm_jstat_gc.csv
B|
H|主な項目
TH|項目|元の値|説明
T|使用中 / アクティブ / アイドル|InUseCount / ActiveCount / IdleCount|アプリが使用中の接続数 / プールに作成済みの物理接続数 / 未使用の接続数
T|使用率|InUseCount ÷ 最大プール|最大プール (max-pool-size) が未設定のときは既定値 20 で計算
T|待ち / 取得失敗|WaitCount / BlockingFailureCount|空き接続を待った取得要求 / 取得タイムアウト (blocking-timeout-wait-millis 超過。アプリでは IJ000453 等)
T|平均取得 / 最大待ち|AverageGetTime / MaxWaitTime|接続の取得時間 (待ちを含む) / 空きを待った最大時間 (ミリ秒。統計の有効化以降の累積値)
T|ヒープ (jstat)|EU + S0U + S1U + OU|Eden・Survivor・Old の使用量の合計。容量は EC + S0C + S1C + OC
T|GC 時間率|GCT の増分 ÷ 経過秒|アプリが GC で止まっていた時間の割合 (G1 の Concurrent は並行処理の停止部分)。目安 5% 未満
B|
H|所見の判定
TH|重要度|内容|判定基準
T|■ 要対応|接続の取得失敗・Full GC|期間中の増分 > 0
T|▲ 注意|接続待ち・プールの上限到達・Old 使用率 ≧ 85%・GC 時間率 ≧ 5%|期間中に 1 回以上
T|● 参考|接続プール使用率 ≧ 80% (待ちなし)・プール統計が無効|－
B|
H|注意事項
P|・収集用の jboss-cli・jstat の JVM も同じコンテナの CPU・メモリを使います (いずれも 1 コアの 1% 未満・RSS 合計 150〜200MB 程度)。本番と比べる試験では、収集の有無で条件を揃えてください。
P|・-Dwildfly.datasources.statistics-enabled=true (データソースの統計の有効化) と、プール統計の実行時有効化によるわずかな負荷があります。
P|・jstat は /tmp が書き込めないと JVM の性能カウンタを読めません。その場合は「ヒープ・GC（jstat）」シートが作られません。
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
