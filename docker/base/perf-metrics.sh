#!/bin/sh
# =============================================================================
# perf-metrics.sh
#
# 性能試験 (JMeter など) の間、JBoss EAP 8.1 / OpenJDK 21 (Red Hat build) を動かす
# コンテナの中で次の統計を一定間隔で取得し、readonlyRootFilesystem=true でも
# 書き込める EFS 上のタスク専用ディレクトリへ CSV で書き出す。
#
#   ① jboss-cli : DB 接続プール (statistics=pool / jdbc)、Undertow (HTTP リスナー・
#                 セッション)、IO ワーカー、EJB / EE スレッドプール、トランザクション、
#                 JVM (platform-mbean: ヒープ・スレッド・クラス・GC・メモリプール)
#   ② jstat     : ヒープ各領域の容量・使用量と GC 回数・時間 (jstat -gc)
#      jcmd      : ヒープとメタスペースの詳細 (jcmd GC.heap_info)、JVM のバージョン・フラグ
#   ③ top / ps  : ロードアベレージ・CPU・メモリ、プロセス別 CPU / RSS / スレッド数、
#                 JBoss のスレッド別 CPU (top -H)、コンテナ (cgroup) の CPU・メモリ
#
# 性能試験用の実装はこの 1 ファイルにまとめている。
#   ビルド時: front / back の Dockerfile が、ビルド引数 PERF_METRICS_ENABLED=true の
#             ときだけこのファイルを残し `perf-metrics.sh install` で ps / top を入れる。
#             false (既定) ならファイルを削除する (本番と同じ中身になる)。
#   起動時  : efs-entrypoint.sh が PERF_METRICS_ENABLED=true のときだけこのファイルを
#             読み込み (.) perf_metrics_start を呼ぶ。収集は別プロセスで行い、
#             JBoss の起動を待たせない。収集側で何が起きても JBoss の起動は止めない。
#
# 【出力先】 (readonlyRootFilesystem=true でも書き込める場所)
#   JBoss のログと同じタスク専用ディレクトリの下の perf/ に書く。
#     /mnt/logs/<Component_name>/logs/<Service_Name>/mid/<LOG_ID>/perf/
#     例) /mnt/logs/front/logs/intra-web/mid/current/perf/
#         (current はエントリポイントが「最後に起動したタスクの <LOG_ID>」へ張る目印)
#   current は全タスクが共有して張り替えるリンクなので、書き込みの経路には使わない
#   (server.log を実体パスへ固定 (pin) しているのと同じ理由。docs/LOG_ROTATION.md)。
#   エントリポイントが決めた実体パス JBOSS_LOG_DIR (= mid/<LOG_ID>) を使い、取れない
#   ときだけ起動時点の current を実体パスに解決して使う。
#   FIFO・PID ファイルなどの作業ファイルはタスクローカルの書き込み可能領域
#   (/tmp → standalone/tmp → 出力先/.work の順に、書けるところ) に置く。
#
# 【出力ファイル】 (UTF-8・LF・1 行目が見出しの CSV。timestamp はコンテナの TZ、epoch は UNIX 秒)
#   perf_meta.csv          収集条件・環境 (category,key,value)
#   collector.log          収集プロセスのログ
#   ① cli_datasource.csv   データソース別 接続プール・JDBC 統計と設定 (最小/最大プール)
#      cli_undertow.csv     HTTP リスナー別 リクエスト数・エラー数・転送量・処理時間
#      cli_session.csv      デプロイメント別 HTTP セッション
#      cli_io_worker.csv    IO ワーカー (リクエスト処理スレッド) の使用数・キュー
#      cli_threadpool.csv   EJB / EE (managed-executor) スレッドプール
#      cli_transactions.csv トランザクション件数
#      cli_jvm.csv          ヒープ・非ヒープ・スレッド・クラス・サーバー状態、CLI 応答時間
#      cli_gc.csv           GC 別 回数・累積時間 (platform-mbean)
#      cli_mempool.csv      メモリプール別 使用量
#   ② jvm_jstat_gc.csv     jstat -gc (KB / 秒)。PERF_JSTAT_OPTIONS で gccause なども追加可
#      jvm_jcmd_heap.csv    jcmd GC.heap_info を数値化したもの
#      jvm_jcmd.log         jcmd の出力そのもの
#      jvm_info.txt         jcmd VM.version / VM.flags (起動時に 1 回)
#      jvm_threaddump/      PERF_THREAD_DUMP_INTERVAL を指定したときのスレッドダンプ
#   ③ os_top_summary.csv   top の要約行 (ロードアベレージ・CPU 内訳・メモリ)
#      os_top_process.csv   top のプロセス行 (区間の %CPU・RES など)
#      os_top_threads.csv   JBoss のスレッド別 %CPU 上位 (PERF_TOP_THREADS 件)
#      os_top.log           top の出力そのもの (パスワードらしき値は伏せる)
#      os_ps.csv            ps のプロセス一覧 (スレッド数・RSS・累積 CPU 秒)
#      os_cgroup.csv        コンテナ (cgroup v1/v2) の CPU 使用量・スロットリング・メモリ
#
# 【使い方】
#   ビルド    : docker build --build-arg PERF_METRICS_ENABLED=true ... docker/front
#   起動      : 何もしなくてよい (エントリポイントが起動時に読み込む)
#   状態確認  : perf-metrics.sh status            (ECS Exec でコンテナに入って実行)
#   停止      : perf-metrics.sh stop
#   スレッドダンプ: perf-metrics.sh threaddump    (その場で 1 回取得)
#   まとめる  : perf-metrics.sh pack              (perf/ を tar 1 ファイルにする)
#   レポート  : perf/ を Windows に持ってきて New-PerfMetricsReport.ps1 で Excel にする
#
# 【設定】 (環境変数。ビルド時の ENV・タスク定義の environment のどちらでも指定できる)
#   PERF_METRICS_ENABLED     true | false     収集するか (ビルド引数から ENV に焼き込む)
#   PERF_INTERVAL            10               各収集の既定の間隔 (秒)
#   PERF_CLI_INTERVAL        PERF_INTERVAL    ① jboss-cli の間隔 (秒)
#   PERF_JSTAT_INTERVAL      PERF_INTERVAL    ② jstat の間隔 (秒)
#   PERF_JCMD_INTERVAL       60               ② jcmd の間隔 (秒)。jcmd は 1 回ごとに JVM を
#                                             起動するため既定を長めにしている (10 も指定可)
#   PERF_TOP_INTERVAL        PERF_INTERVAL    ③ top の間隔 (秒)
#   PERF_PS_INTERVAL         PERF_INTERVAL    ③ ps / cgroup の間隔 (秒)
#   PERF_COLLECTORS          "cli jstat jcmd top ps"   動かす収集 (空白区切り)
#   PERF_DURATION            0                収集する秒数 (0 = コンテナが止まるまで)
#   PERF_OUTPUT_DIR          (自動)           出力先を明示するとき
#   PERF_WORK_DIR            (自動)           作業ディレクトリを明示するとき
#   PERF_TZ                  (コンテナの TZ)  タイムスタンプのタイムゾーン (例 Asia/Tokyo)
#   PERF_ENABLE_STATISTICS   true             standalone.sh に -Dwildfly.statistics-enabled=true
#                                             を足す (Undertow・トランザクション等の統計)
#   PERF_CLI_ENABLE_DS_STATS true             データソースのプール統計を実行時に有効化する
#                                             (statistics=pool / jdbc の実行時属性。設定ファイルは
#                                              変えず、再起動で元に戻る)
#   PERF_CLI_MODE            persistent       persistent (jboss-cli を 1 つ常駐させる) | oneshot
#   PERF_CLI_CONTROLLER      (自動)           例 remote+http://127.0.0.1:9990
#   PERF_CLI_USER / PERF_CLI_PASSWORD         ローカル認証が使えないときの管理ユーザー
#   PERF_CLI_TIMEOUT         30               1 回の取得のタイムアウト (秒)
#   PERF_CLI_HEAP            64m              jboss-cli の最大ヒープ
#   PERF_JSTAT_OPTIONS       gc               jstat のオプション (空白区切り。例 "gc gccause")
#   PERF_JCMD_COMMANDS       GC.heap_info     jcmd のコマンド (カンマ区切り)
#   PERF_THREAD_DUMP_INTERVAL 0               スレッドダンプの間隔 (秒。0 = 取らない)
#   PERF_TOP_THREADS         10               JBoss のスレッド別 CPU を上位何件記録するか (0 = 取らない)
#   PERF_WAIT_TIMEOUT        900              JBoss の JVM / 管理インタフェースを待つ秒数
#
# 【注意】
#   - 収集用に jboss-cli (常駐 1 つ・最大ヒープ 64MB)、jstat (常駐 1 つ・32MB) の JVM が
#     増える。RSS で合計 200MB 前後を見込み、タスクのメモリに余裕を持たせること。
#     収集プロセス自身の CPU / RSS も os_top_process.csv / os_ps.csv に記録される。
#   - jstat / jcmd は /tmp (hsperfdata・アタッチ用ソケット) が書き込めないと JVM に
#     接続できない。readonlyRootFilesystem=true では /tmp にボリュームを当てること。
#     接続できない場合もヒープ・GC は ① の platform-mbean の値で確認できる。
#   - 本ファイルは /bin/sh (POSIX) で動く。エントリポイント (set -eu) から読み込まれる
#     ため、トップレベルでは関数と既定値の定義しか行わない。
# =============================================================================

PERF_METRICS_VERSION="1.0.0"

# --- 既定値 (エントリポイントから読み込まれても安全な代入だけ) ------------------
: "${PERF_METRICS_ENABLED:=false}"
: "${PERF_INTERVAL:=10}"
: "${PERF_CLI_INTERVAL:=${PERF_INTERVAL}}"
: "${PERF_JSTAT_INTERVAL:=${PERF_INTERVAL}}"
: "${PERF_JCMD_INTERVAL:=60}"
: "${PERF_TOP_INTERVAL:=${PERF_INTERVAL}}"
: "${PERF_PS_INTERVAL:=${PERF_INTERVAL}}"
: "${PERF_COLLECTORS:=cli jstat jcmd top ps}"
: "${PERF_DURATION:=0}"
: "${PERF_OUTPUT_DIR:=}"
: "${PERF_WORK_DIR:=}"
: "${PERF_TZ:=}"
: "${PERF_ENABLE_STATISTICS:=true}"
: "${PERF_CLI_ENABLE_DS_STATS:=true}"
: "${PERF_CLI_MODE:=persistent}"
: "${PERF_CLI_CONTROLLER:=}"
: "${PERF_CLI_USER:=}"
: "${PERF_CLI_PASSWORD:=}"
: "${PERF_CLI_TIMEOUT:=30}"
: "${PERF_CLI_CONNECT_TIMEOUT:=60}"
: "${PERF_CLI_HEAP:=64m}"
: "${PERF_CLI_JAVA_OPTS:=}"
: "${PERF_JSTAT_OPTIONS:=gc}"
: "${PERF_JCMD_COMMANDS:=GC.heap_info}"
: "${PERF_THREAD_DUMP_INTERVAL:=0}"
: "${PERF_TOP_THREADS:=10}"
: "${PERF_WAIT_TIMEOUT:=900}"
: "${PERF_TOOL_JAVA_OPTS:=-Xms8m -Xmx32m -XX:+UseSerialGC -XX:TieredStopAtLevel=1 -XX:-UsePerfData -Xss512k}"
: "${PERF_INSTALL_PACKAGES:=true}"

# エントリポイントが standalone.sh の直後へ足す引数 (統計の有効化)。
# 標準の構成は statistics-enabled="${wildfly.statistics-enabled:false}" の式を使って
# いるため、このシステムプロパティ 1 つで Undertow・トランザクション・EJB・
# (式を使っている) データソースの統計がまとめて有効になる。
PERF_SERVER_OPTS=""
if [ "${PERF_ENABLE_STATISTICS}" = "true" ]; then
    PERF_SERVER_OPTS="-Dwildfly.statistics-enabled=true"
fi

PM_SELF="${PERF_METRICS_SCRIPT:-/usr/local/bin/perf-metrics.sh}"

# =============================================================================
# 共通ヘルパー
# =============================================================================
pm_say() {
    printf '[perf-metrics] %s\n' "$*"
}

# ログ: collector.log へ。WARN / ERROR は標準エラー (CloudWatch Logs) にも出す。
pm_log() {
    _pm_ll="$(date '+%Y-%m-%d %H:%M:%S' 2>/dev/null) [perf-metrics] [${PM_NAME:-main}] $1 $2"
    if [ -n "${PM_LOG_FILE:-}" ]; then
        printf '%s\n' "${_pm_ll}" >> "${PM_LOG_FILE}" 2>/dev/null
    fi
    case "$1" in
        WARN|ERROR) printf '%s\n' "${_pm_ll}" >&2 ;;
    esac
    return 0
}

# 同じ警告を何度も出さない (キーごとに 1 回)
pm_warn_once() {  # key message
    case " ${PM_WARNED:-} " in
        *" $1 "*) return 0 ;;
    esac
    PM_WARNED="${PM_WARNED:-} $1"
    pm_log WARN "$2"
}

pm_is_uint() {
    case "$1" in
        ""|*[!0-9]*) return 1 ;;
    esac
    return 0
}

# 正の整数でなければ既定値に戻す
pm_uint_or() {  # value default
    if pm_is_uint "$1" && [ "$1" -gt 0 ]; then printf '%s' "$1"; else printf '%s' "$2"; fi
}

# ディレクトリへ実際に書き込めるか (EROFS / EACCES の両方を実書き込みで判定)
pm_is_writable() {
    [ -d "$1" ] || return 1
    _pm_probe="$1/.perf-metrics-writetest.$$"
    if ( : > "${_pm_probe}" ) 2>/dev/null; then
        rm -f "${_pm_probe}" 2>/dev/null
        return 0
    fi
    return 1
}

# 現在時刻: PM_TS (YYYY-MM-DD HH:MM:SS) と PM_EPOCH (UNIX 秒)
pm_now() {
    _pm_n="$(date '+%Y-%m-%d %H:%M:%S|%s')"
    PM_TS="${_pm_n%|*}"
    PM_EPOCH="${_pm_n#*|}"
}

pm_ms() {
    _pm_v="$(date +%s%N 2>/dev/null)"
    case "${_pm_v}" in
        *N*|"") echo $(( $(date +%s) * 1000 )) ;;
        *) echo $(( _pm_v / 1000000 )) ;;
    esac
}

# 次の「間隔の倍数の時刻」まで待つ (各収集の時刻をそろえる)。終了時刻を過ぎたら 1 を返す。
pm_sleep_tick() {  # interval
    _pm_s="$(date +%s)"
    _pm_nx=$(( (_pm_s / $1 + 1) * $1 ))
    if [ "${PM_END_EPOCH:-0}" -gt 0 ] && [ "${_pm_nx}" -ge "${PM_END_EPOCH}" ]; then
        return 1
    fi
    sleep $(( _pm_nx - _pm_s ))
    return 0
}

pm_ended() {
    [ "${PM_END_EPOCH:-0}" -gt 0 ] && [ "$(date +%s)" -ge "${PM_END_EPOCH}" ]
}

# 常駐させるコマンドの前に付ける「残り時間で止める」指定 (PERF_DURATION を指定したときだけ)
pm_timeout_prefix() {
    if [ "${PM_END_EPOCH:-0}" -gt 0 ] && command -v timeout >/dev/null 2>&1; then
        printf 'timeout %s' "$(( PM_END_EPOCH - $(date +%s) + 1 ))"
    fi
}

# CSV ファイルが無い (または空) なら見出し行を書く
pm_csv_init() {  # file header
    if [ ! -s "$1" ]; then
        printf '%s\n' "$2" >> "$1"
    fi
}

pm_meta() {  # category key value
    if [ -n "${PM_OUT:-}" ]; then
        printf '%s,%s,%s\n' "$1" "$2" "$(printf '%s' "$3" | sed -e 's/"/""/g' -e 's/^/"/' -e 's/$/"/')" \
            >> "${PM_OUT}/perf_meta.csv" 2>/dev/null
    fi
    return 0
}

# タイムスタンプのタイムゾーン。tzdata が無いイメージでも Asia/Tokyo は JST-9 で代替する。
pm_setup_tz() {
    case "${PERF_TZ}" in
        "") ;;
        UTC|utc) TZ=UTC0; export TZ ;;
        JST|jst|Asia/Tokyo)
            if [ -e /usr/share/zoneinfo/Asia/Tokyo ]; then TZ=Asia/Tokyo; else TZ=JST-9; fi
            export TZ ;;
        *) TZ="${PERF_TZ}"; export TZ ;;
    esac
}

# JBoss EAP サーバー JVM の PID (java で、jboss-modules.jar と org.jboss.as.standalone を含む)
pm_find_jboss_pid() {
    for _pm_d in /proc/[0-9]*; do
        [ -r "${_pm_d}/cmdline" ] || continue
        _pm_c="$(tr '\0' ' ' 2>/dev/null < "${_pm_d}/cmdline")" || continue
        case "${_pm_c}" in
            *org.jboss.as.cli*) continue ;;
            *jboss-modules.jar*org.jboss.as.standalone*) ;;
            *) continue ;;
        esac
        _pm_a0="${_pm_c%% *}"
        case "${_pm_a0##*/}" in
            java|java.*) printf '%s\n' "${_pm_d#/proc/}"; return 0 ;;
        esac
    done
    return 1
}

# JBoss の JVM が現れるまで待つ (PERF_WAIT_TIMEOUT 秒)。PM_JPID に入れる。
pm_wait_jboss() {
    _pm_dl=$(( $(date +%s) + PERF_WAIT_TIMEOUT ))
    while :; do
        PM_JPID="$(pm_find_jboss_pid)" && [ -n "${PM_JPID}" ] && return 0
        pm_ended && return 1
        if [ "$(date +%s)" -ge "${_pm_dl}" ]; then
            pm_log WARN "JBoss EAP の JVM が ${PERF_WAIT_TIMEOUT} 秒たっても見つかりません。この収集 (${PM_NAME}) を終了します"
            return 1
        fi
        sleep 2
    done
}

pm_jboss_alive() {
    [ -n "${PM_JPID:-}" ] && [ -r "/proc/${PM_JPID}/cmdline" ] && \
        tr '\0' ' ' 2>/dev/null < "/proc/${PM_JPID}/cmdline" | grep -q 'org.jboss.as.standalone'
}

# JBoss の JVM が使っている java の実体 (ツールは同じ JDK のものを使う)
pm_java_of_jboss() {
    _pm_j=""
    if [ -n "${PM_JPID:-}" ]; then
        _pm_j="$(readlink "/proc/${PM_JPID}/exe" 2>/dev/null)"
    fi
    if [ -z "${_pm_j}" ] || [ ! -x "${_pm_j}" ]; then
        if [ -n "${JAVA_HOME:-}" ] && [ -x "${JAVA_HOME}/bin/java" ]; then
            _pm_j="${JAVA_HOME}/bin/java"
        else
            _pm_j="$(command -v java 2>/dev/null)"
        fi
    fi
    [ -n "${_pm_j}" ] || return 1
    printf '%s\n' "${_pm_j}"
}

# jstat / jcmd の起動方法を決める。PM_TOOL_BIN (実行ファイル) か PM_TOOL_MAIN (java -m の主クラス)。
#   1) JBoss と同じ JDK の bin  2) PATH  3) java -m jdk.jcmd/<主クラス>
#   (Red Hat の OpenJDK 21 runtime イメージは jcmd はあるが jstat は無い。jdk.jcmd
#    モジュールは入っているので java -m で jstat を起動できる)
pm_resolve_tool() {  # jstat | jcmd
    PM_TOOL_BIN=""; PM_TOOL_MAIN=""; PM_TOOL_NAME="$1"
    PM_TOOL_JAVA="$(pm_java_of_jboss)" || PM_TOOL_JAVA=""
    if [ -n "${PM_TOOL_JAVA}" ]; then
        _pm_jb="$(readlink -f "${PM_TOOL_JAVA}" 2>/dev/null || printf '%s' "${PM_TOOL_JAVA}")"
        if [ -x "${_pm_jb%/java}/$1" ]; then PM_TOOL_BIN="${_pm_jb%/java}/$1"; return 0; fi
    fi
    if command -v "$1" >/dev/null 2>&1; then PM_TOOL_BIN="$(command -v "$1")"; return 0; fi
    if [ -n "${PM_TOOL_JAVA}" ] && \
       env -u JAVA_TOOL_OPTIONS -u _JAVA_OPTIONS -u JDK_JAVA_OPTIONS "${PM_TOOL_JAVA}" -XX:-UsePerfData --list-modules 2>/dev/null \
       | grep -q '^jdk\.jcmd@'; then
        case "$1" in
            jstat) PM_TOOL_MAIN="sun.tools.jstat.Jstat" ;;
            jcmd)  PM_TOOL_MAIN="sun.tools.jcmd.JCmd" ;;
        esac
        [ -n "${PM_TOOL_MAIN}" ] && return 0
    fi
    return 1
}

pm_tool_desc() {
    if [ -n "${PM_TOOL_BIN}" ]; then printf '%s' "${PM_TOOL_BIN}"; else printf '%s -m jdk.jcmd/%s' "${PM_TOOL_JAVA}" "${PM_TOOL_MAIN}"; fi
}

# 解決したツールを、アプリ用の JVM オプション (javaagent 等) を外して起動する
pm_tool_exec() {  # args...
    (
        unset JAVA_TOOL_OPTIONS _JAVA_OPTIONS JDK_JAVA_OPTIONS JAVA_OPTS
        LC_ALL=C; export LC_ALL
        # -Dperf.metrics.collector は目印 (ps / top で手動実行の jstat / jcmd と区別するため)
        _pm_mark="-Dperf.metrics.collector=${PM_TOOL_NAME:-tool}"
        if [ -n "${PM_TOOL_BIN}" ]; then
            _pm_jo="-J${_pm_mark}"
            for _pm_o in ${PERF_TOOL_JAVA_OPTS}; do _pm_jo="${_pm_jo} -J${_pm_o}"; done
            # shellcheck disable=SC2086
            exec ${PM_TOOL_TIMEOUT:-} "${PM_TOOL_BIN}" ${_pm_jo} "$@"
        fi
        # shellcheck disable=SC2086
        exec ${PM_TOOL_TIMEOUT:-} "${PM_TOOL_JAVA}" "${_pm_mark}" ${PERF_TOOL_JAVA_OPTS} -m "jdk.jcmd/${PM_TOOL_MAIN}" "$@"
    )
}

# hsperfdata (jstat が読む性能カウンタ) が /tmp にあるか
pm_hsperf_exists() {  # pid
    for _pm_h in /tmp/hsperfdata_*/"$1"; do
        [ -e "${_pm_h}" ] && return 0
    done
    return 1
}

# 指定 PID とその子孫を止める
pm_kill_tree() {  # pid
    _pm_list=" $1 "
    _pm_pass=0
    while [ "${_pm_pass}" -lt 3 ]; do
        for _pm_d in /proc/[0-9]*; do
            _pm_st="$(cat "${_pm_d}/stat" 2>/dev/null)" || continue
            _pm_st="${_pm_st##*) }"
            # shellcheck disable=SC2086
            set -- ${_pm_st}
            case "${_pm_list}" in
                *" $2 "*)
                    case "${_pm_list}" in *" ${_pm_d#/proc/} "*) ;; *) _pm_list="${_pm_list}${_pm_d#/proc/} " ;; esac
                    ;;
            esac
        done
        _pm_pass=$((_pm_pass + 1))
    done
    # shellcheck disable=SC2086
    kill -TERM ${_pm_list} 2>/dev/null
    sleep 1
    # shellcheck disable=SC2086
    kill -KILL ${_pm_list} 2>/dev/null
    return 0
}

# 作業ディレクトリの候補 (タスクローカルで書ける所)
pm_work_candidates() {
    [ -n "${PERF_WORK_DIR}" ] && printf '%s\n' "${PERF_WORK_DIR}"
    printf '%s\n' "/tmp/perf-metrics"
    printf '%s\n' "${JBOSS_HOME:-/opt/jboss-eap}/standalone/tmp/perf-metrics"
    [ -n "${PM_OUT:-}" ] && printf '%s\n' "${PM_OUT}/.work"
    return 0
}

# 状態ファイル (status / stop が使う) を探して読み込む
pm_load_state() {
    for _pm_w in $(pm_work_candidates); do
        if [ -r "${_pm_w}/perf-metrics.state" ]; then
            # shellcheck disable=SC1090
            . "${_pm_w}/perf-metrics.state"
            return 0
        fi
    done
    return 1
}

# パスワードらしき値を伏せる (ps / top の出力用 sed 式)
PM_MASK_SED='s/\([Pp][Aa][Ss][Ss][Ww][Oo][Rr][Dd][A-Za-z0-9_.-]*[=:]\)[^ ,]*/\1********/g'

# =============================================================================
# 起動 (efs-entrypoint.sh から呼ばれる)
# =============================================================================
# 出力先を決める。成功したら PM_OUT に入れる。
pm_resolve_output_dir() {
    PM_OUT=""
    PM_OUT_NOTE=""
    if [ -n "${PERF_OUTPUT_DIR}" ]; then
        PM_OUT="${PERF_OUTPUT_DIR}"
        PM_OUT_NOTE="PERF_OUTPUT_DIR で指定"
    elif [ -n "${JBOSS_LOG_DIR:-}" ] && [ -d "${JBOSS_LOG_DIR}" ]; then
        # エントリポイントが決めた、このタスク専用の実体ディレクトリ (mid/<LOG_ID>)
        PM_OUT="${JBOSS_LOG_DIR}/perf"
        PM_OUT_NOTE="JBoss のログと同じタスク専用ディレクトリ"
    elif [ -n "${EFS_LOG_DIR:-}" ] && [ -d "${EFS_LOG_DIR}/mid/current" ]; then
        # 起動時点の current を実体パスへ解決する (以後 current が張り替わっても影響しない)
        _pm_cur="$(cd "${EFS_LOG_DIR}/mid/current" 2>/dev/null && pwd -P)"
        [ -n "${_pm_cur}" ] && PM_OUT="${_pm_cur}/perf"
        PM_OUT_NOTE="起動時点の mid/current を実体パスへ解決"
    elif [ -d "${JBOSS_HOME:-/opt/jboss-eap}/standalone/log" ]; then
        _pm_cur="$(cd "${JBOSS_HOME:-/opt/jboss-eap}/standalone/log" 2>/dev/null && pwd -P)"
        [ -n "${_pm_cur}" ] && PM_OUT="${_pm_cur}/perf"
        PM_OUT_NOTE="standalone/log の解決先"
    fi
    [ -n "${PM_OUT}" ] || return 1
    mkdir -p "${PM_OUT}" 2>/dev/null
    pm_is_writable "${PM_OUT}"
}

pm_resolve_work_dir() {
    PM_W=""
    for _pm_w in $(pm_work_candidates); do
        mkdir -p "${_pm_w}" 2>/dev/null || continue
        if pm_is_writable "${_pm_w}"; then PM_W="${_pm_w}"; return 0; fi
    done
    return 1
}

# ECS タスクメタデータ (v4) から主な項目を perf_meta.csv へ
pm_meta_ecs() {
    [ -n "${ECS_CONTAINER_METADATA_URI_V4:-}" ] || return 0
    command -v curl >/dev/null 2>&1 || return 0
    _pm_tj="$(curl -fsS --max-time 3 "${ECS_CONTAINER_METADATA_URI_V4}/task" 2>/dev/null)" || return 0
    for _pm_k in TaskARN Cluster Family Revision AvailabilityZone LaunchType; do
        _pm_v="$(printf '%s' "${_pm_tj}" | sed -n "s/.*\"${_pm_k}\"[[:space:]]*:[[:space:]]*\"\\([^\"]*\\)\".*/\\1/p" | head -n 1)"
        [ -n "${_pm_v}" ] && pm_meta ecs "${_pm_k}" "${_pm_v}"
    done
    _pm_v="$(printf '%s' "${_pm_tj}" | tr -d '\n ' | sed -n 's/.*"Limits":{\([^}]*\)}.*/\1/p' | head -n 1)"
    [ -n "${_pm_v}" ] && pm_meta ecs TaskLimits "${_pm_v}"
    return 0
}

pm_banner_line() {
    pm_say "$1"
    [ -n "${PM_LOG_FILE:-}" ] && printf '%s [perf-metrics] [start] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$1" >> "${PM_LOG_FILE}" 2>/dev/null
    return 0
}

pm_has_collector() {
    case " ${PERF_COLLECTORS} " in *" $1 "*) return 0 ;; esac
    return 1
}

perf_metrics_start() {
    # エントリポイントの set -e / -u をこの関数の中だけ外す (収集側の都合で起動を止めない)
    _pm_saved_flags="$-"
    set +eu
    PM_NAME=start
    pm_perf_start_body "$@"
    _pm_rc=$?
    case "${_pm_saved_flags}" in *e*) set -e ;; esac
    case "${_pm_saved_flags}" in *u*) set -u ;; esac
    return "${_pm_rc}"
}

pm_perf_start_body() {
    pm_setup_tz
    PERF_INTERVAL="$(pm_uint_or "${PERF_INTERVAL}" 10)"
    PERF_CLI_INTERVAL="$(pm_uint_or "${PERF_CLI_INTERVAL}" "${PERF_INTERVAL}")"
    PERF_JSTAT_INTERVAL="$(pm_uint_or "${PERF_JSTAT_INTERVAL}" "${PERF_INTERVAL}")"
    PERF_JCMD_INTERVAL="$(pm_uint_or "${PERF_JCMD_INTERVAL}" 60)"
    PERF_TOP_INTERVAL="$(pm_uint_or "${PERF_TOP_INTERVAL}" "${PERF_INTERVAL}")"
    PERF_PS_INTERVAL="$(pm_uint_or "${PERF_PS_INTERVAL}" "${PERF_INTERVAL}")"
    PERF_CLI_TIMEOUT="$(pm_uint_or "${PERF_CLI_TIMEOUT}" 30)"
    PERF_WAIT_TIMEOUT="$(pm_uint_or "${PERF_WAIT_TIMEOUT}" 900)"
    pm_is_uint "${PERF_DURATION}" || PERF_DURATION=0
    pm_is_uint "${PERF_THREAD_DUMP_INTERVAL}" || PERF_THREAD_DUMP_INTERVAL=0
    pm_is_uint "${PERF_TOP_THREADS}" || PERF_TOP_THREADS=10

    if ! pm_resolve_output_dir; then
        pm_say "WARN: 出力先 '${PM_OUT:-(決められません)}' に書き込めないため、性能試験用メトリクス収集を行いません (JBoss は起動します)" >&2
        pm_say "WARN: JBOSS_LOG_DIR='${JBOSS_LOG_DIR:-}' EFS_LOG_DIR='${EFS_LOG_DIR:-}' PERF_OUTPUT_DIR='${PERF_OUTPUT_DIR}'" >&2
        return 1
    fi
    if ! pm_resolve_work_dir; then
        pm_say "WARN: 作業ディレクトリ (/tmp・standalone/tmp・出力先/.work) に書き込めないため、収集を行いません" >&2
        return 1
    fi
    rm -f "${PM_W}"/cli.in "${PM_W}"/cli.out "${PM_W}"/*.pid "${PM_W}/perf-metrics.state" 2>/dev/null
    PM_LOG_FILE="${PM_OUT}/collector.log"

    pm_now
    PM_START_EPOCH="${PM_EPOCH}"
    PM_END_EPOCH=0
    [ "${PERF_DURATION}" -gt 0 ] && PM_END_EPOCH=$(( PM_START_EPOCH + PERF_DURATION ))
    PM_SELF="${PERF_METRICS_SCRIPT:-${PM_SELF}}"

    # 収集プロセスへ渡す値
    export PM_OUT PM_W PM_LOG_FILE PM_START_EPOCH PM_END_EPOCH PM_SELF TZ \
        PERF_INTERVAL PERF_CLI_INTERVAL PERF_JSTAT_INTERVAL PERF_JCMD_INTERVAL PERF_TOP_INTERVAL \
        PERF_PS_INTERVAL PERF_COLLECTORS PERF_DURATION PERF_TZ PERF_CLI_ENABLE_DS_STATS PERF_CLI_MODE \
        PERF_CLI_CONTROLLER PERF_CLI_USER PERF_CLI_PASSWORD PERF_CLI_TIMEOUT PERF_CLI_CONNECT_TIMEOUT \
        PERF_CLI_HEAP PERF_CLI_JAVA_OPTS PERF_JSTAT_OPTIONS PERF_JCMD_COMMANDS PERF_THREAD_DUMP_INTERVAL \
        PERF_TOP_THREADS PERF_WAIT_TIMEOUT PERF_TOOL_JAVA_OPTS PERF_ENABLE_STATISTICS

    # 収集条件 (レポートの「収集情報」に出る)
    pm_meta perf version "${PERF_METRICS_VERSION}"
    pm_meta perf start_time "${PM_TS}"
    pm_meta perf start_epoch "${PM_EPOCH}"
    pm_meta perf utc_offset "$(date +%z)"
    pm_meta perf output_dir "${PM_OUT}"
    pm_meta perf work_dir "${PM_W}"
    pm_meta perf collectors "${PERF_COLLECTORS}"
    pm_meta perf cli_interval "${PERF_CLI_INTERVAL}"
    pm_meta perf jstat_interval "${PERF_JSTAT_INTERVAL}"
    pm_meta perf jcmd_interval "${PERF_JCMD_INTERVAL}"
    pm_meta perf top_interval "${PERF_TOP_INTERVAL}"
    pm_meta perf ps_interval "${PERF_PS_INTERVAL}"
    pm_meta perf duration "${PERF_DURATION}"
    pm_meta perf jstat_options "${PERF_JSTAT_OPTIONS}"
    pm_meta perf jcmd_commands "${PERF_JCMD_COMMANDS}"
    pm_meta perf top_threads "${PERF_TOP_THREADS}"
    pm_meta perf thread_dump_interval "${PERF_THREAD_DUMP_INTERVAL}"
    pm_meta perf statistics_opts "${PERF_SERVER_OPTS}"
    pm_meta container hostname "$(cat /proc/sys/kernel/hostname 2>/dev/null)"
    pm_meta container service_name "${Service_Name:-}"
    pm_meta container component_name "${Component_name:-}"
    pm_meta container component_role "${COMPONENT_ROLE:-}"
    pm_meta container log_id "$(basename "$(dirname "${PM_OUT}")")"
    pm_meta container jboss_home "${JBOSS_HOME:-}"
    pm_meta container cpus_online "$(getconf _NPROCESSORS_ONLN 2>/dev/null)"
    pm_meta container mem_total_kb "$(sed -n 's/^MemTotal:[[:space:]]*\([0-9]*\).*/\1/p' /proc/meminfo 2>/dev/null)"
    pm_meta container os "$(sed -n 's/^PRETTY_NAME="\{0,1\}\([^"]*\)"\{0,1\}$/\1/p' /etc/os-release 2>/dev/null)"
    pm_meta container command "$(printf '%s ' "$@" | sed "${PM_MASK_SED}")"
    pm_meta_ecs

    {
        echo "PM_OUT='${PM_OUT}'"
        echo "PM_W='${PM_W}'"
        echo "PM_LOG_FILE='${PM_LOG_FILE}'"
        echo "PM_START_TS='${PM_TS}'"
    } > "${PM_W}/perf-metrics.state" 2>/dev/null

    # 有効であることを起動ログで分かるようにする
    _pm_dur="コンテナ停止まで"
    [ "${PERF_DURATION}" -gt 0 ] && _pm_dur="${PERF_DURATION} 秒"
    _pm_thr="スレッド別 CPU なし"
    [ "${PERF_TOP_THREADS}" -gt 0 ] && _pm_thr="JBoss のスレッド別 CPU 上位 ${PERF_TOP_THREADS} 件"
    _pm_off() { pm_has_collector "$1" || printf ' (無効)'; }
    pm_banner_line "======================================================================"
    pm_banner_line " 性能試験用メトリクス収集: 有効  (perf-metrics.sh ${PERF_METRICS_VERSION} / PERF_METRICS_ENABLED=true)"
    pm_banner_line " 出力先 : ${PM_OUT}"
    pm_banner_line "          (${PM_OUT_NOTE}。最後に起動したタスクは mid/current/perf から辿れます)"
    pm_banner_line " ① jboss-cli : ${PERF_CLI_INTERVAL} 秒間隔$(_pm_off cli)  DB 接続プール / Undertow / IO ワーカー / スレッドプール / トランザクション / JVM"
    pm_banner_line " ② jstat     : ${PERF_JSTAT_INTERVAL} 秒間隔$(_pm_off jstat)  (-$(printf '%s' "${PERF_JSTAT_OPTIONS}" | sed 's/ / -/g'))   jcmd: ${PERF_JCMD_INTERVAL} 秒間隔$(_pm_off jcmd)  (${PERF_JCMD_COMMANDS})"
    pm_banner_line " ③ top / ps  : ${PERF_TOP_INTERVAL} / ${PERF_PS_INTERVAL} 秒間隔$(_pm_off top)$(_pm_off ps)  (${_pm_thr}、cgroup)"
    if [ -n "${PERF_SERVER_OPTS}" ]; then
        pm_banner_line " 統計の有効化: ${PERF_SERVER_OPTS} を standalone.sh へ追加"
    fi
    pm_banner_line " 収集時間 : ${_pm_dur}   作業ディレクトリ: ${PM_W}"
    pm_banner_line " 状態確認 : perf-metrics.sh status   (ECS Exec で実行)"
    pm_banner_line " ※ 性能試験用のイメージです。本番は PERF_METRICS_ENABLED=false でビルドしてください"
    pm_banner_line "======================================================================"

    # 前提の確認 (足りないものは警告し、その収集だけ止める)
    if pm_has_collector top || pm_has_collector ps; then
        for _pm_c in top ps; do
            if pm_has_collector "${_pm_c}" && ! command -v "${_pm_c}" >/dev/null 2>&1; then
                pm_say "WARN: ${_pm_c} がありません (procps-ng 未導入)。③ の ${_pm_c} 取得は行いません" >&2
            fi
        done
    fi
    if pm_has_collector jstat || pm_has_collector jcmd; then
        if ! pm_is_writable /tmp; then
            pm_say "WARN: /tmp に書き込めないため jstat / jcmd は JVM に接続できません (hsperfdata・アタッチ用ソケットを /tmp に作るため)。" >&2
            pm_say "WARN: タスク定義で /tmp に書き込み可能ボリュームを当ててください。ヒープ・GC は ① (cli_jvm.csv / cli_gc.csv) で代替できます" >&2
        fi
    fi
    if pm_has_collector cli && [ ! -f "${JBOSS_HOME:-/opt/jboss-eap}/bin/jboss-cli.sh" ]; then
        pm_say "WARN: ${JBOSS_HOME:-/opt/jboss-eap}/bin/jboss-cli.sh がありません。① の取得は行いません" >&2
    fi

    # 収集プロセスを起動 (JBoss の起動とは並行して動き、JVM の出現を待ってから取得する)
    for _pm_c in ${PERF_COLLECTORS}; do
        case "${_pm_c}" in
            cli|jstat|jcmd|top|ps) ;;
            *) pm_say "WARN: PERF_COLLECTORS の不明な値を無視します: ${_pm_c}" >&2; continue ;;
        esac
        /bin/sh "${PM_SELF}" run "${_pm_c}" < /dev/null > /dev/null &
        echo "$!" > "${PM_W}/${_pm_c}.pid" 2>/dev/null
        pm_log INFO "収集プロセスを起動しました: ${_pm_c} (pid=$!)"
    done
    return 0
}

# =============================================================================
# ① jboss-cli
# =============================================================================
# 取得するのは管理モデルの読み取りだけ。jboss-cli は常駐セッション (FIFO) で 1 つだけ
# 起動し、間隔ごとに JVM を起動し直すことはしない (起動のたびに CPU を 1 秒以上使うため)。
# 各操作は try / catch で囲み、無いサブシステム (ejb3 など) があっても他は取得する。

pm_cli_q() {  # id operation
    printf 'echo PERF_B %s\ntry\n%s\ncatch\necho PERF_F %s\nend-try\n' "$1" "$2" "$1"
}

pm_cli_quote() {
    case "$1" in
        *[!A-Za-z0-9_.-]*) printf '"%s"' "$(printf '%s' "$1" | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g')" ;;
        *) printf '%s' "$1" ;;
    esac
}

pm_cli_sample_block() {
    _pm_pm="/core-service=platform-mbean"
    for _pm_t in ds xa; do
        if [ "${_pm_t}" = ds ]; then _pm_r=data-source; else _pm_r=xa-data-source; fi
        pm_cli_q "dsp.${_pm_t}" "/subsystem=datasources/${_pm_r}=*/statistics=pool:read-resource(include-runtime=true)"
        pm_cli_q "dsj.${_pm_t}" "/subsystem=datasources/${_pm_r}=*/statistics=jdbc:read-resource(include-runtime=true)"
        pm_cli_q "dsc.${_pm_t}" "/subsystem=datasources/${_pm_r}=*:read-resource(resolve-expressions=true)"
    done
    pm_cli_q web.stat "/subsystem=undertow:read-attribute(name=statistics-enabled,resolve-expressions=true)"
    for _pm_t in http https ajp; do
        pm_cli_q "web.${_pm_t}" "/subsystem=undertow/server=*/${_pm_t}-listener=*:read-resource(include-runtime=true,resolve-expressions=true)"
    done
    pm_cli_q ses.dep "/deployment=*/subsystem=undertow:read-resource(include-runtime=true)"
    pm_cli_q ses.sub "/deployment=*/subdeployment=*/subsystem=undertow:read-resource(include-runtime=true)"
    pm_cli_q iow "/subsystem=io/worker=*:read-resource(include-runtime=true,resolve-expressions=true)"
    pm_cli_q tpe.ejb "/subsystem=ejb3/thread-pool=*:read-resource(include-runtime=true,resolve-expressions=true)"
    pm_cli_q tpe.ee "/subsystem=ee/managed-executor-service=*:read-resource(include-runtime=true,resolve-expressions=true)"
    pm_cli_q tpe.ees "/subsystem=ee/managed-scheduled-executor-service=*:read-resource(include-runtime=true,resolve-expressions=true)"
    pm_cli_q tx "/subsystem=transactions:read-resource(include-runtime=true,resolve-expressions=true)"
    pm_cli_q srv.state ":read-attribute(name=server-state)"
    pm_cli_q srv.suspend ":read-attribute(name=suspend-state)"
    pm_cli_q rc "/subsystem=request-controller:read-attribute(name=active-requests)"
    pm_cli_q mem "${_pm_pm}/type=memory:read-resource(include-runtime=true)"
    pm_cli_q thr.count "${_pm_pm}/type=threading:read-attribute(name=thread-count)"
    pm_cli_q thr.peak "${_pm_pm}/type=threading:read-attribute(name=peak-thread-count)"
    pm_cli_q thr.daemon "${_pm_pm}/type=threading:read-attribute(name=daemon-thread-count)"
    pm_cli_q cls "${_pm_pm}/type=class-loading:read-resource(include-runtime=true)"
    pm_cli_q os "${_pm_pm}/type=operating-system:read-resource(include-runtime=true)"
    pm_cli_q rt.uptime "${_pm_pm}/type=runtime:read-attribute(name=uptime)"
    pm_cli_q gc "${_pm_pm}/type=garbage-collector/name=*:read-resource(include-runtime=true)"
    # memory-pool はワイルドカードの read-resource ができない (WFLYPMB0010) ため属性ごとに読む
    pm_cli_q mp.usage "${_pm_pm}/type=memory-pool/name=*:read-attribute(name=usage)"
    pm_cli_q mp.peak "${_pm_pm}/type=memory-pool/name=*:read-attribute(name=peak-usage)"
    pm_cli_q mp.type "${_pm_pm}/type=memory-pool/name=*:read-attribute(name=type)"
}

pm_cli_meta_block() {
    _pm_pm="/core-service=platform-mbean"
    pm_cli_q m.product-name ":read-attribute(name=product-name)"
    pm_cli_q m.product-version ":read-attribute(name=product-version)"
    pm_cli_q m.release-version ":read-attribute(name=release-version)"
    pm_cli_q m.server-name ":read-attribute(name=name)"
    pm_cli_q m.launch-type ":read-attribute(name=launch-type)"
    pm_cli_q m.running-mode ":read-attribute(name=running-mode)"
    pm_cli_q m.config-file "/core-service=server-environment:read-attribute(name=config-file)"
    pm_cli_q m.node-name "/core-service=server-environment:read-attribute(name=node-name)"
    pm_cli_q m.vm-name "${_pm_pm}/type=runtime:read-attribute(name=vm-name)"
    pm_cli_q m.vm-vendor "${_pm_pm}/type=runtime:read-attribute(name=vm-vendor)"
    pm_cli_q m.vm-version "${_pm_pm}/type=runtime:read-attribute(name=vm-version)"
    pm_cli_q m.jvm-start-time "${_pm_pm}/type=runtime:read-attribute(name=start-time)"
    pm_cli_q m.available-processors "${_pm_pm}/type=operating-system:read-attribute(name=available-processors)"
}

# 解析用 awk (DMR 形式の応答を字句解析し、ワイルドカードの各要素ごとに CSV 行へ)
pm_cli_write_awk() {
    cat > "${PM_W}/cli.awk" <<'PERF_AWK_CLI'
function trim(s) { sub(/^[ \t\r\n]+/, "", s); sub(/[ \t\r\n]+$/, "", s); return s }
function csvq(s) { if (s ~ /[",\r\n]/) { gsub(/"/, "\"\"", s); return "\"" s "\"" } return s }
function norm(v) {
  v = trim(v)
  if (v == "undefined") return ""
  sub(/^expression /, "", v); sub(/^big (decimal|integer) /, "", v)
  if (v ~ /^-?[0-9]+L$/) v = substr(v, 1, length(v) - 1)
  return v
}
function header_all() {
  HD["cli_datasource.csv"] = "timestamp,epoch,ds_type,ds_name,jndi_name,enabled,statistics_enabled,min_pool_size,max_pool_size,blocking_timeout_ms," POOLC "," JDBCC
  HD["cli_undertow.csv"] = "timestamp,epoch,server,listener_type,listener,statistics_enabled," WEBC
  HD["cli_session.csv"] = "timestamp,epoch,deployment," SESC
  HD["cli_io_worker.csv"] = "timestamp,epoch,worker," IOWC
  HD["cli_threadpool.csv"] = "timestamp,epoch,pool_type,pool,active,current,largest,queue,completed,tasks,rejected,hung,max"
  HD["cli_transactions.csv"] = "timestamp,epoch,statistics_enabled," TXC
  HD["cli_jvm.csv"] = "timestamp,epoch,cli_ms,cli_ok,server_state,suspend_state,active_requests,uptime_ms,heap_used,heap_committed,heap_max,heap_init,nonheap_used,nonheap_committed,nonheap_max,thread_count,peak_thread_count,daemon_thread_count,loaded_class_count,total_loaded_class_count,unloaded_class_count,available_processors,system_load_average"
  HD["cli_gc.csv"] = "timestamp,epoch,collector,collection_count,collection_time_ms"
  HD["cli_mempool.csv"] = "timestamp,epoch,pool,type,used,committed,max,peak_used"
}
BEGIN {
  POOLC = "ActiveCount,AvailableCount,InUseCount,IdleCount,MaxUsedCount,WaitCount,MaxWaitCount,BlockingFailureCount,TimedOut,CreatedCount,DestroyedCount,AverageGetTime,MaxGetTime,TotalGetTime,AverageBlockingTime,MaxWaitTime,TotalBlockingTime,AverageUsageTime,MaxUsageTime,TotalUsageTime,AverageCreationTime,MaxCreationTime,TotalCreationTime,AveragePoolTime,MaxPoolTime,TotalPoolTime"
  JDBCC = "PreparedStatementCacheAccessCount,PreparedStatementCacheHitCount,PreparedStatementCacheMissCount,PreparedStatementCacheCurrentSize,PreparedStatementCacheAddCount,PreparedStatementCacheDeleteCount"
  WEBC = "request-count,error-count,bytes-received,bytes-sent,processing-time,max-processing-time,record-request-start-time"
  SESC = "context-root,active-sessions,sessions-created,expired-sessions,rejected-sessions,max-active-sessions,highest-session-count,session-avg-alive-time,session-max-alive-time"
  IOWC = "io-thread-count,busy-task-thread-count,core-pool-size,max-pool-size,queue-size,task-core-threads,task-max-threads"
  TXC = "number-of-transactions,number-of-committed-transactions,number-of-aborted-transactions,number-of-inflight-transactions,number-of-timed-out-transactions,number-of-heuristics,number-of-application-rollbacks,number-of-resource-rollbacks,number-of-system-rollbacks,number-of-nested-transactions,average-commit-time"
  header_all()
  if (mode == "header") {
    for (f in HD) {
      p = dir "/" f; has = 0
      if ((getline line < p) > 0) has = 1
      close(p)
      if (!has) { print HD[f] >> p; close(p) }
    }
    exit
  }
  D = 0; BLK = ""; NDS = 0
}
# ---- 字句解析 ----
function lex(line,    i, n, c, s, j, q) {
  NT = 0; n = length(line); i = 1
  while (i <= n) {
    c = substr(line, i, 1)
    if (c == " " || c == "\t" || c == "\r") { i++; continue }
    if (c == "\"") {
      s = ""; j = i + 1
      while (j <= n) {
        q = index(substr(line, j), "\"")
        if (q == 0) { s = s substr(line, j); j = n + 1; break }
        s = s substr(line, j, q - 1); j = j + q - 1
        if (substr(s, length(s), 1) == "\\" && substr(s, length(s) - 1, 1) != "\\") { s = substr(s, 1, length(s) - 1) "\""; j++; continue }
        break
      }
      gsub(/\\\\/, "\\", s)
      NT++; TT[NT] = "S"; TV[NT] = s; i = j + 1; continue
    }
    if (c == "=" && substr(line, i + 1, 1) == ">") { NT++; TT[NT] = "A"; TV[NT] = ""; i += 2; continue }
    if (index("{}[](),", c) > 0) { NT++; TT[NT] = c; TV[NT] = c; i++; continue }
    j = i
    while (j <= n && index(" \t\r{}[](),\"", substr(line, j, 1)) == 0) j++
    NT++; TT[NT] = "W"; TV[NT] = substr(line, i, j - i); i = j
  }
}
# ---- 構文の追跡 (スタック) ----
function path(d,    p, k) { p = ""; for (k = 2; k <= d; k++) p = p (p == "" ? "" : "/") FK[k]; return p }
function push(t, key) {
  D++; FT[D] = t; FK[D] = key; ST[D] = 0; VB = ""
  if (t == "{" && D == 3 && FK[2] == "result" && FT[2] == "[" && key == "#") { WILD = 1; split("", R); split("", A); EOUT = "" }
  if (D == 1) { WILD = 0; ROUT = ""; split("", R); split("", A) }
}
function pop(    t) {
  t = FT[D]
  if (t == "{" && D == 3 && FK[2] == "result" && FT[2] == "[" && FK[3] == "#") { if (EOUT == "success" || EOUT == "") emit() }
  if (D == 1 && !WILD && ROUT == "success") emit()
  D--
  if (D > 0 && (FT[D] == "{" || FT[D] == "(")) ST[D] = 3
}
function scalar(key, v,    full, p) {
  v = norm(v)
  if (FT[D] == "(") { p = path(D - 1); if (p ~ /\/address$/ || p == "address") A[key] = v; return }
  p = path(D); full = (p == "" ? key : p "/" key)
  if (full == "outcome") { ROUT = v; return }
  if (full == "result") { R["_value"] = v; return }
  if (full == "result/#/outcome") { EOUT = v; return }
  if (full == "result/#/result") { R["_value"] = v; return }
  if (substr(full, 1, 16) == "result/#/result/") { full = substr(full, 17); gsub(/\//, ".", full); R[full] = v; return }
  if (substr(full, 1, 7) == "result/" && substr(full, 1, 9) != "result/#/") { full = substr(full, 8); gsub(/\//, ".", full); R[full] = v; return }
}
function tok(t, v,    f) {
  if (D == 0) { if (t == "{") push("{", ""); return }
  f = FT[D]
  if (f == "{" || f == "(") {
    if (ST[D] == 0) {
      if (t == "S" || t == "W") { CK[D] = v; ST[D] = 1; return }
      if (t == "}" || t == ")") pop()
      return
    }
    if (ST[D] == 1) { if (t == "A") { ST[D] = 2; VB = ""; VN = 0 } return }
    if (ST[D] == 2) {
      if (t == "{" || t == "[" || t == "(") { if (VN == 0) { ST[D] = 3; push(t, CK[D]) } return }
      if (t == "S" || t == "W") { VB = (VN == 0 ? v : VB " " v); VN++; return }
      if (t == ",") { scalar(CK[D], VB); ST[D] = 0; return }
      if (t == "}" || t == ")") { scalar(CK[D], VB); pop(); return }
      return
    }
    if (ST[D] == 3) {
      if (t == ",") { ST[D] = 0; return }
      if (t == "}" || t == ")") pop()
      return
    }
    return
  }
  if (f == "[") {
    if (t == "]") { pop(); return }
    if (t == "{" || t == "[" || t == "(") { push(t, "#"); return }
    return
  }
}
# ---- 1 レコード (応答 1 つ、またはワイルドカードの要素 1 つ) を表へ ----
function addent(tab, ent) { if (!((tab, ent) in SEEN)) { SEEN[tab, ent] = 1; ORD[tab, ++NO[tab]] = ent } }
function emit(    k, ty, nm, ent, s, lt) {
  if (BLK ~ /^ds[pjc]\.(ds|xa)$/) {
    ty = (substr(BLK, 5) == "xa") ? "xa-data-source" : "data-source"
    nm = A[ty]; if (nm == "") return
    ent = substr(BLK, 5) SUBSEP nm; addent("ds", ent); s = substr(BLK, 3, 1)
    for (k in R) V["ds", ent, s, k] = R[k]
    return
  }
  if (BLK ~ /^web\.(http|https|ajp)$/) {
    lt = substr(BLK, 5); nm = A[lt "-listener"]; if (nm == "") return
    ent = A["server"] SUBSEP lt SUBSEP nm; addent("web", ent)
    for (k in R) V["web", ent, k] = R[k]
    return
  }
  if (BLK == "ses.dep" || BLK == "ses.sub") {
    nm = A["deployment"]; if (nm == "") return
    if (BLK == "ses.sub") nm = nm "/" A["subdeployment"]
    addent("ses", nm); for (k in R) V["ses", nm, k] = R[k]
    return
  }
  if (BLK == "iow") { nm = A["worker"]; if (nm == "") return; addent("iow", nm); for (k in R) V["iow", nm, k] = R[k]; return }
  if (BLK ~ /^tpe\./) {
    if (BLK == "tpe.ejb") { ty = "ejb3"; nm = A["thread-pool"] }
    else if (BLK == "tpe.ee") { ty = "ee-executor"; nm = A["managed-executor-service"] }
    else { ty = "ee-scheduled"; nm = A["managed-scheduled-executor-service"] }
    if (nm == "") return
    ent = ty SUBSEP nm; addent("tpe", ent); for (k in R) V["tpe", ent, k] = R[k]
    return
  }
  if (BLK == "gc") { nm = A["name"]; if (nm == "") return; addent("gc", nm); for (k in R) V["gc", nm, k] = R[k]; return }
  if (BLK ~ /^mp\./) {
    nm = A["name"]; if (nm == "") return; addent("mp", nm)
    if (BLK == "mp.type") { V["mp", nm, "type"] = R["_value"]; return }
    for (k in R) V["mp", nm, (BLK == "mp.usage" ? "usage." : "peak-usage.") k] = R[k]
    return
  }
  if (BLK ~ /^m\./) { if ("_value" in R) print "jboss," substr(BLK, 3) "," csvq(R["_value"]) >> metaf; return }
  for (k in R) J[BLK, k] = R[k]
}
function g(tab, ent, k) { return ((tab, ent, k) in V) ? V[tab, ent, k] : "" }
function jv(b, k) { return ((b, k) in J) ? J[b, k] : "" }
function cols(tab, ent, list, sub1,    n, a, i, s) {
  n = split(list, a, ","); s = ""
  for (i = 1; i <= n; i++) s = s "," csvq(sub1 == "" ? g(tab, ent, a[i]) : g(tab, ent SUBSEP sub1, a[i]))
  return s
}
/^PERF_B / { BLK = $2; D = 0; next }
/^PERF_F / { FAIL[$2] = 1; NFAIL++; FAILS = FAILS " " $2; next }
/^PERF_[SE] / { BLK = ""; D = 0; next }
BLK == "" { next }
{
  lex($0)
  for (ti = 1; ti <= NT; ti++) tok(TT[ti], TV[ti])
}
END {
  if (mode == "header") exit
  pre = csvq(ts) "," epoch
  if (mode == "meta") exit
  # ① データソース
  for (i = 1; i <= NO["ds"]; i++) {
    ent = ORD["ds", i]; split(ent, e, SUBSEP)
    st = g("ds", ent SUBSEP "p", "statistics-enabled")
    print pre "," e[1] "," csvq(e[2]) "," csvq(g("ds", ent SUBSEP "c", "jndi-name")) "," g("ds", ent SUBSEP "c", "enabled") "," st \
      "," g("ds", ent SUBSEP "c", "min-pool-size") "," g("ds", ent SUBSEP "c", "max-pool-size") "," g("ds", ent SUBSEP "c", "blocking-timeout-wait-millis") \
      cols("ds", ent, POOLC, "p") cols("ds", ent, JDBCC, "j") >> (dir "/cli_datasource.csv")
    if (enlist != "" && (st == "false" || g("ds", ent SUBSEP "j", "statistics-enabled") == "false")) {
      print (e[1] == "xa" ? "xa-data-source" : "data-source") "\t" e[2] > enlist
    }
  }
  # Undertow リスナー
  for (i = 1; i <= NO["web"]; i++) {
    ent = ORD["web", i]; split(ent, e, SUBSEP)
    print pre "," csvq(e[1]) "," e[2] "," csvq(e[3]) "," jv("web.stat", "_value") cols("web", ent, WEBC, "") >> (dir "/cli_undertow.csv")
  }
  for (i = 1; i <= NO["ses"]; i++) { ent = ORD["ses", i]; print pre "," csvq(ent) cols("ses", ent, SESC, "") >> (dir "/cli_session.csv") }
  for (i = 1; i <= NO["iow"]; i++) { ent = ORD["iow", i]; print pre "," csvq(ent) cols("iow", ent, IOWC, "") >> (dir "/cli_io_worker.csv") }
  for (i = 1; i <= NO["tpe"]; i++) {
    ent = ORD["tpe", i]; split(ent, e, SUBSEP)
    if (e[1] == "ejb3") {
      r = g("tpe", ent, "active-count") "," g("tpe", ent, "current-thread-count") "," g("tpe", ent, "largest-thread-count") "," g("tpe", ent, "queue-size") \
        "," g("tpe", ent, "completed-task-count") "," g("tpe", ent, "task-count") "," g("tpe", ent, "rejected-count") ",," g("tpe", ent, "max-threads")
    } else {
      r = g("tpe", ent, "active-thread-count") "," g("tpe", ent, "thread-count") "," g("tpe", ent, "largest-thread-count") "," g("tpe", ent, "current-queue-size") \
        "," g("tpe", ent, "completed-task-count") "," g("tpe", ent, "task-count") ",," g("tpe", ent, "hung-thread-count") "," g("tpe", ent, "max-threads")
    }
    print pre "," e[1] "," csvq(e[2]) "," r >> (dir "/cli_threadpool.csv")
  }
  if (("tx", "number-of-transactions") in J) {
    n = split(TXC, a, ","); s = ""
    for (k = 1; k <= n; k++) s = s "," jv("tx", a[k])
    print pre "," jv("tx", "statistics-enabled") s >> (dir "/cli_transactions.csv")
  }
  print pre "," cli_ms "," ok "," jv("srv.state", "_value") "," jv("srv.suspend", "_value") "," jv("rc", "_value") "," jv("rt.uptime", "_value") \
    "," jv("mem", "heap-memory-usage.used") "," jv("mem", "heap-memory-usage.committed") "," jv("mem", "heap-memory-usage.max") "," jv("mem", "heap-memory-usage.init") \
    "," jv("mem", "non-heap-memory-usage.used") "," jv("mem", "non-heap-memory-usage.committed") "," jv("mem", "non-heap-memory-usage.max") \
    "," jv("thr.count", "_value") "," jv("thr.peak", "_value") "," jv("thr.daemon", "_value") \
    "," jv("cls", "loaded-class-count") "," jv("cls", "total-loaded-class-count") "," jv("cls", "unloaded-class-count") \
    "," jv("os", "available-processors") "," jv("os", "system-load-average") >> (dir "/cli_jvm.csv")
  for (i = 1; i <= NO["gc"]; i++) { ent = ORD["gc", i]; print pre "," csvq(ent) "," g("gc", ent, "collection-count") "," g("gc", ent, "collection-time") >> (dir "/cli_gc.csv") }
  for (i = 1; i <= NO["mp"]; i++) {
    ent = ORD["mp", i]
    print pre "," csvq(ent) "," g("mp", ent, "type") "," g("mp", ent, "usage.used") "," g("mp", ent, "usage.committed") "," g("mp", ent, "usage.max") "," g("mp", ent, "peak-usage.used") >> (dir "/cli_mempool.csv")
  }
  if (failf != "") { print NFAIL + 0 "\t" FAILS > failf }
}
PERF_AWK_CLI
}

# jboss-cli.sh を、サーバーと同じ java・小さなヒープで起動する (サブシェルから exec)
pm_cli_exec() {  # cli-file
    unset JAVA_TOOL_OPTIONS _JAVA_OPTIONS JDK_JAVA_OPTIONS
    JAVA_OPTS="${PM_CLI_JAVA_OPTS}"
    export JAVA_OPTS JBOSS_HOME
    [ -n "${PM_CLI_JAVA:-}" ] && { JAVA="${PM_CLI_JAVA}"; export JAVA; }
    set -- --connect "--controller=${PM_CLI_CTRL}" "--timeout=15000" "--command-timeout=${PERF_CLI_TIMEOUT}" "--file=$1"
    [ -n "${PERF_CLI_USER}" ] && set -- "$@" "--user=${PERF_CLI_USER}"
    [ -n "${PERF_CLI_PASSWORD}" ] && set -- "$@" "--password=${PERF_CLI_PASSWORD}"
    exec "${PM_CLI}" "$@"
}

pm_cli_resolve() {
    JBOSS_HOME="${JBOSS_HOME:-}"
    # サーバーの起動引数から JBOSS_HOME・管理インタフェースのアドレス / ポートを読む
    _pm_addr=""; _pm_port=""; _pm_off=0; _pm_prev=""
    if [ -r "/proc/${PM_JPID}/cmdline" ]; then
        _pm_args="$(tr '\0' '\n' 2>/dev/null < "/proc/${PM_JPID}/cmdline")"
        _pm_oifs="${IFS}"; IFS='
'
        for _pm_a in ${_pm_args}; do
            case "${_pm_a}" in
                -Djboss.home.dir=*) [ -z "${JBOSS_HOME}" ] && JBOSS_HOME="${_pm_a#*=}" ;;
                -Djboss.bind.address.management=*|-bmanagement=*) _pm_addr="${_pm_a#*=}" ;;
                -Djboss.management.http.port=*) _pm_port="${_pm_a#*=}" ;;
                -Djboss.socket.binding.port-offset=*) _pm_off="${_pm_a#*=}" ;;
            esac
            [ "${_pm_prev}" = "-bmanagement" ] && _pm_addr="${_pm_a}"
            _pm_prev="${_pm_a}"
        done
        IFS="${_pm_oifs}"
    fi
    JBOSS_HOME="${JBOSS_HOME:-/opt/jboss-eap}"
    PM_CLI="${JBOSS_HOME}/bin/jboss-cli.sh"
    [ -f "${PM_CLI}" ] || return 1
    pm_is_uint "${_pm_port}" || _pm_port=9990
    pm_is_uint "${_pm_off}" || _pm_off=0
    case "${_pm_addr}" in ""|0.0.0.0|::|"[::]"|localhost) _pm_addr=127.0.0.1 ;; esac
    PM_CLI_CTRL="${PERF_CLI_CONTROLLER:-remote+http://${_pm_addr}:$(( _pm_port + _pm_off ))}"
    PM_CLI_JAVA="$(pm_java_of_jboss)" || PM_CLI_JAVA=""
    # 先頭の -Dperf.metrics.collector は目印 (ps / top で手動実行の jboss-cli と区別するため)
    PM_CLI_JAVA_OPTS="-Dperf.metrics.collector=jboss-cli ${PERF_CLI_JAVA_OPTS:--Xms16m -Xmx${PERF_CLI_HEAP} -XX:+UseSerialGC -XX:TieredStopAtLevel=1 -XX:-UsePerfData -Xss512k -XX:ReservedCodeCacheSize=32m -Djava.awt.headless=true}"
    return 0
}

pm_cli_alive() { [ -n "${PM_CLI_PID:-}" ] && kill -0 "${PM_CLI_PID}" 2>/dev/null; }

pm_cli_stop() {
    if [ "${PM_CLI_FD:-0}" = 1 ]; then
        printf 'quit\n' >&7 2>/dev/null
        exec 7>&-
        PM_CLI_FD=0
    fi
    if pm_cli_alive; then
        sleep 1
        pm_cli_alive && pm_kill_tree "${PM_CLI_PID}"
    fi
    [ -n "${PM_CLI_PID:-}" ] && wait "${PM_CLI_PID}" 2>/dev/null
    PM_CLI_PID=""
    rm -f "${PM_W}/cli.in" 2>/dev/null
    return 0
}

# 目印の行が出るまで待つ。0=出た 1=CLI が終了した 2=タイムアウト
pm_cli_wait() {  # marker timeout
    _pm_dl=$(( $(date +%s) + $2 ))
    while :; do
        grep -qx "$1" "${PM_W}/cli.out" 2>/dev/null && return 0
        if [ "${PM_CLI_MODE}" = persistent ]; then
            pm_cli_alive || return 1
        fi
        [ "$(date +%s)" -ge "${_pm_dl}" ] && return 2
        sleep 0.2 2>/dev/null || sleep 1
    done
}

pm_cli_start() {
    pm_cli_stop
    : > "${PM_W}/cli.out" || return 1
    [ "${PM_CLI_MODE}" = persistent ] || return 0
    if ! mkfifo -m 600 "${PM_W}/cli.in" 2>/dev/null; then
        pm_warn_once nofifo "FIFO を作れないため jboss-cli を毎回起動する方式 (oneshot) に切り替えます (${PM_W})"
        PM_CLI_MODE=oneshot
        return 0
    fi
    ( pm_cli_exec "${PM_W}/cli.in" ) >> "${PM_W}/cli.out" 2>> "${PM_W}/cli.err" < /dev/null &
    PM_CLI_PID=$!
    # FIFO を読み書き両用で開いておく (CLI が EOF を受けて終わらないように)
    exec 7<> "${PM_W}/cli.in" || { pm_cli_stop; return 1; }
    PM_CLI_FD=1
    printf 'echo PERF_READY\n' >&7
    pm_cli_wait PERF_READY "${PERF_CLI_CONNECT_TIMEOUT}"
    _pm_rc=$?
    if [ "${_pm_rc}" -ne 0 ]; then
        pm_cli_stop
        return 1
    fi
    pm_log INFO "jboss-cli の常駐セッションを開始しました (pid=${PM_CLI_PID}, ${PM_CLI_CTRL})"
    return 0
}

# 1 回分の問い合わせ。成功で 0。
pm_cli_exchange() {  # seq block-file timeout
    : > "${PM_W}/cli.out"
    if [ "${PM_CLI_MODE}" = persistent ]; then
        pm_cli_alive || return 1
        { printf 'echo PERF_S %s\n' "$1"; cat "$2"; printf 'echo PERF_E %s\n' "$1"; } >&7 2>/dev/null || return 1
        pm_cli_wait "PERF_E $1" "$3"
        _pm_rc=$?
        [ "${_pm_rc}" -eq 0 ] && return 0
        [ "${_pm_rc}" -eq 2 ] && pm_log WARN "jboss-cli が ${3} 秒以内に応答しません。セッションを作り直します"
        pm_cli_stop
        return 1
    fi
    { printf 'echo PERF_S %s\n' "$1"; cat "$2"; printf 'echo PERF_E %s\n' "$1"; } > "${PM_W}/oneshot.cli"
    ( pm_cli_exec "${PM_W}/oneshot.cli" ) >> "${PM_W}/cli.out" 2>> "${PM_W}/cli.err" < /dev/null &
    _pm_op=$!
    _pm_dl=$(( $(date +%s) + $3 + PERF_CLI_CONNECT_TIMEOUT ))
    while kill -0 "${_pm_op}" 2>/dev/null; do
        if [ "$(date +%s)" -ge "${_pm_dl}" ]; then pm_kill_tree "${_pm_op}"; break; fi
        sleep 0.2 2>/dev/null || sleep 1
    done
    wait "${_pm_op}" 2>/dev/null
    grep -qx "PERF_E $1" "${PM_W}/cli.out"
}

pm_run_cli() {
    PM_NAME=cli
    PM_CLI_MODE="${PERF_CLI_MODE}"
    case "${PM_CLI_MODE}" in persistent|oneshot) ;; *) PM_CLI_MODE=persistent ;; esac
    PM_CLI_FD=0; PM_CLI_PID=""
    trap 'pm_cli_stop; exit 0' TERM INT
    pm_wait_jboss || exit 0
    if ! pm_cli_resolve; then
        pm_log WARN "jboss-cli.sh が見つかりません (JBOSS_HOME=${JBOSS_HOME:-})。① の取得を終了します"
        exit 0
    fi
    pm_cli_write_awk
    awk -v mode=header -v dir="${PM_OUT}" -f "${PM_W}/cli.awk" /dev/null
    pm_cli_sample_block > "${PM_W}/sample.cli"
    pm_cli_meta_block > "${PM_W}/meta.cli"
    pm_log INFO "① jboss-cli: ${PERF_CLI_INTERVAL} 秒間隔 (${PM_CLI_MODE}, ${PM_CLI_CTRL}, JAVA_OPTS=${PM_CLI_JAVA_OPTS})"
    pm_meta cli controller "${PM_CLI_CTRL}"
    pm_meta cli mode "${PM_CLI_MODE}"
    _pm_seq=0; _pm_meta_done=0; _pm_fail_run=0; _pm_en_last=0; _pm_en_meta=0; _pm_conn_dl=$(( $(date +%s) + PERF_WAIT_TIMEOUT ))
    while :; do
        pm_ended && break
        if [ "${PM_CLI_MODE}" = persistent ] && ! pm_cli_alive; then
            if ! pm_cli_start; then
                _pm_fail_run=$((_pm_fail_run + 1))
                if [ "${_pm_meta_done}" = 0 ] && [ "$(date +%s)" -ge "${_pm_conn_dl}" ]; then
                    pm_log WARN "管理インタフェース (${PM_CLI_CTRL}) に ${PERF_WAIT_TIMEOUT} 秒たっても接続できません。詳細: ${PM_W}/cli.err"
                    _pm_conn_dl=$(( $(date +%s) + PERF_WAIT_TIMEOUT ))
                fi
                pm_jboss_alive || { PM_JPID="$(pm_find_jboss_pid)" || true; }
                sleep 5
                continue
            fi
        fi
        if [ "${_pm_meta_done}" = 0 ]; then
            _pm_seq=$((_pm_seq + 1))
            if pm_cli_exchange "${_pm_seq}" "${PM_W}/meta.cli" "${PERF_CLI_TIMEOUT}"; then
                awk -v mode=meta -v metaf="${PM_OUT}/perf_meta.csv" -v dir="${PM_OUT}" -f "${PM_W}/cli.awk" "${PM_W}/cli.out"
                _pm_meta_done=1
            elif [ "${PM_CLI_MODE}" = oneshot ]; then
                sleep 5
                continue
            fi
        fi
        pm_now
        _pm_seq=$((_pm_seq + 1))
        _pm_t0="$(pm_ms)"
        _pm_ok=0
        rm -f "${PM_W}/cli.enable" "${PM_W}/cli.fail"
        if pm_cli_exchange "${_pm_seq}" "${PM_W}/sample.cli" "${PERF_CLI_TIMEOUT}"; then _pm_ok=1; fi
        _pm_t1="$(pm_ms)"
        _pm_en=""
        [ "${PERF_CLI_ENABLE_DS_STATS}" = "true" ] && _pm_en="${PM_W}/cli.enable"
        awk -v ts="${PM_TS}" -v epoch="${PM_EPOCH}" -v cli_ms="$((_pm_t1 - _pm_t0))" -v ok="${_pm_ok}" \
            -v dir="${PM_OUT}" -v enlist="${_pm_en}" -v failf="${PM_W}/cli.fail" -v metaf="${PM_OUT}/perf_meta.csv" \
            -f "${PM_W}/cli.awk" "${PM_W}/cli.out" 2>> "${PM_W}/cli.err"
        if [ "${_pm_ok}" = 1 ]; then
            _pm_fail_run=0
        else
            _pm_fail_run=$((_pm_fail_run + 1))
            [ "${_pm_fail_run}" = 1 ] && pm_log WARN "① jboss-cli の取得に失敗しました (seq=${_pm_seq})。再接続して続けます。詳細: ${PM_W}/cli.err"
            pm_jboss_alive || { PM_JPID="$(pm_find_jboss_pid)" || true; }
        fi
        # プール統計が無効なデータソースは実行時に有効化する (60 秒に 1 回まで)
        if [ -s "${PM_W}/cli.enable" ] && [ $(( PM_EPOCH - _pm_en_last )) -ge 60 ]; then
            _pm_en_last="${PM_EPOCH}"
            : > "${PM_W}/admin.cli"
            while IFS="$(printf '\t')" read -r _pm_ty _pm_dn; do
                [ -n "${_pm_dn}" ] || continue
                _pm_ad="/subsystem=datasources/${_pm_ty}=$(pm_cli_quote "${_pm_dn}")"
                pm_cli_q "en.pool" "${_pm_ad}/statistics=pool:write-attribute(name=statistics-enabled,value=true)" >> "${PM_W}/admin.cli"
                pm_cli_q "en.jdbc" "${_pm_ad}/statistics=jdbc:write-attribute(name=statistics-enabled,value=true)" >> "${PM_W}/admin.cli"
                pm_log INFO "データソース ${_pm_dn} のプール統計を実行時に有効化します (設定ファイルは変えません)"
            done < "${PM_W}/cli.enable"
            _pm_seq=$((_pm_seq + 1))
            pm_cli_exchange "${_pm_seq}" "${PM_W}/admin.cli" "${PERF_CLI_TIMEOUT}" \
                || pm_log WARN "データソースの統計を有効化できませんでした"
            if [ "${_pm_en_meta}" = 0 ]; then
                pm_meta cli ds_stats_enabled_at_runtime "$(date '+%Y-%m-%d %H:%M:%S')"
                _pm_en_meta=1
            fi
        fi
        pm_sleep_tick "${PERF_CLI_INTERVAL}" || break
    done
    pm_cli_stop
    pm_log INFO "① jboss-cli の収集を終了しました"
}

# =============================================================================
# ② jstat / jcmd
# =============================================================================
# jstat は「間隔」を指定して 1 つだけ常駐させ、出力の各行に時刻を付けて CSV にする。
pm_jstat_csv() {  # csv-file
    _pm_f="$1"
    _pm_hd=0
    [ -s "${_pm_f}" ] && _pm_hd=1
    set -f
    while IFS= read -r _pm_l; do
        # shellcheck disable=SC2086
        set -- ${_pm_l}
        [ "$#" -gt 0 ] || continue
        if [ "$1" = Timestamp ]; then
            if [ "${_pm_hd}" = 0 ]; then
                shift
                _pm_oifs="${IFS}"; IFS=,
                printf 'timestamp,epoch,uptime_s,%s\n' "$*" >> "${_pm_f}"
                IFS="${_pm_oifs}"
                _pm_hd=1
            fi
            continue
        fi
        case "$1" in [0-9]*) ;; *) continue ;; esac
        pm_now
        _pm_oifs="${IFS}"; IFS=,
        printf '%s,%s,%s\n' "${PM_TS}" "${PM_EPOCH}" "$*" >> "${_pm_f}"
        IFS="${_pm_oifs}"
        pm_ended && break
    done
    set +f
}

pm_jstat_loop() {  # option
    _pm_opt="$1"
    while :; do
        pm_ended && return 0
        pm_jboss_alive || PM_JPID="$(pm_find_jboss_pid)" || { sleep 5; continue; }
        if ! pm_hsperf_exists "${PM_JPID}"; then
            pm_warn_once "hsperf${PM_JPID}" "② jstat: JVM の性能カウンタ (/tmp/hsperfdata_*/${PM_JPID}) がありません。/tmp が書き込めない、または -XX:-UsePerfData / -XX:+PerfDisableSharedMem が指定されています。ヒープ・GC は cli_jvm.csv / cli_gc.csv を参照してください"
            sleep 30
            continue
        fi
        pm_log INFO "② jstat -${_pm_opt} を開始します (pid=${PM_JPID}, ${PERF_JSTAT_INTERVAL} 秒間隔, $(pm_tool_desc))"
        PM_TOOL_TIMEOUT="$(pm_timeout_prefix)" pm_tool_exec "-${_pm_opt}" -t "${PM_JPID}" "${PERF_JSTAT_INTERVAL}s" 2>> "${PM_W}/jstat.err" \
            | pm_jstat_csv "${PM_OUT}/jvm_jstat_${_pm_opt}.csv"
        pm_ended && return 0
        pm_log INFO "② jstat -${_pm_opt} が終了しました。5 秒後に再開します (JVM の再起動・停止の可能性)"
        sleep 5
    done
}

pm_run_jstat() {
    PM_NAME=jstat
    trap 'exit 0' TERM INT
    pm_wait_jboss || exit 0
    if ! pm_resolve_tool jstat; then
        pm_log WARN "② jstat を起動できません (jstat が無く、java に jdk.jcmd モジュールもありません)"
        exit 0
    fi
    pm_meta jvm jstat_command "$(pm_tool_desc)"
    for _pm_o in ${PERF_JSTAT_OPTIONS}; do
        ( pm_jstat_loop "${_pm_o}" ) &
    done
    wait
}

# jcmd GC.heap_info を数値化する awk
pm_jcmd_write_awk() {
    cat > "${PM_W}/heap.awk" <<'PERF_AWK_HEAP'
function kb(s,    n, u) {
  n = s; sub(/[KMGkmgB]+$/, "", n); u = toupper(substr(s, length(n) + 1, 1))
  n = n + 0; if (u == "M") n *= 1024; else if (u == "G") n *= 1048576; else if (u == "B") n = n / 1024
  return int(n)
}
function num(line, word,    re) { re = word "[ ]+[0-9]+[KMGB]?"; if (match(line, re)) { s = substr(line, RSTART, RLENGTH); sub(word "[ ]+", "", s); return kb(s) } return "" }
BEGIN {
  H = "timestamp,epoch,gc,heap_total_kb,heap_committed_kb,heap_reserved_kb,heap_used_kb,region_size_kb,young_regions,young_kb,survivor_regions,survivor_kb,young_gen_total_kb,young_gen_used_kb,old_gen_total_kb,old_gen_used_kb,metaspace_used_kb,metaspace_committed_kb,metaspace_reserved_kb,class_used_kb,class_committed_kb,class_reserved_kb,jcmd_ms"
  if (mode == "header") { print H; exit }
}
/garbage-first heap|Shenandoah Heap|ZHeap|^ *ZGC/ {
  gc = ($0 ~ /garbage-first/) ? "G1" : ($0 ~ /Shenandoah/) ? "Shenandoah" : "Z"
  ht = num($0, "total"); hc = num($0, "committed"); hr = num($0, "reserved"); hu = num($0, "used")
  if (ht == "" && hc != "") ht = hc
  next
}
/region size/ {
  rs = num($0, "region size")
  if (match($0, /[0-9]+ young \([0-9]+[KMG]\)/)) { s = substr($0, RSTART, RLENGTH); split(s, a, /[ ()]+/); yr = a[1]; yk = kb(a[3]) }
  if (match($0, /[0-9]+ survivors \([0-9]+[KMG]\)/)) { s = substr($0, RSTART, RLENGTH); split(s, a, /[ ()]+/); sr = a[1]; sk = kb(a[3]) }
  next
}
/(def new generation|PSYoungGen|par new generation)/ { if (gc == "") gc = ($0 ~ /PSYoungGen/) ? "Parallel" : "Serial"; ygt = num($0, "total"); ygu = num($0, "used"); next }
/(tenured generation|ParOldGen)/ { ogt = num($0, "total"); ogu = num($0, "used"); next }
/^ *Metaspace/ { mu = num($0, "used"); mc = num($0, "committed"); mr = num($0, "reserved"); next }
/^ *class space/ { cu = num($0, "used"); cc = num($0, "committed"); cr = num($0, "reserved"); next }
END {
  if (mode == "header") exit
  if (ht == "" && ygt != "") { ht = ygt + ogt; hu = ygu + ogu }
  if (ht == "" && mu == "") exit
  print ts "," epoch "," gc "," ht "," hc "," hr "," hu "," rs "," yr "," yk "," sr "," sk "," ygt "," ygu "," ogt "," ogu "," mu "," mc "," mr "," cu "," cc "," cr "," ms
}
PERF_AWK_HEAP
}

pm_jcmd_run() {  # out-file command...
    _pm_of="$1"; shift
    # shellcheck disable=SC2086
    pm_tool_exec "${PM_JPID}" "$@" > "${_pm_of}" 2>> "${PM_W}/jcmd.err"
}

pm_threaddump() {
    mkdir -p "${PM_OUT}/jvm_threaddump" 2>/dev/null
    _pm_td="${PM_OUT}/jvm_threaddump/threaddump_$(date '+%Y%m%d_%H%M%S').txt"
    if pm_jcmd_run "${PM_W}/td.tmp" Thread.print -l; then
        { echo "# $(date '+%Y-%m-%d %H:%M:%S') jcmd ${PM_JPID} Thread.print -l"; cat "${PM_W}/td.tmp"; } > "${_pm_td}"
        pm_log INFO "スレッドダンプを取得しました: ${_pm_td}"
        echo "${_pm_td}"
    else
        pm_log WARN "スレッドダンプを取得できませんでした (詳細: ${PM_W}/jcmd.err)"
        return 1
    fi
}

pm_run_jcmd() {
    PM_NAME=jcmd
    trap 'exit 0' TERM INT
    pm_wait_jboss || exit 0
    if ! pm_resolve_tool jcmd; then
        pm_log WARN "② jcmd を起動できません (jcmd が無く、java に jdk.jcmd モジュールもありません)"
        exit 0
    fi
    pm_jcmd_write_awk
    pm_csv_init "${PM_OUT}/jvm_jcmd_heap.csv" "$(awk -v mode=header -f "${PM_W}/heap.awk" /dev/null)"
    pm_meta jvm jcmd_command "$(pm_tool_desc)"
    pm_log INFO "② jcmd: ${PERF_JCMD_INTERVAL} 秒間隔 (${PERF_JCMD_COMMANDS}, $(pm_tool_desc))"
    _pm_info=0; _pm_td_last=0; _pm_first=1
    while :; do
        pm_ended && break
        pm_jboss_alive || PM_JPID="$(pm_find_jboss_pid)" || { sleep 5; continue; }
        if [ "${_pm_info}" = 0 ]; then
            # JVM の版とフラグ (起動直後はアタッチできないことがあるので成功するまで繰り返す)
            if pm_jcmd_run "${PM_W}/jcmd.tmp" VM.version; then
                {
                    echo "# $(date '+%Y-%m-%d %H:%M:%S') jcmd ${PM_JPID} VM.version"; cat "${PM_W}/jcmd.tmp"
                    pm_jcmd_run "${PM_W}/jcmd.tmp2" VM.flags && { echo "# jcmd ${PM_JPID} VM.flags"; sed "${PM_MASK_SED}" "${PM_W}/jcmd.tmp2"; }
                } > "${PM_OUT}/jvm_info.txt"
                pm_meta jvm version "$(sed -n '2p' "${PM_W}/jcmd.tmp")"
                _pm_fl="$(cat "${PM_W}/jcmd.tmp2" 2>/dev/null)"
                for _pm_k in MaxHeapSize InitialHeapSize UseG1GC UseParallelGC UseSerialGC UseZGC UseShenandoahGC MaxGCPauseMillis ActiveProcessorCount MaxRAMPercentage; do
                    _pm_v="$(printf '%s\n' "${_pm_fl}" | tr ' ' '\n' | sed -n "s/^-XX:\\([+-]\\{0,1\\}\\)${_pm_k}\\(=\\(.*\\)\\)\\{0,1\\}$/\\1|\\3/p" | tail -n 1)"
                    case "${_pm_v}" in
                        "") ;;
                        "+|") pm_meta jvm_flag "${_pm_k}" true ;;
                        "-|") pm_meta jvm_flag "${_pm_k}" false ;;
                        *) pm_meta jvm_flag "${_pm_k}" "${_pm_v#*|}" ;;
                    esac
                done
                _pm_info=1
            else
                [ "${_pm_first}" = 1 ] && pm_log INFO "② jcmd: JVM へまだアタッチできません。続けて試します (詳細: ${PM_W}/jcmd.err)"
                _pm_first=0
                if ! pm_is_writable /tmp; then
                    pm_warn_once jcmdtmp "② jcmd: /tmp に書き込めないため JVM へアタッチできません (タスク定義で /tmp にボリュームを当ててください)"
                fi
            fi
        fi
        if [ "${_pm_info}" = 1 ]; then
            pm_now
            _pm_oifs="${IFS}"; IFS=,
            set -f
            # shellcheck disable=SC2086
            set -- ${PERF_JCMD_COMMANDS}
            set +f
            IFS="${_pm_oifs}"
            for _pm_c in "$@"; do
                _pm_t0="$(pm_ms)"
                set -f
                # shellcheck disable=SC2086
                if pm_jcmd_run "${PM_W}/jcmd.tmp" ${_pm_c}; then
                    _pm_t1="$(pm_ms)"
                    { echo "===== ${PM_TS} jcmd ${PM_JPID} ${_pm_c} ($((_pm_t1 - _pm_t0)) ms)"; cat "${PM_W}/jcmd.tmp"; } >> "${PM_OUT}/jvm_jcmd.log"
                    case "${_pm_c}" in
                        GC.heap_info)
                            awk -v ts="${PM_TS}" -v epoch="${PM_EPOCH}" -v ms="$((_pm_t1 - _pm_t0))" -f "${PM_W}/heap.awk" "${PM_W}/jcmd.tmp" >> "${PM_OUT}/jvm_jcmd_heap.csv"
                            ;;
                    esac
                else
                    pm_warn_once "jcmd${_pm_c}" "② jcmd ${PM_JPID} ${_pm_c} に失敗しました (詳細: ${PM_W}/jcmd.err)"
                fi
                set +f
            done
            if [ "${PERF_THREAD_DUMP_INTERVAL}" -gt 0 ] && [ $(( PM_EPOCH - _pm_td_last )) -ge "${PERF_THREAD_DUMP_INTERVAL}" ]; then
                _pm_td_last="${PM_EPOCH}"
                pm_threaddump > /dev/null
            fi
        fi
        if [ "${_pm_info}" = 0 ]; then
            sleep 10
            continue
        fi
        pm_sleep_tick "${PERF_JCMD_INTERVAL}" || break
    done
    pm_log INFO "② jcmd の収集を終了しました"
}

# =============================================================================
# ③ top / ps / cgroup
# =============================================================================
pm_role_awk='
function role(c) {
  if (c ~ /perf\.metrics\.collector=jboss-cli/) return "jboss-cli"
  if (c ~ /perf\.metrics\.collector=jstat/) return "jstat"
  if (c ~ /perf\.metrics\.collector=jcmd/) return "jcmd"
  if (c ~ /org\.jboss\.as\.cli|jboss-cli/) return "cli-other"
  if (c ~ /\[Standalone\]|org\.jboss\.as\.standalone/) return "jboss"
  if (c ~ /standalone\.sh/) return "standalone.sh"
  if (c ~ /sun\.tools\.(jstat|jcmd)|(^|\/)(jstat|jcmd)( |$)/) return "jdk-tool"
  if (c ~ /perf-metrics/) return "perf-metrics"
  if (c ~ /^(\/usr\/bin\/)?timeout [0-9]+ /) return "perf-metrics"
  if (c ~ /^(\/usr\/bin\/)?top( |$)/) return "top"
  if (c ~ /^(\/usr\/bin\/)?ps( |$)/) return "ps"
  return "other"
}
function mask(s,    o, k) {
  o = ""
  while (match(s, /[Pp][Aa][Ss][Ss][Ww][Oo][Rr][Dd][A-Za-z0-9_.-]*[=:][^ ,]*/)) {
    k = substr(s, RSTART, RLENGTH); sub(/[=:].*/, "", k)
    o = o substr(s, 1, RSTART - 1) k substr(s, RSTART + length(k), 1) "********"
    s = substr(s, RSTART + RLENGTH)
  }
  return o s
}
function csvq(s) { if (s ~ /[",\r\n]/) { gsub(/"/, "\"\"", s); return "\"" s "\"" } return s }
'

pm_top_write_awk() {
    cat > "${PM_W}/top.awk" <<PERF_AWK_TOP
${pm_role_awk}
PERF_AWK_TOP
    cat >> "${PM_W}/top.awk" <<'PERF_AWK_TOP'
function fullrole(pid, cmd,    r, c, l) {
  r = role(cmd)
  if (r != "other" || length(cmd) < 200) return r
  if (pid in RC) return RC[pid]
  c = "tr \"\\000\" \" \" < /proc/" pid "/cmdline 2>/dev/null"; l = ""; c | getline l; close(c)
  RC[pid] = (l == "") ? r : role(l)
  return RC[pid]
}
function getn(s, w,    re, t) { re = "[0-9.]+[ ]+" w; if (match(s, re)) { t = substr(s, RSTART, RLENGTH); sub(/[ ]+.*/, "", t); return t } return "" }
function tokb(v,    u, n) { u = substr(v, length(v)); if (u ~ /[0-9]/) return v; n = substr(v, 1, length(v) - 1) + 0
  if (u == "m") n *= 1024; else if (u == "g") n *= 1048576; else if (u == "t") n *= 1073741824; else if (u == "p") n *= 1099511627776
  return sprintf("%.0f", n) }
function stamp(    cmd, l) { cmd = "date \"+%Y-%m-%d %H:%M:%S|%s\""; cmd | getline l; close(cmd); split(l, T, "|"); TS = T[1]; EP = T[2] }
function flush_sum() {
  if (frame > 1 && TS != "")
    print TS "," EP "," L1 "," L5 "," L15 "," TK["total"] "," TK["running"] "," TK["sleeping"] "," TK["stopped"] "," TK["zombie"] \
      "," C["us"] "," C["sy"] "," C["ni"] "," C["id"] "," C["wa"] "," C["hi"] "," C["si"] "," C["st"] \
      "," M["total"] "," M["free"] "," M["used"] "," M["buff"] "," S["total"] "," S["free"] "," S["used"] "," M["avail"] >> sumf
  if (frame > 1 && TS != "") fflush(sumf)
}
BEGIN {
  if (mode == "header") {
    print "timestamp,epoch,load1,load5,load15,tasks_total,tasks_running,tasks_sleeping,tasks_stopped,tasks_zombie,cpu_us,cpu_sy,cpu_ni,cpu_id,cpu_wa,cpu_hi,cpu_si,cpu_st,mem_total_mib,mem_free_mib,mem_used_mib,mem_buffcache_mib,swap_total_mib,swap_free_mib,swap_used_mib,mem_avail_mib"
    print "timestamp,epoch,pid,role,user,pr,ni,virt_kb,res_kb,shr_kb,state,cpu_pct,mem_pct,time,command"
    exit
  }
  frame = 0
}
{ line = mask($0); print line >> rawf }
/^top - / {
  flush_sum(); frame++; inproc = 0; stamp()
  split("", TK); split("", C); split("", M); split("", S)
  s = $0; sub(/.*load average: */, "", s); n = split(s, la, /, */); L1 = la[1]; L5 = la[2]; L15 = la[3]
  fflush(rawf); next
}
/^Tasks:/ { TK["total"] = getn($0, "total"); TK["running"] = getn($0, "running"); TK["sleeping"] = getn($0, "sleeping"); TK["stopped"] = getn($0, "stopped"); TK["zombie"] = getn($0, "zombie"); next }
/^%Cpu/ { s = $0; sub(/^[^:]*:/, "", s); C["us"] = getn(s, "us"); C["sy"] = getn(s, "sy"); C["ni"] = getn(s, "ni"); C["id"] = getn(s, "id"); C["wa"] = getn(s, "wa"); C["hi"] = getn(s, "hi"); C["si"] = getn(s, "si"); C["st"] = getn(s, "st"); next }
/^[KMGTP]iB Mem/ {
  f = 1; u = substr($0, 1, 1); if (u == "K") f = 1 / 1024; else if (u == "G") f = 1024; else if (u == "T") f = 1048576
  M["total"] = getn($0, "total") * f; M["free"] = getn($0, "free") * f; M["used"] = getn($0, "used") * f; M["buff"] = getn($0, "buff/cache") * f; MF = f; next
}
/^[KMGTP]iB Swap/ {
  f = 1; u = substr($0, 1, 1); if (u == "K") f = 1 / 1024; else if (u == "G") f = 1024; else if (u == "T") f = 1048576
  S["total"] = getn($0, "total") * f; S["free"] = getn($0, "free") * f; S["used"] = getn($0, "used") * f; M["avail"] = getn($0, "avail") * f; next
}
/^ *PID +USER/ { inproc = 1; next }
inproc && NF >= 12 && $1 ~ /^[0-9]+$/ {
  if (frame < 2) next
  cmd = $0; for (k = 1; k <= 11; k++) sub(/^ *[^ ]+/, "", cmd); sub(/^ +/, "", cmd); cmd = mask(cmd)
  print TS "," EP "," $1 "," fullrole($1, cmd) "," csvq($2) "," $3 "," $4 "," tokb($5) "," tokb($6) "," tokb($7) "," $8 "," $9 "," $10 "," $11 "," csvq(cmd) >> procf
  fflush(procf); next
}
END { flush_sum() }
PERF_AWK_TOP
}

pm_topth_write_awk() {
    cat > "${PM_W}/topth.awk" <<'PERF_AWK_TOPTH'
function stamp(    cmd, l) { cmd = "date \"+%Y-%m-%d %H:%M:%S|%s\""; cmd | getline l; close(cmd); split(l, T, "|"); TS = T[1]; EP = T[2] }
function csvq(s) { if (s ~ /[",\r\n]/) { gsub(/"/, "\"\"", s); return "\"" s "\"" } return s }
BEGIN { if (mode == "header") { print "timestamp,epoch,rank,tid,thread,state,cpu_pct,time"; exit } frame = 0 }
/^top - / { frame++; inproc = 0; rank = 0; stamp(); next }
/^ *PID +USER/ { inproc = 1; next }
inproc && NF >= 12 && $1 ~ /^[0-9]+$/ {
  if (frame < 2 || rank >= top) next
  name = $0; for (k = 1; k <= 11; k++) sub(/^ *[^ ]+/, "", name); sub(/^ +/, "", name)
  rank++
  print TS "," EP "," rank "," $1 "," csvq(name) "," $8 "," $9 "," $11 >> outf
  fflush(outf)
}
PERF_AWK_TOPTH
}

pm_top_threads_loop() {
    PM_NAME=top-threads
    pm_wait_jboss || return 0
    pm_csv_init "${PM_OUT}/os_top_threads.csv" "$(awk -v mode=header -f "${PM_W}/topth.awk" /dev/null)"
    while :; do
        pm_ended && return 0
        pm_jboss_alive || PM_JPID="$(pm_find_jboss_pid)" || { sleep 5; continue; }
        pm_log INFO "③ top -H (JBoss pid=${PM_JPID} のスレッド別 CPU 上位 ${PERF_TOP_THREADS} 件) を開始します"
        # shellcheck disable=SC2086
        ( LC_ALL=C TERM=dumb exec $(pm_timeout_prefix) top -b -H -p "${PM_JPID}" -d "${PERF_TOP_INTERVAL}" ${PM_TOP_W} ) 2>> "${PM_W}/top.err" \
            | awk -v top="${PERF_TOP_THREADS}" -v outf="${PM_OUT}/os_top_threads.csv" -f "${PM_W}/topth.awk"
        pm_ended && return 0
        sleep 5
    done
}

pm_run_top() {
    PM_NAME=top
    trap 'exit 0' TERM INT
    if ! command -v top >/dev/null 2>&1; then
        pm_log WARN "③ top がありません (procps-ng 未導入)。top の取得を行いません"
        exit 0
    fi
    PM_TOP_W=""
    if LC_ALL=C top -b -n 1 -w 512 > /dev/null 2>&1; then PM_TOP_W="-w 512"; fi
    pm_top_write_awk
    pm_topth_write_awk
    awk -v mode=header -f "${PM_W}/top.awk" /dev/null > "${PM_W}/top.hdr"
    pm_csv_init "${PM_OUT}/os_top_summary.csv" "$(sed -n 1p "${PM_W}/top.hdr")"
    pm_csv_init "${PM_OUT}/os_top_process.csv" "$(sed -n 2p "${PM_W}/top.hdr")"
    pm_meta os top_version "$(top -v 2>&1 | head -n 1 | sed 's/^ *//')"
    if [ "${PERF_TOP_THREADS}" -gt 0 ]; then
        ( pm_top_threads_loop ) &
    fi
    pm_log INFO "③ top: ${PERF_TOP_INTERVAL} 秒間隔"
    while :; do
        pm_ended && break
        # shellcheck disable=SC2086
        ( LC_ALL=C TERM=dumb exec $(pm_timeout_prefix) top -b -c -d "${PERF_TOP_INTERVAL}" ${PM_TOP_W} ) 2>> "${PM_W}/top.err" \
            | awk -v rawf="${PM_OUT}/os_top.log" -v sumf="${PM_OUT}/os_top_summary.csv" -v procf="${PM_OUT}/os_top_process.csv" -f "${PM_W}/top.awk"
        pm_ended && break
        pm_log WARN "③ top が終了しました。5 秒後に再開します (詳細: ${PM_W}/top.err)"
        sleep 5
    done
}

# cgroup (v2 / v1) の CPU・メモリを 1 行の CSV で出す
pm_cgroup_row() {
    _pm_v=""; _pm_mc=""; _pm_ml=""; _pm_an=""; _pm_fi=""; _pm_oom=""; _pm_cu=""; _pm_np=""; _pm_nt=""; _pm_tu=""; _pm_q=""; _pm_pc=""
    if [ -f /sys/fs/cgroup/cgroup.controllers ]; then
        _pm_v=2; _pm_g=/sys/fs/cgroup
        _pm_mc="$(cat ${_pm_g}/memory.current 2>/dev/null)"
        _pm_ml="$(cat ${_pm_g}/memory.max 2>/dev/null)"; [ "${_pm_ml}" = max ] && _pm_ml=""
        _pm_an="$(awk '$1=="anon"{print $2}' ${_pm_g}/memory.stat 2>/dev/null)"
        _pm_fi="$(awk '$1=="file"{print $2}' ${_pm_g}/memory.stat 2>/dev/null)"
        _pm_oom="$(awk '$1=="oom_kill"{print $2}' ${_pm_g}/memory.events 2>/dev/null)"
        _pm_s="$(awk '$1=="usage_usec"{u=$2} $1=="nr_periods"{p=$2} $1=="nr_throttled"{t=$2} $1=="throttled_usec"{tu=$2} END{print u"|"p"|"t"|"tu}' ${_pm_g}/cpu.stat 2>/dev/null)"
        _pm_cu="${_pm_s%%|*}"; _pm_s="${_pm_s#*|}"; _pm_np="${_pm_s%%|*}"; _pm_s="${_pm_s#*|}"; _pm_nt="${_pm_s%%|*}"; _pm_tu="${_pm_s#*|}"
        _pm_q="$(awk '$1!="max"{printf "%.2f", $1/$2}' ${_pm_g}/cpu.max 2>/dev/null)"
        _pm_pc="$(cat ${_pm_g}/pids.current 2>/dev/null)"
    elif [ -d /sys/fs/cgroup/memory ]; then
        _pm_v=1; _pm_g=/sys/fs/cgroup/memory
        _pm_mc="$(cat ${_pm_g}/memory.usage_in_bytes 2>/dev/null)"
        _pm_ml="$(cat ${_pm_g}/memory.limit_in_bytes 2>/dev/null)"; [ "${#_pm_ml}" -ge 19 ] && _pm_ml=""
        _pm_an="$(awk '$1=="total_rss"{print $2}' ${_pm_g}/memory.stat 2>/dev/null)"
        _pm_fi="$(awk '$1=="total_cache"{print $2}' ${_pm_g}/memory.stat 2>/dev/null)"
        _pm_oom="$(awk '$1=="oom_kill"{print $2}' ${_pm_g}/memory.oom_control 2>/dev/null)"
        for _pm_g in /sys/fs/cgroup/cpuacct /sys/fs/cgroup/cpu,cpuacct /sys/fs/cgroup/cpuacct,cpu; do
            if [ -r "${_pm_g}/cpuacct.usage" ]; then _pm_cu=$(( $(cat "${_pm_g}/cpuacct.usage") / 1000 )); break; fi
        done
        for _pm_g in /sys/fs/cgroup/cpu /sys/fs/cgroup/cpu,cpuacct /sys/fs/cgroup/cpuacct,cpu; do
            if [ -r "${_pm_g}/cpu.stat" ]; then
                _pm_s="$(awk '$1=="nr_periods"{p=$2} $1=="nr_throttled"{t=$2} $1=="throttled_time"{tu=int($2/1000)} END{print p"|"t"|"tu}' "${_pm_g}/cpu.stat")"
                _pm_np="${_pm_s%%|*}"; _pm_s="${_pm_s#*|}"; _pm_nt="${_pm_s%%|*}"; _pm_tu="${_pm_s#*|}"
                _pm_q="$(awk -v q="$(cat "${_pm_g}/cpu.cfs_quota_us" 2>/dev/null)" -v p="$(cat "${_pm_g}/cpu.cfs_period_us" 2>/dev/null)" 'BEGIN{if (q > 0 && p > 0) printf "%.2f", q/p}')"
                break
            fi
        done
        _pm_pc="$(cat /sys/fs/cgroup/pids/pids.current 2>/dev/null)"
    fi
    printf '%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n' "${PM_TS}" "${PM_EPOCH}" "${_pm_v}" "${_pm_mc}" "${_pm_ml}" "${_pm_an}" "${_pm_fi}" \
        "${_pm_oom}" "${_pm_cu}" "${_pm_np}" "${_pm_nt}" "${_pm_tu}" "${_pm_q}" "${_pm_pc}"
}

pm_run_ps() {
    PM_NAME=ps
    trap 'exit 0' TERM INT
    pm_csv_init "${PM_OUT}/os_cgroup.csv" "timestamp,epoch,cgroup_version,memory_usage_bytes,memory_limit_bytes,memory_anon_bytes,memory_file_bytes,memory_oom_kill,cpu_usage_usec,cpu_nr_periods,cpu_nr_throttled,cpu_throttled_usec,cpu_quota_cores,pids_current"
    _pm_has_ps=1
    if ! command -v ps >/dev/null 2>&1; then
        pm_log WARN "③ ps がありません (procps-ng 未導入)。cgroup だけを記録します"
        _pm_has_ps=0
    else
        pm_csv_init "${PM_OUT}/os_ps.csv" "timestamp,epoch,pid,ppid,role,user,stat,nlwp,rss_kb,vsz_kb,pcpu,pmem,etimes,cpu_time_s,args"
        cat > "${PM_W}/ps.awk" <<PERF_AWK_PS
${pm_role_awk}
{
  args = \$0; for (k = 1; k <= 11; k++) sub(/^ *[^ ]+/, "", args); sub(/^ +/, "", args); args = mask(args)
  print ts "," epoch "," \$1 "," \$2 "," role(args) "," csvq(\$3) "," \$4 "," \$5 "," \$6 "," \$7 "," \$8 "," \$9 "," \$10 "," \$11 "," csvq(args)
}
PERF_AWK_PS
    fi
    pm_log INFO "③ ps / cgroup: ${PERF_PS_INTERVAL} 秒間隔"
    while :; do
        pm_ended && break
        pm_now
        if [ "${_pm_has_ps}" = 1 ]; then
            LC_ALL=C ps -e -ww -o pid=,ppid=,user=,stat=,nlwp=,rss=,vsz=,pcpu=,pmem=,etimes=,times=,args= 2>> "${PM_W}/ps.err" \
                | awk -v ts="${PM_TS}" -v epoch="${PM_EPOCH}" -f "${PM_W}/ps.awk" >> "${PM_OUT}/os_ps.csv"
        fi
        pm_cgroup_row >> "${PM_OUT}/os_cgroup.csv"
        pm_sleep_tick "${PERF_PS_INTERVAL}" || break
    done
    pm_log INFO "③ ps / cgroup の収集を終了しました"
}

# =============================================================================
# ビルド時の準備 / 運用コマンド
# =============================================================================
pm_install() {
    pm_say "ビルド時の準備 (perf-metrics.sh ${PERF_METRICS_VERSION})"
    _pm_need=""
    command -v ps > /dev/null 2>&1 || _pm_need="procps-ng"
    command -v top > /dev/null 2>&1 || _pm_need="procps-ng"
    command -v awk > /dev/null 2>&1 || _pm_need="${_pm_need} gawk"
    if [ -n "${_pm_need}" ]; then
        if [ "${PERF_INSTALL_PACKAGES}" != "true" ]; then
            pm_say "WARN: ${_pm_need} が必要ですが PERF_INSTALL_PACKAGES=${PERF_INSTALL_PACKAGES} のため導入しません"
        elif command -v microdnf > /dev/null 2>&1; then
            # shellcheck disable=SC2086
            microdnf install -y --nodocs ${_pm_need} && microdnf clean all \
                || pm_say "WARN: ${_pm_need} を導入できませんでした (③ top / ps は動きません)"
        elif command -v dnf > /dev/null 2>&1; then
            # shellcheck disable=SC2086
            dnf install -y --nodocs ${_pm_need} && dnf clean all \
                || pm_say "WARN: ${_pm_need} を導入できませんでした (③ top / ps は動きません)"
        else
            pm_say "WARN: パッケージマネージャーが無いため ${_pm_need} を導入できません"
        fi
        rm -rf /var/cache/yum /var/cache/dnf 2>/dev/null
    fi
    _pm_java="$(pm_java_of_jboss 2>/dev/null)" || _pm_java=""
    for _pm_t in ps top awk mkfifo curl; do
        if command -v "${_pm_t}" > /dev/null 2>&1; then pm_say "  ${_pm_t}: $(command -v "${_pm_t}")"; else pm_say "  ${_pm_t}: なし"; fi
    done
    pm_say "  java: ${_pm_java:-なし}"
    for _pm_t in jstat jcmd; do
        if pm_resolve_tool "${_pm_t}"; then pm_say "  ${_pm_t}: $(pm_tool_desc)"; else pm_say "  ${_pm_t}: なし (② の ${_pm_t} は動きません)"; fi
    done
    [ -f "${JBOSS_HOME:-/opt/jboss-eap}/bin/jboss-cli.sh" ] \
        && pm_say "  jboss-cli.sh: ${JBOSS_HOME:-/opt/jboss-eap}/bin/jboss-cli.sh" \
        || pm_say "  jboss-cli.sh: なし (JBoss EAP 導入前のイメージでは正常。① は実行時に JBOSS_HOME から探します)"
    pm_say "性能試験用メトリクス収集を有効にしました (PERF_METRICS_ENABLED=true)"
    return 0
}

pm_status() {
    if ! pm_load_state; then
        echo "perf-metrics: 実行中の収集が見つかりません (PERF_METRICS_ENABLED=${PERF_METRICS_ENABLED})"
        return 1
    fi
    echo "perf-metrics ${PERF_METRICS_VERSION}"
    echo "  出力先     : ${PM_OUT}"
    echo "  作業ディレクトリ: ${PM_W}"
    echo "  開始       : ${PM_START_TS}"
    for _pm_c in cli jstat jcmd top ps; do
        [ -r "${PM_W}/${_pm_c}.pid" ] || continue
        _pm_p="$(cat "${PM_W}/${_pm_c}.pid")"
        if kill -0 "${_pm_p}" 2>/dev/null; then _pm_s="実行中"; else _pm_s="停止"; fi
        echo "  収集 ${_pm_c}: ${_pm_s} (pid=${_pm_p})"
    done
    echo "  ファイル (行数 / 最終行の時刻):"
    for _pm_f in "${PM_OUT}"/*.csv; do
        [ -f "${_pm_f}" ] || continue
        printf '    %-24s %8s  %s\n' "$(basename "${_pm_f}")" "$(wc -l < "${_pm_f}" | tr -d ' ')" "$(tail -n 1 "${_pm_f}" | cut -d, -f1 | tr -d '"')"
    done
    echo "  直近のログ:"
    tail -n 5 "${PM_LOG_FILE}" 2>/dev/null | sed 's/^/    /'
    return 0
}

pm_stop() {
    pm_load_state || { echo "perf-metrics: 実行中の収集が見つかりません"; return 1; }
    for _pm_c in cli jstat jcmd top ps; do
        [ -r "${PM_W}/${_pm_c}.pid" ] || continue
        _pm_p="$(cat "${PM_W}/${_pm_c}.pid")"
        if kill -0 "${_pm_p}" 2>/dev/null; then
            pm_kill_tree "${_pm_p}"
            echo "停止しました: ${_pm_c} (pid=${_pm_p})"
        fi
    done
    PM_NAME=stop
    pm_log INFO "perf-metrics.sh stop で収集を停止しました"
    return 0
}

pm_pack() {
    pm_load_state || { echo "perf-metrics: 出力先が分かりません (PERF_OUTPUT_DIR を指定してください)"; return 1; }
    _pm_par="$(dirname "${PM_OUT}")"
    _pm_tgz="${_pm_par}/perf_$(basename "${_pm_par}")_$(date '+%Y%m%d_%H%M%S').tar"
    if command -v gzip > /dev/null 2>&1; then
        tar -C "${_pm_par}" -czf "${_pm_tgz}.gz" "$(basename "${PM_OUT}")" && echo "${_pm_tgz}.gz"
    else
        tar -C "${_pm_par}" -cf "${_pm_tgz}" "$(basename "${PM_OUT}")" && echo "${_pm_tgz}"
    fi
}

pm_main() {
    _pm_cmd="${1:-}"
    [ "$#" -gt 0 ] && shift
    PM_SELF="$0"
    case "${_pm_cmd}" in
        run)
            pm_setup_tz
            case "${1:-}" in
                cli) pm_run_cli ;;
                jstat) pm_run_jstat ;;
                jcmd) pm_run_jcmd ;;
                top) pm_run_top ;;
                ps) pm_run_ps ;;
                *) echo "usage: perf-metrics.sh run cli|jstat|jcmd|top|ps" >&2; exit 2 ;;
            esac
            ;;
        install) pm_install ;;
        status) pm_status ;;
        stop) pm_stop ;;
        pack) pm_pack ;;
        threaddump)
            pm_load_state || { echo "perf-metrics: 実行中の収集が見つかりません" >&2; exit 1; }
            PM_NAME=threaddump
            PM_JPID="$(pm_find_jboss_pid)" || { echo "JBoss EAP の JVM が見つかりません" >&2; exit 1; }
            pm_resolve_tool jcmd || { echo "jcmd がありません" >&2; exit 1; }
            pm_threaddump
            ;;
        start)
            perf_metrics_start "$@"
            ;;
        version) echo "perf-metrics.sh ${PERF_METRICS_VERSION}" ;;
        *)
            echo "usage: perf-metrics.sh status | stop | threaddump | pack | install | version" >&2
            exit 2
            ;;
    esac
}

# 直接実行されたときだけコマンドを処理する (エントリポイントから . で読み込まれたときは
# 関数の定義だけで終わる)
case "${0##*/}" in
    perf-metrics.sh|perf-metrics) pm_main "$@"; exit $? ;;
esac
:
