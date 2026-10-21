#!/bin/bash
# =============================================================================
# run_e2e.sh : perf-metrics.sh の結合試験 (Docker)
#
# readonlyRootFilesystem=true の ECS タスクに近い形 (--read-only + 書き込み可能ボリューム)
# で front イメージを起動し、検証用 WAR へ負荷をかけて perf/ の出力を確かめる。
#
#   1. eap-sim (UBI 9 OpenJDK 21 runtime + WildFly 35) → base → front (perf / prod) をビルド
#   2. --read-only で起動 (/tmp・standalone/{configuration,tmp,data} は tmpfs、
#      /mnt/logs・/mnt/data はホストのディレクトリ)
#   3. 起動ログに「性能試験用メトリクス収集: 有効」が出ること
#   4. WAR をデプロイし、並列で負荷をかける (接続プールの待ち・GC を起こす)
#   5. cli_datasource.csv・jvm_jstat_gc.csv に行が増え、接続待ちと GC が記録されること・
#      collector.log に ERROR が無いこと
#   6. PERF_METRICS_ENABLED=false のイメージでは収集が動かないこと
#
# 使い方 (Git Bash / WSL / Linux):
#   WILDFLY_TGZ_DIR=<wildfly-35.0.1.Final.tar.gz のあるディレクトリ> \
#   WORK=<出力を置くディレクトリ> bash test/e2e/run_e2e.sh [負荷をかける秒数 (既定 180)]
#   (RUNTIME_IMAGE=<イメージ> で eap-sim の土台を変えられる。既定は ubi9/openjdk-21-runtime)
# =============================================================================
set -u
export MSYS_NO_PATHCONV=1
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(cd "${HERE}/../.." && pwd)"
LOAD_SEC="${1:-180}"
: "${WILDFLY_TGZ_DIR:?WILDFLY_TGZ_DIR を指定してください}"
: "${WORK:?WORK (出力ディレクトリ) を指定してください}"
NAME=perf-e2e
PORT="${PORT:-18080}"
fail=0
ok()   { echo "  [OK]   $*"; }
ng()   { echo "  [NG]   $*"; fail=1; }

hostpath() { if command -v cygpath >/dev/null 2>&1; then cygpath -w "$1"; else printf '%s' "$1"; fi; }
PY="$(command -v python3 || command -v python)"

echo "== 1. ビルド"
docker build -q -t eap-sim:latest -f "$(hostpath "${HERE}/Dockerfile.eap-sim")" \
    --build-arg RUNTIME_IMAGE="${RUNTIME_IMAGE:-registry.access.redhat.com/ubi9/openjdk-21-runtime:latest}" \
    "$(hostpath "${WILDFLY_TGZ_DIR}")" >/dev/null || exit 1
docker build -q -t myapp-base:perftest --build-arg BASE_IMAGE=eap-sim:latest --build-arg STRICT_SEED=1 "$(hostpath "${ROOT}/docker/base")" >/dev/null || exit 1
for v in true false; do
    tag=perf; [ "$v" = false ] && tag=prod
    docker build -q -t "intra-web-front:${tag}" --build-arg BASE_IMAGE=myapp-base:perftest \
        --build-arg Service_Name=intra-web --build-arg Component_name=front \
        --build-arg PERF_METRICS_ENABLED="$v" "$(hostpath "${ROOT}/docker/front")" >/dev/null || exit 1
done
ok "eap-sim / myapp-base:perftest / intra-web-front:perf / intra-web-front:prod"

run_front() {  # tag extra-args...
    local tag=$1; shift
    docker rm -f "${NAME}" >/dev/null 2>&1
    rm -rf "${WORK}/mnt-logs" "${WORK}/mnt-data"; mkdir -p "${WORK}/mnt-logs" "${WORK}/mnt-data" "${WORK}/war"
    # Linux のホストではコンテナの uid 185 から書けるようにする (EFS のアクセスポイントの代わり)
    chmod 0777 "${WORK}/mnt-logs" "${WORK}/mnt-data"
    "${PY}" "$(hostpath "${HERE}/make_war.py")" "$(hostpath "${WORK}/war")" >/dev/null
    docker run -d --name "${NAME}" --read-only --user 185:0 --cpus 2 --memory 2g \
        --tmpfs /tmp:rw,exec,mode=1777 \
        --tmpfs /opt/jboss-eap/standalone/configuration:rw,uid=185,gid=0,mode=0775 \
        --tmpfs /opt/jboss-eap/standalone/tmp:rw,exec,uid=185,gid=0,mode=0775 \
        --tmpfs /opt/jboss-eap/standalone/data:rw,uid=185,gid=0,mode=0775 \
        -v "$(hostpath "${WORK}/mnt-logs"):/mnt/logs" -v "$(hostpath "${WORK}/mnt-data"):/mnt/data" \
        -v "$(hostpath "${WORK}/war"):/war:ro" \
        -e TZ=Asia/Tokyo -p "${PORT}:8080" "$@" "intra-web-front:${tag}" >/dev/null
    for _ in $(seq 1 60); do
        docker logs "${NAME}" 2>&1 | grep -q 'WFLYSRV0025' && return 0
        sleep 2
    done
    return 1
}

echo "== 2. 性能試験用イメージを --read-only で起動"
run_front perf -e PERF_INTERVAL=5 -e EXTRASLB_TRUSTSTORE_PASSWORD=e2e-secret || { ng "JBoss が起動しません"; docker logs "${NAME}" | tail -30; exit 1; }
ok "JBoss 起動"
docker logs "${NAME}" 2>&1 | grep -q '性能試験用メトリクス収集: 有効' && ok "起動ログに有効表示" || ng "起動ログに有効表示がない"
docker logs "${NAME}" 2>&1 | grep -q 'standalone.sh -Dwildfly.datasources.statistics-enabled=true' && ok "統計有効化の引数" || ng "-Dwildfly.datasources.statistics-enabled=true が付いていない"
docker logs "${NAME}" 2>&1 | grep -E '^awk: |perf-metrics.sh: [0-9]+:|syntax error' && ng "起動ログにエラー"

echo "== 3. WAR をデプロイして ${LOAD_SEC} 秒負荷をかける"
docker exec "${NAME}" /opt/jboss-eap/bin/jboss-cli.sh -c --command="deploy /war/perftest.war" >/dev/null || ng "deploy"
curl -sf "http://localhost:${PORT}/perftest/db.jsp?ms=1" >/dev/null && ok "WAR 応答" || ng "WAR が応答しない"
# 負荷は別コンテナから (ネットワークだけ共有するので、対象コンテナの CPU・メモリには混ざらない)。
# 40 並列 × 150ms 接続を保持 → ExampleDS の最大プール 20 を超えて接続待ちが起きる
docker run --rm --network "container:${NAME}" --entrypoint bash eap-sim:latest -c "
    end=\$(( \$(date +%s) + ${LOAD_SEC} ))
    for w in \$(seq 1 40); do
        ( while [ \$(date +%s) -lt \$end ]; do curl -s -o /dev/null 'http://127.0.0.1:8080/perftest/db.jsp?ms=150&kb=4096'; done ) &
    done
    wait" || ng "負荷をかけられません"
sleep 12

echo "== 4. 出力の確認"
OUT="$(ls -d "${WORK}"/mnt-logs/front/logs/intra-web/mid/2*/perf 2>/dev/null | tail -1)"
[ -n "${OUT}" ] && ok "出力先 ${OUT}" || { ng "perf/ がない"; exit 1; }
for f in cli_datasource jvm_jstat_gc; do
    n=$(( $(wc -l < "${OUT}/${f}.csv" 2>/dev/null || echo 1) - 1 ))
    [ "$n" -gt 3 ] && ok "${f}.csv: ${n} 行" || ng "${f}.csv: ${n} 行"
done
# 列は見出しの名前で探す (CSV の列の並びに依存しない)
col_max() {  # file column -> 列の最大値
    awk -F, -v k="$2" 'NR==1 {for (i = 1; i <= NF; i++) if ($i == k) c = i; next} c && $c+0 > m {m = $c+0} END {print m+0}' "$1"
}
[ "$(col_max "${OUT}/cli_datasource.csv" WaitCount)" -gt 0 ] && ok "接続待ち (WaitCount) を記録" || ng "WaitCount が 0 のまま"
[ "$(col_max "${OUT}/cli_datasource.csv" InUseCount)" -ge 20 ] && ok "使用中が最大プール (20) に到達" || ng "InUseCount が 20 に届かない"
[ "$(col_max "${OUT}/jvm_jstat_gc.csv" YGC)" -gt 0 ] && ok "GC 回数 (YGC) を記録" || ng "YGC が 0 のまま"
grep -q '^jvm,jstat_command,' "${OUT}/perf_meta.csv" && ok "jstat の起動方法: $(sed -n 's/^jvm,jstat_command,//p' "${OUT}/perf_meta.csv")" || ng "perf_meta.csv に jstat_command がない"
grep -q 'trustStorePassword=\*\*\*\*' "${OUT}/perf_meta.csv" && ! grep -q e2e-secret "${OUT}/perf_meta.csv" && ok "起動コマンドのパスワードを伏せ字" || ng "起動コマンドのパスワードが伏せられていない"
grep -q ' ERROR ' "${OUT}/collector.log" && ng "collector.log に ERROR" || ok "collector.log に ERROR なし"
docker exec "${NAME}" perf-metrics.sh status > "${WORK}/status.txt" 2>&1 && ok "perf-metrics.sh status" || ng "status"
docker exec "${NAME}" perf-metrics.sh stop > "${WORK}/stop.txt" 2>&1 && grep -q '停止しました: cli' "${WORK}/stop.txt" \
    && grep -q '停止しました: jstat' "${WORK}/stop.txt" && ok "perf-metrics.sh stop" || ng "stop"
rm -rf "${WORK}/perf_result"; cp -r "${OUT}" "${WORK}/perf_result" 2>/dev/null
docker logs "${NAME}" > "${WORK}/container.log" 2>&1

echo "== 5. PERF_METRICS_ENABLED=false のイメージ"
run_front prod -e PERF_METRICS_ENABLED=true || { ng "prod: JBoss が起動しません"; exit 1; }
docker logs "${NAME}" 2>&1 | grep -q 'perf-metrics.sh がありません' && ok "有効化を要求しても、入っていないことを警告" || ng "警告がない"
docker logs "${NAME}" 2>&1 | grep -q '性能試験用メトリクス収集: 有効' && ng "prod で収集が動いた" || ok "prod では収集しない"
ls -d "${WORK}"/mnt-logs/front/logs/intra-web/mid/2*/perf >/dev/null 2>&1 && ng "prod で perf/ ができた" || ok "prod では perf/ を作らない"
docker rm -f "${NAME}" >/dev/null 2>&1

echo
[ "${fail}" = 0 ] && echo "結果: すべて OK  (出力: ${WORK}/perf_result)" || echo "結果: NG あり"
exit "${fail}"
