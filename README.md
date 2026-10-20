# JMeter_Metrics_Script — 性能試験用メトリクス収集 (ECS / Fargate + JBoss EAP 8.1)

JMeter で性能試験をしている間、ECS タスク (Fargate) のコンテナ
(UBI 9.8 + Red Hat build of OpenJDK 21 + JBoss EAP 8.1) の中で次の統計を一定間隔で取り、
`readonlyRootFilesystem=true` でも書き込める EFS 上のタスク専用ディレクトリへ CSV で出力します。
出力は Windows 11 の PowerShell でモノトーンの Excel レポートにできます。

| 区分 | 取得方法 | 主な内容 |
|---|---|---|
| ① | jboss-cli (常駐セッション 1 つ) | DB 接続プール (使用中・待ち・取得失敗・取得時間・文キャッシュ)、HTTP リスナー (リクエスト数・エラー・転送量・処理時間)、HTTP セッション、IO ワーカー (リクエスト処理スレッド・キュー)、EJB / EE スレッドプール、トランザクション、JVM (ヒープ・メモリプール・スレッド・クラス・GC) |
| ② | jstat (常駐 1 つ) / jcmd | ヒープ各領域の容量・使用量、GC 回数・時間 (`jstat -gc`)、ヒープとメタスペースの詳細 (`jcmd GC.heap_info`)、JVM の版・フラグ、スレッドダンプ (任意) |
| ③ | top (常駐) / ps / cgroup | ロードアベレージ・CPU 内訳・メモリ、プロセス別 %CPU・RES・スレッド数、JBoss のスレッド別 %CPU、コンテナの CPU 使用率・スロットリング・メモリ・OOM |

- **性能試験用の実装は `docker/base/perf-metrics.sh` の 1 ファイル**にまとめています。
- **コンテナのビルド引数 `PERF_METRICS_ENABLED=true|false` で有効・無効を切り替えます**
  (既定 false。false ならファイル自体を削除し、本番と同じ中身になります)。
- 有効なイメージでは、エントリポイントが起動時にこのファイルを読み込み、起動ログに次の枠を出します。

```
[perf-metrics] ======================================================================
[perf-metrics]  性能試験用メトリクス収集: 有効  (perf-metrics.sh 1.0.0 / PERF_METRICS_ENABLED=true)
[perf-metrics]  出力先 : /mnt/logs/front/logs/intra-web/mid/20261020222607-ldtkjul4/perf
[perf-metrics]           (JBoss のログと同じタスク専用ディレクトリ。最後に起動したタスクは mid/current/perf から辿れます)
[perf-metrics]  ① jboss-cli : 10 秒間隔  DB 接続プール / Undertow / IO ワーカー / スレッドプール / トランザクション / JVM
[perf-metrics]  ② jstat     : 10 秒間隔  (-gc)   jcmd: 60 秒間隔  (GC.heap_info)
[perf-metrics]  ③ top / ps  : 10 / 10 秒間隔  (JBoss のスレッド別 CPU 上位 10 件、cgroup)
[perf-metrics]  統計の有効化: -Dwildfly.statistics-enabled=true を standalone.sh へ追加
[perf-metrics]  収集時間 : コンテナ停止まで   作業ディレクトリ: /tmp/perf-metrics
[perf-metrics]  状態確認 : perf-metrics.sh status   (ECS Exec で実行)
[perf-metrics]  ※ 性能試験用のイメージです。本番は PERF_METRICS_ENABLED=false でビルドしてください
[perf-metrics] ======================================================================
```

## 1. フォルダ構成

```
docker/                 ECS_EFS_Dockerfile_Symboliclink_lite の pin 方式 (正本・本番と同じ CMD=eap) に組み込んだもの
  base/perf-metrics.sh  ★ 性能試験用の実装 (この 1 ファイル)
  base/entrypoint.sh    「6-B. 性能試験用メトリクス収集」(読み込んで呼ぶだけ) を追加
  base/Dockerfile       perf-metrics.sh を置く (動かすかどうかは front / back で決める)
  front/Dockerfile      ビルド引数 PERF_METRICS_ENABLED / PERF_INTERVAL を追加
  back/Dockerfile       同上
  base/entrypoint.taskid.sh  変更なし
report/
  New-PerfMetricsReport.ps1  ★ Excel レポート (Windows 11 / PowerShell 5.1・7。Excel のインストール不要)
ecs/
  taskdef-perf.example.jsonc  性能試験用タスク定義の差分例 (/tmp ボリューム・環境変数)
patches/
  pin_method.patch            元リポジトリの pin_method/ (正本) に当てるパッチ
  container_link_method.patch 元リポジトリのルート (コンテナ専用リンク方式) に当てるパッチ
  make_patches.py             上の 2 つを作り直すスクリプト
test/
  unit/perf_hook_test.sh      エントリポイントの組み込み部分の単体試験 (Docker・JBoss 不要)
  e2e/run_e2e.sh              結合試験 (Docker。--read-only で起動し、負荷をかけて出力を確認)
  e2e/Dockerfile.eap-sim      検証用: UBI 9 OpenJDK 21 runtime + WildFly 35 (EAP 8.1 と同じ WildFly Core 27 系)
  e2e/make_war.py             検証用 WAR (ExampleDS を使い、接続待ちと GC を起こす JSP)
```

## 2. 使い方

### 2-1. ビルド

base は従来どおり。front / back にビルド引数を足します (タグは本番用と分けてください)。

```sh
docker build -t myapp-base:latest --build-arg STRICT_SEED=1 docker/base

# 性能試験用
docker build -t intra-web-front:perf \
  --build-arg BASE_IMAGE=myapp-base:latest \
  --build-arg Service_Name=intra-web \
  --build-arg Component_name=front \
  --build-arg PERF_METRICS_ENABLED=true \
  docker/front

# 本番用 (PERF_METRICS_ENABLED を付けない = false)
docker build -t intra-web-front:latest \
  --build-arg BASE_IMAGE=myapp-base:latest \
  --build-arg Service_Name=intra-web \
  --build-arg Component_name=front \
  docker/front
```

| ビルド引数 | 既定 | 内容 |
|---|---|---|
| `PERF_METRICS_ENABLED` | `false` | `true`: perf-metrics.sh を残し、`ps` / `top` が無ければ procps-ng を導入 (UBI 9 の openjdk-21-runtime には無い)。`false`: perf-metrics.sh を削除。それ以外の値はビルドエラー |
| `PERF_INTERVAL` | `10` | 各収集の既定の間隔 (秒)。タスク定義の environment で上書きできる |

`true` のビルドではログに導入結果が出ます (UBI 9.8 / OpenJDK 21.0.12 の例)。
`jstat` は runtime イメージに無いため `java -m jdk.jcmd/sun.tools.jstat.Jstat` で起動します (同じ JDK のモジュール)。

```
[perf-metrics]   ps: /usr/bin/ps
[perf-metrics]   top: /usr/bin/top
[perf-metrics]   jstat: /usr/lib/jvm/java-21-openjdk-.../bin/java -m jdk.jcmd/sun.tools.jstat.Jstat
[perf-metrics]   jcmd: /usr/lib/jvm/java-21-openjdk-.../bin/jcmd
```

### 2-2. タスク定義

`ecs/taskdef-perf.example.jsonc` の ★ の箇所を足します。

- **`/tmp` に書き込み可能ボリュームを当てる** (jstat / jcmd が JVM に接続するのに必要)。
  当てない場合も ①③ は動き、起動ログに警告が出ます。ヒープ・GC は ① の platform-mbean の値で確認できます。
- 収集用の JVM の分、メモリに 200MB 程度の余裕を持たせる (下の「負荷」参照)。

### 2-3. 出力先

`/mnt/logs/<Component_name>/logs/<Service_Name>/mid/<LOG_ID>/perf/` (例: `/mnt/logs/front/logs/intra-web/mid/current/perf/`)

- JBoss のログと同じ「タスク専用の実体ディレクトリ」です。`mid/current` は全タスクで共有して張り替える目印のため、
  **書き込みには使いません** (server.log を実体パスへ固定 (pin) しているのと同じ理由)。人が見るときは `mid/current/perf` から辿れます。
- 作業ファイル (FIFO・PID ファイル) は `/tmp` → `standalone/tmp` → `perf/.work` のうち書ける所に置きます。

### 2-4. 試験中の操作 (ECS Exec)

```sh
aws ecs execute-command --cluster <cluster> --task <task> --container intra-web-front --interactive --command "perf-metrics.sh status"
```

| コマンド | 内容 |
|---|---|
| `perf-metrics.sh status` | 出力先・各収集の状態・CSV の行数と最終時刻・直近のログ |
| `perf-metrics.sh threaddump` | その場でスレッドダンプを 1 回取る (`perf/jvm_threaddump/`) |
| `perf-metrics.sh stop` | 収集を止める (JBoss はそのまま) |
| `perf-metrics.sh pack` | `perf/` を `mid/<LOG_ID>/perf_<LOG_ID>_<日時>.tar(.gz)` にまとめる |

### 2-5. Excel レポート (Windows 11)

`perf/` (または `pack` の .tar / .tar.gz、`mid/` ごと) を PC に持ってきて実行します。

```powershell
# perf/ を 1 つ
powershell -NoProfile -ExecutionPolicy Bypass -File .\report\New-PerfMetricsReport.ps1 -InputPath .\perf -Open

# JMeter を実行した時間帯だけ (PC のタイムゾーンで指定)
.\report\New-PerfMetricsReport.ps1 .\perf_20261020-abcd1234_20261020_150000.tar -From "2026-10-20 13:00" -To "2026-10-20 14:00"

# mid/ ごと (タスクごとに 1 ブック。current 経由の重複は 1 回だけ処理)
.\report\New-PerfMetricsReport.ps1 .\mid -OutputPath .\reports
```

| パラメーター | 内容 |
|---|---|
| `-InputPath` | perf/、その親 (mid/ など)、`.tar.gz` / `.tar` / `.zip` |
| `-OutputPath` | 出力する .xlsx (複数のときはディレクトリ)。既定はカレントの `perf_report_<LOG_ID>.xlsx` |
| `-From` / `-To` | この時刻の範囲だけを使う |
| `-UseRecordedTime` | PC のタイムゾーンに換算せず、コンテナで記録した時刻で表示 |
| `-MaxRows` | 1 シートの最大行数 (既定 200,000) |
| `-Open` | 作成後に開く |

シート: サマリー (所見・主要指標の最大 / 95% 値 / 平均 / 最小・データソース別・収集プロセスの負荷) /
グラフ (16 種類前後。全グラフで時間軸を共通化) / DB接続プール / HTTP・スレッド / JVM・GC（CLI） /
ヒープ・GC（jstat） / ヒープ詳細（jcmd） / CPU・メモリ（top） / スレッド別CPU（top） / プロセス（ps） /
プロセス一覧（ps） / コンテナ（cgroup） / 収集情報 / 凡例。
フォント Meiryo UI、黒・グレーのモノトーン、線の種類で系列を区別、見出し固定・オートフィルター・
濃淡 (使用率 80% 以上・待ち・Full GC など)・データバー付き、A4 横で印刷できます。

- 大きな処理は C# (Add-Type) で行います。Add-Type が使えない環境では自動で PowerShell だけの処理に切り替わります (結果は同じ・時間は長くなります)。
- 所要時間の目安 (本 PC): 4 分間の試験 1.5 秒、24 時間分 (10 秒間隔・ps 18 万行) 約 80 秒。

## 3. 出力ファイル

| ファイル | 区分 | 内容 |
|---|---|---|
| `perf_meta.csv` | 共通 | 収集条件・環境 (タスク ARN・LOG_ID・JBoss / JVM の版・最大ヒープ・GC など) |
| `collector.log` | 共通 | 収集プロセスのログ (WARN / ERROR は CloudWatch Logs にも出る) |
| `cli_datasource.csv` | ① | データソース別 プール統計 (statistics=pool)・JDBC 統計・最小 / 最大プール |
| `cli_undertow.csv` | ① | HTTP / HTTPS / AJP リスナー別 request-count・error-count・bytes・processing-time |
| `cli_session.csv` | ① | デプロイメント別 HTTP セッション |
| `cli_io_worker.csv` | ① | IO ワーカーの処理中スレッド・最大・キュー |
| `cli_threadpool.csv` | ① | EJB thread-pool / EE managed-executor の使用中・キュー・ハング |
| `cli_transactions.csv` | ① | トランザクション件数 |
| `cli_jvm.csv` / `cli_gc.csv` / `cli_mempool.csv` | ① | ヒープ・非ヒープ・スレッド・クラス・サーバー状態・CLI 応答時間 / GC 別回数・時間 / メモリプール別使用量 |
| `jvm_jstat_gc.csv` | ② | jstat -gc (KB・秒)。`PERF_JSTAT_OPTIONS="gc gccause"` で `jvm_jstat_gccause.csv` も |
| `jvm_jcmd_heap.csv` / `jvm_jcmd.log` / `jvm_info.txt` | ② | GC.heap_info の数値化 / 出力そのもの / VM.version・VM.flags |
| `jvm_threaddump/` | ② | `PERF_THREAD_DUMP_INTERVAL` または `perf-metrics.sh threaddump` のスレッドダンプ |
| `os_top_summary.csv` / `os_top_process.csv` / `os_top.log` | ③ | top の要約行 / プロセス行 (区間の %CPU) / 出力そのもの |
| `os_top_threads.csv` | ③ | JBoss のスレッド別 %CPU 上位 |
| `os_ps.csv` | ③ | ps (スレッド数・RSS・累計 CPU 秒・コマンドライン。役割 jboss / jboss-cli / jstat … を付ける) |
| `os_cgroup.csv` | ③ | cgroup v1 / v2 の CPU 使用量・スロットリング・メモリ・OOM kill |

CSV は UTF-8・LF、1 行目が見出し、`timestamp` (コンテナの TZ。`PERF_TZ` で指定可) と `epoch` (UNIX 秒) 付き。
ps / top / perf_meta.csv の出力では、`password=` などの値を `********` に伏せます。

## 4. 設定 (環境変数)

ビルド時の ENV・タスク定義の environment のどちらでも指定できます (詳細は perf-metrics.sh の先頭)。

| 変数 | 既定 | 内容 |
|---|---|---|
| `PERF_INTERVAL` | 10 | ① ② (jstat) ③ の既定の間隔 (秒) |
| `PERF_CLI_INTERVAL` / `PERF_JSTAT_INTERVAL` / `PERF_TOP_INTERVAL` / `PERF_PS_INTERVAL` | `PERF_INTERVAL` | 個別の間隔 |
| `PERF_JCMD_INTERVAL` | 60 | jcmd の間隔 (1 回ごとに JVM を起動するため長め。10 も指定可) |
| `PERF_COLLECTORS` | `cli jstat jcmd top ps` | 動かす収集 |
| `PERF_DURATION` | 0 | 収集する秒数 (0 = コンテナが止まるまで) |
| `PERF_TZ` | (コンテナの TZ) | CSV の timestamp のタイムゾーン (例 `Asia/Tokyo`。tzdata が無くても JST-9 で代替) |
| `PERF_ENABLE_STATISTICS` | true | `standalone.sh` の直後に `-Dwildfly.statistics-enabled=true` を足す |
| `PERF_CLI_ENABLE_DS_STATS` | true | プール統計が無効なデータソースを実行時に有効化 (statistics=pool / jdbc の実行時属性。設定ファイルは変わらず、サーバー状態も running のまま) |
| `PERF_CLI_USER` / `PERF_CLI_PASSWORD` | (なし) | ローカル認証が使えない場合の管理ユーザー (Secrets Manager 推奨) |
| `PERF_JSTAT_OPTIONS` | `gc` | jstat のオプション (空白区切り) |
| `PERF_JCMD_COMMANDS` | `GC.heap_info` | jcmd のコマンド (カンマ区切り。例 `GC.heap_info,VM.native_memory summary`) |
| `PERF_THREAD_DUMP_INTERVAL` | 0 | スレッドダンプの間隔 (秒。0 = 取らない) |
| `PERF_TOP_THREADS` | 10 | JBoss のスレッド別 %CPU を上位何件記録するか (0 = 取らない) |
| `PERF_METRICS_ENABLED` | ビルド引数の値 | `false` にすると性能試験用イメージのまま収集だけ止められる |

## 5. 既存の実装への組み込み

`docker/` は `ECS_EFS_Dockerfile_Symboliclink_lite` の **pin 方式 (正本、本番と同じ CMD=eap)** の写しに、性能試験用の変更だけを足したものです。
変更は次の 4 か所と新規ファイル 1 つで、`patches/pin_method.patch` と同じ内容です
(元の pin_method/docker に当てると、このリポジトリの docker/ と 1 文字も違わないことを確認済み)。

| ファイル | 変更 |
|---|---|
| `base/perf-metrics.sh` | 新規 (性能試験用の実装すべて) |
| `base/Dockerfile` | `COPY perf-metrics.sh /usr/local/bin/perf-metrics.sh` (置くだけ) |
| `front/Dockerfile`・`back/Dockerfile` | `ARG PERF_METRICS_ENABLED` / `PERF_INTERVAL` → `ENV`、true なら `perf-metrics.sh install`、false なら削除 |
| `base/entrypoint.sh` | 「6-B.」節: `PERF_METRICS_ENABLED=true` のときだけ perf-metrics.sh を読み込み (`.`) `perf_metrics_start` を呼ぶ。統計の引数を standalone.sh の直後へ足す。説明コメント |

```sh
# 元リポジトリで (LF の作業ツリー。Windows の autocrlf 環境では core.autocrlf=false で)
git apply --directory=pin_method <このリポジトリ>/patches/pin_method.patch      # pin 方式 (正本の写し)
git apply <このリポジトリ>/patches/container_link_method.patch                  # コンテナ専用リンク方式 (ルート)
```

- 本番のエントリポイントに組み込む場合も、`exec` の直前に「6-B.」節を写すだけです (`JBOSS_LOG_DIR` に
  タスク専用の実体パスが入っていれば、出力先はその下の perf/ になります)。
- コンテナ専用リンク方式に当てたものも、--read-only での起動と収集を確認しています。

## 6. 負荷と注意事項

- **収集プロセス自身の負荷** (WildFly 35 / 2 vCPU / 5 秒間隔・40 並列の負荷中に測定):
  jboss-cli 常駐 CPU 平均 3% 前後 (1 コア = 100%)・RSS 約 120MB、jstat 常駐 CPU 0.5% 未満・RSS 約 40MB、
  jcmd は 1 回 0.5 秒前後で終わるため top に映らないことが多い (観測値 平均 0〜0.5%。既定は 60 秒に 1 回)、スクリプト類 1% 未満。
  レポートの「収集プロセスの負荷」と「CPU・メモリ（top）」シートで実測値を確認できます。
- jboss-cli は常駐セッション (FIFO で命令を送る) なので、間隔ごとに JVM を起動しません。
  ただしサーバーが高負荷のときは 1 回の取得に数秒かかることがあります (`cli_jvm.csv` の `cli_ms`)。
- `-Dwildfly.statistics-enabled=true` と統計の実行時有効化のわずかな負荷があります。本番と比べる試験では条件を揃えてください。
- Undertow の平均処理時間は、リスナーの `record-request-start-time=true` のときだけ値が入ります。
- 手動で実行した jboss-cli / jstat / jcmd は役割 `cli-other` / `jdk-tool` として収集プロセスと区別します
  (収集用の JVM には目印の `-Dperf.metrics.collector=...` を付けています)。
- コンテナ停止時は収集プロセスも止まります (CSV は 1 行ずつ追記しているため、それまでの分は残ります)。

## 7. 試験

| 試験 | 結果 |
|---|---|
| `test/unit/perf_hook_test.sh` (dash / bash --posix / busybox sh) | PASS=105 FAIL=0 |
| 元の pin 方式の `test/local/entrypoint_test.sh` を、組み込み後のエントリポイントで実行 | PASS=333 FAIL=0 SKIP=0 (組み込み前と同じ) |
| `test/e2e/run_e2e.sh` (Docker、--read-only、WildFly 35 + UBI 9.8 + OpenJDK 21.0.12、40 並列の負荷) | 全項目 OK (接続プールの枯渇と待ち・IO ワーカーの上限到達・Full GC・CPU スロットリングを記録) |
| /tmp を当てない・`PERF_DURATION`・統計が無効なデータソース・コンテナ専用リンク方式 | 想定どおり (警告して ①③ は継続 / 時間で全収集が停止 / 実行時に有効化しサーバーは running のまま / 収集可) |
| レポート: 生成したブックを Excel 16 で開く (修復なし)・PDF で体裁確認・C# 版と PowerShell 版の出力一致 | OK |

```sh
bash test/unit/perf_hook_test.sh
WILDFLY_TGZ_DIR=<wildfly-35.0.1.Final.tar.gz のあるディレクトリ> WORK=<作業ディレクトリ> bash test/e2e/run_e2e.sh 240
```

**未確認のこと:** 実際の ECS / Fargate と JBoss EAP 8.1 の実機では動かしていません
(EAP 8.1 と同じ WildFly Core 27 系の WildFly 35 で確認)。性能試験の前に、検証環境で 1 度起動ログと
`perf-metrics.sh status` を確認してください。PowerShell 7 は未確認です (5.1 で確認)。
