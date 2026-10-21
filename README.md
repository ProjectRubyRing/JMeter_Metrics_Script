# JMeter_Metrics_Script — 性能試験用メトリクス収集 (ECS / Fargate + JBoss EAP 8.1)

JMeter で性能試験をしている間、ECS タスク (Fargate) のコンテナ
(UBI 9.8 + Red Hat build of OpenJDK 21 + JBoss EAP 8.1) の中で次の 2 つだけを一定間隔で取り、
`readonlyRootFilesystem=true` でも書き込める EFS 上のタスク専用ディレクトリへ CSV で出力します。
出力は Windows 11 の PowerShell でモノトーンの Excel レポートにできます。

| 区分 | 取得方法 | 内容 | 出力 |
|---|---|---|---|
| ① | jboss-cli (常駐セッション 1 つ) | DB 接続プール: データソース / XA データソースごとの使用中・アクティブ・アイドル・待ち・取得失敗・取得 / 待ち / 使用時間 (`statistics=pool`) と最小 / 最大プール・取得タイムアウト | `cli_datasource.csv` |
| ② | jstat (常駐 1 つ) | ヒープ各領域 (Eden・Survivor・Old・メタスペース) の容量・使用量、GC 回数・時間 (`jstat -gc`) | `jvm_jstat_gc.csv` |

- **性能試験用の実装は `docker/base/perf-metrics.sh` の 1 ファイル**にまとめています。
- **コンテナのビルド引数 `PERF_METRICS_ENABLED=true|false` で有効・無効を切り替えます**
  (既定 false。false ならファイル自体を削除し、本番と同じ中身になります)。追加のパッケージは不要です。
- 有効なイメージでは、エントリポイントが起動時にこのファイルを読み込み、起動ログに次の枠を出します。

```
[perf-metrics] ======================================================================
[perf-metrics]  性能試験用メトリクス収集: 有効  (perf-metrics.sh 2.0.0 / PERF_METRICS_ENABLED=true)
[perf-metrics]  出力先 : /mnt/logs/front/logs/intra-web/mid/20261021003859-vbkgervr/perf
[perf-metrics]           (JBoss のログと同じタスク専用ディレクトリ。最後に起動したタスクは mid/current/perf から辿れます)
[perf-metrics]  ① jboss-cli : 10 秒間隔  DB 接続プール (cli_datasource.csv)
[perf-metrics]  ② jstat     : 10 秒間隔  ヒープ・GC (jvm_jstat_gc.csv)
[perf-metrics]  統計の有効化: -Dwildfly.datasources.statistics-enabled=true を standalone.sh へ追加
[perf-metrics]  作業ディレクトリ: /tmp/perf-metrics   状態確認: perf-metrics.sh status (ECS Exec で実行)
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
| `PERF_METRICS_ENABLED` | `false` | `true`: perf-metrics.sh を残し、`perf-metrics.sh install` で必要なコマンドを確認する (awk が無いときだけ gawk を導入)。`false`: perf-metrics.sh を削除。それ以外の値はビルドエラー |
| `PERF_INTERVAL` | `10` | ① ② の既定の間隔 (秒)。タスク定義の environment で上書きできる |

`true` のビルドではログに確認結果が出ます。`jstat` は OpenJDK 21 の runtime イメージに無いため、
同じ JDK の `java -m jdk.jcmd/sun.tools.jstat.Jstat` で起動します。

```
[perf-metrics] ビルド時の確認 (perf-metrics.sh 2.0.0)
[perf-metrics]   awk: /usr/bin/awk
[perf-metrics]   mkfifo: /usr/bin/mkfifo
[perf-metrics]   jstat: /usr/lib/jvm/java-21-openjdk/bin/java -m jdk.jcmd/sun.tools.jstat.Jstat
[perf-metrics]   jboss-cli.sh: /opt/jboss-eap/bin/jboss-cli.sh
[perf-metrics] 性能試験用メトリクス収集を有効にしました (PERF_METRICS_ENABLED=true)
```

### 2-2. タスク定義

`ecs/taskdef-perf.example.jsonc` の ★ の箇所を足します。

- **`/tmp` に書き込み可能ボリュームを当てる** (jstat は JBoss が `/tmp/hsperfdata_*` に作る性能カウンタを読むため)。
  当てない場合も ① DB 接続プールは取れ、起動ログに警告が出ます。
- 収集用の JVM (jboss-cli・jstat) の分、メモリに 200MB 程度の余裕を持たせる (下の「負荷」参照)。

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

シート: サマリー (所見・主要指標の最大 / 95% 値 / 平均 / 最小・データソース別・収集情報の抜粋) /
グラフ (データソースごとに接続数・待ち / 取得失敗、ヒープ使用量・GC 回数・GC 時間・メタスペース。全グラフで時間軸を共通化) /
DB接続プール / ヒープ・GC（jstat） / 収集情報 / 凡例。
フォント Meiryo UI、黒・グレーのモノトーン、線の種類で系列を区別、見出し固定・オートフィルター・
濃淡 (使用率 80% 以上・待ち・取得失敗・Full GC・Old 使用率 85% 以上・GC 時間率 5% 以上)・データバー付き、A4 横で印刷できます。

- 所要時間の目安 (PowerShell 7.4): 2 分間の試験 1 秒前後、24 時間分 (10 秒間隔・データソース 3 つ・計 3.4 万行) 約 30 秒。

## 3. 出力ファイル

| ファイル | 内容 |
|---|---|
| `cli_datasource.csv` | ① データソースごとに 1 行 / 回。`ds_type` (ds / xa)・`ds_name`・`jndi_name`・`statistics_enabled`・`min_pool_size`・`max_pool_size`・`blocking_timeout_ms` と、`statistics=pool` の `ActiveCount` `AvailableCount` `InUseCount` `IdleCount` `MaxUsedCount` `WaitCount` `MaxWaitCount` `BlockingFailureCount` `TimedOut` `CreatedCount` `DestroyedCount` と取得 / 待ち / 使用 / 作成 / プール時間 (Average / Max / Total、ミリ秒) |
| `jvm_jstat_gc.csv` | ② `jstat -gc -t` の各行 (`uptime_s` と S0C … GCT。容量・使用量は KB、GC 時間は秒) |
| `perf_meta.csv` | 収集条件 (版・開始時刻・UTC オフセット・間隔・LOG_ID・起動コマンド・jboss-cli の接続先・jstat の起動方法など) |
| `collector.log` | 収集プロセスのログ (WARN / ERROR は CloudWatch Logs にも出る) |

CSV は UTF-8・LF、1 行目が見出し、`timestamp` (コンテナの TZ。`PERF_TZ` で指定可) と `epoch` (UNIX 秒) 付き。
最小 / 最大プールなどの設定値は属性ごとに読み取り (パスワード等は読みません)、未設定の値は空欄になります。
perf_meta.csv の起動コマンドでは、`password=` などの値を `********` に伏せます。

## 4. 設定 (環境変数)

ビルド時の ENV・タスク定義の environment のどちらでも指定できます (詳細は perf-metrics.sh の先頭)。

| 変数 | 既定 | 内容 |
|---|---|---|
| `PERF_INTERVAL` | 10 | ① ② の既定の間隔 (秒) |
| `PERF_CLI_INTERVAL` / `PERF_JSTAT_INTERVAL` | `PERF_INTERVAL` | 個別の間隔 |
| `PERF_COLLECTORS` | `cli jstat` | 動かす収集 |
| `PERF_TZ` | (コンテナの TZ) | CSV の timestamp のタイムゾーン (例 `Asia/Tokyo`。tzdata が無くても JST-9 で代替) |
| `PERF_ENABLE_STATISTICS` | true | `standalone.sh` の直後に `-Dwildfly.datasources.statistics-enabled=true` を足す (標準構成の式 `${wildfly.datasources.statistics-enabled:...}` を使うデータソースは起動時から統計が有効) |
| `PERF_CLI_ENABLE_DS_STATS` | true | それでもプール統計が無効なデータソースを実行時に有効化 (statistics=pool の実行時属性。設定ファイルは変わらず、サーバー状態も running のまま) |
| `PERF_CLI_USER` / `PERF_CLI_PASSWORD` | (なし) | ローカル認証が使えない場合の管理ユーザー (Secrets Manager 推奨) |
| `PERF_CLI_CONTROLLER` | (自動) | 管理インタフェース (既定は JBoss の起動引数から `remote+http://127.0.0.1:9990` などを求める) |
| `PERF_CLI_TIMEOUT` / `PERF_CLI_HEAP` | 30 / 64m | 1 回の取得のタイムアウト (秒) / jboss-cli の最大ヒープ |
| `PERF_WAIT_TIMEOUT` | 900 | JBoss の JVM・管理インタフェースを待つ秒数 |
| `PERF_OUTPUT_DIR` / `PERF_WORK_DIR` | (自動) | 出力先 / 作業ディレクトリを明示するとき |
| `PERF_METRICS_ENABLED` | ビルド引数の値 | `false` にすると性能試験用イメージのまま収集だけ止められる |

### 1.0.0 からの変更 (2.0.0 で取得範囲を絞ってスリム化)

- 取得するのは ① DB 接続プールと ② jstat -gc だけになりました。次は削除しました:
  HTTP リスナー・セッション・IO ワーカー・EJB / EE スレッドプール・トランザクション・JVM (platform-mbean) の jboss-cli 取得、
  jcmd (GC.heap_info・VM.flags)・スレッドダンプ、top / ps / cgroup (procps-ng の導入も不要に)。
- 削除した設定: `PERF_JCMD_*`・`PERF_THREAD_DUMP_INTERVAL`・`PERF_TOP_*`・`PERF_PS_INTERVAL`・`PERF_JSTAT_OPTIONS`・
  `PERF_DURATION` (止めるときは `perf-metrics.sh stop`)・`PERF_CLI_MODE` (常駐セッションのみ)。コマンド `threaddump` も削除。
- 統計の有効化は `-Dwildfly.statistics-enabled=true` (全サブシステム) から `-Dwildfly.datasources.statistics-enabled=true`
  (データソースだけ) に変えました。取得しない Undertow などの統計の負荷をかけないためです。
- `cli_datasource.csv` から JDBC 統計 (文キャッシュ) と `enabled` 列を外しました。レポートは列名で読むため 1.0.0 の出力も読めます。
- 行数: perf-metrics.sh 1,713 → 約 900 行、New-PerfMetricsReport.ps1 2,106 → 約 1,240 行 (C# 版の高速化処理も外し、PowerShell だけにしました)。

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
- 1.0.0 のパッチを当て済みの場合は、`docker/base/perf-metrics.sh` を置き換え、entrypoint.sh・front / back の Dockerfile の
  コメント (統計の引数・procps-ng の記述) を直すだけです (処理の変更はありません)。

## 6. 負荷と注意事項

- **収集プロセス自身の負荷** (WildFly 35 / 10 秒間隔・40 並列の負荷中に測定。1 コア = 100%):
  jboss-cli 常駐 CPU 0.4% 前後・RSS 約 130MB、jstat 常駐 CPU 0.4% 前後・RSS 約 45MB。
- jboss-cli は常駐セッション (FIFO で命令を送る) なので、間隔ごとに JVM を起動しません。
  JBoss が再起動しても jboss-cli は自分で再接続し、jstat は新しい JVM に付け直して収集を続けます
  (止まっていた間は行が抜け、collector.log に失敗と再開が出ます)。
- `-Dwildfly.datasources.statistics-enabled=true` とプール統計の実行時有効化のわずかな負荷があります。本番と比べる試験では条件を揃えてください。
- プール統計を実行時に有効化したデータソースは、有効化以前の値が 0 です (標準構成の式を使うデータソースは起動時から有効)。
- 手動で実行した jboss-cli / jstat と区別できるよう、収集用の JVM には目印の `-Dperf.metrics.collector=...` を付けています。
- コンテナ停止時は収集プロセスも止まります (CSV は 1 行ずつ追記しているため、それまでの分は残ります)。

## 7. 試験

| 試験 | 結果 |
|---|---|
| `test/unit/perf_hook_test.sh` (dash / bash --posix / busybox sh) | PASS=105 FAIL=0 |
| `test/e2e/run_e2e.sh 90` (Docker、--read-only・uid 185、WildFly 35、40 並列の負荷) | 全項目 OK (接続プールの上限到達と待ち・GC・パスワードの伏せ字・status / stop・PERF_METRICS_ENABLED=false で収集しないこと) |
| ローカルの WildFly 35 + OpenJDK 21 で直接実行 | 設定付きデータソース・XA データソースの取得、プール統計の実行時有効化、JBoss の再起動後の再接続、status / stop / pack を確認 |
| 応答の解析 (awk) | mawk・busybox awk・gawk (UBI 9) で同じ結果 |
| レポート (PowerShell 7.4) | 1.0.0 のレポートと共通の列の値が全行一致、LibreOffice で開いて PDF の体裁を確認、tar.gz 入力・-From / -To・片方の CSV だけの場合を確認 |

```sh
bash test/unit/perf_hook_test.sh
WILDFLY_TGZ_DIR=<wildfly-35.0.1.Final.tar.gz のあるディレクトリ> WORK=<作業ディレクトリ> bash test/e2e/run_e2e.sh 240
# registry.access.redhat.com を使えない環境では RUNTIME_IMAGE=<OpenJDK 21 の入ったイメージ> で土台を変えられる
```

**未確認のこと:** 実際の ECS / Fargate と JBoss EAP 8.1 の実機では動かしていません
(EAP 8.1 と同じ WildFly Core 27 系の WildFly 35 で確認)。2.0.0 の結合試験は、registry.access.redhat.com に
接続できない環境で行ったため、openjdk-21-runtime の代わりに ubi9-minimal + OpenJDK 21 (Temurin 21.0.5。runtime
イメージと同じく jstat の実行ファイルを削除) を土台にしています。2.0.0 のレポートは Windows PowerShell 5.1 と Excel では
まだ開いていません (PowerShell だけの処理は 1.0.0 で 5.1 の動作を確認済みの部分と同じです)。
性能試験の前に、検証環境で 1 度起動ログと `perf-metrics.sh status`、レポートの作成を確認してください。
