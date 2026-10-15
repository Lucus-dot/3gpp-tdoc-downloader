# 3GPP 提案（TDoc）下载工具 · 独立包 v2

本文件夹是 3GPP TDoc 工具的**独立打包版**，两大能力：

1. **按清单批量下载**（T0）：从会议 TDoc 清单（xlsx / csv / zip）**批量下载 → 解压 → 按 agenda 分类**，一步到位。
2. **按号直下**（新增）：只知道 **TDoc 号**（`R1-2607107`）或 **规范号**（`38.211`），不用给清单、不用自己找会议目录，**直接定位并下载**。

- **纯标准库**，无需 `pip install`（xlsx 用 `zipfile` + `ElementTree` 自解析，不需要 openpyxl / pandas）。
- **只读源、只新建文件**：绝不修改或删除你的源清单与共享状态。
- **自包含**：本文件夹内即可运行，不依赖原工作流的其它目录。

## 文件清单

| 文件 | 作用 |
|---|---|
| `tdoc_download.py` | **T0 核心**：按清单下载 + 解压 + 按 agenda 分类（唯一实现） |
| `tdoc_extract.py` | 分类引擎（`classify` / `agenda_matches` / `build_field_map`），被 T0 调用 |
| `tdoc_direct.py` | **按号直下**：TDoc / TS / TR 按号直接下载（四级定位，见下） |
| `tdoc_cache.py` | **索引缓存**：会议表 / 规范版本列表落盘复用，避免同日重复抓取（被 `tdoc_direct.py` 调用） |
| `tdoc_direct.env.example` | 按号直下的**配置文件模板**（复制为 `tdoc_direct.env` 即生效） |
| `tdoc_gui.pyw` | **图形界面**：双击打开窗口，三种模式（见下），零第三方依赖（tkinter） |
| `LICENSE` | MIT 许可证 |
| `.gitignore` | 忽略运行产物 / 清单原件 / 本地配置 |

## 环境要求

- Python 3（3.8+ 均可）。
- 命令行模式：无任何第三方依赖。
- GUI 模式：需要 Python 带 **tkinter**。
  （若双击 `.pyw` 打不开，说明默认解释器没带 tkinter，`tdoc_gui.pyw` 会自动尝试切换到其它已安装的 Python。）

## 快速开始

### 方式一：图形界面（推荐日常使用）

双击 `tdoc_gui.pyw`，顶部选模式：

| 模式 | 用途 |
|---|---|
| ① 基于清单下载+分类 | 选 TDoc 清单（`.xlsx`）→ 程序**从文件名推断组别+会议号**并联网核对 |
| ② 仅分类(已有包) | 对已下载解压的包按 agenda 分类，不下载 |
| ③ **按号直下** | 直接填 **TDoc / 规范号**（可多个）；**未命中时自动列出候选会议**，双击即可填入「指定会议」再搜 |

### 方式二：命令行

```bash
# ---- 按清单下载（T0）----
python tdoc_download.py --from-tdoclist tdocList.xlsx \
    --base https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_126b/Docs/ --out Docs
python tdoc_download.py --from-tdoclist tdocList.xlsx --tsg RAN1 --meeting TSGR1_126b --out Docs
python tdoc_download.py --classify --package extracted/ --from-tdoclist tdoclist.csv --out Docs

# ---- 按号直下 ----
python tdoc_direct.py --get-tdoc "R1-2607107, RP-261575" --out Docs
python tdoc_direct.py --get-spec "38.211, 38.821" --out Docs
python tdoc_direct.py --direct "R1-2607107, TS 38.211" --out Docs   # 混合输入, 自动识别
python tdoc_direct.py --get-tdoc R1-2607107 --meeting R1-126-bis    # 指定会议, 前后各 ±5 个目录
python tdoc_direct.py --get-tdoc R1-2607107 --ask                   # 未命中时交互输入会议号
python tdoc_direct.py --get-spec 38.211 --spec-ver i40              # 指定版本(或 v18.4.0)
python tdoc_direct.py --get-tdoc R1-2607107 --dry-run               # 只定位, 不下载
python tdoc_direct.py --show-config                                 # 查看生效配置
```

## 按号直下（`tdoc_direct.py`）

### 它怎么找到会议？（不是"猜"，是"查表"）

3GPP 官方**会议表** `DynaReport/Meetings-<code>.htm` 每行给出【会议目录名 + TDoc 号区间 + 日期】——
所以是**查表定位，零网络试探**。其中 `code = TDoc 号前缀`（`R1-…`→`R1`，`RP-…`→`RP`），
**与工作组名字无关**；`RAN` / `SA` / `CT` 这类组名式的表是空壳，`RP` / `SP` / `CP` 才有数据。

### 四级定位

| 级别 | 触发 | 说明 |
|---|---|---|
| **L0 指定会议窗口** | `--meeting <会议>` | 在会议表定位该会议，取**前后各 ±radius（默认 5）**个目录成窗口，按号码接近度排序后探测 |
| **L1 区间命中** | 默认 | TDoc 号落入某会议区间 → 直接探测该会议目录 |
| **L2 遍历兜底** | 默认（可关） | 全表按号码距离升序逐目录探测，命中即停（默认只试前 `SCAN_LIMIT=10` 个） |
| **L3 手工兜底** | `--base <URL>` | 直接给 Docs 目录 URL |

> 定位到候选后一律 **HEAD 校验**（3GPP 对不存在路径返回 403，存在返回 200），
> 因为**区间是"包围盒"不是"枚举"**，区间内可能有未分配的号。

### 未命中怎么办

未命中时程序会打印**按号码接近度的前 5 个候选会议**（label / 日期 / 目录名），
**可直接用作 `--meeting` 的值**（GUI 里双击候选即可自动填入）。
也可以加 `--ask`，由程序提示你输入会议号后在**其前后各 ±5 个目录**内再搜一次。

### 配置文件（可选）

把 `tdoc_direct.env.example` 复制为 `tdoc_direct.env`（或 `.env`），放在**脚本目录**或**当前目录**即可被自动读取；
找不到就用默认值。**命令行参数优先级高于配置文件。**

| 键 | 默认 | 含义 |
|---|---|---|
| `SCAN_LIMIT` | 10 | L2 遍历会议目录上限 |
| `NEIGHBOR_RADIUS` | 5 | L0 指定会议时前后各扫几个目录 |
| `WORKERS` | 8 | HEAD / 下载并发 |
| `TIMEOUT` | 90 | 下载超时（秒） |
| `OUT` | Docs | 输出根目录 |
| `SCAN` | 1 | 是否启用 L2 遍历（1/0） |
| `CACHE` | 1 | 是否启用索引缓存（1/0） |
| `CACHE_TTL` | 604800 | 索引缓存有效期（秒，默认 7 天；0 或负数 = 永不过期） |
| `CACHE_DIR` | 空 | 索引缓存目录（空 = `<下载根目录>/.tdoc_cache`） |

### 索引缓存（可选）

按号定位要读几类**索引页**（3GPP 的会议表单页可达 400KB）：
`DynaReport/Meetings-<code>.htm`（TDoc 号区间 → 会议目录）与 `DynaReport/<digits>.htm`（规范 → 各版本 zip）。
这些索引**同日重复运行时内容不变**，所以会**落盘缓存** —— 换进程、重启 GUI 都不必重抓。

- 默认目录 **`<下载根目录>/.tdoc_cache`**（即 `--out` 指定目录下，**与下载产物同处**）；可用 `CACHE_DIR` 改。仅在该目录不可写时，才退到 `<脚本目录>/.tdoc_cache` 或系统临时目录。
- 默认有效期 **7 天**（`CACHE_TTL=604800`；0 或负数 = 永不过期）。
- 命中缓存时日志会打印：`索引缓存命中 meetings_R1（3.2 小时前抓取；要最新请加 --refresh-cache）`。

| 参数 | 说明 |
|---|---|
| `--no-cache` | 本次不使用缓存（仍会写入新抓到的） |
| `--refresh-cache` | 忽略已有缓存，重新抓取并更新 |
| `--cache-ttl N` | 有效期秒数（默认 604800） |
| `--cache-dir <目录>` | 缓存目录（默认 `<下载根目录>/.tdoc_cache`） |
| `--clear-cache` | 清空索引缓存后退出 |

> **只缓存索引页**。真正的 zip 下载与 **HEAD 存在性校验一律实时** —— 后者是"这个文件到底存不存在"的判据，必须现查。
> 若查不到某号、又怀疑是**新会议 / 新版本**，加 `--refresh-cache` 重试即可（GUI 模式③ 也有对应勾选框）。

### 按号直下 参数表

| 参数 | 说明 |
|---|---|
| `--get-tdoc <号>` | 按 TDoc 号下载，支持多值（逗号/分号/顿号/空格） |
| `--get-spec <号>` | 按 TS/TR 号下载，支持多值 |
| `--direct <混合>` | 混合输入，自动识别 TDoc / 规范号 |
| `--spec-ver` | 规范版本：`i40` / `j50` / `v18.4.0`；默认取最新 |
| `--meeting <会议>` | [L0] 会议目录名 / label / 序号，前后各 ±radius |
| `--radius N` | L0 半径（默认 5，配置 `NEIGHBOR_RADIUS`） |
| `--ask` | 未命中时交互式输入会议号，再走 L0 |
| `--base <URL>` | [L3] 直给 Docs 目录 URL |
| `--out <目录>` | 输出根目录，默认 `Docs` |
| `--workers` / `--timeout` | 并发（默认 8）/ 超时（默认 90） |
| `--scan-limit N` | L2 遍历上限（默认 10） |
| `--scan` / `--no-scan` | 启用 / 关闭 L2 遍历 |
| `--force-scan` | 强制走 L2（跳过 L1，用于自测兜底链路） |
| `--env <路径>` | 指定配置文件 |
| `--show-config` | 打印生效配置后退出 |
| `--dry-run` | 只定位、打印 URL，不下载 |

## 按清单下载（T0）常用参数

| 参数 | 说明 |
|---|---|
| `--from-tdoclist <文件>` | 从 3GPP TDoc 清单（xlsx / csv / zip）读清单，**推荐入口** |
| `--manifest <txt>` | 直接给清单（每行一个 TDoc 号） |
| `--base <URL>` | 完整文档目录 URL（以 `/Docs/` 结尾），**最可靠** |
| `--meeting` / `--tsg` | 会议目录名 + 组别，由工具自动拼 URL |
| `--out <目录>` | 输出目录，默认 `./Docs` |
| `--agenda <前缀>` | 按 agenda 过滤，可多个（逗号/分号/空格/顿号/全角逗号分隔），如 `10.7,9.1` |
| `--classify` | 仅分类模式，搭配 `--package` |
| `--keep-raw` | 保留原始 zip 与 `extracted/`（默认按 TDoc 逐个清理已成功分类的原件） |
| `--workers` / `--timeout` | 并发数（默认 5）/ 单请求超时秒数（默认 90） |
| `--dry-run` | 只打印计划，不实际下载 |
| `--resolve <会议号>` | 联网核对会议真实目录名 |
| `--hints` | 打印目录命名实例，便于对照 |

## 会议目录命名（拼 URL 前务必注意）

3GPP FTP 目录名**不规整**，各组 `bis` 写法不统一，**别凭拼接猜**：

| 组 | 目录前缀 | bis 写法示例 |
|---|---|---|
| RAN1 | `tsg_ran/WG1_RL1` | `TSGR1_126b`（小写 b） |
| RAN2 | `tsg_ran/WG2_RL2` | `TSGR2_91bis`（全拼 bis） |
| RAN3 | `tsg_ran/WG3_Iu` | `TSGR3_91bis` |
| RAN4 | `tsg_ran/WG4_Radio` | `TSGR4_94Bis`（大写 Bis，常带 `_e`） |
| 全会 | `tsg_ran/TSG_RAN` | `TSGR_113` |
| SA / CT | `tsg_sa/TSG_SA`、`tsg_ct/TSG_CT` | 常带城市/日期，如 `TSGS_99_Rotterdam_2023_03` |

> 不确定时用 `--resolve <会议号> --tsg <组别>` 拉真实目录清单核对。
> 3GPP 是 IIS：不存在的路径也返回 **403（不是 404）**，故 403 不等于"无权限"，先确认目录是否存在。
>
> **按号直下不受此困扰** —— 它从会议表的链接里**提取**目录名，不拼 URL。

## 输出物

**按清单下载（T0）** —— 在 `--out` 目录下生成：

- `by_agenda/<agenda>/` —— 按议程分类后的提案文件；
- `tdoclist_for_T1.csv`、`tdoc_index.json` —— 清单与索引；
- `_download_log.csv` —— 下载日志；
- `_failures.csv` —— 未取得的 TDoc（含 agenda / 标题 / 来源 / URL）。

**按号直下** —— 在 `--out` 目录下生成：

- `by_tdoc/<TDOC>.zip` —— 按号下载的 TDoc；
- `specs/<规范文件名>.zip` —— 按号下载的规范；
- `direct_index.json` —— 本次结果索引；
- `_direct_log.csv` —— 逐件日志（utf-8-sig）。

## 行为约定

- HTTP **4xx（403 / 404）视为永久失败，不重试**；仅 5xx / 网络异常才重试退避。
- **预检**：抽样首 / 中 / 尾，若全部 4xx 判为整会议不可达，秒级停止（`--no-preflight` 可关）。
- 默认**按 TDoc 逐个清理**：已成功分类的删掉原始 zip 与 `extracted/<TDoc>/`，未成功分类的**保留**便于排查（日志会打印 `KEPT_RAW:`）。
- 全角字符自动归一化：`--agenda "10.7，9.1"`（全角逗号）、`"10.7、9.1"`（顿号）均可。

---

## 作者与许可证

- 作者：**Lucus**
- 许可证：**MIT License**（© 2026 Lucus，详见 [LICENSE](LICENSE)）

*本包由一组 3GPP TDoc 处理脚本整理拆分而来，可独立使用。*
