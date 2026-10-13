# 3GPP 提案（TDoc）下载工具 · 独立包

本文件夹是 **T0 下载工具** 的独立打包版：把 3GPP 会议的全部提案（TDoc）**按清单批量下载 → 解压 → 按 agenda 分类**，一步到位。

- **纯标准库**，无需 `pip install`（xlsx 用 `zipfile` + `ElementTree` 自解析，不需要 openpyxl / pandas）。
- **只读源、只新建文件**：绝不修改或删除你的源清单与共享状态。
- **自包含**：本文件夹内即可运行，不依赖原工作流的其它目录。

## 文件清单

| 文件 | 作用 |
|---|---|
| `tdoc_download.py` | **核心**：下载 + 解压 + 按 agenda 分类（唯一实现） |
| `tdoc_extract.py` | 分类引擎（`classify` / `agenda_matches` / `build_field_map`），被核心调用 |
| `tdoc_gui.pyw` | **图形界面**：双击打开窗口，零第三方依赖（tkinter），内部调用同目录的 `tdoc_download.py` |

## 环境要求

- Python 3（3.8+ 均可）。
- 命令行模式：无任何第三方依赖。
- GUI 模式：需要 Python 带 **tkinter**。
  （若双击 `.pyw` 打不开，说明默认解释器没带 tkinter，`tdoc_gui.pyw` 会自动尝试切换到其它已安装的 Python。）

## 快速开始

### 方式一：图形界面（推荐日常使用）

双击 `tdoc_gui.pyw`，选好 TDoc 清单（`.xlsx`）即可 —— 程序会**从文件名推断组别与会议号**（如 `TDoc_List_Meeting_RAN1#126-bis.xlsx` → RAN1 / 126-bis）并联网核对。

### 方式二：命令行

```bash
# 用法1（最可靠：直接给文档目录 URL）
python tdoc_download.py --from-tdoclist tdocList.xlsx \
    --base https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_126b/Docs/ --out Docs

# 用法1（给会议目录名 + 组别）
python tdoc_download.py --from-tdoclist tdocList.xlsx --tsg RAN1 --meeting TSGR1_126b --out Docs

# 用法2（仅分类，不下载：对已解压好的包分类）
python tdoc_download.py --classify --package extracted/ --from-tdoclist tdoclist.csv --out Docs

# 只检查计划，不下载
python tdoc_download.py --from-tdoclist tdocList.xlsx --base ... --dry-run
```

## 常用参数

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

## 输出物

在 `--out` 目录下生成：

- `by_agenda/<agenda>/` —— 按议程分类后的提案文件；
- `tdoclist_for_T1.csv`、`tdoc_index.json` —— 清单与索引；
- `_download_log.csv` —— 下载日志；
- `_failures.csv` —— 未取得的 TDoc（含 agenda / 标题 / 来源 / URL，Excel 可直接打开）。

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
