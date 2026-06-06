# Mac 微信聊天记录（包含群聊）文本导出工具

一个面向 **Mac 微信客户端 4.x** 的本地文本归档小工具：先通过 [`wx-cli`](https://github.com/jackwener/wx-cli) 读取自己已登录账号的本地聊天记录，再把指定好友或群聊导出为按时间线排列的 **DOCX**。

这个项目的初衷很朴素：朋友遇到案件，需要导出群聊天记录自证。这里把流程整理成一个尽量通用、注重隐私、适合半自助操作的工具，希望能帮到有需要的人。

## 功能

- 支持 Mac 微信客户端 4.x 的本地聊天记录文本导出。
- 支持群聊和私聊。
- 导出为 DOCX，包含：
  - 首页概览
  - 按日期统计
  - 发言人前 30
  - 完整时间线纯文本记录
- 支持直接从 `wx-cli` 导出，也支持把已有 `wx-cli export --format txt` 文本转为 DOCX。
- 适合配合 Codex、Claude Code、Cursor 等大模型工具半自助操作。

## 隐私原则

本项目默认只在本机运行。

- 不上传聊天记录。
- 不上传微信数据库。
- 不上传 `wx-cli` 密钥。
- 不收集遥测。
- 不包含任何云端 API。
- `.gitignore` 默认排除 `.wx-cli/`、`all_keys.json`、`config.json`、`*.txt`、`*.docx`、`*.pdf` 等敏感或导出文件。

请不要把以下文件提交到 GitHub：

```text
~/.wx-cli/all_keys.json
~/.wx-cli/config.json
任何聊天记录 TXT / DOCX / PDF
任何微信数据库目录
任何截图、证据材料、身份信息
```

## 合规与风险提示

请只导出你自己账号、本机已有、且你有权处理的聊天记录。  
如果用于案件、投诉、劳动争议、交易纠纷等场景，请根据当地法律要求保存原始设备、原始数据和操作过程记录。DOCX 适合阅读和提交材料，但不等同于司法鉴定。

本项目不绕过微信登录，不破解他人账号，不提供远程抓取能力。

## 依赖

1. macOS
2. Mac 微信 4.x，且目标账号已经在本机登录过
3. [`wx-cli`](https://github.com/jackwener/wx-cli)
4. Python 3.10+
5. `python-docx`

安装 Python 依赖：

```bash
python3 -m pip install -r requirements.txt
```

## 第一步：初始化 wx-cli

请先按 `wx-cli` 官方说明初始化本地数据库读取能力：

```bash
npx -y @jackwener/wx-cli init --force
```

Mac 上读取数据库密钥通常需要管理员权限，具体以 `wx-cli` 官方 README 为准。  
成功后，确认能列出会话：

```bash
npx -y @jackwener/wx-cli sessions
```

如果你是多开微信、改名微信、双开 App，可能需要先让 `wx-cli` 指向正确的数据目录。不要把 `.wx-cli` 配置和密钥提交到仓库。

## 第二步：导出指定群为 DOCX

最简单方式：

```bash
python3 scripts/export_chat_docx.py \
  --chat "群名" \
  --output "exports/群名_完整文字聊天记录.docx" \
  --limit 999999
```

如果你已经用 `wx-cli` 导出了 TXT：

```bash
python3 scripts/export_chat_docx.py \
  --source-txt "exports/group_full.txt" \
  --chat "群名" \
  --output "exports/group.docx"
```

## 用 Codex 半自助操作

可以把下面这段直接给 Codex：

```text
请在本机操作，不上传任何聊天记录或密钥。
用 wx-cli 列出会话，帮我找到「某某群」。
然后用本项目 scripts/export_chat_docx.py 导出该群完整文字聊天记录为 DOCX。
导出前检查 .gitignore，确保 all_keys.json、config.json、TXT、DOCX、PDF 不会被提交。
```

## 输出格式

DOCX 中的聊天正文格式大致为：

```text
[2026-06-06 13:10] 张三: 这是一条消息
[2026-06-06 13:12] 李四: 这是另一条消息
```

没有发送者字段的媒体占位消息会保留为：

```text
[2026-06-06 13:12] 系统/无发送者: [图片]
```

## 可选：渲染检查

如果你需要提交或打印，建议用 Word / WPS / LibreOffice 打开 DOCX 检查版式。  
对于特别大的群，DOCX 可能会有几千页，转 PDF 和渲染会比较慢。

## 开源说明

本项目使用 Apache License 2.0。  
本项目依赖外部工具 `wx-cli`，但不复制或内置 `wx-cli` 源码。`wx-cli` 由其原作者和贡献者维护，许可证同为 Apache License 2.0。

