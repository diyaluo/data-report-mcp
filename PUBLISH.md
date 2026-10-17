# PUBLISH checklist — data-report-mcp

> 打包与本地实测已完成。以下发布动作需要 GitHub 凭据/人工确认，主进程核验后执行。

## 0. 已验证（2026-10-17）
- `report_core.py --selftest` PASS
- `test_stdio.py` 端到端：initialize → tools{analyze_table,health} → CSV 分析
  （5 行/4 列/1 重复行/数值列识别/Markdown/真实 xlsx base64 可解压）→ 错误扩展拦截 → ALL ASSERTIONS PASSED
- 干净环境铁律：spawn 子进程必须 `env -u PYTHONPATH`（Hermes runtime 3.14 的
  pydantic_core 损坏会污染独立 venv）。

## 1. 建公开仓库
```bash
cd /root/data/disk/projects/data-report-mcp
rm -rf src/**/__pycache__ __pycache__
git init -b main
git add .
git commit -m "feat: data-report-mcp v0.1.0 — local CSV/XLSX analysis over MCP"
gh repo create diyaluo/data-report-mcp --public --source=. --push
```
确认仓库只含可发布资产（无内部运营文档/凭证）。

## 2. 官方 MCP Registry（OCI 形态）
```bash
# 镜像由 Action 在 tag 时推送到 ghcr.io/diyaluo/data-report-mcp:v0.1.0
git tag v0.1.0 && git push origin v0.1.0
# 等 Action 完成、镜像 public 后：
mcp-publisher registry publish   # 需在有浏览器的机器 `mcp-publisher login github`
```
- server.json 已就绪；qualified name = `io.github.diyaluo/data-report-mcp`
- description ≤100 字符（已核）

## 3. Smithery（MCPB bundle，25MB 限制）
```bash
# 在仓库根（bundle 版 server 由 mcpb init 生成；当前包很薄，stdlib+openpyxl）
npx @anthropic-ai/mcpb init     # 生成 bundle 清单
npx @anthropic-ai/mcpb pack     # 产物必须 < 25MB
# 提交到 Smithery：server 43276307/data-report-mcp（CLI 或网页）
smithery publish
```
- 不要打 pandas/numpy（本工具零 pandas，保持薄包）。

## 4. 发布后验证
- 官方 registry：`mcp get io.github.diyaluo/data-report-mcp --version=latest`
- Smithery owner API：`GET api.smithery.ai/servers/43276307/data-report-mcp/logs`
  看真实调用 `total`（bundle 本地埋点对分发型不可见）。
- README 里的 stdio config 片段用 Claude Desktop 实连一次。

## 5. 复用同一模式的下一批候选
resume-builder（注意 PDF 需 LibreOffice，bundle 会变重，优先做 JSON/docx 工具面）、
excel-toolkit（公式生成+CSV清洗，纯 stdlib+openpyxl，最薄，建议第二个做）、
web-extract-skill（HTML→结构化，纯 stdlib）。
