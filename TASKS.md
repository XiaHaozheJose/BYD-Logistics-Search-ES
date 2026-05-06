# BYD Logistics Search — Task Tracker

## Critical Bug

| ID | Task | Priority | Status | Notes |
|----|------|----------|--------|-------|
| BUG-001 | 新 Excel (Inbound_outbound_overview.xlsx) 上传 500 错误 | P0 | FIXED | 根因：13 sheets / 392K行 / 68MB DB 在 Cloud Run 内存受限环境下 OOM。修复：FTS5 分批构建索引 + 每 sheet 错误隔离 + 内存管理优化 |

## Security Fixes

| ID | Task | Priority | Status | Notes |
|----|------|----------|--------|-------|
| SEC-001 | Template API 无认证 — 任何人可增删改模板 | P0 | FIXED | 所有 `/api/templates` 和 `/api/render` 端点添加 `@require_auth` |
| SEC-002 | SQL 注入 — `/api/search` 的 table 参数直接拼入 SQL | P0 | FIXED | search_engine.py 添加表名正则校验；app.py 添加 model_meta 白名单验证 |
| SEC-003 | XSS — 模板渲染不转义 HTML | P0 | FIXED | template_engine.py 使用 `html.escape()` 转义模板占位符值 |
| SEC-004 | CORS 通配符 `*` | P1 | FIXED | 改为环境变量 `CORS_ORIGINS` 控制，默认仍为 `*`（可在部署时收紧） |

## Reliability Fixes

| ID | Task | Priority | Status | Notes |
|----|------|----------|--------|-------|
| REL-001 | 无限后台线程 — 每次上传创建新 thread 无上限 | P1 | FIXED | 改用 `ThreadPoolExecutor(max_workers=3)` |
| REL-002 | limit/offset 无验证 — 非法值导致 500 | P1 | FIXED | 添加 try/except + 范围限制 (limit ≤ 1000, offset ≥ 0) |
| REL-003 | 错误信息泄露内部细节 | P2 | FIXED | 搜索错误不再暴露 Python 异常信息给前端 |

## Code Quality Fixes

| ID | Task | Priority | Status | Notes |
|----|------|----------|--------|-------|
| QA-001 | pandas 依赖未使用但声明在 requirements.txt | P2 | FIXED | 从 requirements.txt 移除，Docker 镜像减少 ~50MB |
| QA-002 | GCS delete_model_files bug — 注释说删除 DB 实际未删 | P2 | FIXED | 修正函数只删除 Excel（DB 由调用方重新上传） |
| QA-003 | 默认模板 {{#each}} 语法无效 | P2 | FIXED | 替换为有效的默认模板 |
| QA-004 | FTS5 索引大表内存峰值 | P1 | FIXED | 分批构建 FTS 索引（每批 5000 行） |

## Known Issues (Deferred)

| ID | Task | Priority | Status | Notes |
|----|------|----------|--------|-------|
| DEF-001 | 前端代码重复 (public/index.html vs templates/index.html) | P3 | DEFERRED | templates/ 版本缺少认证和许多功能，建议长期废弃 |
| DEF-002 | 无自动化测试 | P3 | DEFERRED | 建议后续添加 pytest 测试 |
| DEF-003 | 无 CI/CD 配置 | P3 | DEFERRED | 建议添加 GitHub Actions |
| DEF-004 | 无结构化日志 | P3 | DEFERRED | 建议引入 python-json-logger |
| DEF-005 | mode 参数未严格验证 | P2 | FIXED | 非法 mode 值自动回退到 "replace" |
