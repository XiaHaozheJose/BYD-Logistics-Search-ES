# BYD Logistics Search — Progress Log

## 2026-05-06 — Full Codebase Audit & Fix

### Diagnosis: 新 Excel 上传 500 错误

**问题描述**: 新版 `Inbound_outbound_overview.xlsx` (34.2 MB, 13 sheets, 392K 行) 上传后报 500 错误。旧版 (23.8 MB, 6 sheets) 能正常上传。

**对比分析**:

| 指标 | 旧 Excel | 新 Excel |
|------|----------|----------|
| 文件大小 | 23.8 MB | 34.2 MB |
| Sheet 数量 | 6 | 13 |
| 最大 Sheet 行数 | 315,310 (Outbound Overview) | 315,298 (Outbound Overview) + 70,799 (OB Overview 25.7-26.1) |
| 总行数 | ~321K | ~392K |
| 生成 SQLite 大小 | ~50 MB (估算) | 68.2 MB (实测) |

**根因**: 本地测试加载成功（74 秒），但 Cloud Run 环境：
1. `tmpfs` 文件系统占用内存：34 MB Excel + 68 MB SQLite = ~102 MB 磁盘 = 内存
2. FTS5 索引构建时对 315K 行做全表 `SELECT ... FROM table` 导致内存峰值
3. 13 sheets 比旧版多一倍，每个 sheet 重新打开 workbook
4. Gunicorn 2 workers 共享有限内存

**修复方案**: 
- FTS5 索引分批构建（5000 行/批），避免大 SELECT 内存峰值
- 每 sheet 添加 try/except 错误隔离
- 增大 BATCH_SIZE 到 2000 提升吞吐

---

### Security Audit Summary

| 严重性 | 发现 | 修复 |
|--------|------|------|
| Critical (3) | 模板API无认证、SQL注入、XSS | 全部修复 |
| Medium (4) | CORS通配符、无限线程、无输入验证、Firebase配置暴露 | 前3个修复，Firebase配置为设计如此 |
| Low (2) | 错误信息泄露、无限线程 | 全部修复 |

### Files Modified

| File | Changes |
|------|---------|
| `data_loader.py` | FTS5 分批构建、per-sheet 错误隔离、BATCH_SIZE=2000、_cell_to_str 健壮性 |
| `app.py` | 模板端点加认证、table 白名单验证、ThreadPoolExecutor、limit/offset 验证、CORS 配置化、mode 验证 |
| `search_engine.py` | 表名正则校验、limit/offset 范围限制 |
| `template_engine.py` | HTML 转义防 XSS、修复默认模板 |
| `gcs_helper.py` | 修复 delete_model_files 注释和逻辑 |
| `requirements.txt` | 移除未使用的 pandas 依赖 |
| `TASKS.md` | 新建任务跟踪文档 |
| `PROGRESS.md` | 新建进度日志 |
