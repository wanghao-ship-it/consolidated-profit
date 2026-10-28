# Web 服务部署速查

从项目根目录执行。首次运行先按根目录 README 配置 `.env` 并启动 PostgreSQL、MinIO。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python scripts\run_db_migrations.py
python scripts\create_user.py admin 管理员 --role ADMIN
python scripts\run_local_web_app.py --data-root .\data
```

浏览器打开 `http://127.0.0.1:8000`。PostgreSQL 业务数据库和 Checkpoint 数据库分别由 `DATABASE_URL` 与 `CHECKPOINT_DATABASE_URL` 配置；MinIO endpoint、bucket 和访问凭据由 `.env` 配置。

网页上传的原始 Excel、每个 LangGraph 节点生成的底稿和正式报表都存入 MinIO。PostgreSQL 文件索引保存对象键、文件哈希、大小、版本、上传人等信息；本地 `data/runs/{run_id}` 是供现有 Excel 计算函数使用的工作区，服务恢复任务时会从 MinIO 重建或校验文件。

LangGraph 以数据库 `workflow_runs.thread_id` 作为 Checkpoint 标识。出现差异后，审核人员使用网页提交修正；系统将审核记录追加到 PostgreSQL，再使用相同 `thread_id` 运行 `Command(resume=...)`。API 不接受客户端自报审核人，审核人来自登录会话。

`FINANCE` 可以创建和启动任务；`REVIEWER` 与 `ADMIN` 可以提交修正。用户通过以下命令创建，密码以交互方式输入，不写入源码：

```powershell
python scripts\create_user.py <用户名> <显示姓名> --role FINANCE
python scripts\create_user.py <用户名> <显示姓名> --role REVIEWER
```

首次创建业务表时执行 Alembic 迁移；LangGraph PostgreSQL Checkpointer 在首次构建时调用 `PostgresSaver.setup()`。服务启动时会把数据库中遗留的 `QUEUED`/`RUNNING` 任务重新入队；`WAITING_REVIEW` 任务维持暂停状态。

真实财务资料、数据库文件、导出报表和 `.env` 均属于运行数据，不应提交到 Git。旧 SQLite 记录和旧 Checkpoint 不会自动迁移。
