# consolidated-profit
基于 Python 和 LangGraph StateGraph 构建财务自动化工作流，采用“资料预检—本年/去年同期分支计算—M1-M7 底稿 生成—销售台账校验—差异修正—零差异发布”的流程。系统结合 FastAPI 接口和 PostgreSQL数据库、MinIO 对象存储，实 现 Excel 资料上传、利润表自动计算、运行状态管理、人工修正和可恢复执行。
