# 单机网页工作台

## 启动

在 PowerShell 中先切换到项目根目录，然后执行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python scripts\run_local_web_app.py --data-root .\data
```

然后打开：

```text
http://127.0.0.1:8000
```

网页会自动生成业务 `run_id` 和 LangGraph `thread_id`，财务人员不需要输入它们。每次运行的上传资料位于 `data\runs\{run_id}\input\`，底稿位于同一任务的 `output\`；网页层的运行索引保存在 `data\app.sqlite`，LangGraph Checkpoint 保存在 `data\langgraph_checkpoints.sqlite`。

## 流程

1. 选择年度、月份和经办人，批量上传资料；
2. 页面自动执行预检，显示识别出的文件；缺件时可补充或替换资料后重新预检；
3. 预检通过后点击“开始计算”，后端串行执行一个工作流；
4. M7 全部为零时直接下载正式报表；
5. 存在差异时，页面显示审核材料、主体拆分和无票证据；
6. 填写正确营业收入、审核人和原因，点击提交继续；
7. 后端使用同一个 `thread_id` 恢复，重算 M5、M6、M7；
8. 仍有差异时进入下一轮审核，归零后才发布正式 `.xls`。

网页服务只适合本机或受控局域网使用。原始上传文件在任务开始后不应修改；如资料需要更换，建议新建任务。
