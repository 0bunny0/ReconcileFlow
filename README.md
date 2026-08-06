# POOWARD ReconcileFlow · Streamlit Edition

接单与出货双模块、人民币与原币双口径的 Excel 差异核对工具。此目录可直接作为一个 GitHub 仓库部署到 Streamlit Community Cloud。

## 功能

- 接单差异、出货差异分别上传，无需一次提交四份系统表。
- 文控登记表在两个模块之间共用一次选择。
- 上传框原生支持点击选择和拖拽上传 `.xlsx` / `.xlsm`。
- 每次核对同时计算人民币、原币结果，并保存在当前用户会话中。
- 只展示非零差异，支持“客户代码 → 订单流水号 → 双侧原始行”逐级下钻。
- 当前口径结果可下载为带核对说明的 Excel 底稿。
- “系统金额”和“文控金额”作为中央对照轴突出显示。

## 目录

```text
.
├── streamlit_app.py          # Streamlit 页面和会话状态
├── reconciliation.py         # Excel 读取、匹配、汇总和导出逻辑
├── requirements.txt          # 云端安装依赖
├── .streamlit/config.toml    # 主题和上传大小配置
├── assets/styles.css         # 界面视觉规范
└── tests/test_reconciliation.py
```

## 部署到 Streamlit Community Cloud

1. 新建一个 GitHub 仓库，把本目录中的文件和文件夹完整上传到仓库根目录。
2. 登录 [Streamlit Community Cloud](https://share.streamlit.io/)，选择 **Create app**。
3. 选择仓库和分支，入口文件填写 `streamlit_app.py`。
4. 在 **Advanced settings** 中选择 Python `3.12`。
5. 点击 **Deploy**。首次部署会根据根目录中的 `requirements.txt` 自动安装依赖。
6. 部署完成后，将生成的网址发给有权限的同事即可；同事不需要安装 Python。

如果表格包含公司敏感数据，请使用私有 GitHub 仓库和受限访问的 Streamlit 工作区，不要把应用设为公开。程序不会主动写入数据库，但上传文件及结果会在当前服务器会话内存中存在；会话结束后缓存失效。

## 本机预览

建议使用 Python 3.12：

```bash
python -m venv .venv
```

Windows：

```bat
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

macOS / Linux：

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

浏览器会打开 `http://localhost:8501`。

## 文件要求

### 接单模块

- 文控登记表：读取名称包含“接单”的工作表。
- 接单金额明细：需要客户代码、订单流水号、接单金额(RMB)、交易金额。
- 接单运费明细：需要客户代码、订单流水号、出货运费(RMB)、出货运费(原币)。

### 出货模块

- 文控登记表：读取名称包含“出货”的工作表。
- 出货金额明细：需要客户代码、订单流水号、出货金额(RMB)、出货金额；也兼容“实际出货金额”字段。
- 出货运费明细：需要客户代码、订单流水号、出货运费(RMB)、出货运费(原币)。

人民币口径使用文控 `VAT PRICE`；原币口径使用文控 `TP-CPO`。差异公式始终为：

```text
差异 = 系统金额 - 文控金额
```

## 会话与多人访问

- 每位访问者拥有独立的 Streamlit Session State，不会看到其他用户当前页面的缓存结果。
- 刷新、网络中断、应用重启或会话超时可能清除结果，请在核对后及时下载 Excel。
- 当前版本没有账号、数据库、历史记录和审批流；如需正式公司级部署，建议增加 SSO、审计日志和受控对象存储。

## 测试

```bash
python -m unittest discover -s tests -v
```
