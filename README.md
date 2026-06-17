使用说明
cd c:\Work\1_Dev\stock\zhuss
pip install -r requirements.txt
playwright install chromium
python fetch_variety_pnl.py

可选参数：
--url：其他账户详情页 URL
--excel：指定输出 xlsx 路径
--headed：有界面浏览器，便于排查登录或结构变更

若站点需登录或反爬加强，抓取可能失败；此时用 --headed 看页面，必要时在脚本里扩展登录或等待逻辑

# 登录保存 Cookie（纯登录，不抓数据）：
python fetch_variety_pnl.py --save-storage-state dpswang_auth.json
自动打开浏览器到登录页
你慢慢输入用户名密码（最长等 5 分钟）
登录成功后自动保存 Cookie 并关闭

# 抓取数据（使用已保存的 Cookie）：
python fetch_variety_pnl.py --storage-state dpswang_auth.json
