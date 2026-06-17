## 依赖安装步骤
```bash
pip install -r requirements.txt
playwright install chromium
```

## 脚本可选参数：
--url：其他账户详情页 URL  
--excel：指定输出 xlsx 路径  
--headed：有界面浏览器，便于排查登录或结构变更  

## 登录保存 Cookie（纯登录，不抓数据）：
```bash
python fetch_variety_pnl.py --save-storage-state dpswang_auth.json
```
1. 自动打开浏览器到登录页
2. 你慢慢输入用户名密码（最长等 5 分钟）
3. 登录成功后自动保存 Cookie 并关闭

## 抓取数据（使用已保存的 Cookie）：
```bash
python fetch_variety_pnl.py --storage-state dpswang_auth.json
```
