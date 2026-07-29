# 國中排課系統

利用 **python Google OR-Tools** 自動產生國中課表。

## 功能

- HTML 建立排課資料
- 自動產生 `input.txt`
- Python 自動排課
- 輸出 Excel 課表
- 排課前衝突檢查


## 排課限制
### 硬限制
> ✅同一門課同一天最多一節。<br>
> ✅嚴格檢查每班是否確定是35節(5天各7節)<br>
> ❌無法排第8節<br>
> ❌不連兩天上體育課<br>
> ✅專科教室同時段最多一班使用：
### 軟限制
> 1️⃣考科避免超過單日5節
> 2️⃣考科避免連續兩天都沒課
> 3️⃣老師避免連上第4、5節
> 4️⃣老師避免連續3節課


## 使用
1. 開啟 `step1courses.html`
2. 填排課需求並下載 `input.txt`
   (step3. step4.擇一方法使用)
3. 線上運行(簡單)：將 `input.txt` 與 `scheduler.py` 一同丟給claude AI(免費)
4. 離線運行(python環境)：將 `input.txt` 與 `scheduler.py` 放在同一資料夾，執行python scheduler.py
完成後會產生schedule_output.xlsx

## 注意事項
若使用python運行可不必擔心個資問題，此python & HTML都可本機離線運行，不會竊取教師個資 <br>
若要使用AI幫忙跑，請勿輸入老師真實姓名。
```
可先使用：
教師01、教師02
教學組長、七忠導師
最後利用離線軟體（word& excel）尋找並取代
```


## 使用技術
- Python
- Google OR-Tools (CP-SAT)
- OpenPyXL
- HTML / JavaScript
- claude & codex (vibe coding)
