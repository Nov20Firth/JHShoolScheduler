# 國中排課系統
作者：王老師<br>
原始專案：[GitHub 專案連結](https://github.com/Nov20Firth/JHShoolScheduler/edit/main/README.md)

利用 **python Google OR-Tools** 自動產生國中excel課表。
系統分為兩個步驟：
1. 使用 `step1courses.html` 整理課程、必排條件與教師不可排時段，產出 `input.txt` 。
2. 使用 `step2scheduler.py` 讀取 `input.txt`，自動排課並產生 Excel 課表。

# 功能與限制
## 硬限制
> ✅同一門課同一天最多一節。<br>
> ✅嚴格檢查每班是否確定是35節(5天各7節)<br>
> ❌無法排第8節<br>
> ❌不連兩天上體育課<br>
> ✅專科教室同時段最多一班使用：<br>
## 軟限制
> 1️⃣考科避免超過單日5節<br>
> 2️⃣考科避免連續兩天都沒課<br>
> 3️⃣老師避免連上第4、5節<br>
> 4️⃣老師避免連續3節課<br>

# 使用
1. 開啟 `step1courses.html`
2. 填排課需求並下載 `input.txt`
   (step3. step4.擇一方法使用)
3. 線上運行(簡單)：將 `input.txt` 與 `step2scheduler.py` 一同丟給claude AI(免費)
4. 離線運行(python環境)：將 `input.txt` 與 `step2scheduler.py` 放在同一資料夾，執行python step2scheduler.py.py
完成後會產生schedule.xlsx

# 隱私說明
以本機 Python 執行時，HTML 與 Python 程式皆在電腦本機處理資料；不會將教師資料上傳至網路。
但若將資料上傳至第三方 AI 或雲端服務，請自行評估該服務的隱私政策，並避免使用真實姓名等個人資料。
```
可先使用：
教師01、教師02
教學組長、七忠導師
最後利用離線軟體（word& excel）尋找並取代
```


# 使用技術
- Python
- Google OR-Tools (CP-SAT)
- OpenPyXL
- HTML / JavaScript
- claude & codex (vibe coding)

- ## 授權與使用規則
本專案歡迎非商業用途的使用、分享與修改。

- ✅可分享。
- ✅可修改。
- ✅衍生作品須保留原作者資訊與原始專案連結。
- ❌不得用於商業用途；包含販售、收費服務、商業部署或納入付費產品。
- 商業使用請先聯絡作者取得書面授權。

作者：王老師
完整授權條款：PolyForm Noncommercial License 1.0.0，請見 [LICENSE](./LICENSE)。
