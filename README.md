# 國中排課系統
## 使用方法
打開helper.html，依照說明填寫學校資訊
# 國中排課系統

利用 **Google OR-Tools** 自動產生國中課表。

## 功能

- HTML 建立排課資料
- 自動產生 `input.txt`
- Python 自動排課
- 輸出 Excel 課表
- 排課前衝突檢查

## 使用

1. 開啟 `排課資料整理小幫手.html`
2. 填排課需求並下載 `input.txt`
3. 將 `input.txt` 與 `scheduler.py` 放在同一資料夾<br>
(step4. step5.擇一方法使用)
4. 若未安裝python環境可將 `input.txt` 與 `scheduler.py` 一同丟給claude AI(免費)
5. 若有安裝python環境，可執行python scheduler.py

完成後會產生

```
schedule_output.xlsx
```
## 注意事項
若使用python運行可不必擔心個資問題，此python & HTML都可本機離線運行 <br>
若要使用AI幫忙跑，請勿直接輸入老師真實姓名。
```
可先使用：
教師01、教師02
教學組長、七忠導師
```
> ❌同種課每天不上第二節<br>
> ❌無法排第8節<br>
> ✅嚴格檢查每班是否確定是35節(5天各7節)<br>
> ❌不連兩天上體育課<br>

## 使用技術

- Python
- Google OR-Tools (CP-SAT)
- OpenPyXL
- HTML / JavaScript
