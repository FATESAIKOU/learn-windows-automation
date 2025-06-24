#  windows 自動化工具集

## 目的

我是一個使用 windows 工程師兼專案管理者，我現在需要自動化 windows 上的工作，但現在我具體還沒有整理

1. 要自動化那些工作
2. 如何自動化這些工作

所以我要先建立一個之後可以乘載這個自動化的框架，這個框架需實現的功能目前我想像有兩個:

1. 將 excel / ppt / spo / markdown 甚至網頁等等應用程式的操作自動化。
2. (將來的擴張方向 可以先不考慮 因為可能乾脆分專案更好) 監視某個 windows app 狀態並公開的 script 集合體

請以這兩個為前提進行實現。

## 需求與使用架構
- Must
    - 使用者能隨時添加一份 script 然後使用公用的入口選單程式啟動它
    - 使用者添加的 script 是個吃 argv 的基於 pywin32 或 pywinauto 程式
    - 使用者添加的 script 可以寫原生 pywin32 或 pywinauto 的 api call(也可以同時利用), 但將來希望把可以自動化包裝的 windows 操作包裝成 utils 所以透過公用入口載入啟動
    - 公用入口應該吃兩個參數
        - 啟動的 script 名
        - 給啟動 script 的參數列
            - 後期這個東西的給法可以考慮互動式介面 但現在先不考慮(Todo)
    - user script 的資料夾結構不該用 pywin32 / pywinauto 這種底層技術做區分 應該考慮使用場合
    - 針對程式碼結構導入品質管理
        - linter
        - 單元測試
            - 接觸 windows API(pywin32/pywinauto) 的部分盡量 mock
            - 針對 windows API(pywin32/pywinauto) 的 mock 需要設計一個可重用 可成長的最基礎架構
            - 測試範圍 = pywin32 / pywinauto 的 wrapping utils
- Nice to have
    - 針對程式碼結構導入品質管理
        - 整合測試
            - 實際呼叫 windows API(或者 pywin32/pywinauto 應用程式開發時常用的官方 mock?)進行測試
            - 測試範圍 = userscript(一份 script 一個 test)
            - 需要能在 github actions 一份 workflows 能解決的程度
            - 過度複雜或者無法自動化測試的話就乾脆先放棄
            - 可以考慮只測 pywinauto 的程式碼的部分 -> 但到時候就要考慮測試 scope 戰略 所以請通知我到底可不可行
    - arg 介面可以考慮導入 typer?
    - 希望能導入某種強型態保證機制
    - 開發一定要在 windows 上很惱人 看看有沒有辦法考慮在 mac 上開發 但測試用某種...docker,mock?

## 使用技術
(這都是我大致設想的技術 可能不是最佳解 你可以幫我改 但要提出理由 以及獲取我的認同)
- 語言: python3.13
- 基礎框架: 無(不使用 Django 那種Web大型框架 但如果有cli指令專用的框架可以考慮)
- 實現層lib: pywin32 or pywinauto
- 套件管理工具: uv
    - 環境建置時請勿使用 uv pip, uv python 這種走回頭路的作法! 非要用給我理由 並且徵得我的同意
- 測試框架: pytest(我感覺原生 pytest 應該可以? 或者你有提案都可以提)
- linter: 沒特別想法 聽說 ruff 特別紅? 同樣你有提案可以提
- ci: github actions
- 執行環境: 原生環境(理論上應該不涉及 server 啟動 應該不需要docker類的東西 當然你可以提案)

(不確定有沒有其他 如果有請列出全部要考慮的項目後一一提案徵得我的意見 但不需要一開始就有完美架構 只要之後可成長就好)

## 實現步驟
(所有步驟都包含要實現項目的 1. 設計 2. 實現 3. 測試 每個階段完成後都需要輸出總結請求 review)
1. 建立程式基底
    - 須包含設定: 套件管理工具 語言 基礎框架 linter ci 執行環境設定
    - 都是最基本的殼 單純就是打通各步驟連結
2. 建立程式基底的單元測試
    - 須包含設定: 測試框架 實現層lib 最小用例程式碼(10~20行等級?) 最小用例配套測試程式碼 mock實現資料夾結構設定
3. 添加基本用例
    - 須包含設定: 入口 script, 複雜一點的 pywin32, pywinauto 用例 script, 針對兩者的測試(單元測試)
4. 總結 README.md
5. 嘗試整合測試(nice to have)
    - 先給我生成幾個測試方案 需提及的內容包含測試的依存環境 測試模組範圍 測試方式概要 然後讓我選擇才進實現
