# Excel VBA筆記

2020年突然被叫去管倉庫(是的，你沒看錯，真的是叫軟體工程師去管倉庫），隨著資料量越來越多，傳統的函數已經漸漸無法負荷每天要處理的資料量。於是興起了研究VBA的念頭，希望能透過程式讓我的工作變得更輕鬆

希望這份筆記可以幫助到想要減少工作時間的你

## 環境設置

Excel在安裝好後不會預設開啟開發人員選項，因此第一件要做的事就是將開發人員選項打開

- 示意圖

![image](./img/開發人員in功能區.jpg)

<br/>

接下來跟著我的步驟做吧！

    1.從「檔案」那邊打開選單
    2.選擇「選項」，將小視窗打開
    3.選擇「自訂功能區」
    4.將「開發人員打勾」
    5.按下確定，然後就會發現「開發人員」的工作區出現了


* 「檔案」在什麼地方

![image](./img/ch0_前置準備/檔案在功能區中的位置.jpg)

- 「選項」的位置

![image](./img/ch0_前置準備/選項的位置.jpg)

- 調整「自訂功能區」

![image](./img/ch0_前置準備/自訂功能區的設定.jpg)

步驟可參考youtube影片

<iframe width="480" height="270" src="https://www.youtube-nocookie.com/embed/Koe_JvnvAcY" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>

<br/>

既然大家都開啟「開發人員」，那就要繼續往下囉

<br/>

## 控制項與巨集

巨集（英文：Macro），是一種批次處理的稱謂

詳細內容請參考 [維基百科](https://zh.wikipedia.org/wiki/%E5%B7%A8%E9%9B%86)

按照我自己的理解就是：你有一系列動作，有一套SOP，電腦會隨著你的SOP幫你做出那些動作

例如：
動作遊戲裡有些招式的指令是↓↘→+AB，巨集就是電腦可以不斷幫你按照這個順序敲擊你的鍵盤，幫你重複執行這個指令

<br/>

Excel中的VBA就是「巨集」這個角色，因此你想做什麼動作，就能寫好VBA後讓電腦幫你重複執行

像我之前要點BOM表，幾千幾萬行的資料，人工核對號稱最少要半天，但是透過VBA的協助，不到1分鐘就處理好了
