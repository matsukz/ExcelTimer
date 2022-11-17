ご覧いただきありがとうございます。<br>
こちらはExcelで簡単に動くシンプルなストップウォッチです。
# 機能の紹介

Sleepを利用しているため、**Windowsでしか動作しません**<br>
機能はシンプルです。<br>

## スタート
「開始」ボタンを押すとD2の数値が１日秒ごとに１増えます。<br>
D2の値が60になるとB2に+1しB2の値は0になります。<br>
厳密にはB2=B2-B2です。<br>

## ストップ/リセット
「開始」ボタンによるカウント中に「停止/初期化」を押すとカウントが一時停止します。（再度「開始」を押すとカウントが再開します）<br>
カウント一時停止中に再度「停止/初期化」ボタンを押すとカウントがリセットされます。

## ▲▼
ボタンを押すと数値がプラス１・マイナス１されます。<br>
不具合回避のためタイマーが動作していると両ボタンとも反応しません。

## その他
不具合回避のためシート内セルの編集はできません。<br>
ロックを解除するには「校閲」⇒「シート保護の解除」を選択してください。<br>
パスワードは**Unlocking**です。

### 参考
<a href=https://liclog.net/sleep-function-vba-macro-catia-v5/>VBAで指定した秒数だけ処理を止める方法【Sleep関数(API)】</a>（閲覧日2022-11-15）

<a href=https://www.sejuku.net/blog/69349>【ExcelVBA入門】DoEventsを使って処理を途中で止める方法を徹底解説！</a>（閲覧日2022-11-15）<br>
