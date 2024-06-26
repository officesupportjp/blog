---
title: OneNote の OCR 機能で読み取りにくい文字列がある
date: '2022-04-21'
tags:
  - ''
id: cl28oxvt6002ku8uj1ose6ej2

---

こんにちは、Office サポート チームの永田です。

デスクトップ版 OneNote は光学式文字認識 (OCR) をサポートしており、ページに挿入された画像からテキストを読み取って、クリップボードに取得することができます。このとき、画像の解像度や、画像の形式によって、期待通りにテキストを読み取れないことがあります。

読み取りに影響する要因は、以下のようなものがあります。
- フォントが小さい文字列
  - 画像に含まれる文字のフォント サイズが小さければ小さいほど、読み取られにくい
- 罫線を含む文字列
  - 文字列の上や下に罫線がある場合、文字が正確に読み取られにくい
- 表形式の文字列
  - 複数列によってレイアウトされている表形式の文字列は、レイアウトを維持して読み取られにくい

# 対応状況

この動作は OneNote の機能上の制限です。

# 回避策

誠に恐れ入りますが、以下の方法を検討ください。

- フォントが小さい文字列
  - 画像の表示倍率を拡大して読み込む。拡大して画像がぼやける場合は、必要に応じてシャープネス機能を利用する。Word の場合、図の書式設定の「シャープネス」機能を利用可能。
- 罫線を含む文字列
  - 文字列の上や下に罫線がある場合、最上段や最下段には出来るだけ罫線を含まないようスクリーンショットを採取して OneNote に貼り付けて読み取る
- 表形式の文字列
  - 読み取り対象画像に対して、一列ずつスクリーンショットを採取して OneNote に貼り付けて読み取る


今回の投稿は以上です。


**本情報の内容（添付文書、リンク先などを含む）は、作成日時点でのものであり、予告なく変更される場合があります。**
