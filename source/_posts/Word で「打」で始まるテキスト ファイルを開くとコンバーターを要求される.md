---
title: Word で「打」で始まるテキスト ファイルを開くとコンバーターを要求される
date: '2024-09-17'
id: cm15ufuis0000f0seec0i7yev
tags:
  - Word
---

こんにちは、Office サポート チームの中村です。
今回の記事では、Word でテキスト ファイルを開くとき、一部の内容においてコンバーターが要求される動作について説明します。
<br>

#### **1\. 現象**
以下の条件を満たすテキスト ファイルを Windows デスクトップ アプリの Word で開くと、「このファイルは新しいバージョンの Microsoft Word で作成されているため、このファイルを開くにはコンバーターが必要です。」というメッセージが表示されます。  
<br>

![](msg.png)

<ファイルの条件>
- ファイル内容の先頭が Shift-JIS の「打」 (0x91C5) で始まること  
- 2 文字目以降に文字が続くこと (制御コードを含め、何らかのデータがあること)  

<br>

#### **2\. 原因**
多くのファイル形式では、先頭数バイトでそのファイル形式を表します。このため Word でファイルを開くとき、ファイルのバイナリ データの先頭数バイトを元にファイル形式を判断しています。  
しかし、テキスト ファイルではこのようなデータ部を持たないため、ユーザーがファイルに入力したテキスト情報を、ファイルの先頭バイトから格納します。  
「打」を Shift-JIS でエンコードした場合の 0x91C5 で始まり、3 バイト目以降にデータが続くと、Word はテキスト ファイルとは別のファイル形式と認識します。この Word で認識されたファイル形式は、現在サポートされる Word デスクトップ アプリで開くことができない形式のため、Word はコンバーターが必要と判断し、先述のメッセージを表示します。  

「打」で始まるテキスト ファイルにおけるこの動作は、Windows デスクトップ アプリの Word における制限事項です。
<br>

#### **3\. 対応方法**

##### **3-1\. ユーザーの手動操作でファイルを開く場合**
「このファイルは新しいバージョンの Microsoft Word で作成されているため、このファイルを開くにはコンバーターが必要です。」のメッセージで [いいえ] を選択し、次に表示される以下のメッセージで [OK] を選択すると、テキスト ファイルとして開くことができます。

![](msg2.png)
<br>

##### **3-2\. プログラムからの操作でファイルを開く場合**
VBA などのプログラムでテキスト ファイルを開く場合は、Documents.Open メソッドの Format パラメーターに wdOpenFormatText を指定して明示的にテキスト ファイルとして開きます。この場合、コンバーターを要求するメッセージの表示を回避し、初めからテキスト ファイルとして開きます。

Documents.Open メソッド (Word)  
[https://learn.microsoft.com/ja-jp/office/vba/api/Word.Documents.Open](https://learn.microsoft.com/ja-jp/office/vba/api/Word.Documents.Open)

<公開情報より抜粋>  
_Format オプション バリアント型 (Variant)  
文書を開くために使用するファイル コンバーターを指定します。 WdOpenFormat 定数のいずれかを使用できます。 既定値は wdOpenFormatAuto です。 外部のファイル形式を指定するには、 FileConverter オブジェクトに OpenFormat プロパティを適用して、この引数で使用する値を決定します。_

WdOpenFormat 列挙 (Word)  
[https://learn.microsoft.com/ja-jp/office/vba/api/Word.WdOpenFormat](https://learn.microsoft.com/ja-jp/office/vba/api/Word.WdOpenFormat)

<公開情報より抜粋>  
_wdOpenFormatText	4	エンコードされていないテキスト形式_
<br>

##### **3-3\. ファイルの保存時に対処する場合**
テキスト ファイルの文字コードを変更して問題なければ、Shift-JIS ではなく UTF-8 など他の文字コードで保存することで回避できます。このときに、開かれたファイルが期待する文字コードで認識されない場合は、ファイルを開くときの以下の画面や、プログラムから開く場合は Documents.Open の Encoding パラメーターで文字コードを変更できます (例えば、UTF-8 の場合は msoEncodingUTF8 を指定します)。

![](CharcodeSelect.png)

Documents.Open メソッド (Word)  
[https://learn.microsoft.com/ja-jp/office/vba/api/Word.Documents.Open](https://learn.microsoft.com/ja-jp/office/vba/api/Word.Documents.Open)

<公開情報より抜粋>  
_Encoding オプション バリアント型 (Variant)  
保存された文書を表示するときに Microsoft Word で使用する、文書のエンコード (コード ページまたは文字セット) を指定します。 使用できる定数は、有効な MsoEncoding クラスの定数です。 有効な MsoEncoding クラスの定数の一覧については、Visual Basic Editor のオブジェクト ブラウザーを参照してください。 既定値はシステム コード ページです。_

MsoEncoding 列挙 (Office)  
[https://learn.microsoft.com/ja-jp/office/vba/api/Office.MsoEncoding](https://learn.microsoft.com/ja-jp/office/vba/api/Office.MsoEncoding)

<br>

今回の投稿は以上です。  
  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**