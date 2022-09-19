---
title: Windows 10 Version 1607 上で Workbook.Close 実行後に Excel がエラーとなってしまうことがある問題について
date: '2019-03-01'
id: cl0mdnody002s5ovs5vpp069a
tags:
  - オートメーション

---

(※ 2017 年 1 月 27 日に Japan Office Developer Support Blog に公開した情報のアーカイブです。)  

こんにちは、Office 開発 サポート チームの 遠藤です。

今回は、Windows 10 への対応を行われているお客様からいくつかいただくもので、Excel がクラッシュしてしまう、という問題について報告したいと思います。

<span style="color:#ff0000">**2017/4/13 update**</span>  
<span style="color:#339966">本 Blog で紹介した問題については、2017 年 4 月から提供を開始しました Windows 10 Creators Update にて修正されました。</span>

この問題は以下の条件下で発生します。

(条件の 1 から 4 は必須条件となります。  
5 については、結果的に何かしらのオブジェクトを破棄するような処理が影響する可能性があり、5  は一例となります。)

1.  OS が Windows 10 Version 1607 (Anniversary Update) であること
2.  Excel をオートメーションで操作していること
3.  Excel 内部から結果的にプリンタ情報を取得する処理を行うこと
4.  "Windows で通常使うプリンターを管理する" が既定値である "オン" になっていること![](image1.png)
5.  Workbook オブジェクトの Close メソッドを呼び出すこと  
    

なお、この問題は任意の Excel バージョンで発生し、以下のコードを VBScript (.vbs) と保存して実行すると再現できます。

```
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True

Set xlBook = xlApp.Workbooks.Add

xlBook.Sheets(1).Range("A1:B2").Value = "Sample"
xlBook.Sheets(1).PageSetup.PrintArea = "A1:B2"
xlBook.Close False

MsgBox "少し待機し、Excel がクラッシュしないかを確認します"  
```

  

4 の "Windows で通常使うプリンターを管理する" は、Windows 10 Version 1511 から追加された機能ですが、Windows 10 Version 1607 で機能が拡張されたことに伴って発生します。

この問題に対しては、今年の春以降に予定している Windows 10 の大型アップデートで対応する予定となっているため、現時点では "Windows で通常使うプリンターを管理する" をオフにすることで対応をお願いできればと思います。

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**