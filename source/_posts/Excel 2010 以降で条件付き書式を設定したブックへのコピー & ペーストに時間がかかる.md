---
title: Excel 2010 以降で条件付き書式を設定したブックへのコピー & ペーストに時間がかかる
date: '2019-02-07'
id: cl0n10dcu000j90vsghpb14v9
tags:
  - コピー＆ペースト

---

(※ 2017 年 8 月 14 日に Office Support Team Blog JAPAN に公開した情報のアーカイブです。)

こんにちは、Office サポートの町口です。  

本記事では、Excel 2010 以降で条件付き書式を設定したブックへのコピー& ペーストに時間がかかる現象について説明します。

  

  

  

現象
--

  

Excel 2007 と比較して、Excel 2010 以降では、条件付き書式を設定ブックへのコピー& ペーストに時間がかかります。

  

  

  

原因
--

  

貼り付けした際、ブック内の条件付き書式の再評価が行われますが、

この再評価処理がExcel 2010 以降でより綿密に行われるようになり、処理量が増えています。

その結果、多くの条件付き書式が設定されているブックでは、Excel 2007 と比較して、この再評価処理に顕著に時間がかかります。

この動作は仕様変更の影響によるものです。

  

  

  

対処方法
----

  

Worksheet.EnableFormatConditionsCalculation プロパティをFalse に設定してから

貼り付けてください。

このプロパティをFalse に設定すると、条件付き書式の再評価が行われません。

  

Worksheet.EnableFormatConditionsCalculation プロパティ  
[https://msdn.microsoft.com/ja-jp/library/microsoft.office.tools.excel.worksheet.enableformatconditionscalculation.aspx](https://msdn.microsoft.com/ja-jp/library/microsoft.office.tools.excel.worksheet.enableformatconditionscalculation.aspx)

  

  

  

参考情報
----

  

本動作と直接的に関連性はなく、Excel 2010 向けの資料となりますが、

条件付き書式が多く含まれるブックにおけるパフォーマンス改善案を以下の資料で説明しています。

  
[2797222](https://support.microsoft.com/ja-jp/help/2797222)  
Excel 2010 で条件付き書式が多く含まれるブックを開き、コピーや貼り付けなど編集をおこなうと入力が遅延したりExcel が応答しない

  

  

  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**