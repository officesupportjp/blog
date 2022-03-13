---
title: Office 365 ProPlus で言語パックの追加に失敗する
date: 2019-04-11
lastupdate: 2019-12-19
id: cl0m75al0001amcvs0tsdeksn
alias: /Office 365 ProPlus で言語パックの追加に失敗する/
---

こんにちは、Office サポートの西川 (直) です。  
  
本記事では、Office 365 ProPlus で言語パックの追加に失敗する事象について説明します。  
  
<u>**\- 現象**  </u>
以下の条件を全て満たす場合、言語パックの追加に失敗する可能性があります。  
  
**条件 1 :**  
Office 365 ProPlus をインストール済みで、後から言語パックを追加する  
  
**条件 2 :**  
インストール済みの Office 365 ProPlus と同一ビルド番号の言語パックを追加する  
  
**条件 3 :**  
言語パックのインストールソースがインターネットではなくローカルフォルダであり、そのフォルダにはインストール済みの Office 365 ProPlus のインストールソースが含まれていない  
  
<u>**2019/12/19** **追記**</u>

<u>**条件 4:**</u>  
<u>**Office 365 クライアント バージョン 1905 未満に更新する**</u>

  
<u>**\- 回避策**</u>  
言語パックのインストールソースに、インストール済みの Office 365 ProPlus のインストールソースの cab ファイルを含めてください。

<u>**2019/12/19** **追記**</u>

<u>Office 展開ツールの Configuration.xml に AllowCdnFallback 属性を指定することでも回避できます。</u>

<u>Office 365 ProPlus での言語の展開の概要</u>
[https://docs.microsoft.com/ja-jp/deployoffice/overview-of-deploying-languages-in-office-365-proplus](https://docs.microsoft.com/ja-jp/deployoffice/overview-of-deploying-languages-in-office-365-proplus)

<u>**\- 補足**</u>  
インストールソースとは、Office 展開ツールの /download オプションにより取得できるモジュールです。

<u>**\- 注意点**</u>

本記事では Office 365 ProPlus を対象としていますが、C2R 版の Office であれば該当する可能性があります。

  
<u>**\- 参考**</u>  
Title : Office 展開ツールの概要  
URL : [https://docs.microsoft.com/ja-jp/deployoffice/overview-of-the-office-2016-deployment-tool](https://docs.microsoft.com/ja-jp/deployoffice/overview-of-the-office-2016-deployment-tool)  
  
Title : Office 365 ProPlus の更新履歴 (日付別の一覧)  
URL : [https://docs.microsoft.com/ja-jp/officeupdates/update-history-office365-proplus-by-date](https://docs.microsoft.com/ja-jp/officeupdates/update-history-office365-proplus-by-date)  

本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。