---
title: >-
  Office 2010-2013 (既定) と Office 365 ProPlus の共存環境で Office 365 ProPlus をバージョン
  1710 に更新すると Excel の関連付けが Office 365 ProPlus に変更される事象について
date: '2019-03-09'
id: cl0m7opbe0017dcvs06fk2wqc
tags:
  - インストール

---

(※ 2017 年 11 月 29 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、Office サポートの西川 (直) です。 本記事では、表題の事象についてご説明いたします。  

  

<u>**現象**</u>

  

Office 2010/2013 (既定) と Office 365 ProPlus の共存環境で Office 365 ProPlus バージョンを 1710 に更新すると Excel の関連付けが Office 365 ProPlus に変更されます。

  

  

再現手順例 :

  

1.  Office 2010 (既定)、Office 365 ProPlus (バージョン 1709 (ビルド 8528.2147)) が共存している環境を用意します
2.  Office 365 ProPlus を バージョン 1710 (ビルド 8625.2132) に更新します
3.  Excel ファイルの関連付けが Office 365 ProPlus の Excel 2016 に変更されます。これにより、Excel ファイルをダブルクリックすると、Excel 2010 ではなく Office 365 ProPlus の Excel 2016 が起動します  
    

  

  

<u>**理由**</u>

  

Version 1710 で Excel のファイルの関連付けに関する仕様変更 (機能拡張) が行われているため、更新時に関連付けが変更されます

  

  

<u>**Excel ファイルの関連付けを Office 2010/2013 に戻す方法**</u>

  

Office 2010/2013で、修復インストールを実施します。

  

  

Title : Office アプリケーションを修復する  
URL : [https://support.office.com/ja-jp/article/7821d4b6-7c1d-4205-aa0e-a6b40c5bb88b](https://support.office.com/ja-jp/article/7821d4b6-7c1d-4205-aa0e-a6b40c5bb88b)

  

  

<u>**補足**</u>

  

Office を更新する際、その更新プログラムに依存して、そのバージョンの Office を正しく使用できるように修復する処理が実行されます。この動きは想定された動作です。

  

もし、何らかの理由で Office の設定情報等が正しくない状態の場合、この仕組みにより、Office の更新時には正しい状態に修復されます。

  

  

  

<u>**参考**</u>

  

Title : 異なるバージョンの Office を同じ PC にインストールして使う
URL : [https://support.office.com/ja-jp/article/6ebb44ce-18a3-43f9-a187-b78c513788bf](https://support.office.com/ja-jp/article/6ebb44ce-18a3-43f9-a187-b78c513788bf)

  

Title : Office 2013 スイートおよび Office 2013 プログラム (MSI による展開) を、他のバージョンの Office を実行しているコンピューターで使用する方法について
URL : [https://support.microsoft.com/ja-jp/help/2784668/how-to-use-office-2013-suites-and-programs-msi-deployment-on-a-compute](https://support.microsoft.com/ja-jp/help/2784668/how-to-use-office-2013-suites-and-programs-msi-deployment-on-a-compute)

  

※ Office 2016 でも参考にしていただけます。

  

  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**