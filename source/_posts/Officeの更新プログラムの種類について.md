---
title: Officeの更新プログラムの種類について
date: '2024-09-02'
tags:
  - 更新

---

こんにちは、Office サポートの西川 (直) です。 
本記事では Office の更新プログラムの種類について解説します。

<br>

Office の更新プログラムの違いについて
--
新機能が追加される機能更新が存在するのは Microsoft 365 Apps クライアントのみです。

| バージョン | 機能更新 |セキュリティ更新 | セキュリティ以外の更新(修正) |
| :----: | :----: | :----: | :----: |
| Microsoft 365 Apps クライアント| ○ | ○ | ○ |
| Office 2019/2021 LTSC | - | ○ | ○ |
| パッケージ版 Office 2016/2019/2021 | - | ○ | ○ |
| Office 2016 | - | ○ | ○ |

<br>

Microsoft 365 Apps クライアント
---
更新は[チャネル](https://learn.microsoft.com/ja-jp/microsoft-365-apps/updates/overview-update-channels)に沿ってバージョン、ビルドを上げることで行われます。また、過去の変更は累積となります。
このため、Excel のみ等、特定の製品のみ更新することは出来ません。また、累積のため、バージョン 2302 から 2308 を飛ばして 2402 に直接更新することも可能であり、2308 の変更も含まれます。

![https://learn.microsoft.com/ja-jp/officeupdates/update-history-microsoft365-apps-by-date](image01.png)

機能更新、セキュリティ更新、セキュリティ以外の更新(修正) の 3 種類がありますが、表のように、チャネルに沿ってバージョンアップを伴う更新とビルドのみの更新の 2 つに分けて公開されています。

同日に複数の更新プログラムが公開されている理由は、過去のバージョンから、機能更新を行わず、セキュリティ更新、セキュリティ以外の更新(修正)のみの更新をサポートするためです。

**1) 機能更新**
**バージョンが上がる更新**が機能更新です。これには機能更新、セキュリティ更新、セキュリティ以外の更新(修正)が含まれます。
例えば、最新チャネルの 6/26 において、バージョンが 2405 から 2406 に上がっておりますので、これは機能更新となります。

![](image02.png)

**2) セキュリティ更新**
**ビルドのみ上がる更新**がセキュリティ更新です。これにはセキュリティ更新、セキュリティ以外の更新(修正)が含まれます。
例えば、最新チャネルの 6/19 では、バージョンは 2405 のまま、ビルドのみ上がっておりますので、これはセキュリティ更新、セキュリティ以外の更新(修正)となります。

![](image03.png)

<br>


**MECM上の表記について**
機能更新は "機能更新プログラム"、セキュリティ更新は "品質更新プログラム" または "拡張品質更新プログラム" と表示されます。

"拡張品質更新プログラム"は、最新バージョンではないセキュリティ更新、および、セキュリティ以外の更新を指します。


![](image04.png)

例えば、上記の例では、バージョン 2402 は "品質更新プログラム" であり、2308 と 2302 は "拡張品質更新プログラム" となります。



<br>


Office 2019/2021 LTSC
---
ボリュームライセンス版の Office 2019/2021 LTSC では、セキュリティ更新、セキュリティ以外の更新(修正)のみが公開されています。
これらの Office はそれぞれ、Microsoft 365 Apps のバージョン 1808、2108 を元に製造されており、基本的な考え方は同一となります。

![https://learn.microsoft.com/ja-jp/officeupdates/update-history-office-2019](image06.png)
![https://learn.microsoft.com/ja-jp/officeupdates/update-history-office-2021](image05.png)



<br>

パッケージ版 Office 2016/2019/2021
---
これらの Office は Microsoft 365 Apps を元に製造されており、基本的な考え方は同一となりますが、
バージョンが上がったとしても機能更新は含まれず、セキュリティ更新、セキュリティ以外の更新(修正)のみが公開されています。

![https://learn.microsoft.com/ja-jp/officeupdates/update-history-office-2019](image07.png)
![https://learn.microsoft.com/ja-jp/officeupdates/update-history-office-2021](image08.png)



<br>

Office 2016
---
ボリュームライセンス版の Office 2016 では、セキュリティ更新、セキュリティ以外の更新(修正)のみが公開されています。
こちらの Office は、[クイック実行形式 (C2R) ではなく Windows インストーラー形式 (MSI)](https://officesupportjp.github.io/blog/cl0m8t2dj00393gvs3r7oar2m/) であり、モジュール単位で更新プログラムが公開されています。

![https://learn.microsoft.com/ja-jp/officeupdates/msp-files-office-2016](image09.png)


**MECM上の表記について**
C2R形式の Office の更新プログラムは、Microsoft Update Catalog 上の "分類" は "更新" で統一されておりますが、
MSI形式では "セキュリティ更新" は "セキュリティ問題の修正プログラム"、"セキュリティ以外の更新(修正)" は "重要な更新" と表記されます。
![](image10.png)

<br>

**\- 関連情報**
[クイック実行形式 (C2R) と Windows インストーラー形式 (MSI) を見分ける方法](https://officesupportjp.github.io/blog/cl0m8t2dj00393gvs3r7oar2m/)

<br>


**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**