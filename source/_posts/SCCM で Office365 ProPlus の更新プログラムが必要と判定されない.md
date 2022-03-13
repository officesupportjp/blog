---
title: SCCM で Office365 ProPlus の更新プログラムが必要と判定されない
date: 2019-03-09
lastupdate: 2020-10-16
id: cl0m82yjj002nvovs5qg4he0l
alias: /SCCM で Office365 ProPlus の更新プログラムが必要と判定されない/
---

(※ 2017 年 8 月 3 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、Office サポートの西川 (直)です。  

  

<span style="color:#ff0000">※ 2018/08/14 : 記事の構成を変更しました。</span>  

<span style="color:#ff0000">※ 2018/12/5 : 記事の構成を変更しました。</span>  

<span style="color:#ff0000">※ 2020/10/16 : 反映に関する説明を追記しました。</span> 

  

SCCM (更新プログラム 1602 以降) では、Office 365 ProPlus の更新プログラムを管理できるようになりました。

  

System Center Configuration Manager を使用して Office 365 ProPlus の更新プログラムを管理する  
[https://support.office.com/ja-jp/article/b4a17328-fcfe-40bf-9202-58d7cbf1cede](https://support.office.com/ja-jp/article/b4a17328-fcfe-40bf-9202-58d7cbf1cede)

  

本記事では、SCCM (System Center Configuration Manager) で Office365 ProPlus の更新プログラムが必要と判定されない現象について説明します。

  

<u>**現象**</u>  

オンプレミス環境の共有フォルダ等、ローカルソースより更新プログラムを配布している環境から、SCCM より更新プログラムを配布するように構成を変更した場合、SCCM クライアント上に該当の更新プログラムが表示されず、更新が適用できない場合があります。  

  

<u>**対処方法**</u>  

以下の手順を 1 から順に実施します。そして、コンピューターを再起動し、**最大 24 時間経過後、**SCCM や Office のスキャンが行われることで現象が改善する可能性があります。

  
  

**手順 1. "Office 365 クライアント管理" を有効にします。**  

System Center Configuration Manager を使用して Office 365 ProPlus の更新プログラムを管理する  
[https://docs.microsoft.com/ja-jp/DeployOffice/manage-updates-to-office-365-proplus-with-system-center-configuration-manager](https://docs.microsoft.com/ja-jp/DeployOffice/manage-updates-to-office-365-proplus-with-system-center-configuration-manager)

※ "Configuration Manager からの更新を受信できるように Office 365 クライアントを有効にする" をご確認ください

  
  

**手順 2.** CDNBaseUrl の値が、更新対象のチャネルの URL となっているか確認し、正しくない場合は修正します。  

キー : HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration  
名前 : CDNBaseUrl  
型 : REG\_SZ  
値 : チャネルの構成状況により、以下の何れかを設定します。  
  

月次チャネル:  
[http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60](http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60)

  

半期チャネル (対象指定) :  
[http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf](http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf)

  
  

半期チャネル:  
[http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114](http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114)

  

  

**手順 3.** グループポリシーにて “更新プログラムのパス” が設定されている場合は “未構成”、もしくは “無効” に設定します。  

場所：\[コンピューターの構成\] - \[管理用テンプレート\] – \[Microsoft Office 2016 (マシン)\] – \[更新\] – “更新プログラムのパス”

  

  

**手順 4**. Office 展開ツール (ODT) の Configuration.xml で UpdatePath 属性 (更新プログラムのパス) を設定している場合は、以下のレジストリを削除します。  

キー : HKEY\_LOCAL\_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration  
名前 : UpdateUrl

  

**手順 5.** グループポリシーにて “チャネルの更新” を “有効” とし、チャネルを設定します。  

場所：\[コンピューターの構成\] - \[管理用テンプレート\] – \[Microsoft Office 2016 (マシン)\] – \[更新\] – “チャネルの更新”  

  

※ 注意点 : “チャネルの更新”  を設定すると、SCCM もしくはインターネットから更新プログラムを取得するようになります。SCCM で管理しない端末に設定する際はご留意ください。  

<span style="color:#ff0000">※ Office 365 ProPlus のインストール時にチャネルが混在していない場合は、手順 5 は不要です。</span>  

  

*   **事象が改善されない場合**

  

Office 展開ツールを使用し、Configure オプションにより、Office を再度インストールしてください。  
[https://docs.microsoft.com/ja-jp/deployoffice/overview-of-the-office-2016-deployment-tool](https://docs.microsoft.com/ja-jp/deployoffice/overview-of-the-office-2016-deployment-tool)

  
  

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**