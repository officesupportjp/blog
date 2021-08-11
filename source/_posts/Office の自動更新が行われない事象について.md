---
title: Office の自動更新が行われない事象について
date: 2019-12-19
---

こんにちは、Office サポートの西川 (直) です。  
  
Office 365 クライアントや Office 2019 を展開している環境で、自動更新が行われないとお問い合わせいただくことがございます。  
以下の項目をご確認いただくことで解消する可能性がございますので、ご参考にしていただけますと幸いでございます。  
  
1\. インストール、または、更新した後、24 時間が経過していない  
\----------------------------  
Office は、連続した更新を避けるため、インストール、または、更新が完了した後、24 時間の間は次の更新を行わない動作となっております。  
そのため、インストール、または、更新が完了した後、24 時間が経過していない場合、1 日ほど経過をご観察いただきますようお願いいたします。  
  
\- 参考  
How to manage Office 365 ProPlus Channels for IT Pros  
[https://techcommunity.microsoft.com/t5/Office-365-Blog/How-to-manage-Office-365-ProPlus-Channels-for-IT-Pros/ba-p/795813](https://techcommunity.microsoft.com/t5/Office-365-Blog/How-to-manage-Office-365-ProPlus-Channels-for-IT-Pros/ba-p/795813)  
  
  
2\. ”自動更新” が無効になっている、もしくは、"更新プログラムのパス" が正しく構成されていない  
\----------------------------  
Office は更新プログラムの取得先をグループポリシーや Office 展開ツールにより構成することが可能です。  
以下の設定をご確認いただき、意図した構成となっているかをご確認いただきますようお願いいたします。  
  
Office 365 ProPlus の更新設定を構成する  
[https://docs.microsoft.com/ja-jp/deployoffice/configure-update-settings-for-office-365-proplus](https://docs.microsoft.com/ja-jp/deployoffice/configure-update-settings-for-office-365-proplus)  
\*UpdatePath  
  
Office 展開ツールのオプションの構成  
[https://docs.microsoft.com/ja-jp/deployoffice/configuration-options-for-the-office-2016-deployment-tool](https://docs.microsoft.com/ja-jp/deployoffice/configuration-options-for-the-office-2016-deployment-tool)  
\*UpdatePath  
  
  
3\. "Office 365 クライアント管理" (OfficeMgmtCOM) が構成されている  
\----------------------------  
"Office 365 クライアント管理" (OfficeMgmtCOM) が構成されている場合、Office は自動更新を行いません。  
Office を起動し、\[ファイル\] タブ - "アカウント" に "更新プログラムはシステム管理者が管理します" と記載されている場合、"Office 365 クライアント管理" (OfficeMgmtCOM) が構成されていますので、設定を解除してください。  
  
\- 参考  
Microsoft Endpoint Configuration Manager を使用して Office 365 ProPlus の更新プログラムを管理する  
[https://docs.microsoft.com/ja-jp/deployoffice/manage-office-365-proplus-updates-with-configuration-manager](https://docs.microsoft.com/ja-jp/deployoffice/manage-office-365-proplus-updates-with-configuration-manager)  
  
  
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。