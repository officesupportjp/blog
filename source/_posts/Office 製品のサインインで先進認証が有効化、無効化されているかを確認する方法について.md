---
title: Office 製品のサインインで先進認証が有効化、無効化されているかを確認する方法について
date: 2021-06-17
---

こんにちは、Office サポートの西川 (直) です。  
  
Office 製品のサインインで先進認証が有効化、無効化されているかを確認したいとお問い合わせいただくことがございます。

本投稿では、先進認証の有効化、無効化を確認する方法についてご案内いたします。

**【確認方法について】**

Office のサインイン画面を表示することで確認することが可能です。

**\- 先進認証が有効の場合 (以下の何れかが表示されます)**

**![](image1.png)![](image2.png)**

**\- 先進認証が無効の場合**

![](image3.png)

**【先進認証を有効化、無効化する方法について】**

**\- Office 2013 製品の場合**

Office 2013 製品では、既定で先進認証が無効です。

以下のレジストリキーを設定することで、先進認証を有効にしていただけます。

キー: HKEY\_CURRENT\_USER\\SOFTWARE\\Microsoft\\Office\\15.0\\Common\\Identity  
名前: EnableADAL  
種類: REG\_DWORD  
データ: 1

名前: Version  
種類: REG\_DWORD  
データ: 1

※ Office 2013 では、Microsoft Update により更新プログラムを最新に適用していただかないと、上記のレジストリを設定しても先進認証が有効にならない場合がございます。

※ 本レジストリキーは HKEY\_CURRENT\_USER 配下のため、レジストリエディターで値を編集される場合、別の管理者ユーザーなどでレジストリエディターを起動しないでください。

**\- Office 2016、2019、Microsoft 365 Apps (Office 365 ProPlus) 製品の場合**

Office 2016 以降となるこれらの製品では、既定で先進認証が有効です。

Microsoft 365 Apps 製品では将来的に廃止される可能性はございますが、以下のレジストリの設定により、先進認証を無効化していただけます。

キー: HKEY\_CURRENT\_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Common\\Identity  
名前: EnableADAL  
種類: REG\_DWORD  
データ: 0

※ 本レジストリキーは HKEY\_CURRENT\_USER 配下のため、レジストリエディターで値を編集される場合、別の管理者ユーザーなどでレジストリエディターを起動しないでください。

**【参考資料】**

EnableADAL の説明や各製品の詳細につきましては、以下の資料をご参考にいただきますようお願いいたします。

Title : 2013、Office 2016、および 2019 Office 2019 クライアント アプリOfficeのしくみ

URL : [https://docs.microsoft.com/ja-jp/microsoft-365/enterprise/modern-auth-for-office-2013-and-2016?view=o365-worldwide](https://docs.microsoft.com/ja-jp/microsoft-365/enterprise/modern-auth-for-office-2013-and-2016?view=o365-worldwide)

Title : Windows デバイスの Office 2013 の先進認証を有効にする[](https://docs.microsoft.com/ja-jp/microsoft-365/enterprise/modern-auth-for-office-2013-and-2016?view=o365-worldwide)URL : [https://docs.microsoft.com/ja-jp/microsoft-365/admin/security-and-compliance/enable-modern-authentication?view=o365-worldwide](https://docs.microsoft.com/ja-jp/microsoft-365/admin/security-and-compliance/enable-modern-authentication?view=o365-worldwide)

**\- 関連記事**  
  
Title : Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について  
URL : [https://social.msdn.microsoft.com/Forums/ja-JP/40ba1467-0b54-48d5-9db0-4d3aedca2b0a/office?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/40ba1467-0b54-48d5-9db0-4d3aedca2b0a/office?forum=officesupportteamja)

  

今回の投稿は以上です。

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**