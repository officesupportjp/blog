---
title: Office のサインインのトラブルシュートについて
date: 2020-07-09
lastupdate: 2021-05-10
---

こんにちは、Office サポートの西川 (直) です。 

Office のサインインを行うときに、特定のユーザーだけサインインが成功しないとお問い合わせいただくことがございます。

この場合、以下の公開情報のとおり、サインインおよびライセンス情報のリセットをお試しいただくことで、解消する可能性がございます。

Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする

[https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state)

ただし、上記の公開情報には手順が多く含まれているため、以下を指針としてご対応ください。

1\. Office からサインアウトし、Office を全て終了します。

2\. 公開情報内の "OLicenseCleanup.vbs" をご実施ください

※ コマンドプロンプトを右クリック→管理者権限で起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

※ ログオンユーザーに管理者権限がない場合、スクリプトを 2 回実行する必要があります。ログオンユーザーではダブルクリックで OLicenseCleanup.vbs を実行します。その後、コマンドプロンプトを右クリック→管理者権限として別の管理者ユーザーのアカウントで起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

<以降、[クイック実行形式 (C2R) の Office 2016/2019/Office 365 ProPlus (Microsoft 365 Apps) のみ](https://social.msdn.microsoft.com/Forums/ja-JP/57e5d81e-ef69-4c1f-9ef0-932d03d0e7ce/1246312452124831246323455348922441824335-c2r-12392-windows?forum=officesupportteamja)\>

3\. 公開情報内の "社内参加から資格情報をクリアする" のとおり、\[職場または学校にアクセスする\] からアカウントを切断してください。

4\. (1 の手順でサインアウトを実施できていない場合)"Signoutofwamaccounts.ps1" をご実施ください。

  もし、"Signoutofwamaccounts.ps1" が文字化けしている場合、後述の内容をお試しください。

5\. OS からサインアウトしてください

6\. OS にサインインし、Office にサインインできるかをお試しください。

■ 補足 : \[職場または学校にアクセスする\] のアカウントの削除について

\===========================

Windows 10 上で Microsoft 365 Apps for enterprise 等のクイック実行版の Office をご利用されている場合、Windows 10 の「職場または学校にアクセスする」に登録されているアカウントを使用する動きを行う場合があります。

「職場または学校にアクセスする」に何らかの問題が生じている場合、以下の手順をお試しください。

\- 手順

1\. 以下のサイトから WPJCleanUp.zip を対象のクライアント端末にダウンロードします。  
[https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip](https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip)  

2\. クライアント端末に ＜対象ユーザー＞ でログオンし、WPJCleanUp.zip を任意のフォルダに展開します。

3\. CleanupTool に含まれている WPJCleanp.cmd を実行します。

※ダブルクリックにて実行時、\[Windows によって PC が保護されました\] と表示された場合は \[詳細情報\] をクリックし、表示された \[実行\] をクリックします。  
4\. クライアント端末を再起動します。  

  

**\- signoutofwamaccounts.ps1**

```
if(-not \[Windows.Foundation.Metadata.ApiInformation,Windows,ContentType=WindowsRuntime\]::IsMethodPresent("Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager", "FindAllAccountsAsync"))
{
    throw "This script is not supported on this Windows version. Please, use CleanupWPJ.cmd."
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime

Function AwaitAction($WinRtAction) {
  $asTask = (\[System.WindowsRuntimeSystemExtensions\].GetMethods() | ? { $\_.Name -eq 'AsTask' -and $\_.GetParameters().Count -eq 1 -and !$\_.IsGenericMethod })\[0\]
  $netTask = $asTask.Invoke($null, @($WinRtAction))
  $netTask.Wait(-1) | Out-Null
}

Function Await($WinRtTask, $ResultType) {
  $asTaskGeneric = (\[System.WindowsRuntimeSystemExtensions\].GetMethods() | ? { $\_.Name -eq 'AsTask' -and $\_.GetParameters().Count -eq 1 -and $\_.GetParameters()\[0\].ParameterType.Name -eq 'IAsyncOperation\`1' })\[0\]
  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  $netTask.Result
}

$provider = Await (\[Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime\]::FindAccountProviderAsync("https://login.microsoft.com", "organizations")) (\[Windows.Security.Credentials.WebAccountProvider,Windows,ContentType=WindowsRuntime\])

$accounts = Await (\[Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime\]::FindAllAccountsAsync($provider, "d3590ed6-52b3-4102-aeff-aad2292ab01c")) (\[Windows.Security.Authentication.Web.Core.FindAllAccountsResult,Windows,ContentType=WindowsRuntime\])

$accounts.Accounts | % { AwaitAction ($\_.SignOutAsync("d3590ed6-52b3-4102-aeff-aad2292ab01c")) }
```

**\- 関連記事**

Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について

[https://social.msdn.microsoft.com/Forums/ja-JP/40ba1467-0b54-48d5-9db0-4d3aedca2b0a/office?forum=officesupportteamja](https://social.msdn.microsoft.com/Forums/ja-JP/40ba1467-0b54-48d5-9db0-4d3aedca2b0a/office?forum=officesupportteamja)

**2020/10/28 :** **\[職場または学校にアクセスする\] のアカウントの削除についてを追加しました。**

**2021/5/10 : ログオンユーザーに管理者権限がない場合の手順を変更しました。**

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**