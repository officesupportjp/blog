---
title: Office のサインインのトラブルシュートについて
date: 2020-07-09
lastupdate: 2021-05-10
id: cl0m6umvg001gi4vsgppfcmk0
alias: /Office のサインインのトラブルシュートについて/
---

こんにちは、Office サポートの西川 (直) です。 

Office のサインインを行うときに、特定のユーザーだけサインインが成功しないとお問い合わせいただくことがございます。


この場合、以下の公開情報のとおり、サインインおよびライセンス情報のリセットをお試しいただくことで、解消する可能性がございます。

Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする

[https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state)

ただし、上記の公開情報には手順が多く含まれているため、以下を指針としてご対応ください。

<br>

**-手順**

1\. Office からサインアウトし、Office、Teams、OneDrive 同期クライアントを全て終了します。

2\. 公開情報内の "OLicenseCleanup.vbs" をご実施ください

※ コマンドプロンプトを右クリック→管理者権限で起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

※ ログオンユーザーに管理者権限がない場合、スクリプトを 2 回実行する必要があります。ログオンユーザーではダブルクリックで OLicenseCleanup.vbs を実行します。その後、コマンドプロンプトを右クリック→管理者権限として別の管理者ユーザーのアカウントで起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

<br>

<以降、[クイック実行形式 (C2R) の Office のみご実施ください](https://officesupportjp.github.io/blog/%E3%82%AF%E3%82%A4%E3%83%83%E3%82%AF%E5%AE%9F%E8%A1%8C%E5%BD%A2%E5%BC%8F%20(C2R)%20%E3%81%A8%20Windows%20%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%A9%E3%83%BC%E5%BD%A2%E5%BC%8F%20(MSI)%20%E3%82%92%E8%A6%8B%E5%88%86%E3%81%91%E3%82%8B%E6%96%B9%E6%B3%95/)\>

3\. 公開情報内の "社内参加から資格情報をクリアする" のとおり、\[職場または学校にアクセスする\] からアカウントを切断してください。
※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

![](image01.png)  

4\. (1 の手順でサインアウトを実施できていない場合)"Signoutofwamaccounts.ps1" をご実施ください。

  もし、"Signoutofwamaccounts.ps1" が文字化けしている場合、本記事下部のスクリプトをご利用ください。


5\. 以下の 1 から 5 の手順を実施します。

[https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/sign-in-issues](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/sign-in-issues)

```
1. エクスプローラーを開き、アドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts
2. Ctrl + A キーを押して、すべてを選択します。
3. 選択したファイルを右クリックして、[削除] を選択します。
4. エクスプローラーのアドレス バーに次の場所を入力します。%LOCALAPPDATA%\Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts
5. すべてのファイルを選択して削除します。
```

6\. OS からサインアウト、または、OS を再起動してください

7\. OS にサインインし、Office にサインインできるかをお試しください。

<br>

**■ 補足 : \[職場または学校にアクセスする\] のアカウントの削除について**

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

<br>

**\- signoutofwamaccounts.ps1**

```
if(-not [Windows.Foundation.Metadata.ApiInformation,Windows,ContentType=WindowsRuntime]::IsMethodPresent("Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager", "FindAllAccountsAsync"))
{
    throw "This script is not supported on this Windows version. Please, use CleanupWPJ.cmd."
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime

Function AwaitAction($WinRtAction) {
  $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod })[0]
  $netTask = $asTask.Invoke($null, @($WinRtAction))
  $netTask.Wait(-1) | Out-Null
}

Function Await($WinRtTask, $ResultType) {
  $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  $netTask.Result
}

$provider = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime]::FindAccountProviderAsync("https://login.microsoft.com", "organizations")) ([Windows.Security.Credentials.WebAccountProvider,Windows,ContentType=WindowsRuntime])

$accounts = Await ([Windows.Security.Authentication.Web.Core.WebAuthenticationCoreManager,Windows,ContentType=WindowsRuntime]::FindAllAccountsAsync($provider, "d3590ed6-52b3-4102-aeff-aad2292ab01c")) ([Windows.Security.Authentication.Web.Core.FindAllAccountsResult,Windows,ContentType=WindowsRuntime])

$accounts.Accounts | % { AwaitAction ($_.SignOutAsync("d3590ed6-52b3-4102-aeff-aad2292ab01c")) }
```

<br>

**\- 関連記事**

Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について  
[https://officesupportjp.github.io/blog/Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について](https://officesupportjp.github.io/blog/Office%20%E3%81%AE%E3%82%B5%E3%82%A4%E3%83%B3%E3%82%A4%E3%83%B3%E3%81%A7%E3%83%8D%E3%83%83%E3%83%88%E3%83%AF%E3%83%BC%E3%82%AF%E6%8E%A5%E7%B6%9A%E3%81%8C%E3%81%AA%E3%81%84%E6%97%A8%E3%81%AE%E3%83%A1%E3%83%83%E3%82%BB%E3%83%BC%E3%82%B8%E3%81%8C%E8%A1%A8%E7%A4%BA%E3%81%95%E3%82%8C%E3%82%8B%E4%BA%8B%E8%B1%A1%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/)

<br>

**2020/10/28 :** **\[職場または学校にアクセスする\] のアカウントの削除についてを追加しました。**

**2021/5/10 : ログオンユーザーに管理者権限がない場合の手順を変更しました。**

**2022/3/16 : キャッシュ削除手順を追加しました。**


**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**