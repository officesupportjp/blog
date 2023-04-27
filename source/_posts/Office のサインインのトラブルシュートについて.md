---
title: Office のサインインのトラブルシュートについて
date: '2020-07-09'
lastupdate: '2023-04-27'
id: cl0m6umvg001gi4vsgppfcmk0
tags:
  - サインイン、認証

---

こんにちは、Office サポートの西川 (直) です。 

特定のユーザーだけ Office のサインインが成功しないとお問い合わせいただくことがございます。
この場合、[Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) の内容にて解消する可能性がございます。

しかし、Office には様々なバージョンがあり、この資料の他にも後述の**参考情報**のような情報もございますので、
ご利用されている製品によって、以下を指針としてご対応ください。

<br>

Office 2013、2016 の場合
--
1. Office からサインアウトし、Office を終了します。
2. [OLicenseCleanup.zip](https://download.microsoft.com/download/e/1/b/e1bbdc16-fad4-4aa2-a309-2ba3cae8d424/OLicenseCleanup.zip) をダウンロードして任意のフォルダに解凍します。
3. ダブルクリックで OLicenseCleanup.vbs を実行します。処理はバックグラウンドで終了いたします。
※ 管理者権限は不要です。他の管理者ユーザーで実行すると正しくキャッシュが削除出来ませんので事象が発生しているログオンユーザーの権限で実行してください。
4. Office のサインインの事象が解消するか、ご確認ください。

<br>

Office 2019 以降、または、Microsoft 365 Apps の場合
--
1. Office からサインアウトし、Office を終了します。Teams、OneDrive 同期クライアントも全て終了します。
2. [OLicenseCleanup.zip](https://download.microsoft.com/download/e/1/b/e1bbdc16-fad4-4aa2-a309-2ba3cae8d424/OLicenseCleanup.zip) をダウンロードして任意のフォルダに解凍します。
3. 環境によって以下の対応を実施します。

**Office 2019 以降の場合**
ダブルクリックで OLicenseCleanup.vbs を実行します。処理はバックグラウンドで終了いたします。
※ 管理者権限は不要です。他の管理者ユーザーで実行すると正しくキャッシュが削除出来ませんので事象が発生しているログオンユーザーの権限で実行してください。

**Microsoft 365 Apps の場合**
- ログオンユーザーに管理者権限がある場合
コマンドプロンプトを右クリック→管理者権限で起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

- ログオンユーザーに管理者権限がないけれども別の管理者ユーザーが利用できる場合
スクリプトを 2 回実行する必要があります。ログオンユーザーではダブルクリックで OLicenseCleanup.vbs を実行します。その後、コマンドプロンプトを右クリック→管理者権限として別の管理者ユーザーのアカウントで起動し、コマンドプロンプト上で OLicenseCleanup.vbs を実行してください。実行しても完了ダイアログ等は表示されず、処理はバックグラウンドで終了いたします。

- ログオンユーザーに管理者権限がなく別の管理者ユーザーも利用できない場合
まずはダブルクリックで OLicenseCleanup.vbs を実行してください。事象が解消しない場合は管理者権限での実行もご検討ください


4. 以下の \[職場または学校にアクセスする\] からアカウントを切断してください。
※ Azure AD や AD に接続済みの項目を "切断" しますと、ドメインから外れてしまいますので、Azure AD や AD に接続済み の項目は切断しないようにご注意ください。

![](image01.png)  

5. 以下のフォルダ配下のファイルを全て削除します。

`%LOCALAPPDATA%\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts`
`%LOCALAPPDATA%\Packages\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\AC\TokenBroker\Accounts`
`%LOCALAPPDATA%\Microsoft\TokenBroker\Cache`
`%LOCALAPPDATA%\Microsoft\OneAuth`
`%LOCALAPPDATA%\Microsoft\IdentityCache`
※ 事象が発生しているユーザーにてご実施ください。

6. OS からサインアウト、または、OS を再起動してください
7. Office のサインインの事象が解消するか、ご確認ください。

<br>

Office 2019 以降、または、Microsoft 365 Apps で事象が解消しない場合
--
1. Office、Teams、OneDrive 等のアプリケーションを終了します。
2. ログオンユーザーの権限で PowerShell を起動し、以下のコマンドを実施します。
※ 管理者権限は不要です。また、管理者権限を有する別のユーザーで実施しますと正しく処理が行われません。

```
Add-AppxPackage -Register "C:\Windows\SystemApps\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\Appxmanifest.xml" -DisableDevelopmentMode -ForceApplicationShutdown
```

なお、事象がコンシューマー向けの Microsoft アカウントで生じている場合には以下のコマンドを実施します。

```
Add-AppxPackage -Register "C:\Windows\SystemApps\Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy\Appxmanifest.xml" -DisableDevelopmentMode -ForceApplicationShutdown
```

3. OS を再起動します。
4. Office のサインインの事象が解消するか、ご確認ください。


<br>

**補足**
--
- Office Home & Business 2016 等、コンシューマー製品の場合には "Office 2019 以降の場合" と同じ対処をご実施ください。

- 「職場または学校にアクセスする」に何らかの問題が生じている、または、コマンド等でアカウントを切断する場合、以下の手順をお試しください。
1. 以下のサイトから [WPJCleanUp.zip](https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip) を対象のクライアント端末にダウンロードします。  
2. クライアント端末に ＜対象ユーザー＞ でログオンし、WPJCleanUp.zip を任意のフォルダに展開します。
3. CleanupTool に含まれている WPJCleanp.cmd を実行します。
※ダブルクリックにて実行時、\[Windows によって PC が保護されました\] と表示された場合は \[詳細情報\] をクリックし、表示された \[実行\] をクリックします。  

- [Microsoft 365 Apps for enterprise のライセンス認証の状態をリセットする](https://docs.microsoft.com/ja-jp/office/troubleshoot/activation/reset-office-365-proplus-activation-state) の Signoutofwamaccounts.ps1 が文字化けしている場合、以下をご利用ください。

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

**\- 関連情報**
[Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について](https://officesupportjp.github.io/blog/cl0m75al4001gmcvse36w64dv/index.html)

<br>

**\- 参考情報**
[Microsoft 365 Apps をライセンス認証する際のサインインの問題](https://learn.microsoft.com/ja-jp/office/troubleshoot/activation/sign-in-issues)

[デバイス ID とデスクトップ仮想化](https://learn.microsoft.com/ja-jp/azure/active-directory/devices/howto-device-identity-virtual-desktop-infrastructure)

[Microsoft 365 サービスに接続しようとしたときの Office アプリケーションの認証の問題を修正します](https://learn.microsoft.com/ja-jp/microsoft-365/troubleshoot/authentication/automatic-authentication-fails)

[テナントの移行後、OneDrive for Business を同期できない](https://learn.microsoft.com/ja-jp/sharepoint/troubleshoot/sync/cant-sync-after-migration)

<br>

<span style="color:#ff0000">**2020/10/28  Update**</span>  
<span style="color:#339966">\[職場または学校にアクセスする\] のアカウントの削除についてを追加しました</span>

<span style="color:#ff0000">**2021/5/10  Update**</span>  
<span style="color:#339966">\ログオンユーザーに管理者権限がない場合の手順を変更しました。</span>

<span style="color:#ff0000">**2022/3/16  Update**</span>  
<span style="color:#339966">キャッシュ削除手順を追加しました。</span>

<span style="color:#ff0000">**2023/2/22  Update**</span>  
<span style="color:#339966">記事を簡略化しました。</span>

<span style="color:#ff0000">**2023/4/26  Update**</span>  
<span style="color:#339966">キャッシュ削除内容を追加しました。</span>

<span style="color:#ff0000">**2023/4/27  Update**</span>  
<span style="color:#339966">キャッシュ削除内容を追加しました。</span>

**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**