---
title: Office 365 Sway を PowerShell を使用して指定ユーザー単位で有効化、無効化 (停止) する方法について
date: 2019-02-08
lastupdate: 2019-03-22
id: cl0n10de1001d90vs6xzb5ynb
alias: /Office 365 Sway を PowerShell を使用して指定ユーザー単位で有効化、無効化 (停止) する方法について/
---

(※ 2016 年 8 月 16 日に Office Support Team Blog JAPAN に公開した情報のアーカイブです。)  
  
こんにちは、Office サポートの橋村です。  
Office 365 で公開されている Sway アプリケーションについて、以前に PowerShell スクリプトを使用して、テナント全体のユーザーのライセンスを一括して有効化/無効化する方法をご案内いたしました。本記事では、テナント全体のユーザーではなく、指定したユーザーのみのライセンスを有効化/無効化する方法をご案内いたします。  
  
**目次**  
  
1\. Sway のライセンスについて  
2\. PowerShell のインストール  
3\. ユーザーリストの作成  
4\. PowerShell で指定ユーザーのみを有効化する方法  
5\. PowerShell で指定ユーザーのみを無効化する方法  
6\. ユーザーリストをスクリプトファイルと別のディレクトリに格納して実行する方法  
7\. 関連情報  
  
  
**1\. Sway のライセンスについて**  
Sway のライセンスはユーザー単位で有効化/無効化を設定します。例えば、Office 365 E3 の場合は、下記の Office 365 管理センターのユーザー管理のページにて、ライセンスの詳細項目を表示することで確認、もしくは、設定することができます。  

\[Office 365 管理センター\] - \[ユーザー\] - \[アクティブなユーザー\] - \[製品ライセンス\]の編集

  
**2\. PowerShell のインストール**  
PowerShell を利用するには、以下のページに従って PowerShell の環境を構築する必要があります。  
  
タイトル : PowerShell の導入について  
アドレス : [https://officesupportjp.github.io/blog/PowerShell の導入について/](https://officesupportjp.github.io/blog/PowerShell%20%E3%81%AE%E5%B0%8E%E5%85%A5%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/)  
  
  
**3\. ユーザーリストの作成**  
テキストファイルを用いて、Sway ライセンスを有効化/無効化したいユーザーのリストを作成します。リスト構造はシンプルで、指定ユーザーのアカウントを改行して入力するだけです。  

```
user01@testtenant.onmicrosoft.com
user02@testtenant.onmicrosoft.com
user03@testtenant.onmicrosoft.com
user04@testtenant.onmicrosoft.com
user05@testtenant.onmicrosoft.com
```

このリストに入力されているユーザーのみを対象に、後述の Sway ライセンスの有効化/無効化処理が実行されます。今回の例では、作成するリストを **UserList.txt** というファイル名で保存します。  
  
  
**4\. PowerShell で指定ユーザーのみを有効化する方法**  
下記のソースコードをコピーし、ps1 ファイルとして保存します。今回の例では、作成するスクリプトを **EnableSwayLicense\_UserUnit.ps1** というファイル名で保存します。  

```
# サービス名の設定
$service_name = "SWAY"

connect-msolservice -credential $msolcred

# 改行単位で UPN が記述されたユーザー一覧から取得
$users = New-Object System.Collections.ArrayList
$list = Get-Content "UserList.txt"

foreach ($temp in $list) {
    $users = $users + (Get-MsolUser -UserPrincipalName $temp)
}

# ユーザー毎のライセンス/サービスの確認と変更
foreach ($user in $users) {
    write-host ("Processing " + $user.UserPrincipalName)
    $licensetype = $user | Select-Object -ExpandProperty Licenses | Sort-Object { $_.Licenses }

    foreach ($license in $user.Licenses) {
        write-host (" " + $license.accountskuid)

        $includeService = $false
        $disableplan = @()

        foreach ($row in $($license.ServiceStatus)) {
            if ( $row.ServicePlan.ServiceName -eq $service_name ) { $includeService=$true }
            if ( $row.ProvisioningStatus -eq "Disabled" -and $row.ServicePlan.ServiceName -ne $service_name) {
                $disableplan += $row.ServicePlan.ServiceName
            }
        }

        # 指定したサービスが含まれる SKU の場合、現在の設定に加え、指定したサービスを有効化 (Enable)
        if ($includeService) {
            write-host (" found Target service in " + $license.accountskuid)
            write-host (" current disabled services : " + $disableplan )
            $x = New-MsolLicenseOptions -AccountSkuId $license.accountskuid -DisabledPlans $disableplan
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $x
        }
    }
}
```

<スクリプトファイルの実行>  
スクリプトファイルを実行する前に、先ほど作成したユーザーリスト (**UserList.txt**) を ps1ファイルを保存したフォルダに格納します。(**EnableSwayLicense\_UserUnit.ps1** と **UserList.txt** を同一ディレクトリに格納します。) その後、PowerShell を起動し、ps1 ファイルを保存したフォルダに移動し、下記の様に実行します。(下記は c:\\temp 配下に ps1 ファイルを格納した場合の例です。)  

```
cd c:\temp
.\EnableSwayLicense_UserUnit.ps1
```

認証ダイアログが起動しますので、ご利用中のテナントの管理者アカウント、および、パスワードを入力します。  
  
有効化処理の内容が表示されますので、PowerShell のプロンプトが戻るまで待ちます。  
  
処理完了後、ブラウザにて Office 365 を表示していた場合は一度ブラウザを終了させ、再度 Office 365 にサインインします。  
  
ユーザーのライセンス付与状況を確認し、Sway 製品のライセンスがオンに設定されていることを確認します。また、アプリケーション ランチャー上の Sway のタイルが表示されていることを確認します。  
  
  
**5\. PowerShell で指定ユーザーのみを無効化する方法**  
下記のソースコードをコピーし、ps1 ファイルとして保存します。今回の例では、作成するスクリプトを **DisableSwayLicense\_UserUnit.ps1** というファイル名で保存します。  

```
# サービス名の設定
$service_name = "SWAY"

connect-msolservice -credential $msolcred

# 改行単位で UPN が記述されたユーザー一覧から取得
$users = New-Object System.Collections.ArrayList
$list = Get-Content "UserList.txt"

foreach ($temp in $list) {
    $users = $users + (Get-MsolUser -UserPrincipalName $temp)
}

# ユーザー毎のライセンス/サービスの確認と変更
foreach ($user in $users) {
    write-host ("Processing " + $user.UserPrincipalName)
    $licensetype = $user | Select-Object -ExpandProperty Licenses | Sort-Object { $_.Licenses }

    foreach ($license in $user.Licenses) {
        write-host (" " + $license.accountskuid)

        $includeService = $false
        $disableplan = @()

        foreach ($row in $($license.ServiceStatus)) {
            if ( $row.ServicePlan.ServiceName -eq $service_name ) { $includeService=$true }
            if ( $row.ProvisioningStatus -eq "Disabled" -and $row.ServicePlan.ServiceName -ne $service_name) {
                $disableplan += $row.ServicePlan.ServiceName
            }
        }

        # 指定したサービスが含まれる SKU の場合、現在の設定に加え、指定したサービスを無効化 (Disable)
        if ($includeService) {
            $disableplan += $service_name
            write-host (" found SWAY service in " + $license.accountskuid)
            write-host (" disabled services : " + $disableplan )
            $x = New-MsolLicenseOptions -AccountSkuId $license.accountskuid -DisabledPlans $disableplan
            Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -LicenseOptions $x
        }
    }
}
```

<スクリプトファイルの実行>  
スクリプトファイルを実行する前に、先ほど作成したユーザーリスト (**UserList.txt**) を ps1ファイルを保存したフォルダに格納します。(**DisableSwayLicense\_UserUnit.ps1** と **UserList.txt** を同一ディレクトリに格納します。) その後、PowerShell を起動し、ps1 ファイルを保存したフォルダに移動し、下記の様に実行します。(下記は c:\\temp 配下に ps1 ファイルを格納した場合の例です。)  

```
cd c:\temp
.\DisableSwayLicense_UserUnit.ps1
```

認証ダイアログが起動しますので、ご利用中のテナントの管理者アカウント、および、パスワードを入力します。  
  
有効化処理の内容が表示されますので、PowerShell のプロンプトが戻るまで待ちます。  
  
処理完了後、ブラウザにて Office 365 を表示していた場合は一度ブラウザを終了させ、再度 Office 365 にサインインします。  
  
ユーザーのライセンス付与状況を確認し、Sway 製品のライセンスがオフに設定されていることを確認します。また、アプリケーション ランチャー上の Sway のタイルが表示されていないことを確認します。  
  
  
**6\. ユーザーリストをスクリプトファイルと別のディレクトリに格納して実行する方法**  
ユーザーリストをスクリプトファイルと別のディレクトリに格納して実行する場合、スクリプトファイルの一部を編集する必要があります。スクリプトファイルにユーザーリストの絶対パスを記述します。  
  
<スクリプトファイルの編集>  
ps1 ファイル (**EnableSwayLicense\_UserUnit.ps1**、もしくは、**DisableSwayLicense\_UserUnit.ps1**) をメモ帳などのテキストエディタでひらき、8 行目に以下のコマンドがあることを確認します。  

```
$list = Get-Content "UserList.txt"
```

コマンドのパラメータである、"**UserList.txt**" を削除します。  
なお、"**UserList.txt**" を削除する際、前後のダブルクォーテーション (") も削除します。  
  
\[3. ユーザーリストの作成\] で作成したユーザーリストのパスをコピーします。エクスプローラーでユーザーリストが格納されているディレクトリ (フォルダ) に移動し、"**UserList.txt**" のアイコンを \[Shift\] キーを押しながら右クリックします。ポップアップの一覧から \[パスのコピー\] を選択します。(クリップボードにパスがコピーされます。)  
  
編集中のスクリプトファイルにもどり、先ほど削除した "**UserList.txt**" の箇所に、上記でコピーしたパスを貼りつけます。貼り付けたパスの前後にダブルクォーテーション (") が付いていることを確認します。  

```
例) $list = Get-Content "C:\Office365\Sway\UserList.txt"
```

スクリプトファイルを保存して、テキストエディタをとじます。  
  
以降のスクリプトファイルの実行手順は、上述 \[4. PowerShell で指定ユーザーのみを有効化する方法\]、\[5. PowerShell で指定ユーザーのみを無効化する方法\] のとおりです。  
  
  
**7\. 関連情報**  
タイトル : Office 365 Sway を PowerShell を使用して一括で有効化、無効化 (停止) する方法について  
アドレス : [https://officesupportjp.github.io/blog/Office 365 の Microsoft Forms を PowerShell を使用して一括で有効化、無効化 (停止) する方法について/](https://officesupportjp.github.io/blog/Office%20365%20%E3%81%AE%20Microsoft%20Forms%20%E3%82%92%20PowerShell%20%E3%82%92%E4%BD%BF%E7%94%A8%E3%81%97%E3%81%A6%E4%B8%80%E6%8B%AC%E3%81%A7%E6%9C%89%E5%8A%B9%E5%8C%96%E3%80%81%E7%84%A1%E5%8A%B9%E5%8C%96%20(%E5%81%9C%E6%AD%A2)%20%E3%81%99%E3%82%8B%E6%96%B9%E6%B3%95%E3%81%AB%E3%81%A4%E3%81%84%E3%81%A6/)