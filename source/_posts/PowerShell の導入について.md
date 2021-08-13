---
title: PowerShell の導入について
date: 2019-03-09
---

(※ 2016 年 6 月 30 日に Japan Office Support Blog に公開した情報のアーカイブです。)

こんにちは、Office サポートです。  
  
本記事では、スクリプトを処理する時に必要となる PowerShell の導入手順についてご案内します。Office 365 にて公開されている製品において、ユーザー管理の一括処理等でスクリプトコマンドを実行する時などにご参照ください。  
  
**目次**  
1\. PowerShell の導入  
2\. 関連情報  
  
  
**1\. PowerShell の導入**  

1\. 64 Bit OS が動作する Windows 10/8.1/8/7 の PC を準備します。  
  
2\. Windows 7 の PC に導入する場合は、前提条件を満たすために最新の Windows Update をすべて適用します。  
  
Windows Update 適用後に OS を再起動し、その後、下記リンクから Windows Management Framework 3.0 をインストールし、 Windows PowerShell 3.0 を使用できるようにします。(Windows 10/8.1/8 の PC 場合は、本手順は不要です。)  
タイトル : Windows Management Framework 3.0  
アドレス : http://www.microsoft.com/en-us/download/details.aspx?id=34595  

  

  
  
3\. 下記の弊社 TechNet ページを参照し、Microsoft Online Services サインイン アシスタント、Windows PowerShell 用 Microsoft Azure Active Directory モジュール (64 ビット バージョン) をインストールします。  

タイトル : Connect to Office 365 PowerShell  
アドレス : https://technet.microsoft.com/ja-jp/library/dn975125.aspx

  
  
4\. PowerShell の[実行ポリシー](https://technet.microsoft.com/ja-jp/library/hh847748.aspx)を変更していない場合は、インストールした Windows PowerShell の Microsoft Azure Active Directory モジュールを管理者権限で起動し、下記コマンドを実行します。([【参考 : 実行ポリシー】](https://technet.microsoft.com/ja-jp/library/hh847748.aspx)既に変更済みの場合は、本手順は不要です。)  

```
Set-ExecutionPolicy RemoteSigned
```

  

なお、設定されている[実行ポリシー](https://technet.microsoft.com/ja-jp/library/hh847748.aspx)を確認するには下記のコマンドを実行します。([実行ポリシー](https://technet.microsoft.com/ja-jp/library/hh847748.aspx)を取得してプロンプトに表示します。)  

```
Get-ExecutionPolicy
```

  
  
**2.関連情報**  
  
関連ブログ記事をご案内いたします。  

タイトル : Windows PowerShell ファースト ステップ ガイド  
アドレス : [https://technet.microsoft.com/ja-jp/library/aa973757(VS.85).aspx](https://technet.microsoft.com/ja-jp/library/aa973757(VS.85).aspx)

  

タイトル : PowerShell の学習  
アドレス : [https://technet.microsoft.com/ja-jp/library/cc281945(v=sql.105).aspx](https://technet.microsoft.com/ja-jp/library/cc281945(v=sql.105).aspx)

  

タイトル : PowerShell スクリプト  
アドレス : [https://msdn.microsoft.com/ja-jp/powershell/scripting/powershell-scripting](https://msdn.microsoft.com/ja-jp/powershell/scripting/powershell-scripting)