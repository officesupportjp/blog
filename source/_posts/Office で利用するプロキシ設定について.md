---
title: Office で利用するプロキシ設定について
date: '2024-1-5'
lastupdate: '2024-1-5'
id: clr09r8tu00002wkjcubph9b8
tags:
 - サインイン、認証

---

こんにちは、Office サポートの西川 (直)です。  

本記事では、Office で利用するプロキシ設定について説明します。


説明
--

Office には、オンライン画像の挿入やオンラインテンプレートの取得などインターネット通信を利用する機能がございます。
このプロキシ設定はログオンユーザーの OS のプロキシ設定 (WinINET) が利用されます。WinHTTP のプロキシ設定は利用されません。

ただし、Office 2019 以降や Microsoft 365 クライアントで存在する自動更新の処理等の一部の機能については、以下の順でそれぞれのプロキシ設定を参照することがございます。

1. WinHTTP (存在しない場合、スキップされます。また、構成していても、処理によっては参照しない場合がございます)
2. SYSTEM ユーザーの OS のプロキシ設定 (WinINET)
3. ログオンユーザーの OS のプロキシ設定 (WinINET) (ただし、ユーザーがログオンしている場合に限ります)

最終的には 3 が利用されるため、1 や 2 を構成せずとも 3 を構成しておくことで通信は出来るものとなります。
なお、この動作を無効化したり順番を変更することは出来ません。また、処理内容によっては 1 と 2 を使用しない場合もございます。


FAQ
--
**Office が利用するエンドポイントを教えてください**
[Office 365 の URL と IP アドレスの範囲](https://learn.microsoft.com/ja-jp/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide)をご確認ください。

**Office でネットワーク接続がない旨のメッセージが表示されます**
[Office のサインインでネットワーク接続がない旨のメッセージが表示される事象について](https://officesupportjp.github.io/blog/cl0m75al4001gmcvse36w64dv/)をご確認ください。

**必要なサービスが無効になっていますと表示されます**
[必要なサービスが無効になっていますと表示される現象について](https://officesupportjp.github.io/blog/clmk7jget0001hckj66na16pi/)をご確認ください。

**GPO プロキシ設定を配布していますが、2 の SYSTEM ユーザーのプロキシ設定は変更されますか**
GPO は OS ログオン時に設定されますため、2 の SYSTEM ユーザーのプロキシ設定は変更されません。
2 の SYSTEM ユーザーのプロキシ設定を変更するには、[BitsAdmin コマンド](https://docs.microsoft.com/ja-jp/windows/win32/bits/bitsadmin-tool)を利用する方法があります。

プロキシサーバー "192.168.1.1:8080" を指定する場合の例)
bitsadmin /util /setieproxy localsystem MANUAL_PROXY 192.168.1.1:8080 ""

バイパスリストを含め指定する場合の例)
bitsadmin /util /setieproxy localsystem MANUAL_PROXY 192.168.1.1:8080 "*.microsoft.com"

pac ファイル http://server/proxy.pac を指定する場合の例)
bitsadmin /util /setieproxy localsystem AUTOSCRIPT http://server/proxy.pac

設定をリセットする方法)
bitsadmin /util /setieproxy localsystem NO_PROXY

設定を確認する方法)
bitsadmin /util /getieproxy localsystem



**本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。**