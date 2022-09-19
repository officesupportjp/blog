---
title: >-
  暗号化された Office ファイルを保護ビューで開き、保護ビューが閉じた後、クライアントの %localappdata% フォルダ配下 Packages
  フォルダに oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダが残存する
date: '2022-09-06'
id: cl7pvphjn0000a46ygd3m4nro
tags:
  - ''

---

こんにちは、Office サポート チームです。  
この記事では、暗号化された Office ファイル (*) を保護ビューで開き、保護ビューが閉じた後、   
クライアントのローカル フォルダ %localappdata%\Packages 配下に oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダ (x は任意の英数字) が残存する事象について説明します。  
  
# 事象の詳細

Office アプリでは、暗号化された Office ファイルであるかどうかに関わらず、ファイルを保護ビューで開く場合、以下のフォルダを作成します。  

%localappdata%\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダ   
(例: C:\Users\<ユーザー名>\AppData\Local\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx)  

この oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダ配下に、元のファイルや元のファイルに紐づく情報を一時ファイルとしてコピー/作成した上で、保護ビューでファイルを読み取り専用で開きます。  

この動作は、ユーザーが明示的に [編集を有効にする] を実行しない限り、リスクを軽減しながらファイルの内容を確認できるセキュリティを考慮した実装です。  

なお、Office ファイルが暗号化されていない場合は、[編集を有効にする] をクリックしたタイミングで、%localappdata%\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダは削除されます。  

一方、暗号化された Office ファイル (*) の場合、[編集を有効にする] をクリックしても、%localappdata%\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダは削除されず、残ります。  

![](image1.png)  
![](image2.png)  


(*) 暗号化された Office ファイルとは、以下のような操作などで暗号化されたファイルを指します。  
・Azure Information Protection 統合ラベル付けクライアントや Office アプリで秘密度ラベルを設定したファイル  
・Office アプリの [ファイルの保護] - [アクセスの制限] からアクセス制限ありを設定したファイル  
・SharePoint の Information Rights Management (IRM) を有効に設定したライブラリからダウンロードされたファイル  


# 原因
この問題は、Microsoft 365 Apps にて確認されており、現在、製品開発部門で調査中です。  
新たな情報が確認された場合は、この記事に追記する形で公開していく予定です。

  
# 回避案
エクスプローラーを起動し、残っている %localappdata%\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx フォルダ  
(例: C:\Users\<ユーザー名>\AppData\Local\Packages\oice_xx_xxxxxxxx_xxxxxxxx_xxxx) を選択し、手動で削除してください。  

今回の投稿は以上です。  
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。







