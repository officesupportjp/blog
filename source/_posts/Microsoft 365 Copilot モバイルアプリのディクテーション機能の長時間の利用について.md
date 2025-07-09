---
title: Microsoft 365 Copilot モバイルアプリのディクテーション機能の長時間の利用について
date: '2025-07-09'
id: cmcvaqw140003hc6y0qnedgdp
---

こんにちは、Office サポート チームです。
この記事では、Microsoft 365 Copilot モバイルアプリでのディクテーション機能を  
長時間利用する際に発生する事象についてご案内いたします。

# 事象の詳細
Microsoft 365 Copilot モバイルアプリでは、[作成]-[ディクテーション] より  
音声の録音とその文字起こしを行うことが可能です。
  

録音中はリアルタイムに文字起こしが表示されます。
また、録音中録音データは端末にキャッシュされ、  
録音終了時に自動的に音声データとその文字起こしデータファイルが自身の OneDrive 上にアップロードされます。
  

本機能について、75 分以上録音を行う場合に、リアルタイムの文字起こしが中断され、
録音終了後も OneDrive へのデータのアップロードが失敗する動作が確認されております。
  
また、60 分など長時間録音を継続する場合においても、
何らかのエラーにより録音が中断される場合などが確認されております。

# 対応案
上記の通り録音中の音声データは端末にキャッシュされるため、端末のストレージ容量次第でディクテーション機能の利用に問題が生じる可能性がございます。  
また、ディクテーション機能利用時には安定したネットワーク接続が必須となり、機能の安定した利用にはネットワークの影響もうけます。
  
上記を踏まえ、長時間の録音は控え、最低限の時間ごとに録音いただくことをお勧めいたします。
  
参考: Record voice notes in Office Mobile
URL: [https://techcommunity.microsoft.com/blog/microsoft365insiderblog/record-voice-notes-in-office-mobile-for-ios/4213168](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Ftechcommunity.microsoft.com%2Fblog%2Fmicrosoft365insiderblog%2Frecord-voice-notes-in-office-mobile-for-ios%2F4213168&data=05%7C02%7Ckokisugiura%40microsoft.com%7C4e2660413f9140f57f2708ddb9408c5d%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C638870409089899235%7CUnknown%7CTWFpbGZsb3d8eyJFbXB0eU1hcGkiOnRydWUsIlYiOiIwLjAuMDAwMCIsIlAiOiJXaW4zMiIsIkFOIjoiTWFpbCIsIldUIjoyfQ%3D%3D%7C0%7C%7C%7C&sdata=q1WpUtuykCg2Ps0KzzBJN6rTja8xZFVaj3f8H2MzSc4%3D&reserved=0)
  
なお、録音後自動的に音声データ等が OneDrive にアップロードされない場合は以下のような方法で音声データを取得できるかお試しいただければと存じます。
- [検索] タブより最近使ったファイルとして表示される音声データを開き、表示される UI から OneDrive へアップロードする
- [検索] タブより最近使ったファイルとして表示される音声データを開き、OS の [共有] の仕組みを利用し、OneDrive アプリや Teams アプリなど任意のその他アプリに共有する
  
音声データの文字起こしについては以下機能の利用もご検討いただければと存じます。
  
参考: レコーディングの文字起こし
URL: [https://support.microsoft.com/ja-jp/office/7fc2efec-245e-45f0-b053-2a97531ecf57](https://support.microsoft.com/ja-jp/office/7fc2efec-245e-45f0-b053-2a97531ecf57)

  
今回の投稿は以上です。
本情報の内容 (添付文書、リンク先などを含む) は、作成日時点でのものであり、予告なく変更される場合があります。

