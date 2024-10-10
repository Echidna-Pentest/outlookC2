# outlookC2

これは検証用のoutlookを利用したC2フレームワークです。クライアントでoutlookを実行した状態でbeacon.ps1、サーバ側でoutlookC2Server.pyを実行すると、クライアントのoutlookのメールを介してC2通信することができます。

![](img/demo.gif)

多くのマルウェアやC2フレームワークが通信にhttp/https、もしくはDNSを利用するためか、それらと比較するとSMTP/IMAPを利用するC2通信は注目されていないと感じます。しかし、SMTP/IMAPをC2通信に利用するマルウェア/Actorは少ないですが存在します (主にdata exfiltration)。

また、SMTP/IMAPによるC2通信が脅威・脆弱性となるケースが存在します。例えばWeb分離というセキュリティソリューションを利用すれば、クライアントから外部への直接的なhttp/https通信は遮断することができます。こういった環境下ではSMTP/IMAPをC2通信に利用するマルウェアがある場合、大きな脅威となりえると考えます。

PowershellでOutlookを操作することができるため、このC2フレームワークはクライアント側のOutlookを操作してサーバと通信します。Outlookを利用する理由は、最も企業で利用されているメールクライアント（特にWeb分離を導入できるような大きな組織では特にその傾向がある）であるためです。

## 使い方
### Client
beacon.ps1を実行、もしくは.netのソースコードをコンパイルして実行してください。C2のサーバとして利用するメールアドレスがserverAddressという変数で設定されているため、利用するメールアドレスに変更してください。

`$serverAddress = "attackerSender@testmail.com"`

### Server
サーバを起動しなくても、単純にgmail等からclientのメールアドレスにメールを送付するだけでC2として機能することはできます。しかし、毎回メールを送信するのは面倒、かつシェルをシミュレートするために簡単なGUIツールを作成しています、

outlookC2Server.pyに認証情報が記載されているので、攻撃側メールアドレスの認証情報、攻撃対象のメールアドレスを設定してください。

```
smtp_server = 'smtp.gmail.com'
imap_server = 'imap.gmail.com'
port = 587
#username = input("Your email address: ")
username = "attackerSender@testmail.com"
password = getpass.getpass('Password: ')
#recipient = input("Target email address: ")
recipient = "victimRecipient@testmail.com"
```

その後、サーバを実行してGUIからSubject、Contentsに実行したいC2コマンドを入力してSend Emailボタンを押すと、上記の設定したattackerSenderからvictimRecipient宛にメールが送付されます。

`python3 outlookC2Server.py`

![alt text](img/c2server.png)

## 対応しているC2コマンド

| C2 Command | Description |
| ---- | ---- |
| Download {Filepath}| 指定されたファイルを添付してメールでC2アドレスに送信 |
| Filepath| 指定されたファイルを添付してメールでvictimアドレスに送信、outlookBeacon.ps1は受信したファイルをC:\Windows\Tasksにドロップ |
| search {Keyword} | Keywordを含むメールを受信トレイから検索 |
| forward | C2アドレスに今後受信メールを送信するルール作成 |
| listFolders | すべての受信フォルダを取得して、結果をC2アドレスに送信 |
| getFolders {FolderName} | FolderNameのメールをzipにしてC2アドレスに送信 |
| Other | Powershellコマンドを実行して、結果をC2アドレスにメール送信、例whoamI, ipconfig |

カンマ区切りで複数コマンドを送信できます。
`whoami; listFolders; net user`


## outlookBeacon.ps1の処理の流れ

1. レジストリを操作してOutlookの通知をオフ
2. 起動しているOutlookを監視して、設定されているアドレス(C2アドレス)からメールが来ているかチェック
3. 送られてきたC2コマンドに応じてコマンドを実行、結果をC2アドレスにメールを送信
4. C2アドレスから送受信したメールを削除


## AV/EDRの検知状況

自身の環境で検証したところ、AV/EDRでの検知はなかった。EDRは検知条件が複雑なため、C2コマンドや実行状況にもよるが、一般的なhttp/https/dns/tcpのリバースシェルと比較すると検知されにくいと考える。


### 一般的なリバースシェルの動き

一般的なリバースシェルの処理の流れは以下の通り。

1. リバースシェルが定期的にC2サーバにリクエストを送信
2. C2サーバが指令(コマンド等)を含んだレスポンスを返信
3. リバースシェルが指令を実行して、結果をC2サーバに再度送信

![alt text](img/generalC2.png)


1でC2リクエストを定期的に発生することが多い（Sleepの長さを設定できることは多い）。
また、追加ファイルを書き込む場合は、マルウェア自身、もしくはwgetやbitsadmin, certutil等のマルウェアでよく利用されるプロセスからファイルが書き込まれるため、AVに検知されることがある(インジェクション等を行うことで親プロセスは変更される)。

![alt text](img/generalC2ProcessTree.png)


### 今回のoutlookC2の動き

outlookC2はプロセスを監視しているだけで、定期的なC2とのトラフィックは発生しない。

![alt text](img/outlookC2.png)

追加ファイルの書き込みに関しても、書き込むプロセスはOutlook(正規な利用法でも読み書きが頻繁に行われる)となり、outboundの通信も発生しないため、AV/EDRに検知される可能性も低いと考える。もし、mimikatzのような実行ファイルを送付したい場合は、パスワード付きzip等を利用すれば検知されないと考える。
また、クライアントの起動済みのOutlookプロセスを利用するため、SMTPサーバ-のクレデンシャルも不要となる。

![alt text](img/image-1.png)




## outlookC2の欠点


欠点としては、通知をオフにする処理を追加しているものの、メールクライアントによっては通知をオフできなかったりするため、AVよりもユーザに不審に思われるリスクがある。これの対策としては、例えばユーザが無視するであろう広告メールに添付画像を付与して、そこにC2からの指示をsteganographyのようなテクニックを利用して不審に思われないようにする

もう１つの欠点は企業によっては送付先のメールアドレスドメインを制限している場合がある(例えば、Gmail等のフリーアドレスなど)。この場合は、Botnetのような侵害済みのドメインのクレデンシャルからメールを送付する必要がある。もしくは、内部環境の横展開に利用する場合は組織内のドメインアドレスからのメールとなるため、メールアドレスのチェックは無効になると考えられる。


## 検知ルール

SMTP/IMAPをC2通信に利用するケースは少ないためか、http/https/dnsと比較して多くないが、少ない検知ルールは以下の通り。

- Splunk (Gsuite Outbound Email With Attachment To External Domain)
Gsuite Outbound Email With Attachment To External Domain (Not outlook)

https://research.splunk.com/cloud/dc4dc3a8-ff54-11eb-8bf7-acde48001122/

- Elastic (Suspicious Inter-Process Communication via Outlook)

https://www.elastic.co/guide/en/security/current/suspicious-inter-process-communication-via-outlook.html

- CrowdStrike

EDRによってはEmail Collectionの攻撃を検知する場合があります。例えばCrowdStrikeではプロセスの流れによっては下記のように検知する場合もありました。しかし、検知しないこともあるので、EDR以外でも検知する仕組みを考えることが重要です。

![alt text](img/image.png)



## 参考にしたサイト

- BadOutlook

Web分離環境を対象にしたものではないが、最も類似なツール。
Outlookを利用してシェルコードを実行するPOCを提供。

https://github.com/aahmad097/BadOutlook

- SharpGmailC2

SMTPとIMAPによるGmailプロセスを利用したC2コミュニケーション。

https://github.com/reveng007/SharpGmailC2

- AzureOutlookC2 (2021)

MicrosoftのGraph APIを利用したC2通信。

https://github.com/boku7/azureOutlookC2

- Sans Article

Outlookを利用したC2通信を技術解説したSANSの記事。

https://isc.sans.edu/diary/29180
