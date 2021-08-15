# このプログラムについて
弁理士さんが、実務において特許経過情報をカンタンに取得できるようにしたプログラムです。  
その後の対応のために、拒絶理由通知書のダウンロードについても自動的に行うようしています。  
* 拒絶理由通知書の構文解析については、別プログラムとして開発中です  

なお、読み込み対象となるWebページには随時変更が入っているため、このプログラムの動作保証は致しません。  
また、繰り返しの読み込みによるサーバーへの負荷等の問題がありますので、節度を持ってお使いください。  
  
# 実行手順
1. Python3のインストールが必要  
　Anacondaなどの仮想環境を利用することをお勧めします。  
  
2.Firefoxのインストールが必要  
  
3.Firefoxの拡張geckodriverのインストールが必要 バージョンは最新にすること  
　参考：https://rougeref.hatenablog.com/entry/2019/07/10/095006  
　geckodriverにWindowsのパスを通す必要があります。  
  
4.pythonモジュールのインストールが必要  
　現在のディレクトリのアドレス欄に「cmd」を入力します。  
　コマンドラインを開きます。  
　下記コマンドをコマンドラインに記入してください。  
　"pip install -r requirements.txt"  
　※Anacondaなどの仮想環境を利用することを推奨します。  
  
5.pythonの実行  
　1.AutoDetectの実行  
　　下記コマンドをコマンドラインに記入してください。  
　　"python auto_detect_odp.py"  
  
# 実行上の注意  
　5-1.AutoDetectの実行  
　　CLI（コマンドラインインターフェース）上での表示のみとなります。  
　　・PatentScopeの読み込みについて  
　　　認証画面が出現した場合、手入力で対応するフローにしています。初回対応時に入力することで、以降の処理は問題なく進みます。  
　　　※特許No.あたり1回出現します  
  
# ライセンス
このプログラムはMITライセンスに準拠しています。  
Copyright (c) 2021 A-z. & R  
Released under the MIT license  
https://opensource.org/licenses/mit-license.php  
  
・作者  
A-z.：[特許実務の集約を目指す　from A to Z for IP](https://a2z4ip.hatenablog.com/)  
R：[Rei-Farms](http://rei-farms.jp/)  
