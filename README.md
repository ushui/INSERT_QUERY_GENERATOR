# INSERT QUERY GENERATOR
スプレッドシートに入力したデータから、INSERT文を生成するプログラムです。  
大量に用意したデータをデータベースへ格納したいとき、INSERT文を1つ1つ手作りしている方にオススメです。  
  
## 使い方
使い方を二通り用意しています。  
プログラムはどちらでも使用できますが、ブックファイルやプログラムの他人への配布を考えた際に使い勝手が異なります。  

### ブックファイルにモジュールを追加する場合
Excelブックファイルに対し、モジュールとして追加することで使用できるようにする方法です。  
ブックファイルを閲覧できる方であれば誰でもプログラムの実行を行えますが、ブックファイルごとに追加しなければならず、ブックファイルの拡張子が「xlsm」になり、閲覧時にマクロを有効にしなければなりません。  
追加方法は下記をご覧ください。  

1. このリポジトリから「[INSERT QUERY GENERATOR MODULE](https://github.com/ushui/INSERT_QUERY_GENERATOR/raw/master/VBA/InsertQueryGeneratorModule.bas)」をダウンロードする。  

2. Excelブックファイルを開く。  

3. [開発] タブの [Visual Basic] をクリックする。  
[開発] タブがない場合は[[開発] タブを表示する - Office サポート](https://support.office.com/ja-jp/article/-%E9%96%8B%E7%99%BA-%E3%82%BF%E3%83%96%E3%82%92%E8%A1%A8%E7%A4%BA%E3%81%99%E3%82%8B-e1192344-5e56-4d45-931b-e5fd9bea2d45)を参照してください。  

4. [ファイル] から [ファイルのインポート]をクリックする。  

5. ダウンロードした「InsertQueryGeneratorModule.bas」を選択し、 [開く] をクリックする。  

### Excelにアドインを追加する方法  
Excelに対し、アドインとして追加することで使用できるようにする方法です。  
アドインが追加されている環境ではプログラムの実行を行えますが、そうでない環境では実行できません。  
追加方法は下記をご覧ください。  

1. このリポジトリから「[INSERT QUERY GENERATOR](https://github.com/ushui/INSERT_QUERY_GENERATOR/raw/master/VBA/INSERT QUERY GENERATOR.xlam)」をダウンロードする。  

2. エクスプローラを開いて「C:\Users\{UserName}\AppData\Roaming\Microsoft\AddIns」に移動する。  

3. フォルダに対して、ダウンロードした「INSERT QUERY GENERATOR.xlam」をコピーする。  

4. Excelを開く。  

5. [ファイル] タブをクリックし、[オプション] をクリックして、[アドイン] カテゴリをクリックする。  

6. [管理] ボックスの一覧の [Excel アドイン] をクリックし、[設定] をクリックする。  

7. 表示された [アドイン] ダイアログ ボックスの [有効なアドイン] ボックスに「INSERT QUERY GENERATOR」があることを確認し、横のチェックボックスをオンにして、[OK] をクリックする。  

アドインに関しての詳細は、下記をご覧ください。  
[Excel でアドインを追加または削除する - Office サポート](https://support.office.com/ja-jp/article/excel-%E3%81%A7%E3%82%A2%E3%83%89%E3%82%A4%E3%83%B3%E3%82%92%E8%BF%BD%E5%8A%A0%E3%81%BE%E3%81%9F%E3%81%AF%E5%89%8A%E9%99%A4%E3%81%99%E3%82%8B-0af570c4-5cf3-4fa9-9b88-403625a0b460)  

## 動作環境について
「Microsoft Excel 2007」以降に対応しております。  

## 開発環境・開発言語について
開発環境：Visual Basic Editor  
開発言語：Visual Basic for Applications  

## ソースコードについて
MIT License に準拠します。  

***
2018/07/16 新規作成  
***
作成者： ushui（ゆーしゅい）  
Twitter: [@kaede_hrc](https://twitter.com/kaede_hrc)  
