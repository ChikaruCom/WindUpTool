# WindUpTool

- 対象: Only Windows 10 & 11
- 必要ツール: MS Office(一部の機能)

「WindUpTool」は、Windows 10および11上でMS Officeを必要とするツールで、右クリックメニューに追加されるショートカットを通じて、ファイルやフォルダの管理をサポートします。PDF結合やWordファイルのPDF変換、CSVからTSV変換などを行います。どなたでもご自由にご利用いただけます。(※minicondaの現況のライセンスに依存します、リリース時点では無償で利用可能です。パッケージ管理はpipを利用しています。)


## TODO:

- 利用イメージ動画提供予定
- 利用イメージWEBサイト提供予定


## インストール方法

1-install_wut.batを管理者権限で実行してください。必要ツールのダウンロードやインストールを行います。レジストリの変更を伴いますので、必要に応じて1-install_wut.pyのコードのレビューを行ってください。


## 背景・目的

環境を揃えることで安定した確実な作業ができるようになります。インストールを簡略化することでパソコンに不慣れな方でも素早く導入することが可能です。スクリプトの提供になるため簡単に修正が可能です。学習コストや導入コストも少なめです。本格的なRPAや自動化の実装の前段階の勉強としてもオススメです。このプロジェクトはデジタルの日曜大工のようなイメージです。細かい部品から動きのあるGUIまで大小問わず100個の仕掛けの制作を目指しています。基本的な工具やアプリを制作し誰もが使える状態を提供することでそれぞれの日常の負担が少しでも軽くなることを願っています。


## 1-install_wut.bat

- "%UserProfile%\WindUpTool"の作成
- miniconda3のダウンロード&インストール: "%UserProfile%\WindUpTool\Miniconda3"
- 仮想環境: py3.12_wut の作成(python 3.12)
- "%UserProfile%\WindUpTool\karakuri"の作成
- 右クリックメニューの追加: Wind Up Tool
	- データを右クリックした場合: 個別のデータ用のショートカット表示
	- フォルダを右クリックした場合: ディレクトリを対象としたショートカット表示
- 必要モジュールのインストール
	- pandas=2.2.2
	- PyPDF2=3.0.1
	- pywin32==306
	- psutil==6.0.0
	- pyperclip==1.9.0


### 利用イメージ:

【ファイルを右クリックした場合】 <br>
![ファイル右クリックイメージ](https://github.com/user-attachments/assets/265863c5-f839-423f-8a35-adf409b74538)

【ディレクトリを右クリックした場合】<br>
![ディレクトリ右クリックイメージ](https://github.com/user-attachments/assets/e3c3ace2-e605-498a-9d9a-5746ea266c74)


## 基本フォルダ構成

- `c:\users\{user_name}\WindUpTool`
- `c:\users\{user_name}\WindUpTool\miniconda3`
- `c:\users\{user_name}\WindUpTool\karakuri`
- `c:\users\{user_name}\WindUpTool\assets`
- `c:\users\{user_name}\WindUpTool\output`


## src:

- _script_run: 実行環境の確認機能
- _wut_any2sjis: テキストファイルのエンコードをsjisに変換し接尾辞に_sjisを追加
- _wut_csv2tsv: CSVデータをTSVデータに変換
- _wut_pdfs2pdf_dir: フォルダ内のPDFファイルを一覧表示し結合が可能
	- ファイル名に`_`が含まれている場合、最後の"_"以降の名称を削除
	- ファイル名に`_`が含まれていない場合、接尾辞に"-merged"が追加
- _wut_word2pdf_dir: フォルダ内のWORDファイルを一覧表示しPDF化
	- `c:\users\{user_name}\WindUpTool\output\{日時}`にデータ出力
	- 出力時にエクスプローラ表示

