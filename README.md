# WindUpTool

- target: Only Windows 10 & 11
- require: MS Office

「WindUpTool」は、Windows 10および11上でMS Officeを必要とするツールで、右クリックメニューに追加されるショートカットを通じて、ファイルやフォルダの管理をサポートします。PDF結合やWordファイルのPDF変換、CSVからTSV変換などを行います。どなたでもご自由にご利用いただけます。(※minicondaの現況に依存します。パッケージ管理はpipを利用しています。)


## インストール方法

1-install_wut.batを管理者権限で実行してください。必要ツールのダウンロードやインストールを行います。レジストリの変更を伴いますので、必要に応じて1-install_wut.pyのコードのレビューを行ってください。


## 背景・目的

環境を揃えることで安定した確実な作業ができるようになります。インストールを簡略化することでパソコンに不慣れな方でも素早く導入することが可能です。スクリプトによる提供になるため簡単に修正や更新の学習コストや導入コストも少なめです。本格的なRPAや自動化の実装の前段階の勉強としてもオススメです。小さなデジタルのDIYで、細かい部品から動きのあるGUIまで大小問わず100個のカラクリの制作を目指しています。基本的な工具やアプリを制作し誰もが使える状態を提供することでそれぞれの負担が少しでも少なくなることを願っています。


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

