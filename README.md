# webScrapingITmedia

ITmedia NEWSの情報を取得するwebスクレイピングツールを格納している。<br>

## 概要

Pythonを使用してWEBスクレイピングツールを作成した。<br>
以下のITmedia NEWSから速報情報を取得し、EXCELに一覧化する。<br>
https://www.itmedia.co.jp/news/<br>

## プロジェクト構造

webScraping<br>
　├chromedriver.exe（webdriver）<br>
　├main.py<br>
　└output<br>
　　　└output.xlsx（スクレイピング結果出力ファイル）<br>

## 使用外部ライブラリ

・urllib：urlを引数に、url先の情報をhtml形式で取得するためのライブラリ<br>
・BeautifulSoup：html形式の情報を解析するためのライブラリ<br>
・selenium：仮想ブラウザ操作でスクレイピングをすることができるライブラリ。ログインが必要なサイトや特定のボタンをクリックしないと取得できない情報などを取得することが可能<br>
・openpyxl：pythonでエクセルを操作するためのライブラリ<br>


## 操作方法

【前提】<br>
・Pythonで作成したプロジェクトのためPythonが使用PCにインストールされている必要がある<br>
・webdriverはChromeを使用している（プロジェクト直下のchromedriver.exeが該当）<br>
　※実際に使用しているChromeのバージョンとwebdriverのバージョンを合わせなければいけない<br>
　　格納されているdriverは123.0.6312.122であるためChromeは123系である必要がある<br>

【操作】<br>
①main.pyを実行（プロキシ関係でエラーが出る場合、環境変数「no_proxy」に「localhost,127.0.0.1」を設定する）<br>
②webdriberにより自動でChromeが開き、必要な情報を取得する処理が走る<br>
③output.xlsxに取得した情報が出力される<br>

## 注意点

出力した速報情報の記事内容が途中で途切れることがあり、全文表示されない記事が多々ある。<br>
様々な方法で試したが（処理の間にsleepを入れて間隔をあけたり使用するライブラリを変えてみたり）結果は変わらず。<br>
全文がEXCELに表示されていない場合は出力されたリンク先を参照して本文を見るように。<br>

