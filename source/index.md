# Python in Excelをより便利に使える実践プラクティスの紹介

```{raw} html

<a rel="license" href="http://creativecommons.org/licenses/by/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by/4.0/88x31.png" /></a>
<small>This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by/4.0/">Creative Commons Attribution 4.0 International License</a>.</small>
```

## はじめに

### 自己紹介

* Ryuji Tsutsui@ryu22e
* さくらインターネット株式会社所属
* Python歴は14年くらい（主にDjango）
* Python Boot Camp、Shonan.py、GCPUG Shonanなどコミュニティ活動もしています
* 著書（共著）：『[Python実践レシピ](https://gihyo.jp/book/2022/978-4-297-12576-9)』

### 今​日話したい​こと

* ExcelブックにPythonコードを埋め込めるExcelの機能「Python in Excel」の紹介
* Python in Excelを実際に使う上で便利なテクニックの紹介

### この​トークの​対象者

* Excelの基本的な使い方が分かる人
* Pythonの基礎的な文法が分かる人

### この​トークで​得られる​こと

* Python in Excelの概要
* Python in Excelを実際に使う際にどう書けばいいのか

### トークの​構成

* Python in Excelの概要について軽く説明
* Python in Excel実践プラクティスの紹介

## Python in Excelの概要について軽く説明

### Python in Excelとは

* ExcelブックにPythonコードを埋め込めるExcelの機能
* セルに書かれたデータを読み取り、Pandasで加工したり、加工した内容をグラフで描画できる

### Python in Excelの仕組み

```{figure} _static/img/python-in-excel-image.jpg
:alt: イメージ図

イメージ図
```

### 導入方法

* TODO Python in Excelを利用できるプラットフォームを書く

### 使い方

任意のセルで`=PY(`と入力するか、Ctrl + Alt + Shift + Pを入力。

TODO 画面スクリーンショットを貼る

## Python in Excel実践プラクティスの紹介

### 実践プラクティスの情報ソースについて紹介

Python in Excelを便利に使いこなすためのプラクティスを紹介しているサイトを紹介します。
本トークの内容は、これらのサイトをベースに作成しています。

### 元のデータはPython以外を使って加工しない

通常のプログラミングと同様に、Python in Excelにおいても再利用性が高いコードを書くためのテクニックがあります。ここでは、プラクティスを使わなかった場合、使った場合のExcelブックを比較し、プラクティスがどのように役立つのかを説明します。

### コードは複数のセルに分割して書く

Python in Excelでは、ある程度長いコードを書く場合、1つのセルにすべての処理を書いてしまうと可読性が落ちてしまいます。そんなときは、コードを複数のセルに分割して書く方法をお勧めします。
複数のセルに分割して書く場合、留意しなければならないのは「セルの実行順序」です。ここでは、Python in Excelが複数セルに書かれたPythonコードをどの順番で実行するかを説明します。

### 各セルに「文字列リテラルのコメント」を書く

Pythonは「#」より右側の文字列はコメントとして扱われます。Python in Excelでもコメントを書くことができますが、もっとコードを読みやすくするためのテクニックがあります。
ここでは、「文字列リテラルのコメント」の使い方、「文字列リテラルのコメント」が通常のコメントより優れている点について紹介します。

### グラフの描画時に日本語フォントを指定する

Python in Excelではseaborn、matplotlibなどを使ってグラフを描画することができますが、デフォルトではグラフ中の日本語が文字化けします。ここでは、グラフに日本語を表示させる方法について紹介します。

### その他Tips

その他の細かいTipsについて紹介します。

## 最後に

### まとめ

### ご清聴ありがとうございました
