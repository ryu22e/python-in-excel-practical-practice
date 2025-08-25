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

### 使い方

任意のセルで`=PY(`と入力するか、Ctrl + Alt + Shift + Pを入力。

TODO 画面スクリーンショットを貼る

### 導入方法（2025年9月時点）

現在、Excelは以下のプラットフォーム用のアプリケーションがある。

* Excel for Windows
* Excel for Mac
* Excel on the web
* Excel for iPad / iPhone / Android

参考: <https://support.microsoft.com/ja-jp/office/excel-%E3%81%A7%E3%81%AE-python-%E3%81%AE%E5%8F%AF%E7%94%A8%E6%80%A7-781383e6-86b9-4156-84fb-93e786f7cab0>

### Excel for Windowsの場合

以下のサブスクリプションで利用可能。

* Enterprise / Business
* Family / Personal

### Excel for Macの場合

以下のサブスクリプションで利用可能。

* Enterprise / Business
* Family / Personal（ただしプレビュー版）

### Excel on the webの場合

以下のサブスクリプションで利用可能。

* Enterprise / Business
* Family / Personal（ただしプレビュー版）

### Excel for iPad / iPhone / Androidの場合

非対応（Pythonコードを含むExcelブックを表示できるが、再計算はできない）。

## Python in Excel実践プラクティスの紹介

### 実践プラクティスの情報ソースについて紹介

Python in Excelを便利に使いこなすためのプラクティスを紹介しているサイトを紹介します。
本トークの内容は、これらのサイトをベースに作成しています。

```{revealjs-break}
```

* [Microsoft公式サイト（日本語）](https://support.microsoft.com/ja-jp/office/python-in-excel-%E3%81%AE%E6%A6%82%E8%A6%81-55643c2e-ff56-4168-b1ce-9428c8308545)
* [Anacondaの公式ブログ（英語）](https://www.anaconda.com/blog)

### 元のデータはPython以外を使って加工しない

Python Boot Camp開催実績:

TODO 画面スクリーンショットを貼る

これをPython in Excelで扱う際、不要なデータがいくつか含まれている。

### 前述の例の中で不要なデータ

たとえば、各イベントの参加者数の推移をグラフにする際は、以下のデータは不要。

* 「2018年⁠」⁠、「⁠2019年」など年度を示す行
* 中止になったイベントの行
* 「参加者」の数字の末尾にある「名」

### 不要なデータをどうやって取り除くか？

* Excelブックを直接編集して、不要なデータを削除しないほうがいい
* 直接編集することで、別の目的で統計処理を行うのが難しくなる場合がある
* 不要なデータはPythonコードを使って取り除く

### Pythonコードの例

```{revealjs-code-block} python
# データの読み込み
# xl()関数の第1引数でテーブル名を、第2引数でテーブルのヘッダーを除いて取得するかを指定
df = xl("PyCamp開催実績[#すべて]", headers=True)

# 採番されていない行（年度を表す行）を除去
# df["No."]は「No.」列の値を取得
filtered_df = df[pd.to_numeric(df["No."], errors="coerce").notnull()]
# イベントが中止になった行を除去
filtered_df = filtered_df[~filtered_df["イベント"].str.contains("中止")]

# 参加者の「名」を除去して数値に変換
filtered_df["参加者"] = filtered_df["参加者"].str.replace("名", "")
filtered_df["参加者"] = filtered_df["参加者"].astype(int)
```

### コードは複数のセルに分割して書く

悪い例:

### 前述の例の何が問題か

1. 長いコードだと自分が読みたいコードだけにフォーカスしにくい（毎回先頭から読む必要がある）
2. 共通のロジックを再利用しにくい（他の区の年齢内訳を取得する場合、ほぼ同じコードを書く必要がある）
3. コードの途中の変数の内容を確認しにくい（Python in Excelではbreakpoint()関数を使ったデバッグはできない）

### 改善例

前述のコードを複数のセルに分割して書く。

TODO 画面スクリーンショットを貼る

### 複数セルにPythonコードを書いた場合の実行順序

以下の順番でコードが実行される:

TODO 画面スクリーンショットを貼る

### 各セルに「文字列リテラルのコメント」を書く

悪い例:

TODO 画面スクリーンショットを貼る

### 前述の例の何が問題か

* `#`のコメントだと、該当セルをクリックしないと内容が読めない
* Excelブックを開いたときに処理全体の概要を把握しにくい

### では、どうするか

最後の行に「文字列リテラルのコメント」を書く。

```{revealjs-code-block} python
df = xl("日常生活圏域等別データ[#すべて]", headers=True)
"データの読み込み"
```

### 「⁠文字列リテラルのコメント」とは

* Python in Excelでは、最後の行の値がセルに表示される仕様
* これを利用して、最後の行に文字列リテラルを書き、コメントの代わりに使う

### 「⁠文字列リテラルの​コメント」を使うとこうなる

TODO 画はは面スクリーンショッ貼る

### グラフの描画時に日本語フォントを指定する

Python in Excelではseaborn、matplotlibなどを使ってグラフを描画することができますが、デフォルトではグラフ中の日本語が文字化けします。ここでは、グラフに日本語を表示させる方法について紹介します。

### その他Tips

その他の細かいTipsについて紹介します。

## 最後に

### まとめ

### ご清聴ありがとうございました
