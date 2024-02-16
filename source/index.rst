###########################
Python in Excelの紹介（仮）
###########################
.. raw:: html

   <a rel="license" href="http://creativecommons.org/licenses/by/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by/4.0/88x31.png" /></a><br /><small>This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by/4.0/">Creative Commons Attribution 4.0 International License</a>.</small>

はじめに
========

自己紹介
--------

* Ryuji Tsutsui @ryu22e
* `一般社団法人PyCon JP Association運営メンバー <https://www.pycon.jp/committee/members.html#ryuji-tsutsui>`_
* Python歴は13年くらい（主にDjango）
* Python Boot Camp、Shonan.py、GCPUG Shonanなどコミュニティ活動もしています
* 著書（共著）：『 `Python実践レシピ <https://gihyo.jp/book/2022/978-4-297-12576-9>`_ 』

今日話したいこと
----------------

Excelの新機能「Python in Excel」について紹介します。

セミナーの対象者
----------------

以下のいずれかに該当する人。

* Excelを使ってデータ解析、集計を行っている人
* Pythonを使ってデータ解析、集計を行っている人
* Excelの新機能に興味がある人

（※Pythonの経験がなくても大丈夫です）

このセミナーをやるモチベーション
--------------------------------

* 私は、Python Boot CampというPythonチュートリアルイベントで講師をやっている
* Pythonに興味を持つ人を増やしたい
* Python in Excelを知ってもらうことで、Pythonの使い道をイメージできる人がいるかも？ という期待がある

セミナーの構成
--------------

* Python in Excelの概要
* Python in Excelの使い方（基本編）
* Python in Excelの使い方（応用編）
* Copilot for Microsoft 365サポートについて紹介

Python in Excelの概要
=====================

Python in Excelとは
-------------------

* ``=PY()`` というExcel関数を使って、ExcelのセルにPythonコードを埋め込める
* 「VBAをPythonで書けるようにする機能」ではない
* 2024年3月2日現在、Python in Excelはプレビュー段階の機能で、Windows版Excel（Excel for Windows）のみで利用可能

導入方法(1)
-----------

プレビュー段階の機能を利用するには、「Microsoft 365 Insider」のベータチャネルにサインアップする。

導入方法(2)
-----------

.. video:: _static/mp4/how-to-install-python-in-excel.mp4

どんな仕組みか
--------------

.. figure:: python-in-excel-image.*
   :alt: PythonコードはMicrosoft Cloud上で実行される

   PythonコードはMicrosoft Cloud上で実行される

セキュリティについて
--------------------

他人が書いた不正なコードの実行を防ぐため、以下の制限がある。

* Excelの外にあるローカルリソースへのアクセス
* ネットワークアクセス
* 数式、グラフ、ピボットテーブル、マクロ、VBA コードなど、Excelブック内の他のプロパティへのアクセス

外部リソースにアクセスできないのは結構不便では？
------------------------------------------------

Power Queryを使って、外部リソースのデータをセルに取り込んでからPythonで読み取ることができる。

Python in Excelの使い方（基本編）
=================================

簡単な計算をやってみる
----------------------

デモをやります。

「出力形式」とは
----------------

``=PY()`` Excel関数の出力形式には、以下の2種類がある。

* Pythonオブジェクト（デフォルト）
* Excelの値

Pythonオブジェクト
------------------

* Pythonコードの実行結果をそのまま埋め込む出力形式
* 以下のようにひし形が2つ重なったアイコンが表示される

ひし形のアイコンからオブジェクトに関する情報を確認
--------------------------------------------------

.. figure:: check-information-about-objects.*
   :alt: オブジェクトに関する情報を確認

   オブジェクトに関する情報を確認

Excelの値
---------

* 出力結果を人間に見せる際に使う出力形式
* 後述するグラフを作成する際にはこれを使う

グラフを作成してみる
--------------------

デモをやります。

「コアライブラリ」とは
----------------------

* Python in ExcelではAnacondaに同梱されているライブラリの一部が利用できる
* よく使うライブラリはimport文を書かずに使える
* これを「コアライブラリ」と呼ぶ

コアライブラリの一覧
--------------------

`Excel のオープンソース ライブラリと Python - Microsoft サポート <https://support.microsoft.com/ja-jp/office/excel-%E3%81%AE%E3%82%AA%E3%83%BC%E3%83%97%E3%83%B3%E3%82%BD%E3%83%BC%E3%82%B9-%E3%83%A9%E3%82%A4%E3%83%96%E3%83%A9%E3%83%AA%E3%81%A8-python-c817c897-41db-40a1-b9f3-d5ffe6d1bf3e>`_ を参照。

Python in Excelの使い方（応用編）
=================================

Python in Excelに関する情報ソース(1)
------------------------------------

Microsoft公式サイト（日本語）

https://support.microsoft.com/ja-jp/office/python-in-excel-%E3%81%AE%E6%A6%82%E8%A6%81-55643c2e-ff56-4168-b1ce-9428c8308545

Python in Excelに関する情報ソース(2)
------------------------------------

Anacondaのチュートリアル動画（英語）

https://freelearning.anaconda.cloud/get-started-with-python-in-excel-course

Python in Excelに関する情報ソース(3)
------------------------------------

Anacondaの公式ブログ（英語）

https://www.anaconda.com/resource-topic/python-in-excel

Python in Excelのベストプラクティスとは
---------------------------------------

Load raw data, convert once, and reuse

https://freelearning.anaconda.cloud/get-started-with-python-in-excel-course/113133

つまり、どういうことか
----------------------

* データはそのままだと使いにくい場合がよくあるが、直接加工しない方がいい
* 直接加工してしまうと再利用が難しくなるので
* データの加工はPythonで行う

横浜市のオープンデータを使ってみる
----------------------------------

横浜市が公開している有効求人倍率のデータをグラフ化してみる。

https://www.city.yokohama.lg.jp/city-info/yokohamashi/tokei-chosa/portal/opendata/rodo-kyujin.html

Copilot for Microsoft 365サポートについて紹介
=============================================

Copilot for Microsoft 365とは
-----------------------------

Microsoft 365 アプリ（Word、Excel、PowerPoint、Outlook、Teamsなど）と連携して大規模言語モデル (LLM) を利用できるアプリ。

導入方法
--------

TODO あとで書く

実際に使ってみる
----------------

デモをやります。

最後に
======

まとめ
------

TODO あとで書く

ご清聴ありがとうございました
----------------------------

.. figure:: thank-you-for-your-attention.*
   :alt: AIが考えた「Python in Excelのパワーのおかげで爆速で仕事を進めるビジネスマン」

   AIが考えた「Python in Excelのパワーのおかげで爆速で仕事を進めるビジネスマン」
