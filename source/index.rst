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
* 2024年3月2日現在、Python in Excelはプレビュー段階の機能で、Windows版Excel（Excel for Windows）のみで利用可能

導入方法
--------

プレビュー段階の機能を利用するには、「Microsoft 365 Insider」のベータチャネルにサインアップする。

TODO スクショを入れる

どんな仕組みか
--------------

* PythonコードはMicrosoft Cloud上で実行される
* ローカル環境にPythonのインストールは不要

TODO イメージ図を入れる

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

簡単な足し算をやってみる
------------------------

TODO あとで書く

「出力形式」とは
----------------

TODO あとで書く

グラフを作成してみる
--------------------

TODO あとで書く

Python in Excelの使い方（応用編）
=================================

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

TODO あとで書く

Copilot for Microsoft 365サポートについて紹介
=============================================

Copilot for Microsoft 365とは
-----------------------------

TODO あとで書く

導入方法
--------

TODO あとで書く

実際に使ってみる
----------------

TODO あとで書く

最後に
======

まとめ
------

TODO あとで書く

ご清聴ありがとうございました
----------------------------

TODO あとで書く
