README

# 前提

とあるVBAが仕込まれたフォーマットExcelがあり、  
それによって作成されたExcelがあり、  
フォーマットが変更になった場合を想定。  

# 概要

対象ディレクトリのExcel全てを読み出し、各ExcelのVBAを実行する。  
このため、そのExcelのVBAの処理内容に大きく依存する。  

## 使い方

ディレクトリ構成：  
┠0-tool  
┃┠format  
┃┃┗format.xlsm  
┃┠work  
┃┗tool.xlsm（本ツール)  
┃  
┠1-fromExcel  
┃  
┗2-toExcel  

準備：  

1. \format\format.xlsmという名称で元となるフォーマットファイルを格納。  
1. フォーマット変更したいファイルを「1-fromExcel」に格納  
※ サブディレクトリ可。  

実行：

1. VBAを実行。  
2. 「2_toExcel」に修正後ファイルが格納される。  

# VBA詳細仕様

1. "work"、"2-toExcel"ディレクトリ配下のファイルを全削除。  
1. "1-fromExcel"にある修正対象Excelを探索。階層もぐりつつ全量。  
1. format.xlsmをworkディレクトリにコピー。  
1. 対象Excelの記述データを\work\format.xlsmにコピー。  
※上記は固定名称なので、シート名が異なったり、他のシートが存在しても、その分はコピーできない。  
1. format.xlsmのVBAを走らせ、再フォーマットを行わせる。  
1. コピーと再フォーマットが完了したformat.xlsmに対して以下を行う。  
  6.1. 修正後ディレクトリ"2-toExcel"に移動。  
  6.2. ファイル名も修正対象ファイル名に変更。  
1. 次のファイルに修正対象を移す。以降"1."から繰り返し。
