# VBA
mapM

VBA用のなんちゃってHaskellモジュール(32bit Office Only)

mapM.cpp をdllでコンパイル＆ビルドし、以下の関数をdefファイル等でエクスポート  
	Dimension = Dimension  
	mapM = mapM  
	zipWith = zipWith  
	foldl = foldl  
	foldr = foldr  
	foldl1 = foldl1  
	foldr1 = foldr1  
	scanl = scanl  
	scanr = scanr  
	scanl1 = scanl1  
	scanr1 = scanr1  
	count_if = count_if  

以下のテキストファイルはVBAソースコード。
標準モジュールにそのままコピーすればOK。  
  Declare.txt  
  mapM.txt  
  vector.txt  
  showM.txt  
Declare.txtにあるDeclare文の「Lib "mapM.dll"」部分は、dllの保存フォルダに合わせて適宜書き換える。

