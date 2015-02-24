# VBA
mapM  

VBA用のなんちゃってHaskellモジュール(32bit Office Only)  

mapM.cpp をコンパイル＆ビルドしdll化、以下の関数をdefファイル等でエクスポート  
	Dimension = Dimension  
	simple_invoke = simple_invoke  
	mapL = mapL  
	mapR = mapR  
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
  printM.txt  
Declare.txtにあるDeclare文の「Lib "mapM.dll"」部分は、dllの保存フォルダに合わせる。  

'========= 実行例はtest.txt ======================  
2015/2/24  
左辺値参照関連の関数群を破棄した。（variantRef、forward_as_tuple等）  
これによる影響は、配列をデータで埋める3種類の関数fillM、fillRow、fillColがSubになったことである。  
大きな配列を値で返すことが非効率であるという理由から元々Subであったが、配列を左辺値参照で返す方法が  
見つかったためFunction化し、関連する様々な部品も用意した。  
しかし左辺値参照変数を配列に代入するとデータが壊れるというVBAの挙動により安全な利用が見込めなくなった。  
