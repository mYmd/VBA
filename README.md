# VBA
mapM  

VBA用のなんちゃってHaskellモジュール(32bit Office Only)  
variantRef関数によってVariant変数を左辺値参照っぽく使えて、  
forward_as_tupleによってそれを配列化することもできる。  

mapM.cpp をコンパイル＆ビルドしdll化、以下の関数をdefファイル等でエクスポート  
	Dimension = Dimension  
	variantRef = variantRef  
	isVariantRef = isVariantRef  
	variantDeRef = variantDeRef  
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
