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
  showM.txt  
Declare.txtにあるDeclare文の「Lib "mapM.dll"」部分は、dllの保存フォルダに合わせる。  

'========= 実行例（VBEイミディエイトから）======================  
printM mapL(p_log, Array(1,2,3,4,5,6,7))  
printM zipWith(p_add, Array(1,2,3,4,5), Array(10, 100, 1000, 100, 10))  
?foldl(p_minus, 0, iota(1, 100)) ' = (...(((0-1)-2)-3)-...-100  
?foldr(p_minus, 0, iota(1, 100)) ' = 0-(1-(2-(3-...(99-100)))...)  

'円周率  
N=10000  
points = zip(mapL(p_rnd, repeat(0, N), 1), mapL(p_rnd, repeat(0, N), 1))  
?4 * count_if(p_less, mapL(p_distance, points, Array(0, 0)), 1.0) / N  

'ロジスティック漸化式  
N = 50  
init = 0.1 : r = 3.754  
m = scanl(p_applyAs1st, init, repeat(Array(p_Logistic, r), N))  
printM m  
m = scanr(p_set1stParam, init, repeat(Array(p_Logistic, r), N))  
printM m  

'フィボナッチ数列  
N = 50  
m = unzip(scanl(p_applyTogether, Array(0,1), repeat(p_fibonacci, N)), 1)(0)  
printM m  
m = unzip(scanr(p_setBothParams, Array(0,1), repeat(p_fibonacci, N)), 1)(0)  
printM m  
