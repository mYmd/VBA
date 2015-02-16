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

'========= 実行例（VBEイミディエイトから）======================  
printM mapL(p_log, Array(1,2,3,4,5,6,7))  
printM zipWith(p_plus, Array(1,2,3,4,5), Array(10, 100, 1000, 100, 10))  
? foldl(p_minus, 0, iota(1, 100)) ' = (...(((0-1)-2)-3)-...-100  
? foldr(p_minus, 0, iota(1, 100)) ' = 0-(1-(2-(3-...(99-100)))...)  
? mapL(p_applyFun, 3, Array(p_minus, 5, Null))  
? mapL(p_applyFun, 3, Array(p_minus, Null, 5))  

'円周率をシミュレーションによって確率的に求める  
N=10000  
points = zip(mapL(p_rnd, repeat(0, N), 1), mapL(p_rnd, repeat(0, N), 1))  
? 4 * count_if(p_less, mapL(p_distance, points, Array(0, 0)), 1.0) / N  

'ロジスティック漸化式  
N = 50  
init = 0.1 : r = 3.754  
m = scanl(p_applyFun, init, repeat(Array(p_Logistic, r, Null), N))  
printM m  
m = scanr(p_setParam, init, repeat(Array(p_Logistic, r, Null), N))  
printM m  

'フィボナッチ数列  
N = 50  
m = unzip(scanl(p_applyFun, Array(0,1), repeat(Array(p_fibonacci, Null, Null), N)), 1)(0)  
printM m  
m = unzip(scanr(p_setParam, Array(0,1), repeat(Array(p_fibonacci, Null, Null), N)), 1)(0)  
printM m  
'fibonacci関数が不要になった  
m = unzip(scanl(p_applyFun2by2 , Array(0,1), repeat(Array(p_secondArg, p_plus), N)), 1)(0)  
printM m  
