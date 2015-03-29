# VBA
VBA用のなんちゃってHaskellモジュール(32bit Office Only)  
mapやzipWithやfoldやscan系の関数の真似事によってユーザーコードからループを  
排除しようとする試み。  

πを確率的に求めるコードがループなしの3行で書ける。（-> test.txt : vbaUnit）  
N = 10000  
Points = zip(mapF(p_rnd(, 1), repeat(0, N)), mapF(p_rnd(, 1), repeat(0, N)))  
printM Array("π≒", 4 * count_if(p_less(, 1#), mapF(p_distance, Points)) / N)  

FizzBuzz は２行くらい  
m = Array(Array(p_mod(, 15), Null, "FizzBuzz"), Array(p_mod(, 5), Null, "Buzz"), Array(p_mod(, 3), placeholder, "Fizz"))  
printM foldl1(p_replaceNull, product_set(p_if_else, iota(1, 100), m), 2)  

素数列の生成は 次の2.3.を繰り返し適用することで得られる（効率は考慮外）  
1. m = Array(2, 3, 5)  '初期  
2. z = iota(2, m(UBound(m)) ^ 2)  
3. m = filterR(z, mapF(p_isPrime(, m), z))  

mapのネストや引数の束縛を実装したので、もっと巧みなことがきるのではないかと  
考えているが、そこまでの知性がない。  
///////////////////////////////////////////////////////////////////////////////  
mapM.cppとvbSort.cpp をコンパイル＆ビルドしdll化、以下の関数をdefファイル等でエクスポート  
	Dimension = Dimension  
	placeholder = placeholder  
	is_placeholder = is_placeholder  
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
	stdsort = stdsort  	

(dllバイナリはhttp://home.b07.itscom.net/m-yamada/VBA/mapM.dll)  
以下のbasファイルはVBAソースコード。
標準モジュールにそのまま取り込む。  
  declare_module.bas（Declare文のみ）  
  mapM_module.bas（中心となるモジュール）  
  vector_module.bas（その他配列操作）  
  printM_module.bas（配列表示）  
  sort_module.bas（ソートとlower_bound等）  
  test_module.bas（サンプルプログラム：Sub vbaUnit）  
（declare.basにあるDeclare文の「Lib "mapM.dll"」部分はdllの保存フォルダに合わせてパスを補記。）  

'=============================================================  
2015/3/17  
VBAモジュールの拡張子をtxtからbasへ変更。  

2015/3/12  
ソートとlower_bound等を追加
(vbSort.cpp , sort_module.txt)

2015/3/6  
count_ifをC++側から削除し、VBA側の通常関数にした。  
slashR, slashC　をそれぞれ filterR, filterC に名称変更した。  

2015/3/5  
引数をキャプチャする方式に変更。  
これに伴い、ユーザコードではmapLやmapRに代わってmapFを主に使用するように変更。  

2015/3/4  
サンプルにFizzBuzzを追加  

2015/2/24  
左辺値参照関連の関数群を破棄した。（variantRef、forward_as_tuple等）  
これによる影響は、配列をデータで埋める3種類の関数fillM、fillRow、fillColがSubになったことである。  
大きな配列を値で返すことが非効率であるという理由から元々Subであったが、配列を左辺値参照で返す  
方法が見つかったためFunction化し、関連する様々な部品も用意した。  
しかし左辺値参照変数を配列に代入するとデータが壊れるというVBAの挙動により安全な利用が見込め  
なくなった。  
