# VBA
VBA用のなんちゃってHaskellモジュール(32bit Office Only)  
mapやzipWithやfoldやscan系の関数の真似事によってユーザーコードからループを  
排除しようとする試み。  

< test_module.bas のサンプルプログラム(vbaUnit)より >  
πを確率的に求めるコードがループなしの3行で書ける。  
    N = 10000  
    Points = zip(mapF(p_rnd(, 1), repeat(0, N)), mapF(p_rnd(, 1), repeat(0, N)))  
    printM Array("π≒", 4 * count_if(p_less(, 1#), mapF(p_distance(, Array(0, 0)), Points)) / N)  

FizzBuzz は２行くらい  
m = Array(Array(p_mod(, 15), Null, "FizzBuzz"), Array(p_mod(, 5), Null, "Buzz"), Array(p_mod(, 3), placeholder, "Fizz"))  
printM foldl1(p_replaceNull, product_set(p_if_else, iota(1, 100), m), 2)  

素数列の生成は 次の2.3.を繰り返し適用することで得られる（効率は考慮外）  
1. m = Array(2, 3, 5)  '初期  
2. z = iota(2, m(UBound(m)) ^ 2)  
3. m = filterR(z, mapF(p_isPrime(, m), z))  

単純なニュートン法による方程式の求根は、(x1, f(x)) から (x2, f(x2)) を出力する１ステップを  
表す関数を作り、繰り返し適用する(関数合成 foldl_Funs)ことで求める  
foldl_Funs(初期値, repeat(p_Newton_Raphson(, Array(f, df/dx)), 回数))  

&lt;/ test_module.bas のサンプルプログラム(vbaUnit)より >  

mapのネストや引数の束縛を実装したので、もっと巧みなことがきるのではないかと  
考えているが、そこまでの知性がない。  

///////////////////////////////////////////////////////////////////////////////  
C++ファイルは４つ  
mapM.cppとvbSort.cppとVBA_NestFunc.hppとVBA_NestFunc.cpp  
をdllとしてコンパイル＆ビルド  
以下の関数をdefファイル等でエクスポート  

	Dimension = Dimension  
	placeholder = placeholder  
	is_placeholder = is_placeholder  
	unbind_invoke = unbind_invoke  
	bind_invoke = bind_invoke  
	mapF_imple = mapF_imple  
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

(mapF.defおよびDeclare宣言はdll名をmapM.dllとする前提にしている)  
(dllバイナリはhttp://home.b07.itscom.net/m-yamada/VBA/mapM.dll)  

以下のbasファイルはVBAソースコード。
標準モジュールにそのまま取り込む。  
  Haskell_0_declare.bas（Declare文のみ）  
  Haskell_1_Core.bas（中心となるモジュール）  
  Haskell_2_stdFun.bas（Haskell的なリスト操作）  
  Haskell_3_printM.bas（配列表示）  
  Haskell_4_vector.bas（その他配列操作）  
  Haskell_5_sort.bas（ソートとlower_bound等）  
  test_module.bas（サンプルプログラム：Sub vbaUnit）  
（declare.basにあるDeclare文の「Lib "mapM.dll"」部分はdllの保存フォルダに合わせてパスを補記。）  

'=============================================================  
2015/4/11  
バグ修正とともにC++側ファイル追加  
VBA_NestFunc.hppとVBA_NestFunc.cpp  

2015/4/10  
関数構造を大幅に変更した  
これによって関数合成がかなり自然に書けるようになった  

2015/4/8  
モジュール名称を全体的に変更  
Haskell_1_Core に以下の関数を追加  
repeat_while, repeat_while_not, generate_while, generate_while_not  

2015/4/7  
unfoldrを追加してみたがあまり使い道がなさそうなのでtestのみ  

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
