# VBA
VBA用のなんちゃってHaskellモジュール(32bit Office Only)  
mapやzipWithやfoldやscan系の関数の真似事によってユーザーコードからループを  
排除しようとする試み。  
πを確率的に求めるコードがループなしの3行で書ける。（-> test.txt : vbaUnit）  

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

以下のテキストファイルはVBAソースコード。
標準モジュールにそのままコピーする。  
  declare_module.txt（Declare文のみ）  
  mapM_module.txt（中心となるモジュール）  
  vector_module.txt（その他配列操作）  
  printM_module.txt（配列表示）  
  sort_module.ext（ソートとlower_bound等）  
  test_module.txt（サンプルプログラム：Sub vbaUnit）  
（Declare.txtにあるDeclare文の「Lib "mapM.dll"」部分はdllの保存フォルダに合わせてパスを補記。）  

'=============================================================  
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
