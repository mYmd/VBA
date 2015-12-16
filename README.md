# VBA
VBA用のなんちゃってHaskellモジュール  

・関数型プログラミングのサポート  
　　関数を合成して新しい関数オブジェクトとして返す機能  
　　引数を束縛することによる関数の部分適用  
　　関数オブジェクトの実行（遅延評価あり）  
　　関数を関数の引数にする  

・リスト操作によるループ削減  
　　mapやzipWithやfoldやscan系の関数を用意  

・汎用的なソート関数  
　　ユーザー定義の比較関数を渡してソートできる  

VBAHaskellの紹介　その1　（最初はmapF）  
[http://qiita.com/mmYYmmdd/items/c731edf943acc0a0ebe9]  

///////////////////////////////////////////////////////////////////////////////  
'=============================================================  
 構成要素  
・C++ソースをdllとしてコンパイル＆ビルド  
エクスポートする関数はdefファイルで定義  

(Declare宣言はdll名をmapM.dllとする前提にしている)  
(32bit dllバイナリはhttp://home.b07.itscom.net/m-yamada/VBA/mapM.dll)  
(64bit dllバイナリはhttp://home.b07.itscom.net/m-yamada/VBA/mapM64.dll)  

以下のbasファイルはVBAのソースコード。(Office 2010～)
標準モジュールにそのまま取り込む。  
  Haskell_0_declare.bas（Declare文のみ）  
  Haskell_1_Core.bas（中心となるモジュール）  
  Haskell_2_stdFun.bas（Haskell的なリスト操作）  
  Haskell_3_printM.bas（配列表示）  
  Haskell_4_vector.bas（その他配列操作）  
  Haskell_5_sort.bas（ソートとlower_bound等）  
  Haskell_6_iterator.bas（イテレータ）  
  misc_random.bas（乱数）  
  test_module.bas（サンプルプログラム：Sub vbaUnit）  
  misc_*.bas（test_moduleで使用）  
（declare.basにあるDeclare文の「Lib "mapM.dll"」部分はdllの保存フォルダに合わせてパスを補記。）  
2010以前のOfficeでは、Haskell_1_Coreモジュールに2カ所ある LongPtr をLong に変更し、  
Declare文についている 'PtrSafe'宣言をすべて削除れば使用可能。  
'=============================================================  
2015/10/24  
make_funPointer_with_3_parametersを追加  

2015/10/14  
乱数モジュールを追加  

2015/10/01  
lambdaExpr関数を廃止  
yield式を追加  
VS2013の範囲でのC++11化推進  

2015/7/18  
VBAモジュール Haskell_6_iterator を追加  

2015/7/05  
VBAモジュール misc_ratio を追加  

2015/6/13  
64bit Officeに対応  

2015/6/08  
safearrayRefクラスの導入によるリファクタリング  

2015/5/31  
lambdaExpr関数を追加  

2015/5/16  
新しいプレースホルダ ph_1 と ph_2 を追加  

2015/5/6  
bind_invokeの廃止とmoveVariantの追加  
テストモジュールに木構造のテストを追加  

2015/4/26  
functionExprオブジェクトの構築過程をリファクタリング  
API関数にrepeat_impleを追加  

2015/4/14  
APIにfind_impleを追加してfind_pred関数の実装をdll側に移動  

2015/4/12  
find_pred関数をHaskell_1_Coreに追加  

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
