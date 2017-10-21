# VBA
VBA用の疑似Haskellモジュール<br>

・関数型プログラミングのサポート<br>
　　関数を合成して新しい関数オブジェクトとして返す機能<br>
　　引数を束縛することによる関数の部分適用<br>
　　関数オブジェクトの実行（遅延評価あり）<br>
　　関数を関数の引数にする<br>
<br>
・リスト操作によるループ削減<br>
　　mapやzipWithやfoldやscan系の関数を用意<br>
<br>
・汎用的なソート関数<br>
　　ユーザー定義の比較関数を渡してソートできる<br>
<br>
VBAHaskellの紹介　その1　（最初はmapF）<br>
[http://qiita.com/mmYYmmdd/items/c731edf943acc0a0ebe9]<br>
<br>
関数リファレンス<br>
http://home.b07.itscom.net/m-yamada/vh_Manual/VBAHaskell_reference.htm<br>
<br>
///////////////////////////////////////////////////////////////////////////////<br>
=============================================================<br>
 構成要素<br>
基本的にはここのbinフォルダにある「VBAHaskellほぼ全部入り.xlsm」に記載している通り。<br>
https://github.com/mYmd/VBA/blob/master/bin/VBAHaskell%E3%81%BB%E3%81%BC%E5%85%A8%E9%83%A8%E5%85%A5%E3%82%8A.xlsm<br>
 もしくは<br>
http://home.b07.itscom.net/m-yamada/VBA/VBAHaskell%E3%81%BB%E3%81%BC%E5%85%A8%E9%83%A8%E5%85%A5%E3%82%8A.xlsm<br>
<br>
手動でやる方法<br>
<br>
・C++ソースをdllとしてコンパイル＆ビルド<br>
エクスポートする関数はdefファイルで定義<br>
(Declare宣言はdll名をmapM.dllとする前提にしている)<br>
<br>
以下のbasファイル／clsファイルはVBAのソースコード。(Office 2010～)<br>
標準モジュールとクラスモジュールにそのまま取り込む。<br>
Haskell_0_declare.bas（Declare文のみ）<br>
Haskell_1_Core.bas（中心となるモジュール）<br>
Haskell_2_stdFun.bas（Haskell的なリスト操作）<br>
Haskell_3_printM.bas（配列表示）<br>
Haskell_4_vector.bas（その他配列操作）<br>
Haskell_5_sort.bas（ソートとlower_bound等）<br>
misc_random.bas（乱数）<br>
test_module.bas（サンプルプログラム：Sub vbaUnit）<br>
misc_*.bas（test_moduleで使用）<br>
vh_stdvec.cls<br>
vh_pipe.cls（パイプライン演算子）<br>
<br>
declare.basにあるDeclare文の「Lib "mapM.dll"」部分はdllの保存フォルダに合わせてパスを補記するか<br>
SetDefaultDllDirectoriesとAddDllDirectoryを用いてサーチパスに追加する。（VBAHaskellほぼ全部入り.xlsm 参照）<br>
2010以前のOfficeでは、Haskell_1_Coreモジュールに2カ所ある LongPtr をLong に変更し、<br>
Declare文についている 'PtrSafe'宣言をすべて削除すれば使用可能。<br>
=============================================================<br>
2017/10/01<br>
C++APIにfind_best_imple、VBAにfind_best_predを追加<br>
<br>
2017/1/15<br>
dllバイナリを外部URLからここのbinフォルダに変更<br> 
<br>
2016/10/26<br>
C++APIにself_zipWithを追加（1次元配列の離れた要素間で2項操作を適用する関数）<br>
mapM.defとmisc_utility.basにそれを反映<br>
<br>
2016/7/4<br>
vh_pipe.cls を追加<br>
<br>
2016/6/5<br>
vh_stdvec.cls をいろいろ変更<br>
<br>
2016/3/23<br>
ファイル classMemberCopy.cpp を追加<br>
<br>
2016/3/7<br>
Haskell_6_iterator.bas を廃止して vh_stdvec.cls を追加<br>
<br>
2016/1/9<br>
C++APIにchangeLBoundを追加（SafeArrayのLBoundを変更する関数）<br>
mapM.defとHaskell_0_declare.basにそれを反映<br>
<br>
2015/12/24<br>
大半の関数をnoexceptとする<br>
（エラーが出るコンパイラでは削除もしくはマクロでの対応が必要）<br>
<br>
2015/10/24<br>
make_funPointer_with_3_parametersを追加<br>
<br>
2015/10/14<br>
乱数モジュールを追加<br>
<br>
2015/10/01<br>
lambdaExpr関数を廃止<br>
yield式を追加<br>
VS2013の範囲でのC++11化<br>
<br>
2015/7/18<br>
VBAモジュール Haskell_6_iterator を追加<br>
<br>
2015/7/05<br>
VBAモジュール misc_ratio を追加<br>
<br>
2015/6/13<br>
64bit Officeに対応<br>
<br>
2015/6/08<br>
safearrayRefクラスの導入によるリファクタリング<br>
<br>
2015/5/31<br>
lambdaExpr関数を追加<br>
<br>
2015/5/16<br>
新しいプレースホルダ ph_1 と ph_2 を追加<br>
<br>
2015/5/6<br>
bind_invokeの廃止とmoveVariantの追加<br>
テストモジュールに木構造のテストを追加<br>
<br>
2015/4/26<br>
functionExprオブジェクトの構築過程をリファクタリング<br>
API関数にrepeat_impleを追加<br>
<br>
2015/4/14<br>
APIにfind_impleを追加してfind_pred関数の実装をdll側に移動<br>
<br>
2015/4/12<br>
find_pred関数をHaskell_1_Coreに追加<br>
<br>
2015/4/11<br>
バグ修正とともにC++側ファイル追加<br>
VBA_NestFunc.hppとVBA_NestFunc.cpp<br>
<br>
2015/4/10<br>
関数構造を大幅に変更した<br>
これによって関数合成がかなり自然に書けるようになった<br>
<br>
2015/4/8<br>
モジュール名称を全体的に変更<br>
Haskell_1_Core に以下の関数を追加<br>
repeat_while, repeat_while_not, generate_while, generate_while_not<br>
<br>
2015/4/7<br>
unfoldrを追加してみたがあまり使い道がなさそうなのでtestのみ<br>
<br>
2015/3/17<br>
VBAモジュールの拡張子をtxtからbasへ変更。<br>
<br>
2015/3/12<br>
ソートとlower_bound等を追加
(vbSort.cpp , sort_module.txt)
<br>
2015/3/6<br>
count_ifをC++側から削除し、VBA側の通常関数にした。<br>
slashR, slashC　をそれぞれ filterR, filterC に名称変更した。<br>
<br>
2015/3/5<br>
引数をキャプチャする方式に変更。<br>
これに伴い、ユーザコードではmapLやmapRに代わってmapFを主に使用するように変更。<br>
<br>
2015/3/4<br>
サンプルにFizzBuzzを追加<br>
<br>
2015/2/24<br>
左辺値参照関連の関数群を破棄した。（variantRef、forward_as_tuple等）<br>
これによる影響は、配列をデータで埋める3種類の関数fillM、fillRow、fillColがSubになったことである。<br>
大きな配列を値で返すことが非効率であるという理由から元々Subであったが、配列を左辺値参照で返す<br>
方法が見つかったためFunction化し、関連する様々な部品も用意した。<br>
しかし左辺値参照変数を配列に代入するとデータが壊れるというVBAの挙動により安全な利用が見込めなくなった。<br>
