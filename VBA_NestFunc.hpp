//VBA_NetFunc.hpp
//Copyright (c) 2015 mmYYmmdd

#include <memory>


//使用する唯一のVBAコールバック関数型  VBCallbackFunc  の宣言
//VBAにおけるシグネチャは
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// もしくは
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
typedef VARIANT (__stdcall *VBCallbackFunc)(VARIANT*, VARIANT*);

//VBA配列の次元取得
__int32 __stdcall   Dimension(const VARIANT* pv);

//プレースホルダの生成
VARIANT __stdcall   placeholder();

//プレースホルダの判定
__int32 __stdcall   is_placeholder(const VARIANT* pv);

//bindされていないVBA関数を2引数で呼び出す
VARIANT __stdcall   unbind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2);

//bindされたVBA関数を1引数で呼び出す
VARIANT __stdcall   bind_invoke(VARIANT* bfun, VARIANT* param);

//要素数とLBoundを取得
void    safeArrayBounds(SAFEARRAY* pArray, UINT dim, SAFEARRAYBOUND bounds[]);

VBCallbackFunc  get_bindFun(const VARIANT* bfun, VARIANT& elem1, VARIANT& elem2);

//--------------------------------------------------------
class funcExpr_i    {
public:
    virtual ~funcExpr_i();
    virtual VARIANT* eval(VARIANT*, VARIANT*) = 0;
};
//--------------------------------------------------------
class valueExpr : public funcExpr_i    {
    VARIANT     val;
public:
    valueExpr(VARIANT* v);
    ~valueExpr();
    VARIANT* eval(VARIANT*, VARIANT*);
};
//--------------------------------------------------------
class placeholderExpr : public funcExpr_i    {
public:
    ~placeholderExpr();
    VARIANT* eval(VARIANT*, VARIANT*);
};
//--------------------------------------------------------
class functionExpr : public funcExpr_i    {
    VBCallbackFunc  fun;
    VARIANT         val;
    std::unique_ptr<funcExpr_i> left;
    std::unique_ptr<funcExpr_i> right;
public:
    functionExpr(VBCallbackFunc, VARIANT&, VARIANT&);
    ~functionExpr();
    VARIANT* eval(VARIANT*, VARIANT*);
};

