//VBA_NestFunc.hpp
//Copyright (c) 2015 mmYYmmdd
#include "OAIdl.h"      //wtypes.h
#include <memory>

//VBA配列の次元取得
__int32 __stdcall   Dimension(const VARIANT* pv);

//プレースホルダの生成
VARIANT __stdcall   placeholder(__int32);

//プレースホルダの判定
__int32 __stdcall   is_placeholder(const VARIANT* pv);

//bindされていないVBA関数を2引数で呼び出す
VARIANT __stdcall   unbind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2);

//bindされたVBA関数を1引数で呼び出す
VARIANT __stdcall   bind_invoke(VARIANT* bfun, VARIANT* param);

//要素数とLBoundを取得
void    safeArrayBounds(SAFEARRAY* pArray, UINT dim, SAFEARRAYBOUND bounds[]);

//--------------------------------------------------------
class funcExpr_i    {
public:
    virtual ~funcExpr_i();
    virtual VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) = 0;
};

//--------------------------------------------------------
class functionExpr : public funcExpr_i    {
    //使用する唯一のVBAコールバック関数型  vbCallbackFunc_t
    //VBAにおけるシグネチャは
    // Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
    // もしくは
    // Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
    typedef VARIANT (__stdcall *vbCallbackFunc_t)(VARIANT*, VARIANT*);
    vbCallbackFunc_t  fun;
    VARIANT         val;
    std::unique_ptr<funcExpr_i> left;
    std::unique_ptr<funcExpr_i> right;
public:
        class VBCallbackFunc_   {
            friend class functionExpr;
            vbCallbackFunc_t    fun;
            VARIANT             elem1;
            VARIANT             elem2;
        public:
            VBCallbackFunc_(const VARIANT* bfun);
            ~VBCallbackFunc_();
            operator bool() const;
        };
    //---------------------------------
    functionExpr(VBCallbackFunc_&);
    ~functionExpr();
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0);
};

typedef functionExpr::VBCallbackFunc_   VBCallbackFunc;
