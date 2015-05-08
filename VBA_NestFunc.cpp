//VBA_NestFunc.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "OAIdl.h"      //wtypes.h
#include "VBA_NestFunc.hpp"


//VBA配列の次元取得
__int32 __stdcall Dimension(const VARIANT* pv)
{
	if ( !pv || 0 == (VT_ARRAY & pv->vt ) )
		return	0;
	if ( 0 == (VT_BYREF & pv->vt) )
		return ::SafeArrayGetDim(pv->parray);
	else
		return (pv->pparray)? ::SafeArrayGetDim(*pv->pparray): 0;
}

//プレースホルダの生成
VARIANT __stdcall placeholder()
{
    VARIANT ret;
    ::VariantInit(&ret);
    ret.vt = VT_ERROR;
    ret.scode = 0;
    return ret;
}

//プレースホルダの判定
__int32 __stdcall is_placeholder(const VARIANT* pv)
{
    return ( pv && (pv->vt == VT_ERROR) && pv->scode == 0 ) ? 1 : 0;
}

//要素数とLBoundを取得
void safeArrayBounds(SAFEARRAY* pArray, UINT dim, SAFEARRAYBOUND bounds[3])
{
    for ( ULONG i = 0; i < dim; ++i )
    {
        ::SafeArrayGetLBound(pArray, i+1, &bounds[i].lLbound);
        LONG ub = 0;
        ::SafeArrayGetUBound(pArray, i+1, &ub);
        bounds[i].cElements = 1 + ub - bounds[i].lLbound;
    }
}

//===================================================================
funcExpr_i::~funcExpr_i()
{   }

//--------------------------------------------------------
class valueExpr : public funcExpr_i    {
    VARIANT     val;
public:
    valueExpr(VARIANT* v)
    {
        ::VariantInit(&val);
        ::VariantCopy(&val, v);
    }
    ~valueExpr()    {   }
    VARIANT* eval(VARIANT*, VARIANT*)   {   return &val;    }
};

//--------------------------------------------------------
class placeholderExpr : public funcExpr_i    {
public:
    ~placeholderExpr()  {   }
    VARIANT* eval(VARIANT* x, VARIANT*) {   return x;   }
};

//-------------------------------------------------------------
functionExpr::VBCallbackFunc_::VBCallbackFunc_(const VARIANT* bfun) : fun(0)
{
    ::VariantInit(&elem1);
    ::VariantInit(&elem2);
    if ( 1 != Dimension(bfun) )         return;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & bfun->vt) )?  (bfun->parray): (*bfun->pparray);
    if ( !pArray )                      return;
    SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
    safeArrayBounds(pArray, 1, &bounds);
    if ( bounds.cElements < 4 )         return;
    LONG index = bounds.lLbound + 3;
    {
        VARIANT elem3;
        ::VariantInit(&elem3);
        ::SafeArrayGetElement(pArray, &index, &elem3);
        auto isP = is_placeholder(&elem3);
        ::VariantClear(&elem3);
        if ( !isP )                     return;
    }
    {
        VARIANT elem0;
        ::VariantInit(&elem0);
        index = bounds.lLbound + 0;
        ::SafeArrayGetElement(pArray, &index, &elem0);
        if ( elem0.vt != VT_I4 || elem0.lVal == 0 )
        {
            ::VariantClear(&elem0);
            return;
        }
        fun = reinterpret_cast<vbCallbackFunc_t>(elem0.lVal);
        ::VariantClear(&elem0);
    }
    {
        index = bounds.lLbound + 1;
        ::SafeArrayGetElement(pArray, &index, &elem1);
        index = bounds.lLbound + 2;
        ::SafeArrayGetElement(pArray, &index, &elem2);
    }
}

functionExpr::VBCallbackFunc_::~VBCallbackFunc_()
{
    ::VariantClear(&elem1);
    ::VariantClear(&elem2);
}

functionExpr::VBCallbackFunc_::operator bool() const
{
    return 0 != fun;
}

//--------------------------------------------------------
functionExpr::functionExpr(VBCallbackFunc_& callback) : fun(callback.fun)
{
    ::VariantInit(&val);
    {
        VBCallbackFunc_ fun1(&callback.elem1);
        if ( fun1 )
            left.reset(new functionExpr(fun1));
        else if ( is_placeholder(&callback.elem1) )
            left.reset(new placeholderExpr);
        else
            left.reset(new valueExpr(&callback.elem1));
    }
    {
        VBCallbackFunc_ fun2(&callback.elem2);
        if ( fun2 )
            right.reset(new functionExpr(fun2));
        else if ( is_placeholder(&callback.elem2) )
            right.reset(new placeholderExpr);
        else
            right.reset(new valueExpr(&callback.elem2));
    }
}

functionExpr::~functionExpr()
{
    ::VariantClear(&val);
}

VARIANT* functionExpr::eval(VARIANT* x, VARIANT* y)
{
    if ( fun )
    {
        VARIANT tmp = fun(left->eval(x, x), right->eval(y, y));
        ::VariantClear(&val);
        std::swap(val, tmp);
    }
    return &val;
}
