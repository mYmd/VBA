//VBA_NetFunc.cpp
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

//---------------------------------------
funcExpr_i::~funcExpr_i()
{   }

//---------------------------------------
valueExpr::valueExpr(VARIANT* v)
{
    ::VariantInit(&val);
    ::VariantCopy(&val, v);
}

valueExpr::~valueExpr()
{   }

VARIANT* valueExpr::eval(VARIANT*, VARIANT*)
{
    return &val;
}

//---------------------------------------
placeholderExpr::~placeholderExpr()
{   }

VARIANT* placeholderExpr::eval(VARIANT* x, VARIANT*)
{
    return x;
}

//---------------------------------------

VBCallbackFunc get_bindFun(const VARIANT* bfun, VARIANT& elem1, VARIANT& elem2)
{
    VBCallbackFunc ret = 0;
    if ( 1 != Dimension(bfun) )         return 0;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & bfun->vt) )?  (bfun->parray): (*bfun->pparray);
    if ( !pArray )                      return 0;
    SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
    safeArrayBounds(pArray, 1, &bounds);
    if ( bounds.cElements != 4 )        return 0;
    LONG index = bounds.lLbound + 3;
    {
        VARIANT elem3;
        ::VariantInit(&elem3);
        ::SafeArrayGetElement(pArray, &index, &elem3);
        if ( !is_placeholder(&elem3) )  return 0;
    }{
        VARIANT elem0;
        ::VariantInit(&elem0);
        index = bounds.lLbound + 0;
        ::SafeArrayGetElement(pArray, &index, &elem0);
        if ( elem0.vt != VT_I4 || elem0.lVal == 0 )     return 0;
        ret = reinterpret_cast<VBCallbackFunc>(elem0.lVal);
    }{
        ::VariantInit(&elem1);
        index = bounds.lLbound + 1;
        ::SafeArrayGetElement(pArray, &index, &elem1);
        ::VariantInit(&elem2);
        index = bounds.lLbound + 2;
        ::SafeArrayGetElement(pArray, &index, &elem2);
    }
    return ret;
}

functionExpr::functionExpr(VBCallbackFunc f, VARIANT& elem1, VARIANT& elem2) : fun(f)
{
    ::VariantInit(&val);
    {
        VARIANT elem11, elem12;
        VBCallbackFunc fun1;
        if ( fun1 = get_bindFun(&elem1, elem11, elem12) )
            left.reset(new functionExpr(fun1, elem11, elem12));
        else if ( is_placeholder(&elem1) )
            left.reset(new placeholderExpr);
        else
            left.reset(new valueExpr(&elem1));
    }{
        VARIANT elem21, elem22;
        VBCallbackFunc fun2;
        if ( fun2 = get_bindFun(&elem2, elem21, elem22) )
            right.reset(new functionExpr(fun2, elem21, elem22));
        else if ( is_placeholder(&elem2) )
            right.reset(new placeholderExpr);
        else
            right.reset(new valueExpr(&elem2));
    }
}

functionExpr::~functionExpr()
{
    ::VariantClear(&val);
}

VARIANT* functionExpr::eval(VARIANT* x, VARIANT* y)
{
    ::VariantClear(&val);
    if ( fun )
    {
        VARIANT tmp = fun(left->eval(x, x), right->eval(y, y));
        ::VariantCopy(&val, &tmp);
    }
    return &val;
}
