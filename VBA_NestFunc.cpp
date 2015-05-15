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
VARIANT __stdcall placeholder(__int32 n)
{
    VARIANT ret;
    ::VariantInit(&ret);
    ret.vt = VT_ERROR;
    ret.scode = n;
    return ret;
}

//プレースホルダの判定
__int32 __stdcall is_placeholder(const VARIANT* pv)
{
    return ( pv && (pv->vt == VT_ERROR) && 0 <= pv->scode && pv->scode <= 2 ) ? 1 : 0;
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
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0)   {   return &val;    }
};

//--------------------------------------------------------
class placeholder0 : public funcExpr_i    {
public:
    ~placeholder0()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {
        return ( left_right == 1 )?  x: ( left_right == 2 )?  y:  0;
    }
};

class placeholder1 : public funcExpr_i    {
public:
    ~placeholder1()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {   return x;   }
};

class placeholder2 : public funcExpr_i    {
public:
    ~placeholder2()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {   return y;   }
};

//-------------------------------------------------------------
namespace   {
    int placeholder_num(const VARIANT* pv)
    {
        return ( pv && (pv->vt == VT_ERROR) ) ?  pv->scode :  -1;
    }
}
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
        auto isP = placeholder_num(&elem3);
        ::VariantClear(&elem3);
        if ( isP < 0 )                     return;
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
        int pn = -1;
        if ( fun1 )
            left.reset(new functionExpr(fun1));
        else if ( 0 == (pn = placeholder_num(&callback.elem1)) )
            left.reset(new placeholder0);
        else if ( 1 == pn )
            left.reset(new placeholder1);
        else if ( 2 == pn )
            left.reset(new placeholder2);
        else
            left.reset(new valueExpr(&callback.elem1));
    }
    {
        VBCallbackFunc_ fun2(&callback.elem2);
        int pn = -1;
        if ( fun2 )
            right.reset(new functionExpr(fun2));
        else if ( 0 == (pn = placeholder_num(&callback.elem2)) )
            right.reset(new placeholder0);
        else if ( 1 == pn )
            right.reset(new placeholder1);
        else if ( 2 == pn )
            right.reset(new placeholder2);
        else
            right.reset(new valueExpr(&callback.elem2));
    }
}

functionExpr::~functionExpr()
{
    ::VariantClear(&val);
}

VARIANT* functionExpr::eval(VARIANT* x, VARIANT* y, int left_right) // = 0
{
    if ( fun )
    {
        VARIANT tmp = fun(  left->eval(x, y, left_right? left_right : 1),
                           right->eval(x, y, left_right? left_right : 2)    );
        ::VariantClear(&val);
        std::swap(val, tmp);
    }
    return &val;
}
