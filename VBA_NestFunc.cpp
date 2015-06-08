//VBA_NestFunc.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
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

//===================================================================
safearrayRef::safearrayRef(const VARIANT* pv)
    :psa(nullptr), pvt(0), dim(0), elemsize(0), it(nullptr)//, size{1,1,1}
{
    ::VariantInit(&val_);
    if (!pv || 0 == (VT_ARRAY & pv->vt))            return;
    psa = (0 == (VT_BYREF & pv->vt))? pv->parray: *pv->pparray;
    if (!psa)                                       return;
    SafeArrayAccessData(psa, reinterpret_cast<void**>(&it));
    dim = SafeArrayGetDim(psa);
    if (!it || 3 < dim)
    {
        dim = 0;
        return;
    }
    elemsize = SafeArrayGetElemsize(psa);
    SafeArrayGetVartype(psa, &pvt);
    val_.vt = pvt | VT_BYREF;   //ここ
    size[0] = size[1] = size[2] = 1;
    for (decltype(dim) i = 0; i < dim; ++i)
    {
        LONG ub = 0, lb = 0;
        ::SafeArrayGetUBound(psa, i+1, &ub);
        ::SafeArrayGetLBound(psa, i+1, &lb);
        size[i] = ub - lb + 1;
    }
}
safearrayRef::~safearrayRef()
{
    if(psa)     SafeArrayUnaccessData(psa);
    VariantClear(&val_);
}

std::size_t safearrayRef::getDim() const
{
    return dim;
}

std::size_t safearrayRef::getSize(std::size_t i) const
{
    return size[i-1];
}

std::size_t safearrayRef::getOriginalLBound(std::size_t i) const
{
    LONG lb = 0;
    ::SafeArrayGetLBound(psa, i, &lb);
    return lb;
}

VARIANT& safearrayRef::operator()(std::size_t i, std::size_t j, std::size_t k)
{
    auto distance = size[0]*size[1]*k + size[0]*j + i;
    if ( pvt == VT_VARIANT)
    {
        return *reinterpret_cast<VARIANT*>(it + distance*elemsize);
    }
    else
    {
        val_.pvarVal = reinterpret_cast<VARIANT*>(it + distance*elemsize);
        return val_;
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
    safearrayRef arRef(bfun);
    if ( 1 != arRef.getDim() )         return;
    if ( arRef.getSize(1) < 4 )        return;
    if ( placeholder_num(&arRef(3)) < 0 )   return;
    VARIANT& elem0 = arRef(0);
    if ( elem0.vt != VT_I4 || elem0.lVal == 0 ) return;
    fun = reinterpret_cast<vbCallbackFunc_t>(elem0.lVal);
    ::VariantCopy(&elem1, &arRef(1));
    ::VariantCopy(&elem2, &arRef(2));
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
