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
        ::SafeArrayGetUBound(psa, static_cast<UINT>(i+1), &ub);
        ::SafeArrayGetLBound(psa, static_cast<UINT>(i+1), &lb);
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
    ::SafeArrayGetLBound(psa, static_cast<UINT>(i), &lb);
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

bool funcExpr_i::isYielder() const
{
    return false;
}

//--------------------------------------------------------
//bindされている値
class valueExpr : public funcExpr_i    {
    VARIANT*     val;
public:
    valueExpr(VARIANT* v) : val(v)    {    }
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0)   {   return val;    }
};

//--------------------------------------------------------
//指定された引数を返すプレースホルダ
class placeholder0 : public funcExpr_i    {
public:
    ~placeholder0()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {
        return      ( left_right == 1 )?  x
                :   ( left_right == 2 )?  y
                :   nullptr;
    }
};

//常に第1引数を返すプレースホルダ
class placeholder1 : public funcExpr_i    {
public:
    ~placeholder1()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {   return x;   }
};

//常に第2引数を返すプレースホルダ
class placeholder2 : public funcExpr_i    {
public:
    ~placeholder2()  {   }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {   return y;   }
};

//-------------------------------------------------------------
class yielder : public funcExpr_i    {
    VARIANT     val;
public:
    yielder(__int32 n) : val(placeholder(n))     {   }
    ~yielder()  {   }
    bool isYielder() const
    {   return true;    }
    VARIANT* eval(VARIANT* x, VARIANT* y, int left_right = 0)
    {   return &val;    }
};
//-------------------------------------------------------------
namespace   {   //util
    //プレースホルダの種類
    int placeholder_num(const VARIANT* pv)
    {   return ( pv && (pv->vt == VT_ERROR) ) ?  pv->scode :  -1;   }
    //---------------------------------------------------------------------------------
    struct VBCallbackStruct   {
        vbCallbackFunc_t    fun;
        VARIANT*            elem1;
        VARIANT*            elem2;
        VBCallbackStruct(const VARIANT* bfun);
    };
    VBCallbackStruct::VBCallbackStruct(const VARIANT* bfun) : fun(nullptr), elem1(nullptr), elem2(nullptr)
    {
        safearrayRef arRef(bfun);
        if ( 1 != arRef.getDim() )         return;
        if ( arRef.getSize(1) < 4 )        return;
        if ( placeholder_num(&arRef(3)) < 0 )   return;
        VARIANT pF;
        VariantInit(&pF);
        if ( S_OK != ::VariantChangeType(&pF, &arRef(0), 0, VT_I8))     return;
        fun = reinterpret_cast<vbCallbackFunc_t>(pF.llVal);
        VariantClear(&pF);
        elem1 = &arRef(1);
        elem2 = &arRef(2);
    }
    //
    auto functionExpr_imple(VARIANT* elem) ->std::unique_ptr<funcExpr_i>
    {
        VBCallbackStruct callback(elem);
        int pn = -1;
        if ( callback.fun )
            return std::make_unique<functionExpr>(callback);
        else
            switch (pn = placeholder_num(elem))
            {
            case -1:    return std::make_unique<valueExpr>(elem);
            case 0:     return std::make_unique<placeholder0>();
            case 1:     return std::make_unique<placeholder1>();
            case 2:     return std::make_unique<placeholder2>();
            default:    return std::make_unique<yielder>(pn % 10);
            }
    };
}

//--------------------------------------------------------

functionExpr::functionExpr(const VARIANT* bfun) : functionExpr(VBCallbackStruct(bfun))
{   }

functionExpr::functionExpr(const VBCallbackStruct& callback) : fun(callback.fun)
{
    ::VariantInit(&val);
    if ( !fun )     return;
    left  = functionExpr_imple(callback.elem1);
    right = functionExpr_imple(callback.elem2);
}

functionExpr::~functionExpr()
{
    ::VariantClear(&val);
}

VARIANT* functionExpr::eval(VARIANT* x, VARIANT* y, int left_right) // = 0
{
    if ( left->isYielder() || right->isYielder() )
    {
        VARIANT ret;
        ::VariantInit(&ret);
        ret.vt = VT_ARRAY | VT_VARIANT;
        SAFEARRAYBOUND Bounds = {4, 0};
        ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &Bounds);
        safearrayRef arOut(&ret);
        arOut(0).vt = VT_I8;
        arOut(0).llVal = reinterpret_cast<decltype(arOut(0).llVal)>(fun);
        ::VariantCopy(&arOut(1), left->eval(x, y, left_right? left_right : 1));
        ::VariantCopy(&arOut(2), right->eval(x, y, left_right? left_right : 2));
        VARIANT tmp = placeholder(0);
        ::VariantCopy(&arOut(3), &tmp);
        ::VariantClear(&val);   //計算した後でクリアしなければダメ
        std::swap(val, ret);
    }
    else if ( fun )
    {
        VARIANT tmp = fun(  left->eval(x, y, left_right? left_right : 1),
                           right->eval(x, y, left_right? left_right : 2)    );
        ::VariantClear(&val);   //計算した後でクリアしなければダメ
        std::swap(val, tmp);
    }
    return &val;
}

bool functionExpr::isValid() const
{
    return fun != nullptr;
}

