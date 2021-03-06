//VBA_NestFunc.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "VBA_NestFunc.hpp"

//VBA配列の次元取得
__int32 __stdcall Dimension(VARIANT const& v) noexcept
{
    if (0 == (VT_ARRAY & v.vt))
        return	0;
    if (0 == (VT_BYREF & v.vt))
        return ::SafeArrayGetDim(v.parray);
    else
        return (v.pparray) ? ::SafeArrayGetDim(*v.pparray) : 0;
}

//プレースホルダの生成
VARIANT __stdcall placeholder(__int32 const n) noexcept
{
    VARIANT ret;
    ::VariantInit(&ret);
    ret.vt = VT_ERROR;
    ret.scode = n;
    return ret;
}

//プレースホルダの判定
__int32 __stdcall is_placeholder(VARIANT const& v) noexcept
{
    return ((v.vt == VT_ERROR) && 0 <= v.scode && (v.scode % 10) <= 2) ? 1 : 0;
}

//===================================================================
VARIANT iVariant(VARTYPE const t) noexcept
{
    VARIANT ret;
    ::VariantInit(&ret);
    ret.vt = t;
    return ret;
}

BSTR getBSTR(VARIANT const& expr) noexcept
{
    if (expr.vt & VT_BYREF)
        return (VT_BSTR == (expr.vt & VT_TYPEMASK) && expr.pbstrVal) ? *expr.pbstrVal : nullptr;
    else
        return (VT_BSTR == (expr.vt & VT_TYPEMASK) && expr.bstrVal) ? expr.bstrVal : nullptr;
}

VARIANT bstrVariant(std::wstring const& s) noexcept
{
    VARIANT ret = iVariant(VT_BSTR);
    ret.bstrVal = ::SysAllocString(s.data());
    return ret;
}

//===================================================================

// SafeArray要素のアクセス
safearrayRef::safearrayRef(VARIANT const& v) noexcept
    :psa{nullptr}, pvt{0}, dim{0}, elemsize{0}, it{nullptr}, size({ 1, 1, 1 })
{
    ::VariantInit(&val_);
    if (0 == (VT_ARRAY & v.vt))             return;
    psa = (0 == (VT_BYREF & v.vt)) ? v.parray : *v.pparray;
    if (!psa)                               return;
    //このAPIのせいでreinterpret_cast
    ::SafeArrayAccessData(psa, reinterpret_cast<void**>(&it));
    dim = ::SafeArrayGetDim(psa);
    if (!it || 3 < dim)
    {
        size[0] = 0;
        return;
    }
    elemsize = ::SafeArrayGetElemsize(psa);
    ::SafeArrayGetVartype(psa, &pvt);
    val_.vt = pvt | VT_BYREF;   //ここ
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
    if (psa)     ::SafeArrayUnaccessData(psa);
    ::VariantClear(&val_);
}

std::size_t safearrayRef::getDim() const noexcept
{
    return dim;
}

std::size_t safearrayRef::getSize(std::size_t i) const noexcept
{
    return size[i-1];
}

std::size_t safearrayRef::getOriginalLBound(std::size_t i) const noexcept
{
    LONG lb = 0;
    ::SafeArrayGetLBound(psa, static_cast<UINT>(i), &lb);
    return lb;
}

VARIANT& safearrayRef::operator()(std::size_t i, std::size_t j, std::size_t k) noexcept
{
    auto distance = size[0]*size[1]*k + size[0]*j + i;
    if (pvt == VT_VARIANT)
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
//bindされている値
class valueExpr : public funcExpr_i {
    VARIANT&     val;
public:
    explicit valueExpr(VARIANT& v) noexcept : val{v} {    }
    valueExpr(valueExpr const&) = delete;
    valueExpr(valueExpr&&) = delete;
    ~valueExpr() = default;
    VARIANT& eval(VARIANT&, VARIANT&, int left_right) noexcept override
    {
        return val;
    }
};

//--------------------------------------------------------
//指定された引数を返すプレースホルダ
class placeholder0 : public funcExpr_i {
public:
    ~placeholder0() = default;
    VARIANT& eval(VARIANT& x, VARIANT& y, int left_right) noexcept override
    {
        return      (left_right == 1) ? x : y;
    }
};

//常に第1引数を返すプレースホルダ
class placeholder1 : public funcExpr_i {
public:
    ~placeholder1() = default;
    VARIANT& eval(VARIANT& x, VARIANT& y, int left_right) noexcept override
    {
        return x;
    }
};

//常に第2引数を返すプレースホルダ
class placeholder2 : public funcExpr_i {
public:
    ~placeholder2() = default;
    VARIANT& eval(VARIANT& x, VARIANT& y, int left_right) noexcept override
    {
        return y;
    }
};

//-------------------------------------------------------------
namespace {   //util
              //プレースホルダの種類
    int placeholder_num(const VARIANT& v) noexcept
    {
        return ((v.vt == VT_ERROR)) ? v.scode : -1;
    }

    //------------------------------------------------------------------
    struct VBCallbackStruct {
        vbCallbackFunc_t    fun;
        VARIANT*            elem1;
        VARIANT*            elem2;
        int                 delay;
        VBCallbackStruct(VARIANT const& bfun) noexcept;
        ~VBCallbackStruct() = default;
        VBCallbackStruct(VBCallbackStruct const&) = delete;
        VBCallbackStruct(VBCallbackStruct&&) = delete;
    };

    //------------------------------------------------------------------
    VBCallbackStruct::VBCallbackStruct(VARIANT const& bfun) noexcept
        : fun{nullptr}, elem1{nullptr}, elem2{nullptr}, delay{0}
    {
        safearrayRef arRef{bfun};
        if (1 != arRef.getDim())                return;
        if (arRef.getSize(1) < 4)               return;
        if (placeholder_num(arRef(3)) < 0)      return;
        auto pF = iVariant();
        if (S_OK != ::VariantChangeType(&pF, &arRef(0), 0, VT_I8))     return;
        fun = reinterpret_cast<vbCallbackFunc_t>(pF.llVal);
        VariantClear(&pF);
        elem1 = &arRef(1);
        elem2 = &arRef(2);
        delay = placeholder_num(arRef(3));//   0, 1, 2
    }

    //
    auto functionExpr_imple(VARIANT& elem, bool delay) noexcept ->std::unique_ptr<funcExpr_i>
    {
        VBCallbackStruct callback{elem};
        try {
            if (callback.fun)
            {
                if (delay)
                    return std::make_unique<innerFunction>(elem, true);
                else
                    return std::make_unique<functionExpr>(callback);
            }
            else
            {
                switch (placeholder_num(elem) % 10)
                {
                case 0:     return std::make_unique<placeholder0>();
                case 1:     return std::make_unique<placeholder1>();
                case 2:     return std::make_unique<placeholder2>();
                default:    return std::make_unique<valueExpr>(elem);
                }
            }
        }
        catch (...) {
            return std::unique_ptr<valueExpr>(nullptr);
        }
    };
}

//--------------------------------------------------------

functionExpr::functionExpr(VARIANT const& bfun) noexcept : functionExpr{VBCallbackStruct{bfun}}
{   }

functionExpr::functionExpr(const VBCallbackStruct& callback) noexcept : fun{callback.fun}
{
    ::VariantInit(&val);
    if (!fun)     return;
    left = functionExpr_imple(*callback.elem1, callback.delay == 1);
    right = functionExpr_imple(*callback.elem2, callback.delay == 2);
}

functionExpr::~functionExpr()
{
    ::VariantClear(&val);
}

VARIANT& functionExpr::eval(VARIANT& x, VARIANT& y, int left_right) noexcept // = 0
{
    if (fun)
    {
        auto tmp = fun( left->eval(x, y, left_right ? left_right : 1),
                        right->eval(x, y, left_right ? left_right : 2)  );
        ::VariantClear(&val);   //計算した後でクリアしなければダメ
        std::swap(val, tmp);
    }
    return val;
}

bool functionExpr::isValid() const noexcept
{
    return fun != nullptr;
}
//-------------------------------------------------------------------------
innerFunction::innerFunction(VARIANT& Val, bool copy) noexcept : val{copy ? myVal : Val}, phn1{-2}, phn2{-2}
{
    ::VariantInit(&myVal);
    if (copy)        ::VariantCopyInd(&myVal, &Val);
}

innerFunction::~innerFunction()
{
    ::VariantClear(&myVal);
}

VARIANT& innerFunction::eval(VARIANT& x, VARIANT& y, int left_right) noexcept
{
    VBCallbackStruct callback{ val };
    eval_imple(*callback.elem1, x, y, left_right, 1);
    eval_imple(*callback.elem2, x, y, left_right, 2);
    return val;
}

void innerFunction::eval_imple(VARIANT& elem, VARIANT& x, VARIANT& y, int left_right, int arg12) noexcept
{
    int& phn = (arg12 == 1)? phn1: phn2;
    std::unique_ptr<innerFunction>& ptr = (arg12 == 1)? arg1: arg2;
    if ( ptr || VBCallbackStruct{elem}.fun )
    {
        if ( !ptr )     ptr = std::make_unique<innerFunction>(elem, false);
        ptr->eval(x, y, left_right);
    }
    else
    {
        if ( phn < -1 )     phn = placeholder_num(elem);
        switch ( phn )
        {
        case 0:
            ::VariantCopy(&elem, &placeholder0{}.eval(x, y, left_right));
            break;
        case 1:
            ::VariantCopy(&elem, &placeholder1{}.eval(x, y, left_right));
            break;
        case 2:
            ::VariantCopy(&elem, &placeholder2{}.eval(x, y, left_right));
            break;
        }
    }
}
