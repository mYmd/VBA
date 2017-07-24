//mapM.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include <list>
#include "VBA_NestFunc.hpp"

namespace {
    //基本型のみ想定したmin
    template <typename T>  T minV(T a, T b) { return (a < b) ? a : b; }

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  functionExpr&   bfun,
                        VARIANT*        init,
                        VARIANT*        matrix,
                        __int32 const   axis,
                        VARIANT&        ret,
                        bool const      left) noexcept; //left==true, right == false

                                        //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun,
                        VARIANT*        init,
                        VARIANT*        matrix,
                        __int32 const   axis,
                        VARIANT&        ret,
                        bool const      left) noexcept; //left==true, right == false

                                        //repeat_while と repeat_while_not と generate_while と generate_while_not の共通処理
    __int32 repeat_imple_0( VARIANT const*  init,
                            functionExpr&   pred,
                            functionExpr&   trans,
                            __int32 const   maxN,
                            VARIANT&        ret,
                            bool const      scan,
                            bool const      stopCondition) noexcept;
}   // namespace

    //bindされていないVBA関数を2引数で呼び出す
VARIANT  __stdcall
unbind_invoke(VARIANT const* bfun, VARIANT* param1, VARIANT* param2) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (func.isValid())
        std::swap(ret, *func.eval(param1, param2));
    return ret;
}

//VARIANT変数のswap
__int32 __stdcall
swapVariant(VARIANT* a, VARIANT* b) noexcept
{
    if (a && b)
    {
        if ( 0 == (a->vt & VT_BYREF) )
        {
            if (0 == (b->vt & VT_BYREF))
            {
                std::swap(*a, *b);
                return 1;
            }
            else
                ::VariantCopyInd(a, b);
        }
        else if (0 == (b->vt & VT_BYREF))
            ::VariantCopyInd(b, a);
    }
    return 0;
}

//SafeArrayのLBoundを変更
void __stdcall changeLBound(VARIANT* pv, __int32 const b) noexcept
{
    if (!pv || 0 == (VT_ARRAY & pv->vt))            return;
    auto psa = (0 == (VT_BYREF & pv->vt)) ? pv->parray : *pv->pparray;
    auto dim = ::SafeArrayGetDim(psa);
    for (decltype(dim) i = 0; i < dim; ++i)
        psa->rgsabound[i].lLbound = b;
}

////************************************************************************************

//配列matrixの各要素にVBA関数を適用する
VARIANT  __stdcall
mapF_imple(VARIANT const* bfun, VARIANT* matrix) noexcept
{
    functionExpr func{bfun};
    if (!matrix || !func.isValid())       return iVariant();
    if (0 == (VT_ARRAY & matrix->vt))
    {
        auto ret = iVariant();
        std::swap(ret, *func.eval(matrix, matrix));
        return ret;
    }
    safearrayRef arIn{matrix};
    if (arIn.getDim() == 0)                      return iVariant();
    std::array<SAFEARRAYBOUND, 3>   Bounds{
        {
            { static_cast<ULONG>(arIn.getSize(1)), 0 },
            { static_cast<ULONG>(arIn.getSize(2)), 0 },
            { static_cast<ULONG>(arIn.getSize(3)), 0 }
        }
    };
    // SAFEARRAY作成
    auto ret = iVariant(VT_ARRAY | VT_VARIANT);
    ret.parray = ::SafeArrayCreate(VT_VARIANT, static_cast<UINT>(arIn.getDim()), Bounds.data());
    safearrayRef arOut{&ret};
    for (ULONG i = 0; i < Bounds[0].cElements; ++i)
    {
        for (ULONG j = 0; j < Bounds[1].cElements; ++j)
        {
            for (ULONG k = 0; k < Bounds[2].cElements; ++k)
            {
                auto& elem = arIn(i, j, k);
                std::swap(arOut(i, j, k), *func.eval(&elem, &elem));
            }
        }
    }
    return      ret;
}

//**************************************************************************

//配列matrix1とmatrix2の各要素に2変数のCallback（vbCallbackFunc_t型のVBA関数）を適用する
VARIANT  __stdcall
zipWith(VARIANT const* bfun, VARIANT* matrix1, VARIANT* matrix2) noexcept
{
    functionExpr func{bfun};
    //----------------------------
    if (!matrix1 || !matrix2 || !func.isValid())      return iVariant();
    if (0 == (VT_ARRAY & matrix1->vt) &&  0 == (VT_ARRAY & matrix2->vt))
    {
        auto ret = iVariant();
        std::swap(ret, *func.eval(matrix1, matrix2));
        return ret;
    }
    safearrayRef arIn1{matrix1};
    safearrayRef arIn2{matrix2};
    if (0 == arIn1.getDim() || 0 == arIn2.getDim() || arIn1.getDim() != arIn2.getDim())
        return iVariant();
    //----------------------------
    std::array<SAFEARRAYBOUND, 3>   minBounds{
        {
            { static_cast<ULONG>(minV(arIn1.getSize(1), arIn2.getSize(1))), 0 },
            { static_cast<ULONG>(minV(arIn1.getSize(2), arIn2.getSize(2))), 0 },
            { static_cast<ULONG>(minV(arIn1.getSize(3), arIn2.getSize(3))), 0 }
        }
    };
    // SAFEARRAY作成
    auto ret = iVariant(VT_ARRAY | VT_VARIANT);
    ret.parray = ::SafeArrayCreate(VT_VARIANT, static_cast<UINT>(arIn1.getDim()), minBounds.data());
    safearrayRef arOut{&ret};
    for (ULONG i = 0; i < minBounds[0].cElements; ++i)
    {
        for (ULONG j = 0; j < minBounds[1].cElements; ++j)
        {
            for (ULONG k = 0; k < minBounds[2].cElements; ++k)
            {
                auto& elem1 = arIn1(i, j, k);
                auto& elem2 = arIn2(i, j, k);
                std::swap(arOut(i, j, k), *func.eval(&elem1, &elem2));
            }
        }
    }
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
VARIANT  __stdcall
foldl(VARIANT const* bfun, VARIANT* init, VARIANT* matrix, __int32 axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !init || !func.isValid())  return ret;
    if (0 == (VT_ARRAY & matrix->vt))
    {
        std::swap(ret, *func.eval(init, matrix));
        return ret;
    }
    fold_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
VARIANT  __stdcall
foldr(VARIANT* const bfun, VARIANT* init, VARIANT* matrix, __int32 axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !init || !func.isValid())      return ret;
    if (0 == (VT_ARRAY & matrix->vt))
    {
        std::swap(ret, *func.eval(matrix, init));
        return ret;
    }
    fold_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldl1(VARIANT* const bfun, VARIANT* matrix, __int32 axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !func.isValid())                     return ret;
    if (0 == (VT_ARRAY & matrix->vt))                   return *matrix;
    fold_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldr1(VARIANT* const bfun, VARIANT* matrix, __int32 axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !func.isValid())                     return ret;
    if (0 == (VT_ARRAY & matrix->vt))                   return *matrix;
    fold_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
VARIANT  __stdcall
scanl(VARIANT* const bfun, VARIANT* init, VARIANT* matrix, __int32 const axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !init || !func.isValid())      return ret;
    if (0 == (VT_ARRAY & matrix->vt))
    {
        std::swap(ret, *func.eval(init, matrix));
        return ret;
    }
    scan_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
VARIANT  __stdcall
scanr(VARIANT* const bfun, VARIANT* init, VARIANT* matrix, __int32 const axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !init || !func.isValid())      return ret;
    if (0 == (VT_ARRAY & matrix->vt))
    {
        std::swap(ret, *func.eval(matrix, init));
        return ret;
    }
    scan_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanl1(VARIANT* const bfun, VARIANT* matrix, __int32 const axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !func.isValid())                 return ret;
    if (0 == (VT_ARRAY & matrix->vt))               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanr1(VARIANT* const bfun, VARIANT* matrix, __int32 const axis) noexcept
{
    auto ret = iVariant();
    functionExpr func{bfun};
    if (!matrix || !func.isValid())                 return ret;
    if (0 == (VT_ARRAY & matrix->vt))               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************
//述語による1次元配列からの検索
__int32  __stdcall
find_imple(VARIANT* const bfun, VARIANT* matrix, __int32 const def) noexcept
{
    if (!bfun || !matrix)                         return def;
    safearrayRef arIn{matrix};
    if (arIn.getDim() != 1)                       return def;
    functionExpr func{bfun};
    if (!func.isValid())                          return def;
    for (std::size_t i = 0; i <arIn.getSize(1); ++i)
    {
        auto& elem = arIn(i);
        auto ret = iVariant();
        ::VariantChangeType(&ret, func.eval(&elem, &elem), 0, VT_I4);
        if (ret.lVal != 0)
        {
            ::VariantClear(&ret);
            return static_cast<__int32>(i + arIn.getOriginalLBound(1));
        }
        ::VariantClear(&ret);
    }
    return      def;
}

//repeat_while と repeat_while_not と generate_while と generate_while_not
VARIANT __stdcall
repeat_imple(VARIANT const*       init,
            VARIANT const*        pred,
            VARIANT const*        trans,
            __int32 const         maxN,
            __int32 const         scan,
            __int32 const         stopCondition) noexcept
{
    auto ret = iVariant();
    if (!init || !pred || !trans)                 return ret;
    functionExpr funcP{pred};
    functionExpr funcT{trans};
    if (!funcP.isValid() || !funcT.isValid())     return ret;
    auto i = repeat_imple_0(init, funcP, funcT, maxN, ret, 0 != scan, 0 != stopCondition);
    return ret;
}

//1次元配列vecの離れた要素間で2項操作を適用する
VARIANT  __stdcall
self_zipWith(VARIANT const* bfun, VARIANT* vec, __int32 shift) noexcept
{
    functionExpr func{bfun};
    //----------------------------
    if (!vec || !func.isValid() || 0 == (VT_ARRAY & vec->vt))   return iVariant();
    safearrayRef arIn1{vec}, arIn2{ vec };      // 2つ必要
    if (1 != arIn1.getDim())                                    return iVariant();
    //----------------------------
    auto const len = static_cast<ULONG>(arIn1.getSize(1));
    if (0 == len)                                               return iVariant();
    if (shift < 0)
        shift = ((1 + (-shift)/len) * len + shift) % len;
    SAFEARRAYBOUND bound{ len, 0 };
    auto ret = iVariant(VT_ARRAY | VT_VARIANT);
    ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &bound);
    safearrayRef arOut{&ret};
    for (ULONG i = 0; i < len; ++i)
    {
        auto& elem1 = arIn1(i, 0, 0);
        auto& elem2 = arIn2((i+shift) % len, 0, 0);
        std::swap(arOut(i), *func.eval(&elem1, &elem2));
    }
    return      ret;
}

//********************************************************************

namespace {

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  functionExpr&   bfun,
                        VARIANT*        init,
                        VARIANT*        matrix,
                        __int32 const   axis,
                        VARIANT&        ret,
                        bool const      left)  noexcept //left==true, right == false
    {
        safearrayRef arIn{matrix};
        auto const dim = static_cast<__int32>(arIn.getDim());
        if (0 == dim)                     return;
        if (axis < 1 || dim < axis)       return;
        int i = 0, j = 0, k = 0;
        auto& index1 = (axis == 1) ? j : i;
        auto& index2 = (axis == 3) ? j : k;
        auto& index = (axis == 1) ? i
            : (axis == 2) ? j
            : k;
        auto const bound1 = static_cast<int>((axis == 1) ? arIn.getSize(2) : arIn.getSize(1));
        auto const bound2 = static_cast<int>((axis == 3) ? arIn.getSize(2) : arIn.getSize(3));
        auto const bound = static_cast<int>((axis == 1) ? arIn.getSize(1)
            : (axis == 2) ? arIn.getSize(2)
            : arIn.getSize(3));
        // SAFEARRAY作成
        std::array<SAFEARRAYBOUND, 2> resultBounds{ { { static_cast<ULONG>(bound1), 0 },
        { static_cast<ULONG>(bound2), 0 } } };
        if (1 != dim)
        {
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = ::SafeArrayCreate(VT_VARIANT, dim-1, resultBounds.data());
        }
        safearrayRef arOut{&ret};
        for (index1 = 0; index1 < bound1; ++index1)
        {
            for (index2 = 0; index2 < bound2; ++index2)
            {
                auto result = iVariant();
                VARIANT* presult = nullptr;
                auto first_time = true;
                if (init)
                {
                    presult = init;
                    first_time = false;
                }
                auto initial_state = true;
                for (index = left ? 0 : bound - 1; left ? index < bound : 0 <= index; index += (left ? 1 : -1))
                {
                    if (first_time)
                    {
                        presult = &arIn(i, j, k);
                        if (presult->vt & VT_BYREF)
                        {
                            ::VariantCopyInd(&result, presult);
                            presult = &result;
                            initial_state = false;
                        }
                        first_time = false;
                    }
                    else
                    {
                        presult = left ? bfun.eval(presult, &arIn(i, j, k)) :
                            bfun.eval(&arIn(i, j, k), presult);
                        initial_state = false;
                    }
                }
                if (presult)
                {
                    if (1 == dim)
                    {
                        if (initial_state)
                            VariantCopy(&ret, presult);
                        else
                            std::swap(ret, *presult);
                    }
                    else
                    {
                        if (initial_state)
                            VariantCopy(&arOut(index1, index2), presult);
                        else
                            std::swap(arOut(index1, index2), *presult);
                    }
                }
                ::VariantClear(&result);
            }
        }
    }

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun,
                        VARIANT*        init,
                        VARIANT*        matrix,
                        __int32 const   axis,
                        VARIANT&        ret,
                        bool const      left) noexcept //left==true, right == false
    {
        safearrayRef arIn{matrix};
        auto const dim = static_cast<__int32>(arIn.getDim());
        if (0 == dim)                         return;
        if (axis < 1 || dim < axis)           return;
        int i = 0, j = 0, k = 0;
        int& index1 = (axis == 1) ? j : i;
        int& index2 = (axis == 3) ? j : k;
        int& index = (axis == 1) ? i : (axis == 2) ? j : k;
        auto const bound1 = static_cast<int>((axis == 1) ? arIn.getSize(2) : arIn.getSize(1));
        auto const bound2 = static_cast<int>((axis == 3) ? arIn.getSize(2) : arIn.getSize(3));
        auto const bound = static_cast<int>((axis == 1) ? arIn.getSize(1) : (axis == 2) ? arIn.getSize(2) : arIn.getSize(3));
        // SAFEARRAY作成
        {
            std::array<SAFEARRAYBOUND, 3> resultBounds{
                {
                    { static_cast<ULONG>(arIn.getSize(1)), 0 },
                    { static_cast<ULONG>(arIn.getSize(2)), 0 },
                    { static_cast<ULONG>(arIn.getSize(3)), 0 }
                }
            };
            if (init)     resultBounds[axis-1].cElements += 1;
            auto retArray = ::SafeArrayCreate(VT_VARIANT, dim, resultBounds.data());
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = retArray;
        }
        safearrayRef arOut{&ret};
        auto adj = [=](std::size_t x) { return (init && left && x == axis) ? 1 : 0; };
        for (index1 = 0; index1 < bound1; ++index1)
        {
            for (index2 = 0; index2 < bound2; ++index2)
            {
                auto result = iVariant();
                VARIANT* presult = nullptr;
                auto first_time = true;
                if (init)
                {
                    presult = init;
                    index = left ? 0 : bound;
                    ::VariantCopy(&arOut(i, j, k), presult);
                    first_time = false;
                }
                for (index = left ? 0 : bound-1;
                    left ? index < bound : 0 <= index;
                    index += (left ? 1 : -1)
                    )
                {
                    if (first_time)
                    {
                        first_time = false;
                        presult = &arIn(i, j, k);
                        if (presult->vt & VT_BYREF)
                        {
                            ::VariantCopyInd(&result, presult);
                            presult = &result;
                        }
                    }
                    else
                    {
                        first_time = false;
                        presult = left ? bfun.eval(presult, &arIn(i, j, k)) :
                            bfun.eval(&arIn(i, j, k), presult);
                    }
                    ::VariantCopy(&arOut(i+adj(1), j+adj(2), k+adj(3)), presult);
                }
            }
        }
    }

    //repeat_while と repeat_while_not と generate_while と generate_while_not の共通処理
    __int32 repeat_imple_0( VARIANT const*  init,
                            functionExpr&   pred,
                            functionExpr&   trans,
                            __int32 const   maxN,
                            VARIANT&        ret,
                            bool const      scan,
                            bool const      stopCondition) noexcept
    {
        auto zero  = iVariant();
        auto check = iVariant();
        ::VariantCopy(&ret, init);
        auto pret = &ret;
        try
        {
            std::list<VARIANT> vlist;
            if (scan)
            {
                vlist.push_back(zero);
                ::VariantCopy(&vlist.back(), pret);
            }
            __int32 count = 0;
            while (maxN < 0 || count < maxN)
            {
                ::VariantChangeType(&check, pred.eval(pret, pret), 0, VT_I4);
                if ((check.lVal != 0) == stopCondition)
                {
                    ::VariantClear(&check);
                    break;
                }
                pret = trans.eval(pret, pret);
                if (scan)
                {
                    vlist.push_back(zero);
                    ::VariantCopy(&vlist.back(), pret);
                }
                ::VariantClear(&check);
                ++count;
            }
            if (scan && 0 < vlist.size())
            {
                ::VariantClear(&ret);
                SAFEARRAYBOUND bound = { static_cast<ULONG>(vlist.size()), 0 };
                ret.vt = VT_ARRAY | VT_VARIANT;
                ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &bound);
                safearrayRef arOut{ &ret };
                LONG index = 0;
                for (auto it = vlist.begin(); it != vlist.end(); ++it, ++index)
                    std::swap(arOut(index), *it);
            }
            else
            {
                ::VariantCopy(&ret, pret);
            }
            return count;
        }
        catch (const std::exception&)
        {
            return 0;
        }
    }

}   // namespace
