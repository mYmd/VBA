//mapM.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "VBA_NestFunc.hpp"
#include <list>

namespace   {
    //基本型のみ想定したmin
    template <typename T>  T minV(T a, T b) { return (a < b)? a: b; }

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32 const   axis    ,
                        VARIANT&        ret     ,
                        bool            left    ); //left==true, right == false

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32 const   axis    ,
                        VARIANT&        ret     ,
                        bool            left    ); //left==true, right == false

    //repeat_while と repeat_while_not と generate_while と generate_while_not の共通処理
    __int32 repeat_imple_0( VARIANT*        init    ,
                            functionExpr&   pred    ,
                            functionExpr&   trans   ,
                            __int32         maxN    ,
                            VARIANT&        ret     ,
                            bool            scan    ,
                            __int32         stopCondition);
}   // namespace

//bindされていないVBA関数を2引数で呼び出す
VARIANT  __stdcall
unbind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( func.isValid() )
        std::swap(ret, *func.eval(param1, param2));
    return ret;
}

//VARIANT変数のswap
__int32 __stdcall
swapVariant(VARIANT* a, VARIANT* b)
{
	if ( a && b )
	{
		std::swap(*a, *b);
        return ( 0 == (VT_BYREF & a->vt) && 0 == (VT_BYREF & b->vt) ) ? 1: 0;
	}    
	return 0;
}

////************************************************************************************

//配列matrixの各要素にVBA関数を適用する
VARIANT  __stdcall
mapF_imple(VARIANT* bfun, VARIANT* matrix)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !func.isValid() )       return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix, matrix));
        return ret;
    }
    safearrayRef arIn(matrix);
    auto pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
    if ( arIn.getDim() == 0 )                      return ret;
    std::array<SAFEARRAYBOUND, 3>   Bounds  {
                                {
                                    { static_cast<ULONG>(arIn.getSize(1)), 0 },
                                    { static_cast<ULONG>(arIn.getSize(2)), 0 },
                                    { static_cast<ULONG>(arIn.getSize(3)), 0 }
                                }
                            };
    // SAFEARRAY作成
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = ::SafeArrayCreate(VT_VARIANT, static_cast<UINT>(arIn.getDim()), Bounds.data());
    safearrayRef arOut(&ret);
    for ( ULONG i = 0; i < Bounds[0].cElements; ++i )
    {
        for ( ULONG j = 0; j < Bounds[1].cElements; ++j )
        {
            for ( ULONG k = 0; k < Bounds[2].cElements; ++k )
            {
                auto& elem = arIn(i, j, k);
                ::VariantCopy(&arOut(i, j, k), func.eval(&elem, &elem));
            }
        }
    }
    return      ret;
}

//**************************************************************************

//配列matrix1とmatrix2の各要素に2変数のCallback（vbCallbackFunc_t型のVBA関数）を適用する
VARIANT  __stdcall
zipWith(VARIANT* bfun, VARIANT* matrix1, VARIANT* matrix2)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    //----------------------------
    if ( !matrix1 || !matrix2 || !func.isValid() )      return ret;
    if (  0 == (VT_ARRAY & matrix1->vt ) &&  0 == (VT_ARRAY & matrix2->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix1, matrix2));
        return ret;
    }
    safearrayRef arIn1(matrix1);
    safearrayRef arIn2(matrix2);
    if ( 0 == arIn1.getDim() || 0 == arIn2.getDim() || arIn1.getDim() != arIn2.getDim() )
        return ret;
    //----------------------------
    std::array<SAFEARRAYBOUND, 3>   minBounds{
                    {
                        { static_cast<ULONG>(minV(arIn1.getSize(1), arIn2.getSize(1))), 0 },
                        { static_cast<ULONG>(minV(arIn1.getSize(2), arIn2.getSize(2))), 0 },
                        { static_cast<ULONG>(minV(arIn1.getSize(3), arIn2.getSize(3))), 0 }
                    }
                };
    // SAFEARRAY作成
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = ::SafeArrayCreate(VT_VARIANT, static_cast<UINT>(arIn1.getDim()), minBounds.data());
    safearrayRef arOut(&ret);
    for ( ULONG i = 0; i < minBounds[0].cElements; ++i )
    {
        for ( ULONG j = 0; j < minBounds[1].cElements; ++j )
        {
            for ( ULONG k = 0; k < minBounds[2].cElements; ++k )
            {
                auto& elem1 = arIn1(i, j, k);
                auto& elem2 = arIn2(i, j, k);
                ::VariantCopy(&arOut(i, j, k), func.eval(&elem1, &elem2));
            }
        }
    }
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
VARIANT  __stdcall
foldl(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !init || !func.isValid() )  return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(init, matrix));
        return ret;
    }
    fold_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
VARIANT  __stdcall
foldr(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !init || !func.isValid() )      return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix, init));
        return ret;
    }
    fold_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldl1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !func.isValid() )                       return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    fold_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !func.isValid() )                       return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    fold_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
VARIANT  __stdcall
scanl(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !init || !func.isValid() )      return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(init, matrix));
        return ret;
    }
    scan_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
VARIANT  __stdcall
scanr(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !init || !func.isValid() )      return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix, init));
        return ret;
    }
    scan_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanl1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !func.isValid() )                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    functionExpr func(bfun);
    if ( !matrix || !func.isValid() )                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************
//述語による1次元配列からの検索
__int32  __stdcall
find_imple(VARIANT* bfun, VARIANT* matrix, __int32 def)
{
    if ( !bfun || !matrix )                         return def;
    safearrayRef arIn(matrix);
    if ( arIn.getDim() != 1 )                       return def;
    functionExpr func(bfun);
    if ( !func.isValid() )                          return def;
    for ( std::size_t i = 0; i <arIn.getSize(1); ++i )
    {
        auto& elem = arIn(i);
        VARIANT ret;
        ::VariantInit(&ret);
        ::VariantChangeType(&ret, func.eval(&elem, &elem), 0, VT_I4);
        if ( ret.lVal != 0 )
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
repeat_imple(   VARIANT*        init    ,
                VARIANT*        pred    ,
                VARIANT*        trans   ,
                __int32         maxN    ,
                __int32         scan    ,
                __int32         stopCondition)
{
    VARIANT ret;
    ::VariantInit(&ret);
    if ( !init || !pred || !trans )                 return ret;
    functionExpr funcP(pred);
    functionExpr funcT(trans);
    if ( !funcP.isValid() || !funcT.isValid() )     return ret;
    auto i = repeat_imple_0(init, funcP, funcT, maxN, ret, 0 != scan, stopCondition);
    return ret;
}

//********************************************************************

namespace   {

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32 const   axis    ,
                        VARIANT&        ret     ,
                        bool const      left    ) //left==true, right == false
    {
        safearrayRef arIn(matrix);
        auto const dim = static_cast<__int32>(arIn.getDim());
        if ( 0 == dim )                     return;
        if ( axis < 1 || dim < axis )       return;
        int i = 0, j = 0, k = 0;
        auto& index1 = (axis == 1) ? j : i;
        auto& index2 = (axis == 3) ? j : k;
        auto& index =    (axis == 1) ? i
                     :   (axis == 2) ? j
                     :   k;
        auto const bound1 = static_cast<int>((axis == 1) ? arIn.getSize(2) : arIn.getSize(1));
        auto const bound2 = static_cast<int>((axis == 3) ? arIn.getSize(2) : arIn.getSize(3));
        auto const bound  = static_cast<int>((axis == 1) ?  arIn.getSize(1)
                                            :(axis == 2 )?  arIn.getSize(2)
                                            :               arIn.getSize(3) );
        // SAFEARRAY作成
        std::array<SAFEARRAYBOUND, 2> resultBounds{ { {bound1, 0},{bound2, 0} } };
        if ( 1 != dim )
        {
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = ::SafeArrayCreate(VT_VARIANT, dim-1, resultBounds.data());
        }
        safearrayRef arOut(&ret);
        for ( index1 = 0; index1 < bound1; ++index1 )
        {
            for ( index2 = 0; index2 < bound2; ++index2 )
            {
                VARIANT result;
                auto presult = &result;
                ::VariantInit(presult);
                auto first_time = true;
                if ( init )
                {
                    ::VariantCopy(presult, init);
                    first_time = false;
                }
                for (index = left? 0: bound - 1; left? index < bound: 0 <= index; index += (left? 1: -1))
                {
                    if ( first_time )
                    {
                        ::VariantCopyInd(presult, &arIn(i, j, k));
                        first_time = false;
                    }
                    else
                    {
                        presult = left ? bfun.eval(presult, &arIn(i, j, k)):
                                         bfun.eval(&arIn(i, j, k), presult);
                    }
                }
                if ( 1 == dim ) VariantCopy(&ret, presult);
                else            VariantCopy(&arOut(index1, index2), presult);
                ::VariantClear(presult);
            }
        }
    }

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32 const   axis    ,
                        VARIANT&        ret     ,
                        bool const      left    ) //left==true, right == false
    {
        safearrayRef arIn(matrix);
        auto const dim = static_cast<__int32>(arIn.getDim());
        if ( 0 == dim )                         return;
        if ( axis < 1 || dim < axis )           return;
        int i = 0, j = 0, k = 0;
        int& index1 = (axis == 1) ? j : i;
        int& index2 = (axis == 3) ? j : k;
        int& index = (axis == 1) ? i : (axis == 2)? j: k;
        auto const bound1 = static_cast<int>((axis == 1) ? arIn.getSize(2) : arIn.getSize(1));
        auto const bound2 = static_cast<int>((axis == 3) ? arIn.getSize(2) : arIn.getSize(3));
        auto const bound = static_cast<int>((axis == 1) ? arIn.getSize(1): (axis == 2 )? arIn.getSize(2): arIn.getSize(3));
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
        safearrayRef arOut(&ret);
        auto adj = [=](std::size_t x){ return (init && left && x == axis) ? 1 : 0; };
        for ( index1 = 0; index1 < bound1; ++index1 )
        {
            for ( index2 = 0; index2 < bound2; ++index2 )
            {
                VARIANT result;
                auto presult = &result;
                ::VariantInit(presult);
                auto first_time = true;
                if ( init )
                {
                    ::VariantCopy(presult, init);
                    index = left ? 0 : bound;
                    ::VariantCopy(&arOut(i, j, k), presult);
                    first_time = false;
                }
                for (   index = left? 0: bound-1;
                        left? index < bound: 0 <= index;
                        index += (left? 1: -1)
                    )
                {
                    if ( first_time )
                    {
                        first_time = false;
                        ::VariantCopyInd(presult, &arIn(i, j, k));
                    }
                    else
                    {
                        first_time = false;
                        VARIANT elem;
                        ::VariantInit(&elem);
                        ::VariantCopyInd(&elem, &arIn(i, j, k));
                        presult = left ? bfun.eval(presult, &elem):
                                         bfun.eval(&elem, presult);
                        ::VariantClear(&elem);
                    }
                    ::VariantCopy(&arOut(i+adj(1), j+adj(2), k+adj(3)), presult);
                }
                ::VariantClear(&result);
            }
        }
    }

        //
        inline bool stopCheck(__int32 a, __int32 b)
        {   return ( a == 0 && b == 0 ) || ( a!= 0 && b != 0 ); }
    
    //repeat_while と repeat_while_not と generate_while と generate_while_not の共通処理
    __int32 repeat_imple_0( VARIANT*        init    ,
                            functionExpr&   pred    ,
                            functionExpr&   trans   ,
                            __int32         maxN    ,
                            VARIANT&        ret     ,
                            bool            scan    ,
                            __int32         stopCondition)
    {
        VARIANT zero, check;
        ::VariantClear(&ret);
        ::VariantInit(&zero);
        ::VariantInit(&check);
        ::VariantCopy(&ret, init);
        auto pret = &ret;
        std::list<VARIANT> vlist;
        if ( scan )
        {
            vlist.push_back(zero);
            ::VariantCopy(&vlist.back(), pret);
        }
        __int32 count = 0;
        while ( maxN < 0 || count < maxN )
        {
            ::VariantChangeType(&check, pred.eval(pret, pret), 0, VT_I4);
            if ( stopCheck(check.lVal, stopCondition) )
            {
                ::VariantClear(&check);
                break;
            }
            pret = trans.eval(pret, pret);
            if ( scan )
            {
                vlist.push_back(zero);
                ::VariantCopy(&vlist.back(), pret);
            }
            ::VariantClear(&check);
            ++count;
        }
        if ( scan && 0 < vlist.size() )
        {
            ::VariantClear(&ret);
            SAFEARRAYBOUND bound = { static_cast<ULONG>(vlist.size()), 0 };
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &bound);
            safearrayRef arOut(&ret);
            LONG index = 0;
            for ( auto it = vlist.begin(); it != vlist.end(); ++it, ++index )
                std::swap(arOut(index), *it);
        }
        else
        {
            ::VariantCopy(&ret, pret);
        }
        return count;
    }
	
}   // namespace
