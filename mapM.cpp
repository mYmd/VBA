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
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ); //left==true, right == false

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
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
    if ( !bfun )    return ret;
    VBCallbackFunc pf(bfun);
    if ( pf )
    {
        functionExpr func(pf);
        std::swap(ret, *func.eval(param1, param2));
    }
    return ret;
}

//sourceのVARIANT変数をtargetのVARIANTへmoveする
VARIANT __stdcall
moveVariant(VARIANT* source)
{
    VARIANT target;
    ::VariantInit(&target);
    std::swap(target, *source);
    return target;
}

////************************************************************************************

//配列matrixの各要素にVBA関数を適用する
VARIANT  __stdcall
mapF_imple(VARIANT* bfun, VARIANT* matrix)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc pf(bfun);
    if ( !matrix || !pf )               return ret;
    functionExpr func(pf);
    if (  0 == (VT_ARRAY & matrix->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix, matrix));
        return ret;
    }
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
    if ( !pArray )                                  return ret;
    UINT dim  = ::SafeArrayGetDim(pArray);
    if ( 0 == dim || 3 < dim )                      return ret;
    SAFEARRAYBOUND bounds[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
    safeArrayBounds(pArray, dim, bounds);
    SAFEARRAYBOUND Bounds[3] = {
                                { bounds[0].cElements, bounds[0].lLbound },
                                { bounds[1].cElements, bounds[1].lLbound },
                                { bounds[2].cElements, bounds[2].lLbound }    };
    // SAFEARRAY作成
    SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, dim, Bounds);
    for ( ULONG i = 0; i < Bounds[0].cElements; ++i )
    {
        for ( ULONG j = 0; j < Bounds[1].cElements; ++j )
        {
            for ( ULONG k = 0; k < Bounds[2].cElements; ++k )
            {
                LONG index[3] = {   static_cast<LONG>(i)+bounds[0].lLbound,
                                    static_cast<LONG>(j)+bounds[1].lLbound,
                                    static_cast<LONG>(k)+bounds[2].lLbound     };
                VARIANT elem;
                ::VariantInit(&elem);
                ::SafeArrayGetElement(pArray, index, &elem);
                ::SafeArrayPutElement(retArray, index, func.eval(&elem, &elem));
                ::VariantClear(&elem);
            }
        }
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = retArray;
    return      ret;
}

//**************************************************************************

//配列matrix1とmatrix2の各要素に2変数のCallback（VBCallbackFunc型のVBA関数）を適用する
VARIANT  __stdcall
zipWith(VARIANT* bfun, VARIANT* matrix1, VARIANT* matrix2)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc pf(bfun);
    //----------------------------
    if ( !matrix1 || !matrix2 || !pf )        return ret;
    functionExpr func(pf);
    if (  0 == (VT_ARRAY & matrix1->vt ) &&  0 == (VT_ARRAY & matrix2->vt ) )
    {
        ::VariantCopy(&ret, func.eval(matrix1, matrix2));
        return ret;
    }
    if (  0 == (VT_ARRAY & matrix1->vt ) ||  0 == (VT_ARRAY & matrix2->vt ) )
                                                                return ret;
    //----------------------------
    SAFEARRAY* pArray1 = ( 0 == (VT_BYREF & matrix1->vt) )?  (matrix1->parray): (*matrix1->pparray);
    SAFEARRAY* pArray2 = ( 0 == (VT_BYREF & matrix2->vt) )?  (matrix2->parray): (*matrix2->pparray);
    if ( !pArray1 || !pArray2 )                                 return ret;
    UINT dim  = ::SafeArrayGetDim(pArray1);
    UINT dim2 = ::SafeArrayGetDim(pArray2);
    if ( 0 == dim || 3 < dim || dim != dim2 )                   return ret;
    SAFEARRAYBOUND bounds1[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
    SAFEARRAYBOUND bounds2[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
    safeArrayBounds(pArray1, dim, bounds1);
    safeArrayBounds(pArray2, dim, bounds2);
    SAFEARRAYBOUND minBounds[3] = {
                        { minV(bounds1[0].cElements, bounds2[0].cElements), bounds1[0].lLbound },
                        { minV(bounds1[1].cElements, bounds2[1].cElements), bounds1[1].lLbound },
                        { minV(bounds1[2].cElements, bounds2[2].cElements), bounds1[2].lLbound }    };
    // SAFEARRAY作成
    SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, dim, minBounds);
    for ( ULONG i = 0; i < minBounds[0].cElements; ++i )
    {
        for ( ULONG j = 0; j < minBounds[1].cElements; ++j )
        {
            for ( ULONG k = 0; k < minBounds[2].cElements; ++k )
            {
                LONG index1[3] = {  static_cast<LONG>(i)+bounds1[0].lLbound,
                                    static_cast<LONG>(j)+bounds1[1].lLbound,
                                    static_cast<LONG>(k)+bounds1[2].lLbound     };
                LONG index2[3] = {  static_cast<LONG>(i)+bounds2[0].lLbound,
                                    static_cast<LONG>(j)+bounds2[1].lLbound,
                                    static_cast<LONG>(k)+bounds2[2].lLbound };
                VARIANT elem1, elem2;
                ::VariantInit(&elem1);
                ::VariantInit(&elem2);
                ::SafeArrayGetElement(pArray1, index1, &elem1);
                ::SafeArrayGetElement(pArray2, index2, &elem2);
                ::SafeArrayPutElement(retArray, index1, func.eval(&elem1, &elem2));
                ::VariantClear(&elem1);
                ::VariantClear(&elem2);
            }
        }
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = retArray;
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（初期値指定あり）
VARIANT  __stdcall
foldl(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc pf(bfun);
    if ( !matrix || !init || !pf )        return ret;
    functionExpr func(pf);
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
    VBCallbackFunc pf(bfun);
    if ( !matrix || !init || !pf )                          return ret;
    functionExpr func(pf);
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
    VBCallbackFunc pf(bfun);
    if ( !matrix || !pf )                                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    functionExpr func(pf);
    fold_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc pf(bfun);
    if ( !matrix || !pf )                                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    functionExpr func(pf);
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
    VBCallbackFunc pf(bfun);
    if ( !matrix || !init || !pf )                          return ret;
    functionExpr func(pf);
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
    VBCallbackFunc pf(bfun);
    if ( !matrix || !init || !pf )                          return ret;
    functionExpr func(pf);
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
    VBCallbackFunc pf(bfun);
    if ( !matrix || !pf )                                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    functionExpr func(pf);
    scan_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc pf(bfun);
    if ( !matrix || !pf )                                   return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    functionExpr func(pf);
    scan_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************
//述語による1次元配列からの検索
__int32  __stdcall
find_imple(VARIANT* bfun, VARIANT* matrix, __int32 def)
{
    if ( !bfun || !matrix )                         return def;
    if (  0 == (VT_ARRAY & matrix->vt ) )           return def;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
    if ( !pArray )                                  return def;
    if ( 1 != ::SafeArrayGetDim(pArray) )           return def;
    VBCallbackFunc pf(bfun);
    if ( !pf )                                      return def;
    functionExpr func(pf);
    SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
    safeArrayBounds(pArray, 1, &bounds);
    SAFEARRAYBOUND Bounds = { bounds.cElements, bounds.lLbound };
    for ( ULONG i = 0; i < Bounds.cElements; ++i )
    {
        LONG index = static_cast<LONG>(i) + bounds.lLbound;
        VARIANT elem, ret;
        ::VariantInit(&elem);
        ::VariantInit(&ret);
        ::SafeArrayGetElement(pArray, &index, &elem);
        ::VariantChangeType(&ret, func.eval(&elem, &elem), 0, VT_I4);
        if ( ret.lVal != 0 )
        {
            ::VariantClear(&elem);
            ::VariantClear(&ret);
            return static_cast<__int32>(index);
        }
        ::VariantClear(&elem);
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
    VBCallbackFunc ppred(pred);
    VBCallbackFunc ptrans(trans);
    if ( !ppred || !ptrans )                        return ret;
    functionExpr funcP(ppred);
    functionExpr funcT(ptrans);
    __int32 i = repeat_imple_0(init, funcP, funcT, maxN, ret, 0 != scan, stopCondition);
    return ret;
}

//********************************************************************

namespace   {

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        if ( !pArray )                                                  return;
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( 0 == dim || 3 < dim )                                      return;
        if ( axis < 1 || static_cast<__int32>(dim) < axis )             return;
        axis -= 1;
        SAFEARRAYBOUND bounds[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
        safeArrayBounds(pArray, dim, bounds);
        // SAFEARRAY作成
        const LONG i = (axis == 0)? 1 : 0;
        const LONG j = (axis == 2)? 1 : 2;
        SAFEARRAYBOUND resultBounds[2] = { { bounds[i].cElements, 0 }, { bounds[j].cElements, 0} };
        SAFEARRAY* retArray = (dim == 1)? 0 :   ::SafeArrayCreate(VT_VARIANT, dim-1, resultBounds);
        LONG index[3];
        for ( index[i] = 0; index[i] < static_cast<LONG>(bounds[i].cElements); ++index[i] )
        {
            for ( index[j] = 0; index[j] < static_cast<LONG>(bounds[j].cElements); ++index[j] )
            {
                VARIANT result;
                VARIANT* presult = &result;
                ::VariantInit(presult);
                bool first_time = true;
                if ( init )
                {
                    ::VariantCopy(presult, init);
                    first_time = false;
                }
                for (   index[axis] = left? 0: static_cast<LONG>(bounds[axis].cElements) - 1;
                        left? index[axis] < static_cast<LONG>(bounds[axis].cElements): 0 <= index[axis];
                        index[axis] += (left? 1: -1)
                    )
                {
                    LONG sourceIndex[3] = { index[0] + bounds[0].lLbound,
                                            index[1] + bounds[1].lLbound,
                                            index[2] + bounds[2].lLbound    };
                    if ( first_time )
                    {
                        ::SafeArrayGetElement(pArray, sourceIndex, presult);
                        first_time = false;
                    }
                    else
                    {
                        VARIANT elem;
                        ::VariantInit(&elem);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     presult = bfun.eval(presult, &elem);
                        else            presult = bfun.eval(&elem, presult);
                        ::VariantClear(&elem);
                    }
                }
                if ( 1 == dim )
                {
                    ::VariantCopy(&ret, presult);
                }
                else
                {
                    LONG targetIndex[2] = { index[i], index[j] };
                    ::SafeArrayPutElement(retArray, targetIndex, presult);
                }
                ::VariantClear(&result);
            }
        }
        if ( 1 != dim )
        {
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = retArray;
        }
    }

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  functionExpr&   bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        if ( !pArray )                                                  return;
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( 0 == dim || 3 < dim )                                      return;
        if ( axis < 1 || static_cast<__int32>(dim) < axis )             return;
        axis -= 1;
        SAFEARRAYBOUND bounds[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
        safeArrayBounds(pArray, dim, bounds);
        // SAFEARRAY作成
        const LONG i = (axis == 0)? 1 : 0;
        const LONG j = (axis == 2)? 1 : 2;
        SAFEARRAYBOUND resultBounds[3] = {  { bounds[0].cElements, 0},
                                            { bounds[1].cElements, 0},
                                            { bounds[2].cElements, 0}   };
        if ( init )     resultBounds[axis].cElements += 1;
        SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, dim, resultBounds);
        LONG index[3];
        for ( index[i] = 0; index[i] < static_cast<LONG>(bounds[i].cElements); ++index[i] )
        {
            for ( index[j] = 0; index[j] < static_cast<LONG>(bounds[j].cElements); ++index[j] )
            {
                VARIANT result;
                VARIANT* presult = &result;
                ::VariantInit(presult);
                bool first_time = true;
                if ( init )
                {
                    ::VariantCopy(presult, init);
                    LONG targetIndex[3] = { index[0], index[1], index[2] };
                    targetIndex[axis] = left? 0: static_cast<LONG>(bounds[axis].cElements);
                    ::SafeArrayPutElement(retArray, targetIndex, presult);
                    first_time = false;
                }
                for (   index[axis] = left? 0: static_cast<LONG>(bounds[axis].cElements) - 1;
                        left? index[axis] < static_cast<LONG>(bounds[axis].cElements): 0 <= index[axis];
                        index[axis] += (left? 1: -1)
                    )
                {
                    LONG sourceIndex[3] = { index[0] + bounds[0].lLbound,
                                            index[1] + bounds[1].lLbound,
                                            index[2] + bounds[2].lLbound    };
                    if ( first_time )
                    {
                        ::SafeArrayGetElement(pArray, sourceIndex, presult);
                        first_time = false;
                    }
                    else
                    {
                        first_time = false;
                        VARIANT elem;
                        ::VariantInit(&elem);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     presult = bfun.eval(presult, &elem);
                        else            presult = bfun.eval(&elem, presult);
                        ::VariantClear(&elem);
                    }
                    LONG targetIndex[3] = { index[0], index[1], index[2] };
                    if ( init && left )     targetIndex[axis] += 1;
                    ::SafeArrayPutElement(retArray, targetIndex, presult);
                }
                ::VariantClear(&result);
            }
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = retArray;
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
        VARIANT* pret = &ret;
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
            SAFEARRAYBOUND bound = { vlist.size(), 0 };
            SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, 1, &bound);
            LONG index = 0;
            for ( auto it = vlist.begin(); it != vlist.end(); ++it, ++index )
            {
                ::SafeArrayPutElement(retArray, &index, &*it);
                ::VariantClear(&*it);
            }
            ::VariantClear(&ret);
            ret.vt = VT_ARRAY | VT_VARIANT;
            ret.parray = retArray;
        }
        else
        {
            ::VariantCopy(&ret, pret);
        }
        return count;
    }

}   // namespace
