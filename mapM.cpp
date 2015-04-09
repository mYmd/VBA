//mapM.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include "OAIdl.h"      //wtypes.h

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

//プレースホルダ・オブジェクトの生成
VARIANT __stdcall placeholder()
{
    VARIANT ret;
    VariantClear(&ret);
    ret.vt = VT_ERROR;
    ret.scode = 0;
    return ret;
}

//プレースホルダ・オブジェクト判定
__int32 __stdcall is_placeholder(const VARIANT* pv)
{
    return ( pv && (pv->vt == VT_ERROR) && pv->scode == 0 ) ? 1 : 0;
}

//使用する唯一のVBAコールバック関数型  VBCallbackFunc  の宣言
//VBAにおけるシグネチャは
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// もしくは
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
typedef VARIANT (__stdcall *VBCallbackFunc)(VARIANT*, VARIANT*);


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

namespace   {
    //基本型のみ想定したmin
    template <typename T>  T minV(T a, T b) { return (a < b)? a: b; }

    VBCallbackFunc is_bindFun(const VARIANT* bfun, VARIANT& elem0, VARIANT& elem1, VARIANT& elem2);

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  VARIANT*        bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ); //left==true, right == false

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple(  VARIANT*        bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ); //left==true, right == false

}   // namespace

////************************************************************************************

//2引数にVBA関数を適用する
VARIANT  __stdcall
bind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    if ( !bfun )    return ret;
    VARIANT elem0, elem1, elem2;
    VBCallbackFunc func = is_bindFun(bfun, elem0, elem1, elem2);
    if ( func )
    {
        VARIANT tmp1 = bind_invoke(&elem1, param1, param1);
        VARIANT tmp2 = bind_invoke(&elem2, param2, param2);
        return (*func)(&tmp1, &tmp2);
    }
    else if ( is_placeholder(bfun) )
    {
        return *param1;
    }
    else
    {
        return *bfun;
    }
}


//配列matrixの各要素にVBA関数を適用する
VARIANT  __stdcall
mapF_imple(VARIANT* bfun, VARIANT* matrix)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    //----------------------------
    if ( !matrix || !func )                         return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )           return bind_invoke(bfun, matrix, matrix);
    
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
    UINT dim  = ::SafeArrayGetDim(pArray);
    if ( 0 == dim || 3 < dim  )                     return ret;
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
                VARIANT result = bind_invoke(bfun, &elem, &elem);
                ::SafeArrayPutElement(retArray, index, &result);
                ::VariantClear(&elem);
                ::VariantClear(&result);
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
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    //----------------------------
    if ( !matrix1 || !matrix2 || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix1->vt ) &&  0 == (VT_ARRAY & matrix2->vt ) )
                                                                return bind_invoke(bfun, matrix1, matrix2);
    if (  0 == (VT_ARRAY & matrix1->vt ) ||  0 == (VT_ARRAY & matrix2->vt ) )
                                                                return ret;
    //----------------------------
    SAFEARRAY* pArray1 = ( 0 == (VT_BYREF & matrix1->vt) )?  (matrix1->parray): (*matrix1->pparray);
    SAFEARRAY* pArray2 = ( 0 == (VT_BYREF & matrix2->vt) )?  (matrix2->parray): (*matrix2->pparray);
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
                VARIANT result = bind_invoke(bfun, &elem1, &elem2);
                ::SafeArrayPutElement(retArray, index1, &result);
                ::VariantClear(&elem1);
                ::VariantClear(&elem2);
                ::VariantClear(&result);
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
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !init || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return bind_invoke(bfun, init, matrix);
    fold_imple(bfun, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
VARIANT  __stdcall
foldr(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !init || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return bind_invoke(bfun, matrix, init);
    fold_imple(bfun, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldl1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !func )                                 return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    fold_imple(bfun, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !func )                                 return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    fold_imple(bfun, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
VARIANT  __stdcall
scanl(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !init || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return bind_invoke(bfun, init, matrix);
    scan_imple(bfun, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
VARIANT  __stdcall
scanr(VARIANT* bfun, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !init || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return bind_invoke(bfun, matrix, init);
    scan_imple(bfun, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanl1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !func )                                 return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    scan_imple(bfun, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanr1(VARIANT* bfun, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VARIANT tmp0, tmp1, tmp2;
    VBCallbackFunc func = is_bindFun(bfun, tmp0, tmp1, tmp2);
    if ( !matrix || !func )                                 return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                   return *matrix;
    scan_imple(bfun, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

namespace   {
    VBCallbackFunc is_bindFun(const VARIANT* bfun, VARIANT& elem0, VARIANT& elem1, VARIANT& elem2)
    {
        if ( 1 != Dimension(bfun) )         return 0;
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & bfun->vt) )?  (bfun->parray): (*bfun->pparray);
        if ( !pArray )                      return 0;
        SAFEARRAYBOUND bounds = {1,0};   //要素数、LBound
        safeArrayBounds(pArray, 1, &bounds);
        if ( bounds.cElements != 4 )        return 0;
        LONG index = bounds.lLbound + 3;
        VARIANT elem3;
        ::VariantInit(&elem3);
        ::SafeArrayGetElement(pArray, &index, &elem3);
        if ( !is_placeholder(&elem3) )      return 0;
        ::VariantInit(&elem0);
        index = bounds.lLbound + 0;
        ::SafeArrayGetElement(pArray, &index, &elem0);
        if ( elem0.vt != VT_I4 || elem0.lVal == 0 )     return 0;
        ::VariantInit(&elem1);
        index = bounds.lLbound + 1;
        ::SafeArrayGetElement(pArray, &index, &elem1);
        ::VariantInit(&elem2);
        index = bounds.lLbound + 2;
        ::SafeArrayGetElement(pArray, &index, &elem2);
        return reinterpret_cast<VBCallbackFunc>(elem0.lVal);
    }

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  VARIANT*        bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( !bfun || 0 == dim || 3 < dim )                             return;
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
                ::VariantInit(&result);
                bool first_time = true;
                if ( init )
                {
                    ::VariantCopy(&result, init);
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
                        ::SafeArrayGetElement(pArray, sourceIndex, &result);
                        first_time = false;
                    }
                    else
                    {
                        VARIANT elem, tmp;
                        ::VariantInit(&elem);
                        ::VariantInit(&tmp);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     tmp = bind_invoke(bfun, &result, &elem);
                        else            tmp = bind_invoke(bfun, &elem, &result);
                        ::VariantClear(&result);
                        result = tmp;               //result = std::move(tmp)と同じ
                        ::VariantClear(&elem);
                    }
                }
                if ( 1 == dim )
                {
                    ::VariantCopy(&ret, &result);
                }
                else
                {
                    LONG targetIndex[2] = { index[i], index[j] };
                    ::SafeArrayPutElement(retArray, targetIndex, &result);
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
    void   scan_imple(  VARIANT*        bfun    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( !bfun || 0 == dim || 3 < dim )                             return;
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
                ::VariantInit(&result);
                bool first_time = true;
                if ( init )
                {
                    ::VariantCopy(&result, init);
                    LONG targetIndex[3] = { index[0], index[1], index[2] };
                    targetIndex[axis] = left? 0: static_cast<LONG>(bounds[axis].cElements);
                    ::SafeArrayPutElement(retArray, targetIndex, &result);
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
                        ::SafeArrayGetElement(pArray, sourceIndex, &result);
                        first_time = false;
                    }
                    else
                    {
                        first_time = false;
                        VARIANT elem, tmp;
                        ::VariantInit(&elem);
                        ::VariantInit(&tmp);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     tmp = bind_invoke(bfun, &result, &elem);
                        else            tmp = bind_invoke(bfun, &elem, &result);
                        ::VariantClear(&result);
                        result = tmp;               //result = std::move(tmp)と同じ
                        ::VariantClear(&elem);
                    }
                    LONG targetIndex[3] = { index[0], index[1], index[2] };
                    if ( init && left )     targetIndex[axis] += 1;
                    ::SafeArrayPutElement(retArray, targetIndex, &result);
                }
                ::VariantClear(&result);
            }
        }
        ret.vt = VT_ARRAY | VT_VARIANT;
        ret.parray = retArray;
    }

}   // namespace
