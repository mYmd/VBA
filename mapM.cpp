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


namespace   {
    //基本型のみ想定したmin
    template <typename T>  T minV(T a, T b) { return (a < b)? a: b; }

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

    //mapLとmapRの共通処理
    VARIANT  mapLR(  __int32         pCallback   ,
                        VARIANT*     matrix      ,
                        VARIANT*     param       ,
                        bool            left        );  //left==true, right == false

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple( VBCallbackFunc   func    ,
                        VARIANT*     init    ,
                        VARIANT*     matrix  ,
                        __int32         axis    ,
                        VARIANT&     ret     ,
                        bool            left    ); //left==true, right == false

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple( VBCallbackFunc   func    ,
                        VARIANT*     init    ,
                        VARIANT*     matrix  ,
                        __int32         axis    ,
                        VARIANT&     ret     ,
                        bool            left    ); //left==true, right == false

}   // namespace

////************************************************************************************

//2引数にCallback（VBCallbackFunc型のVBA関数）を適用する
VARIANT  __stdcall
simple_invoke(__int32 pCallback, VARIANT* param1, VARIANT* param2)
{
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !func )
    {
        VARIANT      ret;
        ::VariantInit(&ret);
        return ret;
    }
    return (*func)(param1, param2);
}

//配列matrixの各要素elemとparamにCallback(elem, param)（VBCallbackFunc型のVBA関数）を適用する
VARIANT  __stdcall
mapL(__int32 pCallback, VARIANT* matrix, VARIANT* param)
{
    return mapLR(pCallback, matrix, param, true);
}

//paramと配列matrixの各要素elemにCallback(param, elem)（VBCallbackFunc型のVBA関数）を適用する
VARIANT  __stdcall
mapR(__int32 pCallback, VARIANT* param, VARIANT* matrix)
{
    return mapLR(pCallback, matrix, param, false);
}

//**************************************************************************

//配列matrix1とmatrix2の各要素に2変数のCallback（VBCallbackFunc型のVBA関数）を適用する
VARIANT  __stdcall
zipWith(__int32 pCallback, VARIANT* matrix1, VARIANT* matrix2)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    //----------------------------
    if ( !matrix1 || !matrix2 || !func )                        return ret;
    if (  0 == (VT_ARRAY & matrix1->vt ) &&  0 == (VT_ARRAY & matrix2->vt ) )
                                                                return (*func)(matrix1, matrix2);
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
                        {minV(bounds1[0].cElements, bounds2[0].cElements), bounds1[0].lLbound},
                        {minV(bounds1[1].cElements, bounds2[1].cElements), bounds1[1].lLbound},
                        {minV(bounds1[2].cElements, bounds2[2].cElements), bounds1[2].lLbound}      };
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
                VARIANT result = (*func)(&elem1, &elem2);
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
foldl(__int32 pCallback, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !init || !func )                                return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return (*func)(init, matrix);
    fold_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
VARIANT  __stdcall
foldr(__int32 pCallback, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !init || !func )                                return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return (*func)(matrix, init);
    fold_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldl1(__int32 pCallback, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !func )                                         return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return *matrix;
    fold_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
VARIANT  __stdcall
foldr1(__int32 pCallback, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !func )                                         return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return *matrix;
    fold_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
VARIANT  __stdcall
scanl(__int32 pCallback, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !init || !func )                                return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return (*func)(init, matrix);
    scan_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
VARIANT  __stdcall
scanr(__int32 pCallback, VARIANT* init, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !init || !func )                                return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return (*func)(matrix, init);
    scan_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanl1(__int32 pCallback, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !func )                                         return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return *matrix;
    scan_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
VARIANT  __stdcall
scanr1(__int32 pCallback, VARIANT* matrix, __int32 axis)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    if ( !matrix || !func )                                         return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )                           return *matrix;
    scan_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

namespace   {

    //mapLとmapRの共通処理
    VARIANT  
    mapLR(  __int32      pCallback   ,
            VARIANT*     matrix      ,
            VARIANT*     param       ,
            bool         left        )   //left==true, right == false
    {
        VARIANT      ret;
        ::VariantInit(&ret);
        VBCallbackFunc  func = reinterpret_cast<VBCallbackFunc>(pCallback);
        //----------------------------
        if ( !matrix || !func )                                     return ret;
        if (  0 == (VT_ARRAY & matrix->vt ) )
            return left? (*func)(matrix, param) : (*func)(param, matrix);
        //----------------------------
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( 0 == dim || 3 < dim )                                  return ret;
        SAFEARRAYBOUND bounds[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
        safeArrayBounds(pArray, dim, bounds);
        // SAFEARRAY作成
        SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, dim, bounds);
        for ( ULONG i = 0; i < bounds[0].cElements; ++i )
        {
            for ( ULONG j = 0; j < bounds[1].cElements; ++j )
            {
                for ( ULONG k = 0; k < bounds[2].cElements; ++k )
                {
                    LONG index[3] = { static_cast<LONG>(i)+bounds[0].lLbound,
                                      static_cast<LONG>(j)+bounds[1].lLbound,
                                      static_cast<LONG>(k)+bounds[2].lLbound    };
                    VARIANT elem;
                    ::VariantInit(&elem);
                    ::SafeArrayGetElement(pArray, index, &elem);
                    VARIANT result = left?   (*func)(&elem, param):
                                                (*func)(param, &elem);
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

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple(  VBCallbackFunc  func    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( !func || 0 == dim || 3 < dim )                             return;
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
                        if ( left )     tmp = (*func)(&result, &elem);
                        else            tmp = (*func)(&elem, &result);
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
    void   scan_imple(  VBCallbackFunc  func    ,
                        VARIANT*        init    ,
                        VARIANT*        matrix  ,
                        __int32         axis    ,
                        VARIANT&        ret     ,
                        bool            left    ) //left==true, right == false
    {
        SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
        UINT dim = ::SafeArrayGetDim(pArray);
        if ( !func || 0 == dim || 3 < dim )                             return;
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
                        if ( left )     tmp = (*func)(&result, &elem);
                        else            tmp = (*func)(&elem, &result);
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
