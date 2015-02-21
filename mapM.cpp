#include "stdafx.h"
#include "OAIdl.h"

//Variant左辺値の参照を返す
tagVARIANT  __stdcall
variantRef(tagVARIANT* lvalueRef)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    if ( lvalueRef )
    {
        if ( (VT_BYREF & lvalueRef->vt) && (VT_VARIANT == (VT_VARIANT & lvalueRef->vt)) )
        {
            ret.vt = VT_BYREF | VT_VARIANT;
            ret.pvarVal = lvalueRef->pvarVal;
    }
        else
        {
            ret.vt = VT_BYREF | VT_VARIANT;
            ret.pvarVal = lvalueRef;
        }
    }
    return ret;
}

    //Variant左辺値の参照外し
    tagVARIANT* variantDeRef_imple(tagVARIANT* vRef)
    {
        return ( vRef && (VT_BYREF & vRef->vt) && (VT_VARIANT == (VT_VARIANT & vRef->vt)) )?
            vRef->pvarVal:
            vRef;
    }
// Variantの左辺値参照かどうかを返す
__int32 __stdcall
isVariantRef(tagVARIANT* value)
{
    return ( value && (VT_BYREF & value->vt) && (VT_VARIANT == (VT_VARIANT & value->vt)) )?
        1:
        0;
}

//Variant左辺値の参照外し
tagVARIANT  __stdcall
variantDeRef(tagVARIANT* vRef)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    tagVARIANT* p = variantDeRef_imple(vRef);
    if ( p )    ::VariantCopyInd(&ret, p);
    return ret;
}


//VBA配列の次元取得
__int32 __stdcall Dimension(tagVARIANT* pvariant)
{
    tagVARIANT* p = variantDeRef_imple(pvariant);
	if ( !p || 0 == (VT_ARRAY & p->vt ) )
		return	0;
	else
		return ::SafeArrayGetDim(p->parray);
}

//使用する唯一のVBAコールバック関数型  VBCallbackFunc  の宣言
//VBAにおけるシグネチャは
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// もしくは
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
typedef tagVARIANT (__stdcall *VBCallbackFunc)(tagVARIANT*, tagVARIANT*);


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
    tagVARIANT  mapLR(  __int32         pCallback   ,
                        tagVARIANT*     matrix      ,
                        tagVARIANT*     param       ,
                        bool            left        );  //left==true, right == false

    //foldl と foldr と foldl1 と foldr1 の共通処理
    void   fold_imple( VBCallbackFunc   func    ,
                        tagVARIANT*     init    ,
                        tagVARIANT*     matrix  ,
                        __int32         axis    ,
                        tagVARIANT&     ret     ,
                        bool            left    ); //left==true, right == false

    //scanl と scanr と scanl1 と scanr1 の共通処理
    void   scan_imple( VBCallbackFunc   func    ,
                        tagVARIANT*     init    ,
                        tagVARIANT*     matrix  ,
                        __int32         axis    ,
                        tagVARIANT&     ret     ,
                        bool            left    ); //left==true, right == false

}   // namespace

////************************************************************************************

//2引数にCallback（VBCallbackFunc型のVBA関数）を適用する
    tagVARIANT
    simple_invoke_imple(VBCallbackFunc func, tagVARIANT* param1, tagVARIANT* param2)
    {
        if ( !func )
        {
            tagVARIANT      ret;
            ::VariantInit(&ret);
            return ret;
        }
        return (*func)(variantDeRef_imple(param1), variantDeRef_imple(param2));
    }

tagVARIANT  __stdcall
simple_invoke(__int32 pCallback, tagVARIANT* param1, tagVARIANT* param2)
{
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    return simple_invoke_imple(func, param1, param2);
}

//配列matrixの各要素elemとparamにCallback(elem, param)（VBCallbackFunc型のVBA関数）を適用する
tagVARIANT  __stdcall
mapL(__int32 pCallback, tagVARIANT* matrix, tagVARIANT* param)
{
    return mapLR(pCallback, matrix, param, true);
}

//paramと配列matrixの各要素elemにCallback(param, elem)（VBCallbackFunc型のVBA関数）を適用する
tagVARIANT  __stdcall
mapR(__int32 pCallback, tagVARIANT* param, tagVARIANT* matrix)
{
    return mapLR(pCallback, matrix, param, false);
}

//**************************************************************************

//配列matrix1とmatrix2の各要素に2変数のCallback（VBCallbackFunc型のVBA関数）を適用する
tagVARIANT  __stdcall
zipWith(__int32 pCallback, tagVARIANT* matrix1_, tagVARIANT* matrix2_)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto matrix1 = variantDeRef_imple(matrix1_);
    auto matrix2 = variantDeRef_imple(matrix2_);
    //----------------------------
    if ( !matrix1 || !matrix2 || !func )                return ret;
    if (  0 == (VT_ARRAY & matrix1->vt ) &&  0 == (VT_ARRAY & matrix2->vt ) )
                                                        return simple_invoke_imple(func, matrix1, matrix2);
    if (  0 == (VT_ARRAY & matrix1->vt ) ||  0 == (VT_ARRAY & matrix2->vt ) )
                                                        return ret;
    //----------------------------
    SAFEARRAY* pArray1 = ( 0 == (VT_BYREF & matrix1->vt) )?  (matrix1->parray): (*matrix1->pparray);
    SAFEARRAY* pArray2 = ( 0 == (VT_BYREF & matrix2->vt) )?  (matrix2->parray): (*matrix2->pparray);
    UINT dim  = ::SafeArrayGetDim(pArray1);
    UINT dim2 = ::SafeArrayGetDim(pArray2);
    if ( 0 == dim || 3 < dim || dim != dim2 )           return ret;
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
                tagVARIANT elem1, elem2;
                ::VariantInit(&elem1);
                ::VariantInit(&elem2);
                ::SafeArrayGetElement(pArray1, index1, &elem1);
                ::SafeArrayGetElement(pArray2, index2, &elem2);
                tagVARIANT result = simple_invoke_imple(func, &elem1, &elem2);
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
tagVARIANT  __stdcall
foldl(__int32 pCallback, tagVARIANT* init_, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto init = variantDeRef_imple(init_);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !init || !func )                    return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return simple_invoke_imple(func, init, matrix);
    fold_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（初期値指定あり）
tagVARIANT  __stdcall
foldr(__int32 pCallback, tagVARIANT* init_, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto init = variantDeRef_imple(init_);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !init || !func )                    return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return simple_invoke_imple(func, matrix, init);
    fold_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左畳み込み（先頭要素を初期値とする）
tagVARIANT  __stdcall
foldl1(__int32 pCallback, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !func )                             return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    fold_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右畳み込み（先頭要素を初期値とする）
tagVARIANT  __stdcall
foldr1(__int32 pCallback, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !func )                             return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    fold_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//3次元までのVBA配列に対する特定の軸に沿った左scan（初期値指定あり）
tagVARIANT  __stdcall
scanl(__int32 pCallback, tagVARIANT* init_, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto init = variantDeRef_imple(init_);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !init || !func )                    return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return simple_invoke_imple(func, init, matrix);
    scan_imple(func, init, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（初期値指定あり）
tagVARIANT  __stdcall
scanr(__int32 pCallback, tagVARIANT* init_, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto init = variantDeRef_imple(init_);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !init || !func )                    return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return simple_invoke_imple(func, matrix, init);
    scan_imple(func, init, matrix, axis, ret, false);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った左scan（先頭要素を初期値とする）
tagVARIANT  __stdcall
scanl1(__int32 pCallback, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !func )                             return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, true);
    return      ret;
}

//3次元までのVBA配列に対する特定の軸に沿った右scan（先頭要素を初期値とする）
tagVARIANT  __stdcall
scanr1(__int32 pCallback, tagVARIANT* matrix_, __int32 axis)
{
    tagVARIANT      ret;
    ::VariantInit(&ret);
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    auto matrix = variantDeRef_imple(matrix_);
    if ( !matrix || !func )                             return ret;
    if (  0 == (VT_ARRAY & matrix->vt ) )               return *matrix;
    scan_imple(func, 0, matrix, axis, ret, false);
    return      ret;
}

//**************************************************************************

//VBA配列matrixの各要素でCallbackによる評価結果がゼロでないものの数
__int32     __stdcall
count_if(__int32 pCallback, tagVARIANT* matrix_, tagVARIANT* additionalParameter_)
{
    auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
    //----------------------------
    auto matrix = variantDeRef_imple(matrix_);
    auto additionalParameter = variantDeRef_imple(additionalParameter_);
    if ( !matrix || !func )                     return 0;
    if (  0 == (VT_ARRAY & matrix->vt ) )       return simple_invoke_imple(func, matrix, additionalParameter).lVal? 1: 0;
    //----------------------------
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & matrix->vt) )?  (matrix->parray): (*matrix->pparray);
    UINT dim = ::SafeArrayGetDim(pArray);
    if ( 0 == dim || 3 < dim )                  return 0;
    SAFEARRAYBOUND bounds[3] = {{1,0}, {1,0}, {1,0}};   //要素数、LBound
    safeArrayBounds(pArray, dim, bounds);
    __int32     ret = 0;
    for ( ULONG i = 0; i < bounds[0].cElements; ++i )
    {
        for ( ULONG j = 0; j < bounds[1].cElements; ++j )
        {
            for ( ULONG k = 0; k < bounds[2].cElements; ++k )
            {
                LONG index[3] = {   static_cast<LONG>(i)+bounds[0].lLbound,
                                    static_cast<LONG>(j)+bounds[1].lLbound,
                                    static_cast<LONG>(k)+bounds[2].lLbound  };
                tagVARIANT elem;
                ::VariantInit(&elem);
                ::SafeArrayGetElement(pArray, index, &elem);
                tagVARIANT result = simple_invoke_imple(func, &elem, additionalParameter);
                if ( result.lVal )  ++ret;
                ::VariantClear(&elem);
                ::VariantClear(&result);
            }
        }
    }
    return      ret;
}

//*****************************************************************

namespace   {

    //mapLとmapRの共通処理
    tagVARIANT  
    mapLR(  __int32         pCallback   ,
            tagVARIANT*     matrix_     ,
            tagVARIANT*     param_      ,
            bool            left        )   //left==true, right == false
    {
        tagVARIANT      ret;
        ::VariantInit(&ret);
        auto  func = reinterpret_cast<VBCallbackFunc>(pCallback);
        auto matrix = variantDeRef_imple(matrix_);
        auto param = variantDeRef_imple(param_);
        //----------------------------
        if ( !matrix || !func )                                     return ret;
        if (  0 == (VT_ARRAY & matrix->vt ) )
            return left? simple_invoke_imple(func, matrix, param) : simple_invoke_imple(func, param, matrix);
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
                    tagVARIANT elem;
                    ::VariantInit(&elem);
                    ::SafeArrayGetElement(pArray, index, &elem);
                    tagVARIANT result = left?   simple_invoke_imple(func, &elem, param):
                                                simple_invoke_imple(func, param, &elem);
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
                        tagVARIANT*     init    ,
                        tagVARIANT*     matrix  ,
                        __int32         axis    ,
                        tagVARIANT&     ret     ,
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
                tagVARIANT result;
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
                        tagVARIANT elem, tmp;
                        ::VariantInit(&elem);
                        ::VariantInit(&tmp);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     tmp = simple_invoke_imple(func, &result, &elem);
                        else            tmp = simple_invoke_imple(func, &elem, &result);
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
                        tagVARIANT*     init    ,
                        tagVARIANT*     matrix  ,
                        __int32         axis    ,
                        tagVARIANT&     ret     ,
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
                tagVARIANT result;
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
                        tagVARIANT elem, tmp;
                        ::VariantInit(&elem);
                        ::VariantInit(&tmp);
                        ::SafeArrayGetElement(pArray, sourceIndex, &elem);
                        if ( left )     tmp = simple_invoke_imple(func, &result, &elem);
                        else            tmp = simple_invoke_imple(func, &elem, &result);
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
