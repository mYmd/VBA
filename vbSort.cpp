#include "stdafx.h"
#include <algorithm>
#include <memory>
#include <OleAuto.h>//<OAIdl.h>

//VBA比較関数型  vbCompFunc  の宣言 (VBCallbackFuncと同一)
//VBAにおけるシグネチャは
// Function comp(ByRef elem As Variant, ByRef dummy As Variant) As Variant
typedef VARIANT (__stdcall *vbCompFunc)(VARIANT*, VARIANT*);

__int32 __stdcall Dimension(const VARIANT* pv);
void safeArrayBounds(SAFEARRAY* pArray, UINT dim, SAFEARRAYBOUND bounds[]);

class compareByVBAfunc   {
    VARIANT*    begin;
    vbCompFunc  comp;
public:
    compareByVBAfunc(VARIANT* pA, vbCompFunc f) : begin(pA), comp(f) { }
    bool operator ()(__int32 i, __int32 j) const
    {
        return (*comp)(begin + i, begin + j).lVal ? true: false;
    }
};

//1次元昇順
class compFunctor  {
    VARIANT*    begin;
public:
    compFunctor(VARIANT* pA) : begin(pA) { }
    bool operator ()(__int32 i, __int32 j) const
    {
        return VARCMP_LT == VarCmp(begin + i, begin + j, LANG_JAPANESE, 0);
    }
};

//2次元昇順
class compDictionaryFunctor  {
    VARIANT*    begin;
    void set(__int32 k, SAFEARRAY*&  pArray, SAFEARRAYBOUND& bound) const
    {
        if ( 1 == Dimension(begin + k) ) 
        {
            pArray = ( 0 == (VT_BYREF & begin[k].vt) )?  (begin[k].parray): (*begin[k].pparray);
            safeArrayBounds(pArray, 1, &bound);
        }
        else
        {
            pArray = nullptr;
        }
    }
public:
    compDictionaryFunctor(VARIANT* pA) : begin(pA) { }
    bool operator ()(__int32 i, __int32 j) const
    {
        SAFEARRAY*  pArray1, *pArray2;
        SAFEARRAYBOUND bound1, bound2;
        set(i, pArray1, bound1);
        set(j, pArray2, bound2);
        if ( pArray1 == nullptr || pArray2 == nullptr )     return false;
        VARIANT Var1, Var2;
        for ( ULONG k = 0; k < bound1.cElements && k < bound2.cElements; ++k )
        {
            auto index1 = static_cast<LONG>(k + bound1.lLbound);
            auto index2 = static_cast<LONG>(k + bound2.lLbound);
            ::SafeArrayGetElement(pArray1, &index1, &Var1);
            ::SafeArrayGetElement(pArray2, &index2, &Var2);
            if ( VARCMP_LT == VarCmp(&Var1, &Var2, LANG_JAPANESE, 0) )
                return true;
            if ( VARCMP_LT == VarCmp(&Var2, &Var1, LANG_JAPANESE, 0) )
                return false;
        }
        return false;
    }
};

VARIANT __stdcall stdsort(VARIANT* array, __int32 pComp)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    if ( 1 != Dimension(array) )            return ret;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & array->vt) )?  (array->parray): (*array->pparray);
    SAFEARRAYBOUND bound = {1, 0};   //要素数、LBound
    safeArrayBounds(pArray, 1, &bound);
    std::unique_ptr<__int32[]>	index(new __int32[bound.cElements]);
    std::unique_ptr<VARIANT[]>	VArray(new VARIANT[bound.cElements]);
    for ( ULONG i = 0; i < bound.cElements; ++i )
    {
        index[i] = static_cast<__int32>(i);
        auto j = static_cast<LONG>(i + bound.lLbound);
        ::SafeArrayGetElement(pArray, &j, &VArray[i]);
    }
    if ( 0 < pComp )
    {
        compareByVBAfunc functor(VArray.get(), reinterpret_cast<vbCompFunc>(pComp));
        std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    else if ( -1 == pComp ) //1次元昇順
    {
        compFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    else if ( -2 == pComp  ) //2次元昇順
    {
        compDictionaryFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    //-------------------------------------------------------
    SAFEARRAYBOUND boundRet = { bound.cElements, 0};   //要素数、LBound
    SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, 1, &boundRet);
    VARIANT elem;
    ::VariantInit(&elem);
    elem.vt = VT_I4;
    for ( LONG i = 0; i < static_cast<LONG>(bound.cElements); ++i )
    {
        elem.lVal = index[i];
        ::SafeArrayPutElement(retArray, &i, &elem);
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = retArray;
    return      ret;
}
