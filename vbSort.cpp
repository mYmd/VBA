//vbSort.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include <algorithm>
#include <OleAuto.h>//<OAIdl.h>
#include "VBA_NestFunc.hpp"

//”äŠrŠÖ”
class compareByVBAfunc   {
    VARIANT*    begin;
    std::shared_ptr<functionExpr> comp;
public:
    compareByVBAfunc(VARIANT* pA, VARIANT* f) : begin(pA)
    {
        VBCallbackFunc pf(f);
        if ( pf )   comp.reset(new functionExpr(pf));
    }
    bool valid() const  { return static_cast<bool>(comp);  }
    bool operator ()(__int32 i, __int32 j) const
    {
        return comp->eval(begin + i, begin + j)->lVal ? true: false;
    }
};

//1ŽŸŒ³¸‡
class compFunctor  {
    VARIANT*    begin;
public:
    compFunctor(VARIANT* pA) : begin(pA) { }
    bool operator ()(__int32 i, __int32 j) const
    {
        return VARCMP_LT == VarCmp(begin + i, begin + j, LANG_JAPANESE, 0);
    }
};

//2ŽŸŒ³¸‡
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

VARIANT __stdcall stdsort(VARIANT* array, __int32 defaultFlag, VARIANT* pComp)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    if ( 1 != Dimension(array) )            return ret;
    SAFEARRAY* pArray = ( 0 == (VT_BYREF & array->vt) )?  (array->parray): (*array->pparray);
    SAFEARRAYBOUND bound = {1, 0};   //—v‘f”ALBound
    safeArrayBounds(pArray, 1, &bound);
    std::unique_ptr<__int32[]>	index(new __int32[bound.cElements]);
    std::unique_ptr<VARIANT[]>	VArray(new VARIANT[bound.cElements]);
    for ( ULONG i = 0; i < bound.cElements; ++i )
    {
        index[i] = static_cast<__int32>(i);
        auto j = static_cast<LONG>(i + bound.lLbound);
        ::SafeArrayGetElement(pArray, &j, &VArray[i]);
    }
    if ( defaultFlag == 1 ) //1ŽŸŒ³¸‡
    {
        compFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    else if ( defaultFlag == 2 ) //2ŽŸŒ³¸‡
    {
        compDictionaryFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    else if ( pComp )
    {
        compareByVBAfunc functor(VArray.get(), pComp);
        if ( functor.valid() )
            std::sort(index.get(), index.get() + bound.cElements, functor);
    }
    //-------------------------------------------------------
    SAFEARRAYBOUND boundRet = { bound.cElements, 0};   //—v‘f”ALBound
    SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, 1, &boundRet);
    VARIANT elem;
    ::VariantInit(&elem);
    elem.vt = VT_I4;
    for ( LONG i = 0; i < static_cast<LONG>(bound.cElements); ++i )
    {
        elem.lVal = index[i] + bound.lLbound;
        ::SafeArrayPutElement(retArray, &i, &elem);
    }
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = retArray;
    return      ret;
}
