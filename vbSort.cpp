//vbSort.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include <algorithm>
#include "VBA_NestFunc.hpp"

//比較関数
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
public:
    compDictionaryFunctor(VARIANT* pA) : begin(pA) { }
    bool operator ()(__int32 i, __int32 j) const
    {
        safearrayRef arr1(begin + i);
        safearrayRef arr2(begin + j);
        if ( arr1.getDim() == 1 || arr2.getDim() == 1 )     return false;
        for ( ULONG k = 0; k < arr1.getSize(1) && k < arr2.getSize(1); ++k )
        {
            if ( VARCMP_LT == VarCmp(&arr1(k), &arr2(k), LANG_JAPANESE, 0) )
                return true;
            if ( VARCMP_LT == VarCmp(&arr2(k), &arr1(k), LANG_JAPANESE, 0) )
                return false;
        }
        return false;
    }
};

VARIANT __stdcall stdsort(VARIANT* array, __int32 defaultFlag, VARIANT* pComp)
{
    VARIANT      ret;
    ::VariantInit(&ret);
    safearrayRef arrIn(array);
    if ( 1 != arrIn.getDim() )          return ret;
    std::unique_ptr<__int32[]>	index(new __int32[arrIn.getSize(1)]);
    std::unique_ptr<VARIANT[]>	VArray(new VARIANT[arrIn.getSize(1)]);
    for (std::size_t i = 0; i < arrIn.getSize(1); ++i)
    {
        index[i] = static_cast<__int32>(i);
        ::VariantInit(&VArray[i]);
        ::VariantCopyInd(&VArray[i], &arrIn(i));
    }
    if ( defaultFlag == 1 ) //1次元昇順
    {
        compFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + arrIn.getSize(1), functor);
    }
    else if ( defaultFlag == 2 ) //2次元昇順
    {
        compDictionaryFunctor   functor(VArray.get());
        std::sort(index.get(), index.get() + arrIn.getSize(1), functor);
    }
    else if ( pComp )
    {
        compareByVBAfunc functor(VArray.get(), pComp);
        if ( functor.valid() )
            std::sort(index.get(), index.get() + arrIn.getSize(1), functor);
    }
    //-------------------------------------------------------
    SAFEARRAYBOUND boundRet = { static_cast<ULONG>(arrIn.getSize(1)), 0};   //要素数、LBound
    SAFEARRAY* retArray = ::SafeArrayCreate(VT_VARIANT, 1, &boundRet);
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = retArray;
    safearrayRef arrOut(&ret);
    VARIANT elem;
    ::VariantInit(&elem);
    elem.vt = VT_I4;
    for ( std::size_t i = 0; i < arrIn.getSize(1); ++i )
    {
        elem.lVal = static_cast<LONG>(index[i] + arrIn.getOriginalLBound(1));
        ::VariantCopy(&arrOut(i), &elem);
    }
    return      ret;
}
