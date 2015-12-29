//vbSort.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include <algorithm>
#include "VBA_NestFunc.hpp"

//比較関数
class compareByVBAfunc   {
    VARIANT*        begin;
    functionExpr*   comp;
public:
    compareByVBAfunc(VARIANT* pA, functionExpr& f) noexcept : begin(pA), comp(&f)    {    }
    compareByVBAfunc(compareByVBAfunc const&) = default;
    compareByVBAfunc(compareByVBAfunc&&) = delete;
    ~compareByVBAfunc() = default;
    bool valid() const noexcept { return comp != nullptr;  }
    bool operator ()(__int32 i, __int32 j) const noexcept
    {
        return comp->eval(begin + i, begin + j)->lVal ? true: false;
    }
};

//1次元昇順
class compFunctor  {
    VARIANT*    begin;
public:
    explicit compFunctor(VARIANT* pA) noexcept : begin(pA) { }
    compFunctor(compFunctor const&) = default;
    compFunctor(compFunctor&&) = delete;
    ~compFunctor() = default;
    bool operator ()(__int32 i, __int32 j) const noexcept
    {
        return VARCMP_LT == VarCmp(begin + i, begin + j, LANG_JAPANESE, 0);
    }
};

//2次元昇順
class compDictionaryFunctor  {
    VARIANT*    begin;
public:
    explicit compDictionaryFunctor(VARIANT* pA) noexcept : begin(pA) { }
    compDictionaryFunctor(compDictionaryFunctor const&) = default;
    compDictionaryFunctor(compDictionaryFunctor&&) = delete;
    ~compDictionaryFunctor() = default;
    bool operator ()(__int32 i, __int32 j) const noexcept
    {
        safearrayRef arr1(begin + i);
        safearrayRef arr2(begin + j);
        if ( arr1.getDim() != 1 || arr2.getDim() != 1 )     return false;
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
    auto index = std::make_unique<__int32[]>(arrIn.getSize(1));
    auto VArray = std::make_unique<VARIANT[]>(arrIn.getSize(1));
    auto refFlag = std::make_unique<bool[]>(arrIn.getSize(1));
    for (std::size_t i = 0; i < arrIn.getSize(1); ++i)
    {
        index[i] = static_cast<__int32>(i);
        ::VariantInit(&VArray[i]);
        if ( VT_BYREF & arrIn(i).vt )
        {
            refFlag[i] = true;
            ::VariantCopyInd(&VArray[i], &arrIn(i)); 
        }
        else
        {
            refFlag[i] = false;
            VArray[i] = arrIn(i);
        }
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
        functionExpr comp(pComp);
        compareByVBAfunc functor(VArray.get(), comp);
        if ( functor.valid() )
            std::sort(index.get(), index.get() + arrIn.getSize(1), functor);
    }
    //-------------------------------------------------------
    SAFEARRAYBOUND boundRet = { static_cast<ULONG>(arrIn.getSize(1)), 0};   //要素数、LBound
    ret.vt = VT_ARRAY | VT_VARIANT;
    ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &boundRet);
    safearrayRef arrOut(&ret);
    VARIANT elem;
    ::VariantInit(&elem);
    elem.vt = VT_I4;
    for ( std::size_t i = 0; i < arrIn.getSize(1); ++i )
    {
        elem.lVal = static_cast<decltype(elem.lVal)>(index[i] + arrIn.getOriginalLBound(1));
        ::VariantCopy(&arrOut(i), &elem);
        if ( refFlag[i] )   ::VariantClear(&VArray[i]);
    }
    return      ret;
}
