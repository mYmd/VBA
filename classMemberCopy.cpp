//classMemberCopy.cpp
//Copyright (c) 2016 mmYYmmdd
#include "stdafx.h"
#include <utility>
#include "OAIdl.h"

namespace   {
    //ReDim可能な配列かどうか
    HRESULT redimCheck(SAFEARRAY * psa);
}

// VBAのクラスオブジェクトの特定のメンバ（オブジェクト型以外）を同一クラスの他のオブジェクトにコピーする
// me     : クラスオブジェクト（ByVal x As ***Class) (As Objectではダメ)
// mbr    : 特定のメンバ（オブジェクト型以外）
// target : 対象オブジェクト（meと同一クラス）
// dir    : 0 < dir : meからtargetへコピー、 dir < 0 : targetからmeへコピー、　dir == 0 : swap
VARIANT_BOOL __stdcall copy_valueMember(IDispatch* me, VARIANT* mbr, IDispatch* target, __int32 dir)
{
    const auto vt = mbr->vt;
    if ( vt & VT_BYREF )
    {
        if ( vt & VT_ARRAY )
        {
            if ( redimCheck(*mbr->pparray) == S_OK )
            {
                auto d = reinterpret_cast<char*>(mbr->pparray) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<SAFEARRAY**>(reinterpret_cast<char*>(target) + d);
                if ( 0 < dir )
                {
                    ::SafeArrayDestroy(*p);
                    ::SafeArrayCopy(*mbr->pparray, p);
                }
                else if ( dir < 0 )
                {
                    ::SafeArrayDestroy(*mbr->pparray);
                    ::SafeArrayCopy(*p, mbr->pparray);
                }
                else
                {
                    std::swap(*p, *mbr->pparray);
                }
            }
            else
            {
                auto d = reinterpret_cast<char*>(*mbr->pparray) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<SAFEARRAY*>(reinterpret_cast<char*>(target) + d);
                if ( 0 < dir )
                {
                    ::SafeArrayCopyData(*mbr->pparray, p);
                }
                else if ( dir < 0 )
                {
                    ::SafeArrayCopyData(p, *mbr->pparray);
                }
                else
                {
                    SAFEARRAY* tmp = nullptr;
                    ::SafeArrayCopy(p, &tmp);
                    ::SafeArrayCopyData(*mbr->pparray, p);
                    ::SafeArrayCopyData(tmp, p);
                    ::SafeArrayDestroy(tmp);
                }
            }
            return -1;
        }
        else switch ( vt % VT_ARRAY )
        {
            case    VT_UI1:
            {
                auto n = reinterpret_cast<char*>(mbr->pbVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<BYTE*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pbVal; }
                else if( dir < 0 )  { *mbr->pbVal = *p; }
                else                { std::swap(*p, *mbr->pbVal); }
                break;
            }
            case    VT_I2:
            {
                auto n = reinterpret_cast<char*>(mbr->piVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<SHORT*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->piVal; }
                else if ( dir < 0 ) { *mbr->piVal = *p; }
                else                { std::swap(*p, *mbr->piVal); }
                break;
            }
            case    VT_I4:
            {
                auto n = reinterpret_cast<char*>(mbr->plVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<LONG*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->plVal; }
                else if ( dir < 0 ) { *mbr->plVal = *p; }
                else                { std::swap(*p, *mbr->plVal); }
                break;
            }
            case    VT_I8:
            {
                auto n = reinterpret_cast<char*>(mbr->pllVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<LONGLONG*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pllVal; }
                else if ( dir < 0 ) { *mbr->pllVal = *p; }
                else                { std::swap(*p, *mbr->pllVal); }
                break;
            }
            case    VT_R4:
            {
                auto n = reinterpret_cast<char*>(mbr->pfltVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<FLOAT*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pfltVal; }
                else if ( dir < 0 ) { *mbr->pfltVal = *p; }
                else                { std::swap(*p, *mbr->pfltVal); }
                break;
            }
            case    VT_R8:
            {
                auto n = reinterpret_cast<char*>(mbr->pdblVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<DOUBLE*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pdblVal; }
                else if ( dir < 0 ) { *mbr->pdblVal = *p; }
                else                { std::swap(*p, *mbr->pdblVal); }
                break;
            }
            case    VT_CY:
            {
                auto n = reinterpret_cast<char*>(mbr->pcyVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<CY*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pcyVal; }
                else if ( dir < 0 ) { *mbr->pcyVal = *p; }
                else                { std::swap(*p, *mbr->pcyVal); }
                break;
            }
            case    VT_DATE:
            {
                auto n = reinterpret_cast<char*>(mbr->pdate) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<DATE*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pdate; }
                else if ( dir < 0 ) { *mbr->pdate = *p; }
                else                { std::swap(*p, *mbr->pdate); }
                break;
            }
            case    VT_BSTR:
            {
                auto n = reinterpret_cast<char*>(mbr->pbstrVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<BSTR*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { ::SysReAllocString(p, *mbr->pbstrVal); }
                else if ( dir < 0 ) { ::SysReAllocString(mbr->pbstrVal, *p); }
                else                { std::swap(*p, *mbr->pbstrVal); }
                break;
            }
            case    VT_BOOL:
            {
                auto n = reinterpret_cast<char*>(mbr->pboolVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<VARIANT_BOOL*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )      { *p = *mbr->pboolVal; }
                else if ( dir < 0 ) { *mbr->pboolVal = *p; }
                else                { std::swap(*p, *mbr->pboolVal); }
                break;
            }
            default:
            {
                return 0;
            }
        }
        return -1;
    }
    else
    {
        auto n = reinterpret_cast<char*>(mbr) - reinterpret_cast<char*>(me);
        auto p = reinterpret_cast<VARIANT*>(reinterpret_cast<char*>(target) + n);
        if ( 0 < dir )      { ::VariantCopy(p, mbr); }
        else if ( dir < 0 ) { ::VariantCopy(mbr, p); }
        else                { std::swap(*p, *mbr); }
        return -1;
    }
}

// VBAのクラスオブジェクトのオブジェクト型のメンバを同一クラスの他のオブジェクトにコピーする
// me     : クラスオブジェクト（ByVal x As ***Class) (As Objectではダメ)
// pmbr   : 特定のメンバ
// target : 対象オブジェクト（meと同一クラス）
// dir    : 0 < dir : meからtargetへコピー、 dir < 0 : targetからmeへコピー、　dir == 0 : swap
// method : メンバであるオブジェクトの引数なしのコピーメソッド名（"clone"等）
VARIANT_BOOL __stdcall copy_objectMember(IDispatch* me, IDispatch** pmbr, IDispatch* target, __int32 dir, VARIANT* method)
{
    auto n = reinterpret_cast<char*>(pmbr) - reinterpret_cast<char*>(me);
    auto p = reinterpret_cast<IDispatch**>(reinterpret_cast<char*>(target) + n);
    if ( 0 == dir )
    {
        std::swap(*p, *pmbr);
    }
    else
    {
        BSTR name = [](VARIANT* nm)->BSTR {
            if ( nm->vt & VT_BYREF )
                return ( (nm->vt & VT_BSTR) && nm->pbstrVal )? *nm->pbstrVal : nullptr;
            else
                return ( (nm->vt & VT_BSTR) && nm->bstrVal )?   nm->bstrVal  : nullptr;
        } (method);
        if ( !name )    return 0;
        VARIANT tmp;
        ::VariantInit(&tmp);
        DISPID dispid;
        (*pmbr)->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
        HRESULT hr = (*pmbr)->Invoke(dispid,
                                     IID_NULL,
                                     LOCALE_SYSTEM_DEFAULT,
                                     DISPATCH_METHOD,
                                     &dp,
                                     &tmp,
                                     NULL,
                                     NULL);
        if ( hr != S_OK )  return 0;
        if ( 0 < dir )      std::swap(tmp.pdispVal, *p);
        else                std::swap(tmp.pdispVal, *pmbr);
        ::VariantClear(&tmp);
    }
    return -1;
}

namespace   {
    //ReDim可能な配列かどうか
    HRESULT redimCheck(SAFEARRAY * psa)
    {
        const auto dim = ::SafeArrayGetDim(psa);
        if ( 0 < dim )
        {
            SAFEARRAYBOUND saboundNew{0, 0};
            LONG lLbound, lUbound;
            auto a1 = ::SafeArrayGetLBound(psa, dim, &lLbound);
            auto a2 = ::SafeArrayGetUBound(psa, dim, &lUbound);
            saboundNew.cElements = 1 + lUbound - lLbound;
            saboundNew.lLbound = lLbound;
            return ::SafeArrayRedim(psa, &saboundNew);
        }
        else
        {
            return S_OK;    //未初期化配列もReDim可能
        }
    }
}
