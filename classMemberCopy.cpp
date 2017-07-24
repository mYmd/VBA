//classMemberCopy.cpp
//Copyright (c) 2016 mmYYmmdd
#include "stdafx.h"
#include <utility>
#include "OAIdl.h"

#if _MSC_VER < 1900
#define noexcept throw()
#endif

namespace   {
    //ReDim可能な配列かどうか
    HRESULT redimCheck(SAFEARRAY * psa) noexcept
    {
        if ( const auto dim = ::SafeArrayGetDim(psa) )
        {
            SAFEARRAYBOUND saboundNew{ 0, 0 };
            LONG lLbound, lUbound;
            auto a1 = ::SafeArrayGetLBound(psa, dim, &lLbound);
            auto a2 = ::SafeArrayGetUBound(psa, dim, &lUbound);
            saboundNew.cElements = 1 + lUbound - lLbound;
            saboundNew.lLbound = lLbound;
            return ::SafeArrayRedim(psa, &saboundNew);
        }
        return S_OK;    //未初期化配列もReDim可能
    }

    // コピー  *me->mem   <===>   *target->mem
    template <typename T>
    void copy_or_swap(T* pV, IDispatch* me, IDispatch* target, __int32 dir) noexcept
    {
        auto n = reinterpret_cast<char*>(pV) - reinterpret_cast<char*>(me);
        auto p = reinterpret_cast<T*>(reinterpret_cast<char*>(target) + n);
        if (0 < dir)        *p = *pV;
        else if (dir < 0)   *pV = *p;
        else                std::swap(*p, *pV);
    }
}

// VBAのクラスオブジェクトの特定のメンバ（オブジェクト型以外）を同一クラスの他のオブジェクトにコピーする
// me     : クラスオブジェクト（ByVal x As ***Class) (As Objectではダメ)
// mbr    : 特定のメンバ（オブジェクト型以外）
// target : 対象オブジェクト（meと同一クラス）
// dir    : 0 < dir : meからtargetへコピー、 dir < 0 : targetからmeへコピー、　dir == 0 : swap
VARIANT_BOOL __stdcall copy_valueMember(IDispatch*  me      ,
                                        VARIANT*    mbr     ,
                                        IDispatch*  target  ,
                                        __int32     dir     ) noexcept
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
        }
        else switch ( vt % VT_ARRAY )
        {
            case    VT_UI1:     copy_or_swap(mbr->pbVal, me, target, dir);      break;
            case    VT_I2:      copy_or_swap(mbr->piVal, me, target, dir);      break;
            case    VT_I4:      copy_or_swap(mbr->plVal, me, target, dir);      break;
            case    VT_I8:      copy_or_swap(mbr->pllVal, me, target, dir);     break;
            case    VT_R4:      copy_or_swap(mbr->pfltVal, me, target, dir);    break;
            case    VT_R8:      copy_or_swap(mbr->pdblVal, me, target, dir);    break;
            case    VT_CY:      copy_or_swap(mbr->pcyVal, me, target, dir);     break;
            case    VT_DATE:    copy_or_swap(mbr->pdate, me, target, dir);      break;
            case    VT_BOOL:    copy_or_swap(mbr->pboolVal, me, target, dir);   break;
            case    VT_BSTR:    /////////////
            {
                auto n = reinterpret_cast<char*>(mbr->pbstrVal) - reinterpret_cast<char*>(me);
                auto p = reinterpret_cast<BSTR*>(reinterpret_cast<char*>(target) + n);
                if ( 0 < dir )          ::SysReAllocString(p, *mbr->pbstrVal);
                else if ( dir < 0 )     ::SysReAllocString(mbr->pbstrVal, *p);
                else                    std::swap(*p, *mbr->pbstrVal);
                break;
            }
            default:                return 0;
        }
    }
    else
    {
        auto n = reinterpret_cast<char*>(mbr) - reinterpret_cast<char*>(me);
        auto p = reinterpret_cast<VARIANT*>(reinterpret_cast<char*>(target) + n);
        if ( 0 < dir )      { ::VariantCopy(p, mbr); }
        else if ( dir < 0 ) { ::VariantCopy(mbr, p); }
        else                { std::swap(*p, *mbr); }
    }
    return -1;
}

// VBAのクラスオブジェクトのオブジェクト型のメンバを同一クラスの他のオブジェクトにコピーする
// me     : クラスオブジェクト（ByVal x As ***Class) (As Objectではダメ)
// pmbr   : 特定のメンバ
// target : 対象オブジェクト（meと同一クラス）
// dir    : 0 < dir : meからtargetへコピー、 dir < 0 : targetからmeへコピー、　dir == 0 : swap
// method : メンバであるオブジェクトの引数なしのコピーメソッド名（"clone"等）
VARIANT_BOOL __stdcall copy_objectMember(   IDispatch*  me      ,
                                            IDispatch** pmbr    ,
                                            IDispatch*  target  ,
                                            __int32     dir     ,
                                            VARIANT*    method  ) noexcept
{
    auto n = reinterpret_cast<char*>(pmbr) - reinterpret_cast<char*>(me);
    auto p = reinterpret_cast<IDispatch**>(reinterpret_cast<char*>(target) + n);
    if ( 0 == dir )
    {
        std::swap(*p, *pmbr);
        return -1;
    }
    else
    {
        BSTR name = (method->vt & VT_BYREF )?
                ( ((method->vt & VT_BSTR) && method->pbstrVal )? *method->pbstrVal : nullptr ) :
                ( ((method->vt & VT_BSTR) && method->bstrVal ) ? method->bstrVal  : nullptr );
        if ( !name )    return 0;
        VARIANT tmp;
        ::VariantInit(&tmp);
        tmp.vt = VT_DISPATCH;
        tmp.pdispVal = nullptr;
        DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
        if (IDispatch*& from_ = (0 < dir) ? *pmbr : *p )
        {
            DISPID dispid;
            from_->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
            HRESULT hr = from_->Invoke(dispid,
                                        IID_NULL,
                                        LOCALE_SYSTEM_DEFAULT,
                                        DISPATCH_METHOD,
                                        &dp,
                                        &tmp,
                                        NULL,
                                        NULL);
            if (hr != S_OK)  return 0;
        }
        std::swap(tmp.pdispVal, (0 < dir) ? *p: *pmbr);
        ::VariantClear(&tmp);
        return -1;
    }
}
