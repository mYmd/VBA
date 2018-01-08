//fileUtil.cpp
//Copyright (c) 2018 mmYYmmdd
#include "stdafx.h"
#undef min
#undef max
#include <fstream>
#include <codecvt>
#include <vector>
#include "VBA_NestFunc.hpp"

//テキストファイルをVBA配列にする
VARIANT __stdcall
textfile2array(VARIANT const& fileName_, __int32 codepage_, __int32 head_n, __int32 head_cut, __int8 toArry)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )    return iVariant();
    auto codepage = static_cast<UINT>(codepage_);
    if ( head_n < 0 )   head_n = static_cast<__int32>((1u << 31) - 1);
    std::ifstream ifs{fileName, std::ios_base::in};
    __int32 count{0};
    std::string str;
    while ( count < head_cut && std::getline(ifs, str) )    ++count;
    std::wstring w;
    std::vector<std::wstring> wVec;
    while ( count < head_n && std::getline(ifs, str) )
    {        
        auto len = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, nullptr, 0);
        w.resize(len? len-1: 0, L'\0');     // len-1
        wVec.push_back(std::wstring());
        if ( ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, &w[0], len) )
            wVec.back() = w;
        ++count;
    }
    if ( toArry )
    {
        std::vector<VARIANT> varVec(wVec.size());
        std::transform(wVec.cbegin(), wVec.cend(), varVec.begin(), 
                       [](std::wstring const& x){ return bstrVariant(x); });
        return vec2VArray(std::move(varVec));
    }
    else
    {
        std::wstring ret;
        for ( auto p = wVec.cbegin(); p < wVec.cend(); ++p )    ret += *p;
        return bstrVariant(ret);
    }
}
//CP_ACP                    0           // default to ANSI code page
//CP_UTF7                   65000       // UTF-7 translation
//CP_UTF8                   65001       // UTF-8 translation
//CP932                     932         // Microsoftコードページ932
    /*
    std::codecvt_utf8<wchar_t, 0x10ffff, std::codecvt_mode>を
    ifstream::imbue に代入する方法はC++17で非推奨になる
    */

//テキストファイルの各行にVBAHaskell関数を適用する
//返り値：適用した行数
__int32 __stdcall
textfile_for_each(VARIANT const& Fn, VARIANT const& fileName_, __int32 codepage_, __int32 head_n, __int32 head_cut)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )            return 0;
    functionExpr func{Fn};
    if ( !func.isValid() )      return 0;
    auto codepage = static_cast<UINT>(codepage_);
    if ( head_n < 0 )   head_n = static_cast<__int32>((1u << 31) - 1);
    std::ifstream ifs{fileName, std::ios_base::in};
    __int32 count{0};
    std::string str;
    while ( count < head_cut && std::getline(ifs, str) )    ++count;
    std::wstring w;
    __int32 ret{0};
    while ( count < head_n && std::getline(ifs, str) )
    {        
        auto len = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, nullptr, 0);
        w.resize(len? len-1: 0, L'\0');     // len-1
        if ( ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, &w[0], len) )
        {
            auto x = bstrVariant(w);
            func.eval(x, x);
            ++ret;
        }
        ++count;
    }
    return ret;
}

//VBA配列をテキストファイルに書き出す
__int32 array2textfile(VARIANT const& array, VARIANT const& fileName_, __int8 utf8, __int8 feed_at_last)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )    return 0;
    safearrayRef ref{array};
    if ( ref.getDim() != 1 )    return 0;
    auto size = ref.getSize(1);
    std::wofstream ofs{fileName, std::ios_base::out | std::ios_base::trunc};
    ofs.imbue(utf8 ?
              std::locale(std::locale::empty(), new std::codecvt_utf8<wchar_t>) :   //C++17で非推奨
              std::locale("", LC_CTYPE)    );
    auto dest = iVariant(VT_BSTR);
    std::size_t i = 0; 
    for ( ; i < size; ++i )
    {
        if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
            ofs << getBSTR(dest);
        if ( feed_at_last || i < size - 1 )
            ofs << L'\n';
    }
    return i;
}
