//fileUtil.cpp
//Copyright (c) 2018 mmYYmmdd
#include "stdafx.h"
#include <fstream>
#include <codecvt>
#include <vector>
#include "VBA_NestFunc.hpp"

namespace   {
    template <typename Fn>
    __int32 fgetws_ex(Fn&&, BSTR, UINT, __int32, __int32);

    template <typename Fn>
    __int32 ifstream_ex(Fn&&, BSTR, UINT, __int32, __int32);

    __int32 fputws_ex(safearrayRef&, BSTR, UINT, bool);

    __int32 ofstream_ex(safearrayRef&, BSTR, UINT, bool);

}

//**************************************************

//テキストファイルをVBA配列にする
VARIANT __stdcall
textfile2array(VARIANT const& fileName_, __int32 codepage_, __int32 head_n, __int32 head_cut, __int8 toArry)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )        return iVariant();
    auto codepage = static_cast<UINT>(codepage_);
    if ( head_n < 0 )       head_n = static_cast<__int32>((1u << 31) - 1);
    std::vector<std::wstring> wVec;
    auto callback = [&](std::wstring const& w) { wVec.push_back(w); };
    //               ANSI ,          UTF-16LE ,              UTF-8
    if ( codepage == 1252 || codepage == 1200 || codepage == 65001 )
        fgetws_ex(callback, fileName, codepage, head_n, head_cut);
    else
        ifstream_ex(callback, fileName, codepage, head_n, head_cut);
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
    if ( !fileName )                return 0;
    functionExpr vbaFunc{Fn};
    bool validFunc = vbaFunc.isValid();
    auto codepage = static_cast<UINT>(codepage_);
    if ( head_n < 0 )   head_n = static_cast<__int32>((1u << 31) - 1);
    auto callback = [&](std::wstring const& w) {
        if ( validFunc )
        {
            auto x = bstrVariant(w);
            vbaFunc.eval(x, x);
        }
    };
    //               ANSI ,          UTF-16LE ,              UTF-8
    if ( codepage == 1252 || codepage == 1200 || codepage == 65001 )
        return fgetws_ex(callback, fileName, codepage, head_n, head_cut);
    else
        return ifstream_ex(callback, fileName, codepage, head_n, head_cut);
}

//VBA配列をテキストファイルに書き出す(UTF-8, UTF-16, ANSI, S-JIS)
__int32 __stdcall
array2textfile(VARIANT const& array, VARIANT const& fileName_, __int32 codepage_, __int8 append)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )            return 0;
    safearrayRef ref{array};
    if ( ref.getDim() != 1 )    return 0;
    auto codepage = static_cast<UINT>(codepage_);
    //               ANSI ,          UTF-16LE ,              UTF-8
    if ( codepage == 1252 || codepage == 1200 || codepage == 65001 )
        return fputws_ex(ref, fileName, codepage, append != 0);
    else
        return ofstream_ex(ref, fileName, codepage, append != 0);
}

//**************************************************

namespace   {
    struct file_closer {
        void operator()(FILE* pF) const  noexcept
        { if (pF)   ::fclose(pF); }
    };

    using fileCloseRAII = std::unique_ptr<FILE, file_closer>;

    std::wstring fgetws_ex_line(FILE* fp, std::size_t n, bool& eof);

    //UTF-16LE(codepage:1200) or UTF-8(codepage:65001)
    //https://msdn.microsoft.com/en-us/library/ee719641.aspx

    template <typename Fn>
    __int32 fgetws_ex(  Fn&&    func    ,
                        BSTR    fileName,
                        UINT    codepage,
                       __int32  head_n  ,
                       __int32  head_cut)
    {
        FILE* fp = nullptr;
        auto err = ::_wfopen_s(&fp,
                               fileName, 
                               (codepage==1200)?    L"rt, ccs=UTF-16LE":
                               (codepage==65001)?   L"rt, ccs=UTF-8":
                                                    L"rt");     //ANSI(1252)
        fileCloseRAII fc_tmp(err? nullptr: fp);
        if ( err || !fp )       return 0;
        std::wstring w(256, L'\0');
        wchar_t* p = nullptr;
        __int32 count{0}, ret{0};
        while ( count < head_cut )
        {
            auto eof = true;
            w = fgetws_ex_line(fp, w.capacity(), eof);
            if ( eof )      break;
            ++count;
        }
        while ( count < head_n )
        {
            auto eof = true;
            w = fgetws_ex_line(fp, w.capacity(), eof);
            if ( eof )      break;
            std::forward<Fn>(func)(w);
            ++count; ++ret;
        }
        return ret;
    }
    //https://msdn.microsoft.com/ja-jp/library/z5hh6ee9.aspx
    //https://msdn.microsoft.com/ja-jp/magazine/mt763237.aspx
    //https://msdn.microsoft.com/ja-jp/library/c565h7xx.aspx

    //改行が出てくるまでバッファを伸ばす
    std::wstring fgetws_ex_line(FILE* fp, std::size_t n, bool& eof)
    {
        std::wstring buf(n, L'\0');
        auto p = std::fgetws(&buf[0], static_cast<int>(buf.size()+1), fp);
        if ( p )
        {
            auto len = std::char_traits<wchar_t>::length(p);
            if ( 0 < len )
            {
                eof = false;
                if ( p[len-1] == L'\n' )
                {
                    buf.resize(len-1);
                    return  buf;
                }
                else    // len == n のはず
                {
                    buf.resize(len);
                    return  buf += fgetws_ex_line(fp, n*2, eof);
                }
            }
        }
        buf.clear();
        return buf;
    }

    template <typename Fn>
    __int32 ifstream_ex(Fn&&        func    ,
                        BSTR        fileName,
                        UINT        codepage,
                        __int32     head_n  ,
                        __int32     head_cut)
    {
        std::ifstream ifs{fileName, std::ios_base::in};
        __int32 count{0}, ret{0};;
        std::string str;
        while ( count < head_cut && std::getline(ifs, str) )    ++count;
        std::wstring w;
        while ( count < head_n && std::getline(ifs, str) )
        {        
            auto len = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, nullptr, 0);
            w.resize(len? len-1: 0, L'\0');     // len-1
            auto re = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, str.data(), -1, &w[0], len);
            std::forward<Fn>(func)(re? w: std::wstring());
            ++count; ++ret;
        }
        return ret;
    }

    //---------------------------------------------------------

    __int32 fputws_ex(safearrayRef& ref, BSTR fileName, UINT codepage, bool append)
    {
        FILE* fp = nullptr;
        auto openmode = append? 
            ((codepage==1200)? L"a+t, ccs=UTF-16LE":
            (codepage==65001)? L"a+t, ccs=UTF-8":   L"a+t"  )
            :
            ((codepage==1200)? L"wt, ccs=UTF-16LE":
            (codepage==65001)? L"wt, ccs=UTF-8":    L"wt"   );
        auto err = ::_wfopen_s(&fp, fileName,  openmode);
        fileCloseRAII fc_tmp(err? nullptr: fp);
        if ( err || !fp )       return 0;
        auto size = ref.getSize(1);
        auto dest = iVariant();
        std::size_t i{0}; 
        for ( ; i < size; ++i )
        {
            if ( ref(i).vt == VT_BSTR )
            {
                if ( std::fputws(getBSTR(ref(i)), fp) < 0 )    break;
            }
            else if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
            {
                if ( std::fputws(getBSTR(dest), fp) < 0 )      break;
            }
            if ( std::fputwc(L'\n', fp) < 0 )
                break;
            ::VariantClear(&dest);
        }
        return static_cast<__int32>(i);
    }

    __int32 ofstream_ex(safearrayRef& ref, BSTR fileName, UINT codepage, bool append)
    {
        auto size = ref.getSize(1);
        std::wofstream ofs{fileName, 
                            append? (std::ios_base::out | std::ios_base::app):
                                    (std::ios_base::out | std::ios_base::trunc)};
        ofs.imbue(std::locale("Japanese", LC_CTYPE));
        auto dest = iVariant();
        std::size_t i{0}; 
        for ( ; i < size; ++i )
        {
            if ( ref(i).vt == VT_BSTR )
                ofs << getBSTR(ref(i));
            else if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
                ofs << getBSTR(dest);
            ofs << L'\n';
            ::VariantClear(&dest);
        }
        return static_cast<__int32>(i);
    }

}
