//fileUtil.cpp
//Copyright (c) 2018 mmYYmmdd
#include "stdafx.h"
#include <fstream>
#include <codecvt>
#include <vector>
#include "VBA_NestFunc.hpp"

namespace   {
    template <typename Fn, typename T>
    __int32 fgets_ex(Fn&&, BSTR, UINT, __int32, __int32, T);

    template <typename Fn>
    __int32 ifstream_ex(Fn&&, BSTR, UINT, __int32, __int32);

    template <typename T>
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
    if ( codepage == 1200 || codepage == 65001 )        // UTF-16LE, UTF-8
        fgets_ex(callback, fileName, codepage, head_n, head_cut, L'a');
    else if ( codepage == 1252 || codepage == 932 )     // ANSI, SHIFT-JIS
        fgets_ex(callback, fileName, codepage, head_n, head_cut, 'a');
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
    if ( codepage == 1200 || codepage == 65001 )    // UTF-16LE, UTF-8
        return fgets_ex(callback, fileName, codepage, head_n, head_cut, L'a');
    else if ( codepage == 1252 || codepage == 932 ) // ANSI, SHIFT-JIS
        return fgets_ex(callback, fileName, codepage, head_n, head_cut, 'a');
    else
        return ifstream_ex(callback, fileName, codepage, head_n, head_cut);
}

//VBA配列をテキストファイルに書き出す(UTF-8, UTF-16, ANSI, S-JIS, ...)
__int32 __stdcall
array2textfile(VARIANT const& array, VARIANT const& fileName_, __int32 codepage_, __int8 append)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )            return 0;
    safearrayRef ref{array};
    if ( ref.getDim() != 1 )    return 0;
    auto codepage = static_cast<UINT>(codepage_);
    if ( codepage == 1200 || codepage == 65001 )    // UTF-16LE, UTF-8
        return fputws_ex<wchar_t>(ref, fileName, codepage, append != 0);
    else if ( codepage == 1252 || codepage == 932 ) // ANSI, SHIFT-JIS
        return fputws_ex<char>(ref, fileName, codepage, append != 0);
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
    
    char* std_fgets(char* Buffer, int BufferCount, FILE* Strm) noexcept
    {   return std::fgets(Buffer, BufferCount, Strm);   }

    wchar_t* std_fgets(wchar_t* Buffer, int BufferCount, FILE* Strm) noexcept
    {   return std::fgetws(Buffer, BufferCount, Strm);    }

    template <typename T>
    std::basic_string<T> fgets_ex_line(FILE* fp, std::size_t n, bool& eof);

    std::wstring const& MultiByteToWideChar_if(std::wstring const& w, UINT)
    {   return w;   }

    std::wstring MultiByteToWideChar_if(std::string const& s, UINT);

    //UTF-16LE(codepage:1200) or UTF-8(codepage:65001)
    //https://msdn.microsoft.com/en-us/library/ee719641.aspx

    template <typename Fn, typename T>
    __int32 fgets_ex(Fn&&      func    ,
                     BSTR      fileName,
                     UINT      codepage,
                     __int32   head_n  ,
                     __int32   head_cut,
                     T         dummy    )
    {
        FILE* fp = nullptr;
        auto err = ::_wfopen_s(&fp,
                               fileName, 
                               (codepage==1200)?    L"rt, ccs=UTF-16LE":
                               (codepage==65001)?   L"rt, ccs=UTF-8":
                               L"rt");     //ANSI, SHIFT-JIS
        fileCloseRAII fc_tmp(err? nullptr: fp);
        if ( err || !fp )       return 0;
        std::basic_string<T> str(256, T{'\0'});
        __int32 count{0}, ret{0};
        while ( count < head_cut )
        {
            auto eof = true;
            str = fgets_ex_line<T>(fp, str.capacity(), eof);
            if ( eof )      break;
            ++count;
        }
        while ( count < head_n )
        {
            auto eof = true;
            str = fgets_ex_line<T>(fp, str.capacity(), eof);
            if ( eof )      break;
            std::forward<Fn>(func)(MultiByteToWideChar_if(str, codepage));
            ++count; ++ret;
        }
        return ret;
    }
    //https://msdn.microsoft.com/ja-jp/library/z5hh6ee9.aspx
    //https://msdn.microsoft.com/ja-jp/magazine/mt763237.aspx
    //https://msdn.microsoft.com/ja-jp/library/c565h7xx.aspx

    //改行が出てくるまでバッファを伸ばす
    template <typename T>
    std::basic_string<T> fgets_ex_line(FILE* fp, std::size_t n, bool& eof)
    {
        std::basic_string<T> buf(n, T{'\0'});
        auto p = std_fgets(&buf[0], static_cast<int>(buf.size()+1), fp);
        if ( p )
        {
            auto len = std::char_traits<T>::length(p);
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
                    return  buf += fgets_ex_line<T>(fp, n*2, eof);
                }
            }
        }
        buf.clear();
        return buf;
    }

    std::wstring MultiByteToWideChar_if(std::string const& s, UINT codepage)
    {
        std::wstring w;
        auto len = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, s.data(), -1, nullptr, 0);
        w.resize(len? len-1: 0, L'\0');
        auto re = ::MultiByteToWideChar(codepage, MB_ERR_INVALID_CHARS, s.data(), -1, &w[0], len);
        if ( !re )  w.clear();
        return w;
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
        while ( count < head_n && std::getline(ifs, str) )
        {        
            std::forward<Fn>(func)(MultiByteToWideChar_if(str, codepage));
            ++count; ++ret;
        }
        return ret;
    }

    //---------------------------------------------------------

    BSTR WideCharToMultiByte_if(BSTR str, std::wstring&, UINT)
    {   return str;    }

    char const* WideCharToMultiByte_if(BSTR, std::string&, UINT); 

    int std_fputs(char const* buffer, FILE* Strm) noexcept
    {   return std::fputs(buffer, Strm);   }

    int std_fputc(int c, FILE* Strm) noexcept
    {   return std::fputc(c, Strm);   }

    int std_fputs(wchar_t const* buffer, FILE* Strm) noexcept
    {   return std::fputws(buffer, Strm);    }

    int std_fputc(wchar_t c, FILE* Strm) noexcept
    {   return std::fputwc(c, Strm);   }

    template <typename T>
    __int32 fputws_ex(safearrayRef& ref, BSTR fileName, UINT codepage, bool append)
    {
        FILE* fp = nullptr;
        auto openmode = std::wstring(append? L"a+t": L"wt");
        if ( codepage==1200 )   openmode += L", ccs=UTF-16LE";
        if ( codepage==65001 )  openmode += L", ccs=UTF-8";
        auto err = ::_wfopen_s(&fp, fileName,  openmode.data());
        fileCloseRAII fc_tmp(err? nullptr: fp);
        if ( err || !fp )       return 0;
        auto size = ref.getSize(1);
        auto dest = iVariant();
        std::size_t i{0}; 
        std::basic_string<T> buf;
        for ( ; i < size; ++i )
        {
            BSTR pp{nullptr};
            if ( ref(i).vt == VT_BSTR )
                pp = getBSTR(ref(i));
            else if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
                pp = getBSTR(dest);
            //
            if ( pp && std_fputs(WideCharToMultiByte_if(pp, buf, codepage), fp) < 0 )  break;
            if ( std_fputc(T{'\n'}, fp) < 0 )       break;
            ::VariantClear(&dest);
        }
        return static_cast<__int32>(i);
    }

    __int32 ofstream_ex(safearrayRef& ref, BSTR fileName, UINT codepage, bool append)
    {
        auto size = ref.getSize(1);
        std::ofstream ofs{fileName, 
                append? (std::ios_base::out | std::ios_base::app):
                (std::ios_base::out | std::ios_base::trunc)};
        auto dest = iVariant();
        std::size_t i{0};
        std::string buf;
        for ( ; i < size; ++i )
        {
            BSTR p{nullptr};
            if ( ref(i).vt == VT_BSTR )
                p = getBSTR(ref(i));
            else if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
                p = getBSTR(dest);
            if ( p )
                ofs << WideCharToMultiByte_if(p, buf, codepage);
            ::VariantClear(&dest);
            if (ofs.fail())     break;
            ofs << L'\n';
        }
        return static_cast<__int32>(i);
    }

    char const* WideCharToMultiByte_if(BSTR p, std::string& buf, UINT codepage) 
    {
        auto b = ::WideCharToMultiByte(codepage, 0, p, -1, nullptr, 0, nullptr, nullptr);
        buf.resize(b);
        b = ::WideCharToMultiByte(codepage, 0, p, -1, &buf[0], b, nullptr, nullptr); 
        return buf.data();
    }

}   //namespace {
