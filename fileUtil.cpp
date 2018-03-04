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

    template <typename PUTS, typename PUTC>
    __int32 fputws_ex(safearrayRef&, BSTR, wchar_t const*, PUTS&&, PUTC&&, char const*);

    char const* WideCharToMultiByte_b(BSTR, std::string&, UINT); 
}

//**************************************************

//テキストファイルをVBA配列にする
VARIANT __stdcall
textfile2array(VARIANT const& fileName_, __int32 Code_Page, __int32 head_n, __int32 head_cut, __int8 toArry)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )        return iVariant();
    auto codepage = static_cast<UINT>(0 <= Code_Page? Code_Page: -Code_Page);
    if ( head_n < 0 )   head_n = (std::numeric_limits<int32_t>::max)();
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
textfile_for_each(VARIANT const& Fn, VARIANT const& fileName_, __int32 Code_Page, __int32 head_n, __int32 head_cut)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )                return 0;
    functionExpr vbaFunc{Fn};
    bool validFunc = vbaFunc.isValid();
    auto codepage = static_cast<UINT>(0 <= Code_Page? Code_Page : -Code_Page);
    if ( head_n < 0 )   head_n = (std::numeric_limits<int32_t>::max)();
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
array2textfile(VARIANT const& array, VARIANT const& fileName_, __int32 Code_Page, __int8 append, __int8 CrLf)
{
    auto const fileName = getBSTR(fileName_);
    if ( !fileName )            return 0;
    safearrayRef ref{array};
    if ( ref.getDim() != 1 )    return 0;
    auto codepage = static_cast<UINT>((0 < Code_Page)? Code_Page: -Code_Page);
    //      UTF-16LE
    if ( codepage == 1200 )
    {
        auto openmode = append? L"a+t, ccs=UTF-16LE": L"wt, ccs=UTF-16LE";
        auto puts = [](BSTR pp, FILE* fp) { return std::fputws(pp, fp); };
        auto putc = [](FILE* fp) { return std::fputwc(L'\n', fp); };
        return fputws_ex(ref, fileName, openmode, puts, putc, nullptr);
    }
    //              UTF-8       ,           ANSI    ,       SHIFT-JIS
    else if ( codepage == 65001 || codepage == 1252 || codepage == 932 )
    {
        auto openmode = append? L"a+b": L"wb";
        char bom[] = {char(0xEF), char(0xBB), char(0xBF), '\0' };
        char const* BOM = (Code_Page==65001)? bom: nullptr;
        std::string buf;
        auto puts = [codepage, &buf](BSTR pp, FILE* fp) {
            return std::fputs(WideCharToMultiByte_b(pp, buf, codepage), fp);
        };
        auto putc = [CrLf](FILE* fp) {
            int ret = (CrLf==0)? 0: std::fputc('\r', fp);
            return (ret < 0)? ret: std::fputc('\n', fp);
        };
        return fputws_ex(ref, fileName, openmode, puts, putc, BOM);
    }
    else
        return 0;
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
    //std::basic_string<T> fgets_ex_line(FILE* fp, std::size_t n, bool& eof);
    bool fgets_ex_line(FILE* fp, std::basic_string<T>& buf);

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
            if ( !fgets_ex_line<T>(fp, str) )   break;
            ++count;
        }
        while ( count < head_n )
        {
            if ( !fgets_ex_line<T>(fp, str) )   break;
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
    std::basic_string<T> fgets_ex_line_imple(FILE*, std::size_t);

    template <typename T>
    bool fgets_ex_line(FILE* fp, std::basic_string<T>& buf)
    {
        buf.resize(buf.capacity());
        auto p = std_fgets(&buf[0], static_cast<int>(buf.size()+1), fp);
        if ( !p )            return false;
        auto const len = std::char_traits<T>::length(p);
        if ( p[len-1] == T{'\n'} )
            buf.resize(len-1);
        else
            buf += fgets_ex_line_imple<T>(fp, buf.size());
        return true;
    }

    template <typename T>
    std::basic_string<T> fgets_ex_line_imple(FILE* fp, std::size_t n)
    {
        std::basic_string<T> buf(n, T{'\0'});
        auto p = std_fgets(&buf[0], static_cast<int>(buf.size()+1), fp);
        if ( !p )   {
            buf.clear();
        } else {
            auto const len = std::char_traits<T>::length(p);
            if ( p[len-1] == T{'\n'} )
                buf.resize(len-1);
            else
                buf += fgets_ex_line_imple<T>(fp, 2 * buf.size());
        }
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
    template <typename PUTS, typename PUTC>
    __int32 fputws_ex(safearrayRef& ref, BSTR fileName, wchar_t const* openmode, PUTS&& putS, PUTC&& putC, char const* BOM)
    {
        FILE* fp = nullptr;
        auto err = ::_wfopen_s(&fp, fileName,  openmode);
        fileCloseRAII fc_tmp(err? nullptr: fp);
        if ( err || !fp )       return 0;
        auto size = ref.getSize(1);
        auto dest = iVariant();
        std::size_t i{0}; 
        if ( BOM )  std::fputs(BOM, fp);
        for ( ; i < size; ++i )
        {
            BSTR pp{nullptr};
            if ( ref(i).vt == VT_BSTR )
                pp = getBSTR(ref(i));
            else if ( S_OK == ::VariantChangeType(&dest, &ref(i), 0, VT_BSTR) )
                pp = getBSTR(dest);
            //
            if ( pp && std::forward<PUTS>(putS)(pp, fp) < 0 )   break;
            if ( std::forward<PUTC>(putC)(fp) < 0 )             break;
            ::VariantClear(&dest);
        }
        return static_cast<__int32>(i);
    }

    char const* WideCharToMultiByte_b(BSTR p, std::string& buf, UINT codepage) 
    {
        auto b = ::WideCharToMultiByte(codepage, 0, p, -1, nullptr, 0, nullptr, nullptr);
        buf.resize(b);
        b = ::WideCharToMultiByte(codepage, 0, p, -1, &buf[0], b, nullptr, nullptr); 
        return buf.data();
    }

}   //namespace {
