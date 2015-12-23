//VBA_NestFunc.hpp
//Copyright (c) 2015 mmYYmmdd
#include "OAIdl.h"      //wtypes.h
#include <memory>
#include <array>

//VBA配列の次元取得
__int32 __stdcall   Dimension(const VARIANT* pv) noexcept;

//プレースホルダの生成
VARIANT __stdcall   placeholder(__int32) noexcept;

//プレースホルダの判定
__int32 __stdcall   is_placeholder(const VARIANT* pv) noexcept;

//bindされていないVBA関数を2引数で呼び出す
VARIANT __stdcall   unbind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2) noexcept;

//--------------------------------------------------------
class safearrayRef  {
    SAFEARRAY*      psa; 
    VARTYPE         pvt;
    std::size_t     dim;
    std::size_t     elemsize;
    char*           it;
    VARIANT         val_;
    std::array<std::size_t, 3>  size;
public:
    explicit safearrayRef(const VARIANT* pv) noexcept;
    ~safearrayRef();
    safearrayRef(safearrayRef const&) = delete;
    safearrayRef(safearrayRef&&) = delete;
    std::size_t getDim() const noexcept;
    std::size_t getSize(std::size_t i) const noexcept;
    std::size_t getOriginalLBound(std::size_t i) const noexcept;
    VARIANT& operator()(std::size_t i, std::size_t j = 0, std::size_t k = 0) noexcept;
};

//--------------------------------------------------------
//関数のインタフェース
class funcExpr_i    {
public:
    virtual ~funcExpr_i() = default;
    virtual bool isYielder() const noexcept;
    virtual VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) = 0;
};

//--------------------------------------------------------
//使用する唯一のVBAコールバック関数型  vbCallbackFunc_t
//VBAにおけるシグネチャは
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// もしくは
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
using vbCallbackFunc_t = VARIANT (__stdcall * )(VARIANT*, VARIANT*);

namespace{
    struct VBCallbackStruct;
}
//
class functionExpr : public funcExpr_i    {
    vbCallbackFunc_t    fun;
    VARIANT             val;
    std::unique_ptr<funcExpr_i> left;
    std::unique_ptr<funcExpr_i> right;
public:
    explicit functionExpr(const VARIANT*) noexcept;
    explicit functionExpr(const VBCallbackStruct&) noexcept;
    functionExpr(functionExpr const&) = delete;
    functionExpr(functionExpr&&) = delete;
    ~functionExpr();
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) noexcept;
    bool isValid() const noexcept;
};
