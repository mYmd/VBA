//VBA_NestFunc.hpp
//Copyright (c) 2015 mmYYmmdd
#include "OAIdl.h"      //wtypes.h
#include <memory>
#include <array>
#include <string>

#if _MSC_VER < 1900
#define noexcept throw()
#endif

//VBA�z��̎����擾
__int32 __stdcall   Dimension(VARIANT const& v) noexcept;

//�v���[�X�z���_�̐���
VARIANT __stdcall   placeholder(__int32) noexcept;

//�v���[�X�z���_�̔���
__int32 __stdcall   is_placeholder(VARIANT const& v) noexcept;

//bind����Ă��Ȃ�VBA�֐���2�����ŌĂяo��
VARIANT __stdcall   unbind_invoke(VARIANT const& bfun, VARIANT& param1, VARIANT& param2) noexcept;

//--------------------------------------------------------
// VARIANT�\���̂̍쐬
VARIANT iVariant(VARTYPE t = VT_EMPTY) noexcept;

// VARIANT�\���̂���BSTR�̎擾
BSTR getBSTR(VARIANT const& expr) noexcept;

// BSTR�^VARIANT�\���̂̍쐬
VARIANT bstrVariant(std::wstring const&) noexcept;

//--------------------------------------------------------
// SafeArray�v�f�̃A�N�Z�X
class safearrayRef {
    SAFEARRAY*      psa;
    VARTYPE         pvt;
    std::size_t     dim;
    std::size_t     elemsize;
    char*           it;
    VARIANT         val_;
    std::array<std::size_t, 3>  size;
public:
    explicit safearrayRef(VARIANT const& v) noexcept;
    ~safearrayRef();
    safearrayRef(safearrayRef const&) = delete;
    safearrayRef(safearrayRef&&) = delete;
    std::size_t getDim() const noexcept;
    std::size_t getSize(std::size_t i) const noexcept;
    std::size_t getOriginalLBound(std::size_t i) const noexcept;
    VARIANT& operator()(std::size_t i, std::size_t j = 0, std::size_t k = 0) noexcept;
};

namespace 
{
    struct SafeArrayUnaccessor {
        void operator()(SAFEARRAY* ptr) const  noexcept
        { ::SafeArrayUnaccessData(ptr); }
    };
    using safearrayRAII = std::unique_ptr<SAFEARRAY, SafeArrayUnaccessor>;

    struct swap_v_t {
        void operator ()(VARIANT& a, VARIANT& b) const noexcept { std::swap(a, b); }
        void operator ()(VARIANT& a, VARIANT&& b) const noexcept { std::swap(a, b); }
    };
}

// �E�Ӓl�R���e�i�����VARIANT�z�񐶐�
template <typename Container_t, typename F>
VARIANT vec2VArray(Container_t&& cont, F&& trans) noexcept
{
    static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter must be a rvalue reference !!");
    SAFEARRAYBOUND rgb = { static_cast<ULONG>(cont.size()), 0 };
    safearrayRAII pArray{::SafeArrayCreate(VT_VARIANT, 1, &rgb)};
    char* it = nullptr;
    ::SafeArrayAccessData(pArray.get(), reinterpret_cast<void**>(&it));
    if (!it)            return iVariant();
    auto const elemsize = ::SafeArrayGetElemsize(pArray.get());
    std::size_t i{0};
    swap_v_t swap_v;
    try
    {
        for (auto p = cont.begin(); p != cont.end(); ++p, ++i)
            swap_v(*reinterpret_cast<VARIANT*>(it + i * elemsize), std::forward<F>(trans)(*p));
        auto ret = iVariant(VT_ARRAY | VT_VARIANT);
        ret.parray = pArray.get();
        try { cont.clear(); }
        catch (...) {}
        return ret;
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

// �E�Ӓl�R���e�i<VARIANT>����� VARIANT�z�񐶐�
template <typename Container_t>
VARIANT vec2VArray(Container_t&& cont) noexcept
{
    static_assert(!std::is_reference<Container_t>::value, "vec2VArray's parameter must be a rvalue reference !!");
    auto trans = [](typename Container_t::reference x) -> typename Container_t::reference   {
        return x;
    };
    return vec2VArray(std::move(cont), trans);
}

//--------------------------------------------------------
//�֐��̃C���^�t�F�[�X
class funcExpr_i {
public:
    virtual ~funcExpr_i() = default;
    virtual VARIANT& eval(VARIANT&, VARIANT&, int left_right) noexcept = 0;
};

//--------------------------------------------------------
//�g�p����B���VBA�R�[���o�b�N�֐��^  vbCallbackFunc_t
//VBA�ɂ�����V�O�l�`����
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// ��������
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
using vbCallbackFunc_t = VARIANT(__stdcall *)(VARIANT&, VARIANT&);

namespace {
    struct VBCallbackStruct;
}

//VBA�֐��̕\��
class functionExpr : public funcExpr_i {
    vbCallbackFunc_t    fun;
    VARIANT             val;
    std::unique_ptr<funcExpr_i> left;
    std::unique_ptr<funcExpr_i> right;
    functionExpr(functionExpr const&) = delete;
    functionExpr(functionExpr&&) = delete;
public:
    explicit functionExpr(VARIANT const&) noexcept;
    explicit functionExpr(const VBCallbackStruct&) noexcept;
    ~functionExpr();
    VARIANT& eval(VARIANT&, VARIANT&, int left_right = 0) noexcept override;
    bool isValid() const noexcept;
};

//�֐��̈����Ƃ��Ă̊֐��i�֐������ł͂Ȃ��j
class innerFunction : public funcExpr_i {
    VARIANT             myVal;
    VARIANT&            val;
    int                 phn1;
    int                 phn2;
    std::unique_ptr<innerFunction>  arg1;
    std::unique_ptr<innerFunction>  arg2;
    innerFunction(innerFunction const&) = delete;
    innerFunction(innerFunction&&) = delete;
    void eval_imple(VARIANT&, VARIANT&, VARIANT&, int, int) noexcept;
public:
    innerFunction(VARIANT&, bool) noexcept;
    ~innerFunction();
    VARIANT& eval(VARIANT&, VARIANT&, int left_right) noexcept override;
};
