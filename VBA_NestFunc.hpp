//VBA_NestFunc.hpp
//Copyright (c) 2015 mmYYmmdd
#include "OAIdl.h"      //wtypes.h
#include <memory>
#include <array>

#if _MSC_VER < 1900
#define noexcept throw()
#endif

//VBA�z��̎����擾
__int32 __stdcall   Dimension(const VARIANT* pv) noexcept;

//�v���[�X�z���_�̐���
VARIANT __stdcall   placeholder(__int32) noexcept;

//�v���[�X�z���_�̔���
__int32 __stdcall   is_placeholder(const VARIANT* pv) noexcept;

//bind����Ă��Ȃ�VBA�֐���2�����ŌĂяo��
VARIANT __stdcall   unbind_invoke(VARIANT* bfun, VARIANT* param1, VARIANT* param2) noexcept;

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
//�֐��̃C���^�t�F�[�X
class funcExpr_i {
public:
    virtual ~funcExpr_i() = default;
    virtual VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) noexcept = 0;
};

//--------------------------------------------------------
//�g�p����B���VBA�R�[���o�b�N�֐��^  vbCallbackFunc_t
//VBA�ɂ�����V�O�l�`����
// Function fun(ByRef elem As Variant, ByRef dummy As Variant) As Variant
// ��������
// Function fun(ByRef elem As Variant, Optional ByRef dummy As Variant) As Variant
using vbCallbackFunc_t = VARIANT(__stdcall *)(VARIANT*, VARIANT*);

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
    explicit functionExpr(const VARIANT*) noexcept;
    explicit functionExpr(const VBCallbackStruct&) noexcept;
    ~functionExpr();
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) noexcept override;
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
    void eval_imple(VARIANT*, VARIANT*, VARIANT*, int, int) noexcept;
public:
    innerFunction(VARIANT*, bool) noexcept;
    ~innerFunction();
    VARIANT* eval(VARIANT*, VARIANT*, int left_right = 0) noexcept override;
};
