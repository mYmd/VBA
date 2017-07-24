//VBA_Random.cpp
//Copyright (c) 2015 mmYYmmdd
#include "stdafx.h"
#include <random>
#include "VBA_NestFunc.hpp"

namespace{
    std::random_device seed_gen;
    std::default_random_engine d_engine{seed_gen()};
}

// ���� seed ���w�肵�ă����_�}�C�Y
// seed�ȗ����������͐����Ƃ��ĕ]���ł��Ȃ��Ƃ��� seed_gen �ɂ��
__int32 __stdcall seed_Engine(VARIANT* seed) noexcept
{
    auto tmp = iVariant();
    try
    {
        if (seed && S_OK == ::VariantChangeType(&tmp, seed, 0, VT_I4))
        {
            d_engine.seed(static_cast<std::default_random_engine::result_type>(tmp.lVal));
            return tmp.lVal;
        }
        else
        {
            auto sd = seed_gen();
            d_engine.seed(sd);
            return static_cast<__int32>(sd);
        }
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

namespace   {
    //���ʃT�u���[�`��
    template <typename DIST, typename F>
    VARIANT dist_imple(__int32 N, DIST dist, F fun)
    {
        auto ret = iVariant();
        if ( N < 1 )
        {
            fun(ret, dist(d_engine));
        }
        else
        {
            ret.vt = VT_ARRAY | VT_VARIANT;
            SAFEARRAYBOUND Bound = { static_cast<ULONG>(N), 0 };
            ret.parray = ::SafeArrayCreate(VT_VARIANT, 1, &Bound);
            safearrayRef arOut{&ret};
            for ( int i = 0; i < N; ++i )
            {
                ::VariantInit(&arOut(i));
                fun(arOut(i), dist(d_engine));
            }
        }
        return ret;
    }
}

// N �̈�l���������𐶐�  �͈�[from, to]���w�� 
VARIANT __stdcall uniform_int_dist(__int32 N, __int32 from, __int32 to) noexcept
{
    try
    {
        return
            dist_imple(N,
                       std::uniform_int_distribution<__int32>{from, to},
                       [](VARIANT& v, __int32 i) { v.vt = VT_I4; v.lVal = i; }
        );
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

// N �̈�l�������𐶐�  �͈�[from, to]���w��
VARIANT __stdcall uniform_real_dist(__int32 N, double from, double to) noexcept
{
    try
    {
        return
            dist_imple(N,
                       std::uniform_real_distribution<>{from, to},
                       [](VARIANT& v, double d) { v.vt = VT_R8; v.dblVal = d; }
        );
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

// N �̐��K���z�������𐶐�  ���� mean �ƕW���΍� stddev ���w��
VARIANT __stdcall normal_dist(__int32 N, double mean, double stddev) noexcept
{
    try
    {
        return
            dist_imple(N,
                       std::normal_distribution<>{mean, stddev},
                       [](VARIANT& v, double d) { v.vt = VT_R8; v.dblVal = d; }
        );
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}

// N �̃x���k�[�C���z�����𐶐�  �m�� prob ���w��
VARIANT __stdcall bernoulli_dist(__int32 N, double prob) noexcept
{
    try
    {
        return
            dist_imple( N,
                        std::bernoulli_distribution{prob},
                        [](VARIANT& v, bool b) { v.vt = VT_I4; v.lVal = b ? 1 : 0; }
                      );
    }
    catch (const std::exception&)
    {
        return iVariant();
    }
}
