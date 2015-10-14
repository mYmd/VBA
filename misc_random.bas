Attribute VB_Name = "misc_random"
'misc_random
'Copyright (c) 2015 mmYYmmdd
Option Explicit
                                                
Declare PtrSafe Function seed_Engine Lib "mapM.dll" _
                                    (Optional ByRef seedN As Variant) As Long

Declare PtrSafe Function uniform_int_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal fromN As Long, _
                            ByVal toN As Long) As Variant

Declare PtrSafe Function uniform_real_dist Lib "mapM.dll" _
                                        (ByVal N As Long, _
                                    ByVal fromD As Double, _
                                ByVal toD As Double) As Variant

Declare PtrSafe Function normal_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal mean As Double, _
                            ByVal stddev As Double) As Variant
            
Declare PtrSafe Function bernoulli_dist Lib "mapM.dll" _
                                    (ByVal N As Long, _
                                ByVal prob As Double) As Variant

