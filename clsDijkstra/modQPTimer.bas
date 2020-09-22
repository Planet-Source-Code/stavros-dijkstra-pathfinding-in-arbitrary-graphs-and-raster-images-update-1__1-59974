Attribute VB_Name = "modQPTimer"
'Functions used for timing purposes.

Option Explicit

Private cuStart As Currency
Private cuFreq  As Currency

Private Declare Function QueryPerformanceCounter Lib "kernel32" (cuPerfCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (cuFrequency As Currency) As Long

Public Function SetRTime() As Long
    If QueryPerformanceFrequency(cuFreq) Then
        SetRTime = 1
    End If
    QueryPerformanceCounter cuStart
End Function

Public Function GetRTime() As Double
Dim cuStop As Currency
    
    QueryPerformanceCounter cuStop
    If cuFreq Then
        GetRTime = (cuStop - cuStart) / cuFreq
    Else
        GetRTime = -1
    End If
End Function
