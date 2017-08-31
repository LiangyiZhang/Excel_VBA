'====================================
'Coding by lyzhang.me @Sidneyzhang
'Compliance with The MIT License
'Copyright 2017 Sidneyzhang
'====================================

Public Testing As Boolean
#If VBA7 And Win64 Then
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Function DelayEx(ms As Long)
    Dim start As Long
    start = GetTickCount
    Do While GetTickCount - start < ms
        If Not Testing Then
        Exit Do
        End If
        DoEvents
    Loop
    StopTest
    MsgBox "考试结束！"
End Function

Sub StartTest()
    Testing = True
    ZhanKai
    Const s As Long = 1000
    DelayEx (s * 3000)
End Sub

Sub StopTest()
    Testing = False
    ZheDie
End Sub

Sub CuoTi()
    If Not Worksheets("SELABAS").Range("B13").Value Then
    MsgBox "请先进行测试！"
    Exit Sub
    End If
    ZhanKai
End Sub

Sub ZheDie()
    Dim hid As Range
    Worksheets("Excel水平测试").Unprotect "10471048"
    For i = 17 To 210
    If Not hid Is Nothing Then
    Set hid = Union(hid, Rows(i))
    Else
    Set hid = Rows(i)
    End If
    Next
    hid.EntireRow.Hidden = True
    Worksheets("Excel水平测试").Protect "10471048"
    Worksheets("SELABAS").Range("B13").Value = True
End Sub

Sub ZhanKai()
    Worksheets("Excel水平测试").Unprotect "10471048"
    Dim hid As Range
    For i = 17 To 210
    If Not hid Is Nothing Then
    Set hid = Union(hid, Rows(i))
    Else
    Set hid = Rows(i)
    End If
    Next
    hid.EntireRow.Hidden = False
    Worksheets("Excel水平测试").Protect "10471048"
End Sub

Sub ChongZhi()
    Worksheets("Excel水平测试").Unprotect "10471048"
    Worksheets("Excel水平测试").Range("H18, H23, H28, H33, H38, H43, H48, H53, H58, H63, H68, H73, H78, H90, H95, F116, F131, D141,D142,D143, F151, F174, D176, F176, H176, D207, E189").Value = ""
    Worksheets("Excel水平测试").Protect "10471048"
    ZheDie
    Worksheets("SELABAS").Range("B13").Value = False
End Sub
