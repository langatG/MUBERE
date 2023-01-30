Attribute VB_Name = "modsecurity"
Option Explicit
'// - Created by the PIUS KIPROTICH SIGEI

Function Decript_String(str As String) As String
    On Error Resume Next
    Dim StrLn As Integer, j As Integer, Ch As String, ChVal As Integer
    Dim EncptStr As String
    StrLn = Len(str)
    For j = 1 To StrLn
        Ch = Mid(str, j, 1)
        ChVal = Asc(Ch)
        EncptStr = EncptStr & Chr(Decript_Char_Value(ChVal))
    Next j
    Decript_String = EncptStr
End Function
Function Encript_String(str As String) As String
    On Error Resume Next
    Dim StrLn As Integer, j As Integer, Ch As String, ChVal As Integer
    Dim EncptStr As String
    StrLn = Len(str)
    For j = 1 To StrLn
        Ch = Mid(str, j, 1)
        ChVal = Asc(Ch)
        EncptStr = EncptStr & Chr(Encript_Char_Value(ChVal))
    Next j
    Encript_String = EncptStr
End Function
Function Encript_Char_Value(CharVal As Integer) As Integer
    Dim C As Integer, MaxCharval As Integer, Encpt As Integer
    C = 32
    MaxCharval = 128
    Encpt = MaxCharval - CharVal + C
    Encript_Char_Value = Encpt
End Function
Function Decript_Char_Value(Encpt As Integer) As Integer
    Dim MaxCharval As Integer, C As Integer
    C = 32
    MaxCharval = 128
    Decript_Char_Value = MaxCharval - Encpt + C
End Function

