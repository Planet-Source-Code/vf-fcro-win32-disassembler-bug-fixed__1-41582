Attribute VB_Name = "ModuleASM"
'***************************
'OPCODE OFFSET STRUCTURE W32
'***************************

'(0-3f) OFFSET
' If XOO = 5 Then  'X1
'    ElseIf XOO = 4 Then  'X2
'        If (YRG Mod 8&) = 4 Then 'Y1
'            If YOO = 5 Then
'            Else
'            End If
'        ElseIf YOO = 5 Then 'Y2
'        Else 'Y3
'        End If
'    Else  'X3
'    End If

'(40-7f) OFFSET
'    If XOO = 4 Then
'       If (YRG Mod 8&) = 4 Then
'       Else
'       End If
'    Else
'    End If

'(80-cf) OFFSET
'   If XOO = 4 Then
'       If (YRG Mod 8&) = 4 Then
'       Else
'       End If
'   Else
'   End If

Public REG16O1() As String
Public SEGOFFSET() As String
Public REGOFFSET4() As String
Public REGOFFSET2() As String
Public REGOFFSET1() As String
Public MATHOFFSET() As String
Public BITOFFSET() As String
Public FLOATOFFSET() As String
Public INTFLOATOFFSET() As String
Public FLOATSTACK() As String
Public FLOATOP1() As String
Public FLOATOP2() As String
Public FLOATOP3() As String
Public FLOATOP4() As String
Public FLOATOP5() As String
Public FLOATOP6() As String
Public FLOATOP7() As String
Public FLOATOP8() As String
Public FLOATOP9() As String
Public FLOATOP10() As String
Public MATH2() As String
Public CFLX() As String
Public CMPS() As String

Public LOOPX() As String
Public JXX() As String
Public SETXX() As String

Public IDCJP() As String
Public MMI() As String
Public MMI2() As String
Public XMM() As String
Public MMX() As String

'************Constants!
Public Const NOL As String = " ["
Public Const NOR As String = ",["

Public Const DPL As String = " DWORD PTR ["
Public Const DPR As String = ",DWORD PTR ["
Public Const WPL As String = " WORD PTR ["
Public Const WPR As String = ",WORD PTR ["
Public Const BPL As String = " BYTE PTR ["
Public Const BPR As String = ",BYTE PTR ["
Public Const QPL As String = " QWORD PTR ["
Public Const QPR As String = ",QWORD PTR ["
Public Const MMPL As String = " MMWORD PTR ["
Public Const MMPR As String = ",MMWORD PTR ["
Public Const XMMPL As String = " XMMWORD PTR ["
Public Const XMMPR As String = ",XMMWORD PTR ["
Public Const FPL As String = " FWORD PTR ["
Public Const FPR As String = ",FWORD PTR ["
Public Const TBPL As String = " TBYTE PTR ["
Public Const TBPR As String = ",TBYTE PTR ["

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Function SByteToStr(ByVal BYTEEX As Byte) As String
If BYTEEX = 0 Then Exit Function
If BYTEEX > &H7F Then
BYTEEX = 256 - BYTEEX
SByteToStr = "-" & Hex(BYTEEX)
Else
SByteToStr = "+" & Hex(BYTEEX)
End If
End Function
Public Function BByteToStr(BYTEEX As Byte) As String
BByteToStr = Left(Hex(BYTEEX And &HF0), 1) & Hex(BYTEEX And &HF)
End Function
Public Function ByteToStr(BYTEEX As Byte) As String
ByteToStr = Hex(BYTEEX)
End Function
Public Function WordToStr(WORDEX As Integer) As String
WordToStr = Hex(WORDEX)
End Function
Public Function SWordToStr(WORDEX As Integer) As String
'SIGNED WORD
If WORDEX = 0 Then Exit Function
Dim X As Integer
SWordToStr = Space(5)
If WORDEX < 0 Then
X = &HFFFF - WORDEX + 1
Mid(SWordToStr, 1, 1) = "-"
Else
Mid(SWordToStr, 1, 1) = "+"
End If
Mid(SWordToStr, 2, 1) = Left(Hex(X And &HF000&), 1)
Mid(SWordToStr, 3, 1) = Left(Hex(X And &HF00&), 1)
Mid(SWordToStr, 4, 1) = Left(Hex(X And &HF0&), 1)
Mid(SWordToStr, 5, 1) = Hex(X And &HF&)
End Function
Public Function SDwordToStr(DWORDEX As Long) As String
'SIGNED DWORD
If DWORDEX = 0 Then Exit Function
If DWORDEX < 0 Then
If Not DWORDEX = &H80000000 Then
DWORDEX = Abs(DWORDEX)
End If
SDwordToStr = "-" & Hex(DWORDEX)
Else
DWORDEX = Abs(DWORDEX)
SDwordToStr = "+" & Hex(DWORDEX)
End If
End Function
Public Function DwordToStr(DWORDEX As Long) As String
DwordToStr = Hex(DWORDEX)
Exit Function
End Function
Public Function GetWordFromList(DATA() As Byte, count As Long) As Integer
CopyMemory GetWordFromList, DATA(count), 2
End Function
Public Function GetDWordFromList(DATA() As Byte, count As Long) As Long
CopyMemory GetDWordFromList, DATA(count), 4
End Function
Public Sub RWDBDump(DATA() As Byte, ByRef count As Long, ByRef VBASE As Long, ByRef size As Byte, ByRef CMD As String)
Dim u As Byte
CMD = DwordToStr(VBASE + count) & " " & "db"
For u = 0 To size - 1
CMD = CMD & " " & ByteToStr(DATA(count + u)) & "h"
If u > 0 Then CMD = CMD & " "
Next u
End Sub

Public Sub RWDump(DATA() As Byte, ByRef count As Long, ByRef VBASE As Long, ByRef size As Byte, ByRef CMD As String)
Dim u As Byte
Dim SSt As String
For u = 0 To size - 1
SSt = SSt & BByteToStr(DATA(count + u)) & " "
Next u
CMD = SSt & vbTab & CMD
End Sub

Public Sub LJoin(ByRef leftPTRConst As String, ByRef DisAssemble As String, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef ret As Byte)
If ret = 1 Then
DisAssemble = COMMD & " " & X1 & "," & X3
Else
DisAssemble = COMMD & leftPTRConst & X1 & X2 & "]," & X3
End If
End Sub
Public Sub RJoin(ByRef RightPTRConst As String, ByRef DisAssemble As String, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef ret As Byte)
If ret = 1 Then
DisAssemble = COMMD & " " & X3 & "," & X1
Else
DisAssemble = COMMD & " " & X3 & RightPTRConst & X1 & X2 & "]"
End If
End Sub
Public Sub MJoin(ByRef leftPTRConst As String, ByRef DisAssemble As String, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef ret As Byte)
If ret = 1 Then
DisAssemble = COMMD & " " & X1
Else
DisAssemble = COMMD & leftPTRConst & X1 & X2 & "]"
End If
End Sub
Public Sub TJoin(ByRef RightPTRConst As String, ByRef DisAssemble As String, ByRef COMMD As String, ByRef PRG As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef ret As Byte)
If ret = 1 Then
DisAssemble = COMMD & " " & PRG & "," & X1 & "," & X3
Else
DisAssemble = COMMD & " " & PRG & RightPTRConst & X1 & X2 & "]," & X3
End If
End Sub
Public Sub TJoin2(ByRef leftPTRConst As String, ByRef DisAssemble As String, ByRef COMMD As String, ByRef PRG As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef ret As Byte)
If ret = 1 Then
DisAssemble = COMMD & " " & X1 & "," & PRG & "," & X3
Else
DisAssemble = COMMD & leftPTRConst & X1 & X2 & "]," & PRG & "," & X3
End If
End Sub

Public Function CalcShortJump(DATA As Byte, ByRef JXS As String, ByRef Start As Long, ByRef ActualAdr As Long) As String
Dim BTemp As Long
If DATA >= &H80 Then
BTemp = (ActualAdr + DATA + 2) - 256&
Else
BTemp = ActualAdr + 2 + DATA
End If
CalcShortJump = JXS & " " & DwordToStr(BTemp)
End Function
Public Function CalcLongJump(DATA() As Byte, ByRef JXS As String, ByRef Start As Long, ByRef ActualAdr As Long) As String
CalcLongJump = JXS & " " & DwordToStr(GetDWordFromList(DATA, Start + 1) + 5 + ActualAdr - &HFFFFFFFF)
End Function
Public Function Check0(DATA() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef CL As Byte) As Byte
Dim S As Long
Dim OUTF As Long
OUTF = UBound(DATA) + 1
Do
S = S + 1
If Start + S = OUTF Then Exit Do
Loop While DATA(Start + S) = 0 And S < 16
If S > 2 Then 'Possible 0 code!
COMMD = "BYTE 0 DUP(" & S & ")"
CL = S
Check0 = 1
End If

End Function
