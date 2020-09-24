VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisAsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****W32 OPCODE DISASSEMBLER WRITTEN BY VANJA FUCKAR EMAIL:INGA@VIP.HR
'PENTIUM 3+PENTIUM 4 OPCODE INCLUDED!!!!!!!!!!!!!!!!!!!!!!!!!!



Private VBASE As Long 'VIRTUAL BASE ADDRESS


Public Property Let BaseAddress(ByVal Address As Long)
VBASE = Address
End Property
Public Property Get BaseAddress() As Long
BaseAddress = VBASE
End Property



Public Function DisassembleBack(DATA() As Byte, ByRef Start As Long, ByRef Backward As Byte) As String
Dim Cu As Long
Dim Xu As Long
Dim BCK As Byte
BCK = 24 'Max!
If Start - BCK < 0 Then BCK = Start

Xu = 0
Begs:
Cu = Start - (BCK - Xu)
Do
DisassembleBack = DisAssemble(DATA, Cu, Backward, 1, 1)
Cu = Cu + Backward
Loop While Cu < Start

If Cu = Start Then
Exit Function
ElseIf Xu = BCK Then
DisassembleBack = Hex(Start + ActiveH - 1) & vbTab & "db " & BByteToStr(DATA(Start - 1)): Backward = 1
Exit Function
Else
Xu = Xu + 1
GoTo Begs
End If

'WINCMD
'4012f9

'mijenjaju se adrese Data()
'4858 117
'4859 64
'4860 255
'4861 37

End Function

Public Function DisAssemble(DATA() As Byte, ByVal Start As Long, ByRef Forward As Byte, Optional ByVal HexAdr As Byte = 1, Optional ByVal BinDump As Byte) As String
On Error GoTo EendX:
Dim XRG As Byte
Dim ZRG As Byte

Dim OPT16 As Byte '16-OPCODE
Dim EXTRAADD As Byte
Dim EXTRA As Byte
Dim CL As Byte 'Command Length
Dim DL As Byte 'Data Length

Dim ISWP As Byte

Dim COMMD As String



Dim X1 As String
Dim X2 As String
Dim X3 As String
Dim ret As Byte

'FLOAT************
Dim FWR As Byte
Dim DB2 As Byte
Dim DB3 As Byte
Dim CHK As String
'*****************
Dim CX3 As Byte
Dim MCs As Byte

BEGGIN:

MCs = DATA(Start)

If Start < UBound(DATA) Then
DB2 = DATA(Start + 1) 'SEGMENTS 0-3f,40-7f,80-bf,c0-ff
FWR = Int((DB2 And &H3F) / 8) 'OCTAL SEGMENTS 0-7
End If

Select Case MCs

'MATH ***********************************************+DOBRO
'Case Is = &H0, &H8, &H10, &H18, &H20, &H28, &H30, &H38
'COMMD = MATHOFFSET(Int(MCs / 8))
'GoTo InTTS2

'MATH ***********************************************+DOBRO
Case &H0  ', &H8, &H10, &H18, &H20, &H28, &H30, &H38
If ISWP = 0 Then
If Check0(DATA, Start, DisAssemble, Forward) = 1 Then
GoTo OutOfF
End If
End If

COMMD = MATHOFFSET(Int(MCs / 8))
GoTo InTTS2

Case Is = &H8, &H10, &H18, &H20, &H28, &H30, &H38
COMMD = MATHOFFSET(Int(MCs / 8))
GoTo InTTS2
'**********************************************************


Case Is = &H1, &H9, &H11, &H19, &H21, &H29, &H31, &H39
COMMD = MATHOFFSET(Int(MCs / 8))
GoTo InTTS4

Case Is = &H2, &HA, &H12, &H1A, &H22, &H2A, &H32, &H3A
COMMD = MATHOFFSET(Int(MCs / 8))
GoTo InTTSR2

Case Is = &H3, &HB, &H13, &H1B, &H23, &H2B, &H33, &H3B
COMMD = MATHOFFSET(Int(MCs / 8))
GoTo InTTSR4



Case Is = &H4, &HC, &H14, &H1C, &H24, &H2C, &H34, &H3C
CL = 2
DisAssemble = MATHOFFSET(Int(MCs / 8)) & " AL," & ByteToStr(DB2)

Case Is = &H5, &HD, &H15, &H1D, &H25, &H2D, &H35, &H3D
If ISWP = 1 Then
CL = 3
DisAssemble = MATHOFFSET(Int(MCs / 8)) & " AX," & WordToStr(GetWordFromList(DATA, Start + 1))
Else
CL = 5
DisAssemble = MATHOFFSET(Int(MCs / 8)) & " EAX," & DwordToStr(GetDWordFromList(DATA, Start + 1))
End If


Case &H6
CL = 1
DisAssemble = "PUSH ES"

Case &H7
CL = 1
DisAssemble = "POP ES"

Case &HE
CL = 1
DisAssemble = "PUSH CS"

Case &H16
CL = 1
DisAssemble = "PUSH SS"

Case &H17
CL = 1
DisAssemble = "POP SS"

Case &H1E
CL = 1
DisAssemble = "PUSH DS"

Case &H1F
CL = 1
DisAssemble = "POP DS"

'26:ES ,2E:CS ,36:SS ,3E:DS ,64:FS ,65:GS ,66:?? (SEGMENTS)
Case &H26
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H2E
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H36
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H3E
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H64
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H65
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H66
ISWP = 1 '16 BYTES!
Start = Start + 1
EXTRA = EXTRA + 1
GoTo BEGGIN

Case &H27
CL = 1
DisAssemble = "DAA"

Case &H2F
CL = 1
DisAssemble = "DAS"

Case &H37
CL = 1
DisAssemble = "AAA"

Case &H3F
CL = 1
DisAssemble = "AAS"

Case &H40 To &H47
CL = 1
If ISWP = 1 Then
DisAssemble = "INC " & REGOFFSET2(MCs And 7)
Else
DisAssemble = "INC " & REGOFFSET4(MCs And 7)
End If


Case &H48 To &H4F
CL = 1
If ISWP = 1 Then
DisAssemble = "DEC " & REGOFFSET2(MCs And 7)
Else
DisAssemble = "DEC " & REGOFFSET4(MCs And 7)
End If


Case &H50 To &H57
CL = 1
If ISWP = 1 Then
DisAssemble = "PUSH " & REGOFFSET2(MCs And 7)
Else
DisAssemble = "PUSH " & REGOFFSET4(MCs And 7)
End If


Case &H58 To &H5F
CL = 1
If ISWP = 1 Then
DisAssemble = "POP " & REGOFFSET2(MCs And 7)
Else
DisAssemble = "POP " & REGOFFSET4(MCs And 7)
End If

Case &H60
CL = 1
If ISWP = 1 Then
DisAssemble = "PUSHA"
Else
DisAssemble = "PUSHAD"
End If


Case &H61
CL = 1
If ISWP = 1 Then
DisAssemble = "POPA"
Else
DisAssemble = "POPAD"
End If


Case &H62 '+
If DB2 >= &HC0 Then GoTo EendX
COMMD = "BOUND"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret
End If



Case &H63
COMMD = "ARPL"
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret




'********RESERVED!!!!!!!!!!!!!!!16 BYTES!
Case &H67
Start = Start + 1
EXTRA = EXTRA + 1
OPT16 = 1
GoTo BEGGIN

Case &H68 '+
If ISWP = 1 Then
CL = 3
DisAssemble = "PUSH " & WordToStr(GetWordFromList(DATA, Start + 1))
Else
CL = 5
DisAssemble = "PUSH " & DwordToStr(GetDWordFromList(DATA, Start + 1))
End If


Case &H69 '+
COMMD = "IMUL"
If ISWP = 1 Then
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
TJoin WPR, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
Else
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
TJoin DPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
End If


Case &H6A '+
CL = 2
If DB2 > &H7F Then
If ISWP = 1 Then
DisAssemble = "PUSH " & WordToStr(DB2 Or &HFF00)
Else
DisAssemble = "PUSH " & DwordToStr(DB2 Or &HFFFFFF00)
End If
Else
DisAssemble = "PUSH " & Hex(DB2)
End If

Case &H6B '+
COMMD = "IMUL"
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
'CX3 = CByte(Val("&H" & X3))
CX3 = CByte("&H" & X3)
If CX3 > &H7F Then X3 = WordToStr(CX3 Or &HFF00)
TJoin WPR, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
'CX3 = CByte(Val("&H" & X3))
CX3 = CByte("&H" & X3)
If CX3 > &H7F Then X3 = DwordToStr(CX3 Or &HFFFFFF00)
TJoin DPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
End If

Case &H6C
DisAssemble = "INS BYTE PTR [EDI],DX"
CL = 1

Case &H6D '+
If ISWP = 1 Then
DisAssemble = "INS WORD PTR [EDI],DX"
Else
DisAssemble = "INS DWORD PTR [EDI],DX"
End If
CL = 1

Case &H6E
DisAssemble = "OUTS DX,BYTE PTR [ESI]"
CL = 1

Case &H6F '+
If ISWP = 1 Then
DisAssemble = "OUTS DX,WORD PTR [ESI]"
Else
DisAssemble = "OUTS DX,DWORD PTR [ESI]"
End If
CL = 1

Case &H70 To &H7F
DisAssemble = CalcShortJump(DB2, JXX(MCs And &HF), Start + 1, VBASE + Start)
CL = 2



Case &HF
'************SC
DB3 = DATA(Start + 2)
FWR = Int((DB3 And &H3F) / 8)
Select Case DB2

Case 0
If FWR = 7 Then GoTo EendX
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = MMI(XRG)
MJoin WPL, DisAssemble, COMMD, X1, X2, ret

Case 1
If DB3 >= &HC0 Then
If Not (FWR = 4) And Not (FWR > 5) Then
GoTo EendX
ElseIf FWR = 7 Then
DisAssemble = "SWAPGS"
CL = 2
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = MMI2(XRG)
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
End If
Else
If FWR = 5 Then GoTo EendX
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = MMI2(XRG)
Select Case FWR
Case 0 To 3
MJoin FPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 4, 6
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
Case Else
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
End Select
End If

Case 2
COMMD = "LAR"
GoTo InIML

Case 3
COMMD = "LSL"
GoTo InIML

Case 5
DisAssemble = "SYSCALL"
CL = 1

Case 6
DisAssemble = "CLTS"
CL = 1

Case 7
DisAssemble = "SYSRET"
CL = 1

Case 8
DisAssemble = "INVD"
CL = 1

Case 9
DisAssemble = "WBINVD"
CL = 1

Case &HB
DisAssemble = "UD2"
CL = 1

Case &HD 'PREFETCH
If DB3 >= &HC0 Then GoTo EendX
If FWR = 0 Then
COMMD = "PREFETCH"
ElseIf FWR = 1 Then
COMMD = "PREFETCHW"
Else
GoTo EendX
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
MJoin NOL, DisAssemble, COMMD, X1, X2, ret


Case &HE
DisAssemble = "FEMMS"
CL = 1

Case &H6C
If ISWP = 1 Then
COMMD = "PUNPCKLQDQ": GoTo InXMMX
Else
GoTo EendX
End If

Case &H6D
If ISWP = 1 Then
COMMD = "PUNPCKHQDQ": GoTo InXMMX
Else
GoTo EendX
End If




Case &H12
If DB3 >= &HC0 Then
COMMD = "MOVHLPS"
Else
If ISWP = 1 Then
COMMD = "MOVLPD"
Else
COMMD = "MOVLPS"
End If
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H13
If DB3 >= &HC0 Then GoTo EendX
If ISWP = 1 Then
COMMD = "MOVLPD"
Else
COMMD = "MOVLPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin QPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H14
If ISWP = 1 Then
COMMD = "UNPCKLPD"
Else
COMMD = "UNPCKLPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret


Case &H15
If ISWP = 1 Then
COMMD = "UNPCKHPD"
Else
COMMD = "UNPCKHPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H16
If ISWP = 1 Then
COMMD = "MOVHPD"
Else
COMMD = "MOVHPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin QPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H17
If DB3 >= &HC0 Then GoTo EendX
If ISWP = 1 Then
COMMD = "MOVHPD"
Else
COMMD = "MOVHPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin QPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H18 'PREFETCH A-2
If DB3 >= &HC0 Then GoTo EendX
If FWR = 0 Then
COMMD = "PREFETCHNTA"
GoTo In18
ElseIf FWR = 1 Then
COMMD = "PREFETCHT0"
GoTo In18
ElseIf FWR = 2 Then
COMMD = "PREFETCHT1"
GoTo In18
ElseIf FWR = 3 Then
COMMD = "PREFETCHT2"
In18:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
Else
GoTo EendX
End If

Case &H28
If ISWP = 1 Then
COMMD = "MOVAPD"
Else
COMMD = "MOVAPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H29
If ISWP = 1 Then
COMMD = "MOVAPD"
Else
COMMD = "MOVAPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H2A
If ISWP = 1 Then
COMMD = "CVTPI2PD"
Else
COMMD = "CVTPI2PS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = XMM(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H2B
If DB3 >= &HC0 Then GoTo EendX
If ISWP = 1 Then
COMMD = "MOVNTPD"
Else
COMMD = "MOVNTPS"
End If
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H2C
If ISWP = 1 Then
COMMD = "CVTTPD2PI"
Else
COMMD = "CVTTPS2PI"
End If
GoTo In2D


Case &H2E
If ISWP = 1 Then
COMMD = "UCOMISD"
Else
COMMD = "UCOMISS"
End If
GoTo In2F

Case &H2F
If ISWP = 1 Then
COMMD = "COMISD"
Else
COMMD = "COMISS"
End If
In2F:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H30
DisAssemble = "WRMSR"
CL = 1

Case &H31
DisAssemble = "RDTSC"
CL = 1

Case &H32
DisAssemble = "RDMSR"
CL = 1

Case &H33
DisAssemble = "RDPMC"
CL = 1

Case &H34
DisAssemble = "SYSENTER"
CL = 1

Case &H35
DisAssemble = "SYSEXIT"
CL = 1

Case &H40
COMMD = "CMOVO"
GoTo InIML

Case &H41
COMMD = "CMOVNO"
GoTo InIML

Case &H42
COMMD = "CMOVB"
GoTo InIML

Case &H43
COMMD = "CMOVAE"
GoTo InIML

Case &H44
COMMD = "CMOVE"
GoTo InIML

Case &H45
COMMD = "CMOVNE"
GoTo InIML

Case &H46
COMMD = "CMOVBE"
GoTo InIML

Case &H47
COMMD = "CMOVA"
GoTo InIML

Case &H48
COMMD = "CMOVS"
GoTo InIML

Case &H49
COMMD = "CMOVNS"
GoTo InIML

Case &H4A
COMMD = "CMOVP"
GoTo InIML

Case &H4B
COMMD = "CMOVNP"
GoTo InIML

Case &H4C
COMMD = "CMOVL"
GoTo InIML

Case &H4D
COMMD = "CMOVGE"
GoTo InIML

Case &H4E
COMMD = "CMOVLE"
GoTo InIML

Case &H4F
COMMD = "CMOVG"
GoTo InIML

Case &H51
If ISWP = 1 Then
COMMD = "SQRTPD"
Else
COMMD = "SQRTPS"
End If
GoTo InSQ

Case &H52
If ISWP = 1 Then GoTo EendX
COMMD = "RSQRTPS"
GoTo InSQ

Case &H53
If ISWP = 1 Then GoTo EendX
COMMD = "RCPPS"
GoTo InSQ

Case &H54
If ISWP = 1 Then
COMMD = "ANDPD"
Else
COMMD = "ANDPS"
End If
GoTo InSQ

Case &H55
If ISWP = 1 Then
COMMD = "ANDNPD"
Else
COMMD = "ANDNPS"
End If
GoTo InSQ

Case &H56
If ISWP = 1 Then
COMMD = "ORPD"
Else
COMMD = "ORPS"
End If
GoTo InSQ

Case &H57
If ISWP = 1 Then
COMMD = "XORPD"
Else
COMMD = "XORPS"
End If
GoTo InSQ

Case &H58
If ISWP = 1 Then
COMMD = "ADDPD"
Else
COMMD = "ADDPS"
End If
GoTo InSQ

Case &H59
If ISWP = 1 Then
COMMD = "MULPD"
Else
COMMD = "MULPS"
End If
GoTo InSQ

Case &H5A
If ISWP = 1 Then
COMMD = "CVTPD2PS"
Else
COMMD = "CVTPS2PD"
End If
GoTo InSQ

Case &H5B
If ISWP = 1 Then
COMMD = "CVTPS2DQ"
Else
COMMD = "CVTDQ2PD"
End If
GoTo InSQ

Case &H5C
If ISWP = 1 Then
COMMD = "SUBPD"
Else
COMMD = "SUBPS"
End If
GoTo InSQ

Case &H5D
If ISWP = 1 Then
COMMD = "MINPD"
Else
COMMD = "MINPS"
End If
GoTo InSQ

Case &H5E
If ISWP = 1 Then
COMMD = "DIVPD"
Else
COMMD = "DIVPS"
End If
GoTo InSQ

Case &H5F
If ISWP = 1 Then
COMMD = "MAXPD"
Else
COMMD = "MAXPS"
End If

InSQ:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H60
COMMD = "PUNPCKLBW"
GoTo InPU

Case &H61
COMMD = "PUNPCKLWD"
GoTo InPU

Case &H62
COMMD = "PUNPCKLDQ"
InPU:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &H63
COMMD = "PACKSSWB"
GoTo InPA

Case &H64
COMMD = "PCMPGTB"
GoTo InPA

Case &H65
COMMD = "PCMPGTW"
GoTo InPA

Case &H66
COMMD = "PCMPGTD"
GoTo InPA

Case &H67
COMMD = "PACKUSWB"
GoTo InPA

Case &H68
COMMD = "PUNPCKHBW"
GoTo InPA

Case &H69
COMMD = "PUNPCKHWD"
GoTo InPA

Case &H6A
COMMD = "PUNPCKHDQ"
GoTo InPA

Case &H6B
COMMD = "PACKSSDW"
GoTo InPA


Case &H6E
COMMD = "MOVD"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = XMM(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
InMD:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = MMX(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &H6F
If ISWP = 1 Then
COMMD = "MOVDQA"
Else
COMMD = "MOVQ"
End If

InPA:
If ISWP = 1 Then
GoTo InXMMX
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H70 '***BYTE OPCODE
If ISWP = 1 Then
COMMD = "PSHUFD"
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
Else
COMMD = "PSHUFW"
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, ZRG, OPT16)
TJoin MMPR, DisAssemble, COMMD, MMX(FWR), X1, X2, X3, ret
End If

Case &H74
COMMD = "PCMPEQB"
If ISWP = 1 Then
GoTo InXMMX
Else
GoTo InPE
End If

Case &H75
COMMD = "PCMPEQW"
If ISWP = 1 Then
GoTo InXMMX
Else
GoTo InPE
End If

Case &H76
COMMD = "PCMPEQD"
If ISWP = 1 Then GoTo InXMMX
InPE:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H77
DisAssemble = "EMMS"
CL = 1

Case &H7E
COMMD = "MOVD"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = XMM(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = MMX(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H7F
If ISWP = 1 Then
COMMD = "MOVDQA"
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
COMMD = "MOVQ"
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
LJoin MMPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H80 To &H8F
DisAssemble = CalcLongJump(DATA, JXX(DB2 And &HF), Start + 1, VBASE + Start)
CL = 5

Case &H90 To &H9F
COMMD = SETXX(DB2 And &HF)
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
MJoin BPL, DisAssemble, COMMD, X1, X2, ret

Case &HA0
DisAssemble = "PUSH FS"
CL = 1

Case &HA1
DisAssemble = "POP FS"
CL = 1

Case &HA2
DisAssemble = "CPUID"
CL = 1

Case &HA3
COMMD = "BT"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HA4
COMMD = "SHLD"
GoTo InSHD

Case &HA5
COMMD = "SHLD"
GoTo InSHD2


Case &HA6
COMMD = "XBTS"
GoTo InIML

Case &HA7
COMMD = "IBTS"
GoTo InBTS


Case &HA8
DisAssemble = "PUSH GS"
CL = 1

Case &HA9
DisAssemble = "POP GS"
CL = 1

Case &HAA
DisAssemble = "RSM"
CL = 1

Case &HAB
COMMD = "BTS"
GoTo InBTS


Case &HAC
COMMD = "SHRD"
InSHD:
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
TJoin2 WPL, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
TJoin2 DPL, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
End If

Case &HAD
COMMD = "SHRD"
InSHD2:
X3 = "CL"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
TJoin2 WPL, DisAssemble, COMMD, REGOFFSET2(FWR), X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
TJoin2 DPL, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
End If

Case &HAE
If DB3 >= &HC0 Then
Select Case FWR
Case 5
DisAssemble = "LFENCE"
CL = 2
Case 6
DisAssemble = "MFENCE"
CL = 2
Case 7
DisAssemble = "SFENCE"
CL = 2
Case Else
GoTo EendX
End Select
Else
Select Case FWR
Case Is = 4, 5, 6
GoTo EendX
Case Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = CFLX(XRG)
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
End Select
End If

Case &HAF
COMMD = "IMUL"
GoTo InIML


Case &HB0
COMMD = "CMPXCHG"
GoTo InLA


Case &HB1
COMMD = "CMPXCHG"
GoTo InBTS

Case &HB2
If DB3 >= &HC0 Then GoTo EendX
COMMD = "LSS"
GoTo InLS


Case &HB3
COMMD = "BTR"
GoTo InBTS


Case &HB4
If DB3 >= &HC0 Then GoTo EendX
COMMD = "LFS"
GoTo InLS

Case &HB5
If DB3 >= &HC0 Then GoTo EendX
COMMD = "LGS"
InLS:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin FPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HB6
COMMD = "MOVZX"
GoTo InBSX


Case &HB7
COMMD = "MOVZX"
GoTo InWSX


Case &HB8
COMMD = "JMPE"
DisAssemble = CalcLongJump(DATA, COMMD, Start + 1, VBASE + Start)
CL = 5

Case &HBB
COMMD = "BTC"
GoTo InBTS

Case &HBC
COMMD = "BSF"
GoTo InIML

Case &HBD
COMMD = "BSR"
InIML:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HBE
COMMD = "MOVSX"
InBSX:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &HBF
COMMD = "MOVSX"
InWSX:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret


Case &HC0
COMMD = "XADD"
InLA:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
X3 = REGOFFSET1(XRG)
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &HC1
COMMD = "XADD"
GoTo InBTS


Case &HC2 'CMPS!
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
If (DB3 Mod &H8&) = 4 Then
If DATA(Start + 3) <= 7 Then
COMMD = CMPS(DATA(Start + 3))
Else
COMMD = "CMPSS"
End If
ElseIf DB3 <= 7 Then
COMMD = CMPS(DB3)
Else
COMMD = "CMPSS"
End If
TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret

Case &HC3
If DB3 >= &HC0 Then GoTo EendX
COMMD = "MOVNTI"
InBTS:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HC4
COMMD = "PINSRW"
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
TJoin WPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
TJoin WPR, DisAssemble, COMMD, MMX(FWR), X1, X2, X3, ret
End If


Case &HC5
COMMD = "PEXTRW"
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
TJoin XMMPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, ZRG, OPT16)
TJoin MMPR, DisAssemble, COMMD, REGOFFSET4(FWR), X1, X2, X3, ret
End If


Case &HC6
If ISWP = 1 Then
COMMD = "SHUFPD"
Else
COMMD = "SHUFPS"
End If
ret = PerformCMD2(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, ZRG, OPT16)
TJoin XMMPR, DisAssemble, COMMD, XMM(FWR), X1, X2, X3, ret

Case &HC7
If (DB3 >= 8 And DB3 <= 15) Or (DB3 >= &H48 And DB3 <= &H4F) Or (DB3 >= &H88 And DB3 <= &H8F) Then
COMMD = "CMPXCHG8B"
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
MJoin QPL, DisAssemble, COMMD, X1, X2, ret
Else
GoTo EendX
End If


Case &HC8 To &HCF
DisAssemble = "BSWAP " & REGOFFSET4(DB2 And 7)
CL = 1

Case &HD1
COMMD = "PSRLW"
GoTo InPSR

Case &HD2
COMMD = "PSRLD"
GoTo InPSR

Case &HD3
COMMD = "PSRLQ"
GoTo InPSR

Case &HD4
COMMD = "PADDQ"
GoTo InPSR

Case &HD5
COMMD = "PMULLW"
GoTo InPSR

Case &HD7
COMMD = "PMOVMSKB"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &HD8
COMMD = "PSUBUSB"
GoTo InPSR

Case &HD9
COMMD = "PSUBUSW"
GoTo InPSR

Case &HDA
COMMD = "PMINUB"
GoTo InPSR

Case &HDB
COMMD = "PAND"
GoTo InPSR

Case &HDC
COMMD = "PADDUSB"
GoTo InPSR

Case &HDD
COMMD = "PADDUSW"
GoTo InPSR

Case &HDE
COMMD = "PMAXUB"
GoTo InPSR

Case &HDF
COMMD = "PANDN"
GoTo InPSR

Case &HE0
COMMD = "PAVGB"
GoTo InPSR

Case &HE1
COMMD = "PSRAW"
GoTo InPSR

Case &HE2
COMMD = "PSRAD"
GoTo InPSR

Case &HE3
COMMD = "PAVGW"
GoTo InPSR

Case &HE4
COMMD = "PMULHUW"
GoTo InPSR

Case &HE5
COMMD = "PMULHW"
GoTo InPSR

Case &HE7
If DB3 >= &HC0 Then GoTo EendX

If ISWP = 1 Then
COMMD = "MOVNTDQ"
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin XMMPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
COMMD = "MOVNTQ"
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
LJoin MMPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &HE8
COMMD = "PSUBSB"
GoTo InPSR

Case &HE9
COMMD = "PSUBSW"
GoTo InPSR

Case &HEA
COMMD = "PMINSW"
GoTo InPSR

Case &HEB
COMMD = "POR"
GoTo InPSR

Case &HEC
COMMD = "PADDSB"
GoTo InPSR

Case &HED
COMMD = "PADDSW"
GoTo InPSR

Case &HEE
COMMD = "PMAXSW"
GoTo InPSR

Case &HEF
COMMD = "PXOR"
GoTo InPSR

Case &HF1
COMMD = "PSLLW"
GoTo InPSR

Case &HF2
COMMD = "PSLLD"
GoTo InPSR

Case &HF3
COMMD = "PSLLQ"
GoTo InPSR

Case &HF4
COMMD = "PMULUDQ"
GoTo InPSR

Case &HF5
COMMD = "PMADDWD"
GoTo InPSR

Case &HF6
COMMD = "PSADBW"
GoTo InPSR

Case &HF8
COMMD = "PSUBB"
GoTo InPSR

Case &HF9
COMMD = "PSUBW"
GoTo InPSR

Case &HFA
COMMD = "PSUBD"
GoTo InPSR

Case &HFB
COMMD = "PSUBQ"
GoTo InPSR

Case &HFC
COMMD = "PADDB"
GoTo InPSR

Case &HFD
COMMD = "PADDW"
GoTo InPSR

Case &HFE
COMMD = "PADDD"
InPSR:

If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, MMX, XRG, OPT16)
X3 = MMX(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &H10
If ISWP = 1 Then
COMMD = "MOVUPD"
Else
COMMD = "MOVUPS"
End If
InXMMX:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H11
If ISWP = 1 Then
COMMD = "MOVUPD"
Else
COMMD = "MOVUPS"
End If
InXMMX2:
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = XMM(XRG)
LJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H2D
If ISWP = 1 Then
COMMD = "CVTPD2PI"
Else
COMMD = "CVTPS2PI"
End If
In2D:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = MMX(XRG)
RJoin XMMPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start + 1, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, XMM, XRG, OPT16)
X3 = MMX(XRG)
RJoin MMPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case Else
GoTo EendX

End Select
CL = CL + 1



Case &H88
COMMD = "MOV"
GoTo InTTS2

Case &H89
COMMD = "MOV"
GoTo InTTS4

Case &H8A
COMMD = "MOV"
GoTo InTTSR2

Case &H8B
COMMD = "MOV"
GoTo InTTSR4

Case &H80
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
COMMD = MATHOFFSET(ZRG)
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H81 '+
If ISWP = 1 Then
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
COMMD = MATHOFFSET(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
COMMD = MATHOFFSET(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H83 '+
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
COMMD = MATHOFFSET(ZRG)
'CX3 = CByte(Val("&H" & X3))
CX3 = CByte("&H" & X3)
If CX3 > &H7F Then X3 = WordToStr(CX3 Or &HFF00)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
COMMD = MATHOFFSET(ZRG)
'CX3 = CByte(Val("&H" & X3))
CX3 = CByte("&H" & X3)
If CX3 > &H7F Then X3 = DwordToStr(CX3 Or &HFFFFFF00)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H84
COMMD = "TEST"
InTTS2:
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
X3 = REGOFFSET1(XRG)
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H85
COMMD = "TEST"
InTTS4:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &H86
COMMD = "XCHG"
InTTSR2:
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
X3 = REGOFFSET1(XRG)
RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H87
COMMD = "XCHG"
InTTSR4:
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If



Case &H8C
COMMD = "MOV"
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = SEGOFFSET(XRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &H8E
COMMD = "MOV"
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = SEGOFFSET(XRG)
RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret

Case &H8D '+
If DB2 >= &HC0 Then
GoTo EendX
Else
COMMD = "LEA"
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin NOR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin NOR, DisAssemble, COMMD, X1, X2, X3, ret
End If
End If

Case &H8F '+
COMMD = "POP"
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
If ISWP = 1 Then
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
Else
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
End If


Case &H90
DisAssemble = "NOP"
CL = 1

Case &H91 To &H97
If ISWP = 1 Then
DisAssemble = "XCHG AX," & REGOFFSET2(MCs And 7)
Else
DisAssemble = "XCHG EAX," & REGOFFSET4(MCs And 7)
End If
CL = 1

Case &H98
If ISWP = 1 Then
DisAssemble = "CBW"
Else
DisAssemble = "CWDE"
End If
CL = 1

Case &H99
If ISWP = 1 Then
DisAssemble = "CWD"
Else
DisAssemble = "CDQ"
End If
CL = 1

Case &H9A
If ISWP = 1 Then
DisAssemble = "CALL " & WordToStr(GetWordFromList(DATA, Start + 3)) & ":" & _
WordToStr(GetWordFromList(DATA, Start + 1))
CL = 5
Else
DisAssemble = "CALL " & WordToStr(GetWordFromList(DATA, Start + 5)) & ":" & _
DwordToStr(GetDWordFromList(DATA, Start + 1))
CL = 7
End If


Case &H9B
DisAssemble = "WAIT"
CL = 1

Case &H9C
If ISWP = 1 Then
DisAssemble = "PUSHF"
Else
DisAssemble = "PUSHFD"
End If
CL = 1

Case &H9D
If ISWP = 1 Then
DisAssemble = "POPF"
Else
DisAssemble = "POPFD"
End If
CL = 1

Case &H9E
DisAssemble = "SAHF"
CL = 1

Case &H9F
DisAssemble = "LAHF"
CL = 1

Case &HA0
DisAssemble = "MOV AL,BYTE PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "]"
CL = 5

Case &HA1
If ISWP = 1 Then
DisAssemble = "MOV AX,WORD PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "]"
Else
DisAssemble = "MOV EAX,DWORD PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "]"
End If
CL = 5

Case &HA2
DisAssemble = "MOV BYTE PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "],AL"
CL = 5

Case &HA3
If ISWP = 1 Then
DisAssemble = "MOV WORD PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "],AX"
Else
DisAssemble = "MOV DWORD PTR [" & DwordToStr(GetDWordFromList(DATA, Start + 1)) & "],EAX"
End If
CL = 5

Case &HA4
DisAssemble = "MOVS BYTE PTR [EDI],BYTE PTR [ESI]"
CL = 1

Case &HA5
If ISWP = 1 Then
DisAssemble = "MOVS WORD PTR [EDI],WORD PTR [ESI]"
Else
DisAssemble = "MOVS DWORD PTR [EDI],DWORD PTR [ESI]"
End If
CL = 1

Case &HA6
DisAssemble = "CMPS BYTE PTR [ESI],BYTE PTR [EDI]"
CL = 1

Case &HA7
If ISWP = 1 Then
DisAssemble = "CMPS WORD PTR [ESI],WORD PTR [EDI]"
Else
DisAssemble = "CMPS DWORD PTR [ESI],DWORD PTR [EDI]"
End If
CL = 1

Case &HA8
DisAssemble = "TEST AL," & ByteToStr(DB2)
CL = 2

Case &HA9
If ISWP = 1 Then
DisAssemble = "TEST AX," & WordToStr(GetWordFromList(DATA, Start + 1))
CL = 3
Else
DisAssemble = "TEST EAX," & DwordToStr(GetDWordFromList(DATA, Start + 1))
CL = 5
End If


Case &HAA
DisAssemble = "STOS BYTE PTR [EDI]"
CL = 1

Case &HAB
If ISWP = 1 Then
DisAssemble = "STOS WORD PTR [EDI]"
Else
DisAssemble = "STOS DWORD PTR [EDI]"
End If
CL = 1

Case &HAC
DisAssemble = "LODS BYTE PTR [ESI]"
CL = 1

Case &HAD
If ISWP = 1 Then
DisAssemble = "LODS WORD PTR [ESI]"
Else
DisAssemble = "LODS DWORD PTR [ESI]"
End If
CL = 1

Case &HAE
DisAssemble = "SCAS BYTE PTR [EDI]"
CL = 1

Case &HAF
If ISWP = 1 Then
DisAssemble = "SCAS WORD PTR [EDI]"
Else
DisAssemble = "SCAS DWORD PTR [EDI]"
End If
CL = 1

Case &HB0 To &HB7
DisAssemble = "MOV " & REGOFFSET1(MCs And 7) & "," & ByteToStr(DB2)
CL = 2

Case &HB8 To &HBF
If ISWP = 1 Then
DisAssemble = "MOV " & REGOFFSET2(MCs And 7) & "," & WordToStr(GetWordFromList(DATA, Start + 1))
CL = 3
Else
DisAssemble = "MOV " & REGOFFSET4(MCs And 7) & "," & DwordToStr(GetDWordFromList(DATA, Start + 1))
CL = 5
End If

Case &HC0
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
COMMD = BITOFFSET(ZRG)
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &HC1
If ISWP = 1 Then
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, ZRG, OPT16)
COMMD = BITOFFSET(ZRG)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, ZRG, OPT16)
COMMD = BITOFFSET(ZRG)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HC2
DisAssemble = "RET " & WordToStr(GetWordFromList(DATA, Start + 1))
CL = 3

Case &HC3
DisAssemble = "RET"
CL = 1

Case &HC4
COMMD = "LES"
GoTo InLDX

Case &HC5
COMMD = "LDS"
InLDX:
If DB2 >= &HC0 Then GoTo EendX
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
X3 = REGOFFSET2(XRG)
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
X3 = REGOFFSET4(XRG)
RJoin FPR, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &HC6
COMMD = "MOV"
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &HC7 '+
COMMD = "MOV"
If ISWP = 1 Then
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HC8
DisAssemble = "ENTER " & WordToStr(GetWordFromList(DATA, Start + 1)) & "," & _
ByteToStr(DATA(Start + 3))
CL = 4

Case &HC9
DisAssemble = "LEAVE"
CL = 1

Case &HCA
DisAssemble = "RETF " & WordToStr(GetWordFromList(DATA, Start + 1))
CL = 3

Case &HCB
DisAssemble = "RETF"
CL = 1

Case &HCC
DisAssemble = "INT 3"
CL = 1

Case &HCD
DisAssemble = "INT " & ByteToStr(DB2)
CL = 2

Case &HCE
DisAssemble = "INTO"
CL = 1

Case &HCF
DisAssemble = "IRETD"
CL = 1

Case &HD0
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "1"
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &HD1 '+
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "1"
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "1"
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If


Case &HD2
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "CL"
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret

Case &HD3 '+
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "CL"
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
COMMD = BITOFFSET(XRG)
X3 = "CL"
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case &HD4
DisAssemble = "AAMB " & ByteToStr(DB2)
CL = 2

Case &HD5
DisAssemble = "AADB " & ByteToStr(DB2)
CL = 2

Case &HD7
DisAssemble = "XLAT BYTE PTR [EBX]"
CL = 1


Case &HFE
If FWR > 1 Then GoTo EendX
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
COMMD = IDCJP(XRG)
MJoin BPL, DisAssemble, COMMD, X1, X2, ret

Case &HFF '+
If FWR = 7 Then GoTo EendX
If DB2 >= &HC0 Then If FWR = 3 Or FWR = 5 Then GoTo EendX
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = IDCJP(XRG)
If FWR = 3 Or FWR = 5 Then
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
Else
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End If
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
COMMD = IDCJP(XRG)
If FWR = 3 Or FWR = 5 Then
MJoin FPL, DisAssemble, COMMD, X1, X2, ret
Else
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
End If
End If


'FLOAT*
Case &HD8
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
If DB2 >= &HC0 Then X1 = "ST(0)," & X1
COMMD = FLOATOFFSET(XRG)
MJoin DPL, DisAssemble, COMMD, X1, X2, ret


Case &HDA
If DB2 >= &HC0 Then
CL = 2
CHK = FLOATOP5(DB2)
If Len(CHK) = 0 Then GoTo EendX
DisAssemble = CHK
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = INTFLOATOFFSET(XRG)
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
End If

Case &HD9
If DB2 >= &HC0 Then
CL = 2
CHK = FLOATOP3(DB2)
If Len(CHK) = 0 Then GoTo EendX
DisAssemble = CHK
Else
If FWR = 1 Then GoTo EendX
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP1(XRG)
Select Case FWR
Case Is = 0, 2, 3
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 4, 6
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
Case Is = 5, 7
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End Select
End If


Case &HDB
If DB2 >= &HC0 Then
CL = 2
CHK = FLOATOP4(DB2)
If Len(CHK) = 0 Then GoTo EendX
DisAssemble = CHK
Else
If FWR = 1 Or FWR = 4 Or FWR = 6 Then GoTo EendX
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP2(XRG)
Select Case FWR
Case Is = 0, 2, 3
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 5, 7
MJoin TBPL, DisAssemble, COMMD, X1, X2, ret
End Select
End If


Case &HDC
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
If DB2 >= &HC0 Then
X1 = X1 & ",ST(0)"
If FWR = 2 Then
COMMD = FLOATOFFSET(XRG) & "2"
ElseIf FWR = 3 Then
COMMD = FLOATOFFSET(XRG) & "3"
ElseIf FWR = 4 Or FWR = 6 Then
COMMD = FLOATOFFSET(XRG + 1)
ElseIf FWR = 5 Or FWR = 7 Then
COMMD = FLOATOFFSET(XRG - 1)
Else
COMMD = FLOATOFFSET(XRG)
End If
Else
COMMD = FLOATOFFSET(XRG)
End If
MJoin QPL, DisAssemble, COMMD, X1, X2, ret


Case &HDD
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
If DB2 >= &HC0 Then
If FWR > 5 Then GoTo EendX
COMMD = FLOATOP7(XRG)
MJoin QPL, DisAssemble, COMMD, X1, X2, ret
Else
If FWR = 1 Or FWR = 5 Then GoTo EendX
COMMD = FLOATOP6(XRG)
Select Case FWR
Case Is = 0, 2, 3
MJoin QPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 4, 6
MJoin NOL, DisAssemble, COMMD, X1, X2, ret
Case 7
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End Select
End If

Case &HDE
If DB2 >= &HC0 Then
If DB2 = &HD9 Then
DisAssemble = "FCOMPP"
CL = 2
ElseIf FWR = 3 Then
GoTo EendX
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP8(XRG)
If FWR <> 2 Then X1 = X1 & ",ST(0)"
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End If
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = INTFLOATOFFSET(XRG)
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End If

Case &HDF
If DB2 >= &HC0 Then
If DB2 = &HE0 Then
DisAssemble = "FNSTSW AX"
CL = 2
ElseIf FWR = 4 Or FWR = 5 Or FWR = 7 Then
GoTo EendX
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP10(XRG)
If FWR = 6 Then COMMD = COMMD & " ST(0),"
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
End If
Else
Select Case FWR
Case Is = 0, 2, 3
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP9(XRG)
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 4, 6
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP9(XRG)
MJoin TBPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 5, 7
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, FLOATSTACK, XRG, OPT16)
COMMD = FLOATOP9(XRG)
MJoin QPL, DisAssemble, COMMD, X1, X2, ret
Case 1
GoTo EendX
End Select
End If

Case &HE0 To &HE3
DisAssemble = CalcShortJump(DB2, LOOPX(MCs And &H3), Start + 1, VBASE + Start)
CL = 2

Case &HE4
DisAssemble = "IN AL," & ByteToStr(DB2)
CL = 2

Case &HE5
If ISWP = 1 Then
DisAssemble = "IN AX," & ByteToStr(DB2)
Else
DisAssemble = "IN EAX," & ByteToStr(DB2)
End If
CL = 2

Case &HE6
DisAssemble = "OUT " & ByteToStr(DB2) & ",AL"
CL = 2

Case &HE7
If ISWP = 1 Then
DisAssemble = "OUT " & ByteToStr(DB2) & ",AX"
Else
DisAssemble = "OUT " & ByteToStr(DB2) & ",EAX"
End If
CL = 2

Case &HE8
DisAssemble = CalcLongJump(DATA, "CALL", Start, VBASE + Start - 1)
CL = 5

Case &HE9
DisAssemble = CalcLongJump(DATA, "JMP", Start, VBASE + Start - 1)
CL = 5

Case &HEA '+
If ISWP = 1 Then
DisAssemble = "JMP " & WordToStr(GetWordFromList(DATA, Start + 3)) & ":" & _
WordToStr(GetWordFromList(DATA, Start + 1))
CL = 5
Else
DisAssemble = "JMP " & WordToStr(GetWordFromList(DATA, Start + 5)) & ":" & _
DwordToStr(GetDWordFromList(DATA, Start + 1))
CL = 7
End If

Case &HEB
DisAssemble = CalcShortJump(DB2, "JMP", Start + 1, VBASE + Start)
CL = 2

Case &HEC
DisAssemble = "IN AL,DX"
CL = 1

Case &HED
If ISWP = 1 Then
DisAssemble = "IN AX,DX"
Else
DisAssemble = "IN EAX,DX"
End If
CL = 1

Case &HEE
DisAssemble = "OUT DX,AL"
CL = 1

Case &HEF
If ISWP = 1 Then
DisAssemble = "OUT DX,AX"
Else
DisAssemble = "OUT DX,EAX"
End If
CL = 1

Case &HF0
Start = Start + 1
EXTRA = EXTRA + 1
EXTRAADD = 1
GoTo BEGGIN

Case &HF2
Start = Start + 1
EXTRA = EXTRA + 1
EXTRAADD = 2
GoTo BEGGIN

Case &HF3
Start = Start + 1
EXTRA = EXTRA + 1
EXTRAADD = 3
GoTo BEGGIN

Case &HF4
DisAssemble = "HLT"
CL = 1

Case &HF5
DisAssemble = "CMC"
CL = 1

Case &HF6
Select Case FWR
Case Is = 0, 1
ret = PerformCMD2(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, ZRG, OPT16)
COMMD = "TEST"
LJoin BPL, DisAssemble, COMMD, X1, X2, X3, ret
Case Is = 2, 3, 5, 7
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
COMMD = MATH2(XRG)
MJoin BPL, DisAssemble, COMMD, X1, X2, ret
Case Is = 4, 6
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET1, XRG, OPT16)
COMMD = MATH2(XRG)
X3 = "AL"
RJoin BPR, DisAssemble, COMMD, X1, X2, X3, ret
End Select


Case &HF7 '+
Select Case FWR
Case Is = 0, 1

If ISWP = 1 Then
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16, ISWP)
COMMD = "TEST"
LJoin WPL, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD3(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16, ISWP)
COMMD = "TEST"
LJoin DPL, DisAssemble, COMMD, X1, X2, X3, ret
End If

Case Is = 2, 3, 5
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = MATH2(XRG)
MJoin WPL, DisAssemble, COMMD, X1, X2, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
COMMD = MATH2(XRG)
MJoin DPL, DisAssemble, COMMD, X1, X2, ret
End If

Case Is = 4, 6, 7
If ISWP = 1 Then
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET2, XRG, OPT16)
COMMD = MATH2(XRG)
X3 = "AX"
RJoin WPR, DisAssemble, COMMD, X1, X2, X3, ret
Else
ret = PerformCMD(DATA, Start, COMMD, X1, X2, X3, CL, DL, REGOFFSET4, REGOFFSET4, REGOFFSET4, XRG, OPT16)
COMMD = MATH2(XRG)
X3 = "EAX"
RJoin DPR, DisAssemble, COMMD, X1, X2, X3, ret
End If

End Select



Case &HF8
DisAssemble = "CLC"
CL = 1

Case &HF9
DisAssemble = "STC"
CL = 1

Case &HFA
DisAssemble = "CLI"
CL = 1

Case &HFB
DisAssemble = "STI"
CL = 1

Case &HFC
DisAssemble = "CLD"
CL = 1

Case &HFD
DisAssemble = "STD"
CL = 1


Case Else
GoTo EendX

End Select

Forward = CL + DL + EXTRA

If EXTRAADD = 1 Then
DisAssemble = "LOCK " & DisAssemble
ElseIf EXTRAADD = 2 Then
DisAssemble = "REPNE " & DisAssemble
ElseIf EXTRAADD = 3 Then
DisAssemble = "REP " & DisAssemble
End If

If BinDump = 1 Then
RWDump DATA, Start, VBASE, Forward, DisAssemble
End If
OutOfF:
If HexAdr = 1 Then
DisAssemble = Hex(VBASE + Start - EXTRA) & vbTab & DisAssemble

End If

Exit Function

EendX:
On Error GoTo 0
DisAssemble = "db " & BByteToStr(DATA(Start - EXTRA)): Forward = 1
If HexAdr = 1 Then
DisAssemble = Hex(VBASE + Start - EXTRA) & " " & DisAssemble
End If

End Function
Friend Function PerformCMD2(DATA() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef ZRG As Byte, ByRef OPT16 As Byte) As Byte
Dim XD As Byte 'Data
Dim XOO As Byte 'OCTAL OFFSET
Dim YD As Byte 'Data2
Dim YOO As Byte 'OCTAL OFFSET
Dim YRG As Byte 'REGS OFFSET
Dim MPO As Byte 'MULTIPLY
XD = DATA(Start + 1)
XOO = XD Mod 8&
XRG = Int((XD And &H3F) / 8)
If OPT16 = 1 Then GoTo OP16X
Dim MPS As String 'Multiplication
If XD >= 0 And XD <= &H3F Then
If XOO = 4 Then
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
ZRG = Int(XD - 5) / 8
If (YRG Mod 8&) = 4 Then
If YOO = 5 Then
CL = 3
DL = 5
X1 = DwordToStr(GetDWordFromList(DATA, Start + CL))
Else
CL = 3
DL = 1
X1 = R1(YOO)
End If
ElseIf YOO = 5 Then
CL = 3
DL = 5
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X1 = R2(YRG Mod 8&)
Else
X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
Else
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&)
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
CL = 3
DL = 1
X1 = R1(YOO)
End If
ElseIf XOO = 5 Then
CL = 2
DL = 5
ZRG = Int(XD - 5) / 8
X1 = DwordToStr(GetDWordFromList(DATA, Start + CL))
Else
CL = 2
DL = 1
ZRG = XRG
X1 = R1(XOO)
End If
X3 = ByteToStr(DATA(Start + CL + DL - 1))
ElseIf XD >= &H40 And XD <= &H7F Then
ZRG = XRG
If XOO = 4 Then
CL = 3
DL = 2
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
X1 = R1(YOO)
X2 = SByteToStr(DATA(Start + CL))
Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SByteToStr(DATA(Start + CL))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(DATA(Start + CL))
End If
End If
Else
CL = 2
DL = 2
X1 = R1(XOO)
X2 = SByteToStr(DATA(Start + CL))
End If
X3 = ByteToStr(DATA(Start + CL + DL - 1))
ElseIf XD >= &H80 And XD <= &HBF Then
If XOO = 4 Then
CL = 3
DL = 5
YD = DATA(Start + 2)
YOO = YD Mod 8
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
X1 = R1(YOO)
Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(DATA, Start + CL))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(DATA, Start + CL))
End If
End If
Else
CL = 2
DL = 5
X1 = R1(XOO)
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
End If
X3 = ByteToStr(DATA(Start + CL + DL - 1))
ZRG = XRG
ElseIf XD >= &HC0 Then
CL = 2
DL = 1
X3 = ByteToStr(DATA(Start + CL))
X1 = R3(XOO)
PerformCMD2 = 1
ZRG = XRG
End If
Exit Function

OP16X:
ZRG = XRG
If XD >= 0 And XD <= &H3F Then
If XOO = 6 Then
CL = 2
DL = 3
X1 = WordToStr(GetWordFromList(DATA, Start + CL))
X3 = ByteToStr(DATA(Start + CL + 2))
Else
X1 = REG16O1(XOO)
CL = 2
DL = 1
X3 = ByteToStr(DATA(Start + CL))
End If
ElseIf XD >= &H40 And XD <= &H7F Then
CL = 2
DL = 2
X2 = SByteToStr(DATA(Start + CL))
X1 = REG16O1(XOO)
X3 = ByteToStr(DATA(Start + CL + 1))
ElseIf XD >= &H80 And XD <= &HBF Then
CL = 2
DL = 3
X2 = "+" & WordToStr(GetWordFromList(DATA, Start + CL))
X1 = REG16O1(XOO)
X3 = ByteToStr(DATA(Start + CL + 2))
ElseIf XD >= &HC0 Then
CL = 2
DL = 1
X3 = ByteToStr(DATA(Start + CL))
X1 = R3(XOO)
PerformCMD2 = 1
End If

End Function
Friend Function PerformCMD3(DATA() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef OPT16 As Byte, Optional ByRef ISWP As Byte) As Byte
Dim XD As Byte 'Data
Dim XOO As Byte 'OCTALNI OFFSET
Dim YD As Byte 'Data2
Dim YOO As Byte 'OCTALNI OFFSET
Dim YRG As Byte 'REGS OFFSET
Dim MPO As Byte 'MULTIPLY
XD = DATA(Start + 1)
XOO = XD Mod 8&
XRG = Int((XD And &H3F) / 8)
If OPT16 = 1 Then GoTo OPT16X
If XD >= 0 And XD <= &H3F Then
If XOO = 5 Then
CL = 1
X2 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If

ElseIf XOO = 4 Then
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
If YOO = 5 Then
CL = 2
X1 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If

Else
CL = 2
X1 = R1(YOO)
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

End If
ElseIf YOO = 5 Then
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X1 = R2(YRG Mod 8&)
Else
X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
CL = 2
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL + 1))
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If

Else
CL = 2

MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&)
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
X1 = R1(YOO)

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

End If
Else
CL = 1
X1 = R1(XOO)

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

End If
ElseIf XD >= &H40 And XD <= &H7F Then
If XOO = 4 Then
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
CL = 3
If (YRG Mod 8&) = 4 Then
X1 = R1(YOO) & SByteToStr(DATA(Start + CL))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SByteToStr(DATA(Start + CL))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(DATA(Start + CL))
End If
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

End If
Else
CL = 2

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

X2 = SByteToStr(DATA(Start + CL))
X1 = R1(XOO)
End If
ElseIf XD >= &H80 And XD <= &HBF Then
If XOO = 4 Then
CL = 2

YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
X1 = R1(YOO)
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL + 1))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If

Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(DATA, Start + CL + 1))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(DATA, Start + CL + 1))
End If
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If
End If
Else
CL = 1
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 5))
DL = 7
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 5))
DL = 9
End If
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL + 1))
X1 = R1(XOO)
End If
ElseIf XD >= &HC0 Then
PerformCMD3 = 1
CL = 1
X1 = R3(XOO)

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

End If
Exit Function
OPT16X:
If XD >= 0 And XD <= &H3F Then
If XOO = 6 Then
CL = 2

X1 = WordToStr(GetWordFromList(DATA, Start + CL))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 2))
DL = 4
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 2))
DL = 6
End If

Else
CL = 2

X1 = REG16O1(XOO)

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL))
DL = 2
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL))
DL = 4
End If

End If
ElseIf XD >= &H40 And XD <= &H7F Then
CL = 2

X1 = REG16O1(XOO)
X2 = SByteToStr(DATA(Start + CL))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If

ElseIf XD >= &H80 And XD <= &HBF Then
CL = 2
X1 = REG16O1(XOO)
X2 = "+" & WordToStr(GetWordFromList(DATA, Start + CL))

If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 2))
DL = 4
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 2))
DL = 6
End If

ElseIf XD >= &HC0 Then
PerformCMD3 = 1
CL = 1

X1 = R3(XOO)
If ISWP = 1 Then
X3 = WordToStr(GetWordFromList(DATA, Start + CL + 1))
DL = 3
Else
X3 = DwordToStr(GetDWordFromList(DATA, Start + CL + 1))
DL = 5
End If
End If
End Function
Friend Function PerformCMD(DATA() As Byte, ByRef Start As Long, ByRef COMMD As String, ByRef X1 As String, ByRef X2 As String, ByRef X3 As String, ByRef CL As Byte, ByRef DL As Byte, R1() As String, R2() As String, R3() As String, ByRef XRG As Byte, ByRef OPT16 As Byte) As Byte
Dim XD As Byte 'Data
Dim XOO As Byte 'OCTAL OFFSET
Dim YD As Byte 'Data2
Dim YOO As Byte 'OCTAL OFFSET
Dim YRG As Byte 'REGS OFFSET
Dim MPO As Byte 'MULTIPLY
XD = DATA(Start + 1)
XOO = XD Mod 8&
XRG = Int((XD And &H3F) / 8)
If OPT16 = 1 Then GoTo OPT16X
If XD >= 0 And XD <= &H3F Then
If XOO = 5 Then
CL = 2
DL = 4
X1 = DwordToStr(GetDWordFromList(DATA, Start + CL))
ElseIf XOO = 4 Then
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
If YOO = 5 Then
CL = 3
DL = 4
X1 = DwordToStr(GetDWordFromList(DATA, Start + CL))
Else
CL = 3
DL = 0
X1 = R1(YOO)
End If
ElseIf YOO = 5 Then
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X1 = R2(YRG Mod 8&)
Else
X1 = R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
CL = 3
DL = 4
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
Else
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&)
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO)
End If
X1 = R1(YOO)
CL = 3
DL = 0
End If
Else
CL = 2
DL = 0
X1 = R1(XOO)
End If
ElseIf XD >= &H40 And XD <= &H7F Then
If XOO = 4 Then
CL = 3
DL = 1
YD = DATA(Start + 2)
YOO = YD Mod 8&
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
X1 = R1(YOO) & SByteToStr(DATA(Start + CL))
Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SByteToStr(DATA(Start + CL))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SByteToStr(DATA(Start + CL))
End If
End If
Else
CL = 2
DL = 1
X2 = SByteToStr(DATA(Start + CL))
X1 = R1(XOO)
End If
ElseIf XD >= &H80 And XD <= &HBF Then
If XOO = 4 Then
CL = 3
DL = 4
YD = DATA(Start + 2)
YOO = YD Mod 8
YRG = Int(YD / 8)
If (YRG Mod 8&) = 4 Then
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
X1 = R1(YOO)
Else
X1 = R1(YOO)
MPO = 2 ^ Int(YD / &H40)
If MPO < 2 Then
X2 = "+" & R2(YRG Mod 8&) & SDwordToStr(GetDWordFromList(DATA, Start + CL))
Else
X2 = "+" & R2(YRG Mod 8&) & "*" & CStr(MPO) & SDwordToStr(GetDWordFromList(DATA, Start + CL))
End If
End If
Else
CL = 2
DL = 4
X1 = R1(XOO)
X2 = SDwordToStr(GetDWordFromList(DATA, Start + CL))
End If
ElseIf XD >= &HC0 Then
PerformCMD = 1
CL = 2
X1 = R3(XOO)
End If
Exit Function
OPT16X:
If XD >= 0 And XD <= &H3F Then
If XOO = 6 Then
CL = 2
DL = 2
X1 = WordToStr(GetWordFromList(DATA, Start + CL))
Else
X1 = REG16O1(XOO)
CL = 1
DL = 1
End If
ElseIf XD >= &H40 And XD <= &H7F Then
CL = 2
DL = 1
X1 = REG16O1(XOO)
X2 = SByteToStr(DATA(Start + CL))
ElseIf XD >= &H80 And XD <= &HBF Then
CL = 2
DL = 2
X2 = "+" & WordToStr(GetWordFromList(DATA, Start + CL))
X1 = REG16O1(XOO)
ElseIf XD >= &HC0 Then
PerformCMD = 1
CL = 2
X1 = R3(XOO)
End If
End Function

Private Sub Class_Initialize()
Static Init_ As Byte
If Init_ = 1 Then Exit Sub
Init_ = 1

ReDim REGOFFSET4(7)
REGOFFSET4(0) = "EAX"
REGOFFSET4(1) = "ECX"
REGOFFSET4(2) = "EDX"
REGOFFSET4(3) = "EBX"
REGOFFSET4(4) = "ESP"
REGOFFSET4(5) = "EBP"
REGOFFSET4(6) = "ESI"
REGOFFSET4(7) = "EDI"

ReDim REGOFFSET2(7)
REGOFFSET2(0) = "AX"
REGOFFSET2(1) = "CX"
REGOFFSET2(2) = "DX"
REGOFFSET2(3) = "BX"
REGOFFSET2(4) = "SP"
REGOFFSET2(5) = "BP"
REGOFFSET2(6) = "SI"
REGOFFSET2(7) = "DI"

ReDim REGOFFSET1(7)
REGOFFSET1(0) = "AL"
REGOFFSET1(1) = "CL"
REGOFFSET1(2) = "DL"
REGOFFSET1(3) = "BL"
REGOFFSET1(4) = "AH"
REGOFFSET1(5) = "CH"
REGOFFSET1(6) = "DH"
REGOFFSET1(7) = "BH"

ReDim SEGOFFSET(7)
SEGOFFSET(0) = "ES"
SEGOFFSET(1) = "CS"
SEGOFFSET(2) = "SS"
SEGOFFSET(3) = "DS"
SEGOFFSET(4) = "FS"
SEGOFFSET(5) = "GS"
SEGOFFSET(6) = "??"
SEGOFFSET(7) = "??"

ReDim MATHOFFSET(7)
MATHOFFSET(0) = "ADD"
MATHOFFSET(1) = "OR"
MATHOFFSET(2) = "ADC"
MATHOFFSET(3) = "SBB"
MATHOFFSET(4) = "AND"
MATHOFFSET(5) = "SUB"
MATHOFFSET(6) = "XOR"
MATHOFFSET(7) = "CMP"

ReDim BITOFFSET(7)
BITOFFSET(0) = "ROL"
BITOFFSET(1) = "ROR"
BITOFFSET(2) = "RCL"
BITOFFSET(3) = "RCR"
BITOFFSET(4) = "SHL"
BITOFFSET(5) = "SHR"
BITOFFSET(6) = "SAL"
BITOFFSET(7) = "SAR"

ReDim FLOATOFFSET(7)
FLOATOFFSET(0) = "FADD"
FLOATOFFSET(1) = "FMUL"
FLOATOFFSET(2) = "FCOM"
FLOATOFFSET(3) = "FCOMP"
FLOATOFFSET(4) = "FSUB"
FLOATOFFSET(5) = "FSUBR"
FLOATOFFSET(6) = "FDIV"
FLOATOFFSET(7) = "FDIVR"

ReDim INTFLOATOFFSET(7)
INTFLOATOFFSET(0) = "FIADD"
INTFLOATOFFSET(1) = "FIMUL"
INTFLOATOFFSET(2) = "FICOM"
INTFLOATOFFSET(3) = "FICOMP"
INTFLOATOFFSET(4) = "FISUB"
INTFLOATOFFSET(5) = "FISUBR"
INTFLOATOFFSET(6) = "FIDIV"
INTFLOATOFFSET(7) = "FIDIVR"

ReDim FLOATSTACK(7)
FLOATSTACK(0) = "ST(0)"
FLOATSTACK(1) = "ST(1)"
FLOATSTACK(2) = "ST(2)"
FLOATSTACK(3) = "ST(3)"
FLOATSTACK(4) = "ST(4)"
FLOATSTACK(5) = "ST(5)"
FLOATSTACK(6) = "ST(6)"
FLOATSTACK(7) = "ST(7)"

ReDim FLOATOP1(7)
FLOATOP1(0) = "FLD" '- DWORD PTR
'FLOATOP1(1) = "" '-----NO INSTR.
FLOATOP1(2) = "FST" '- DWORD PTR
FLOATOP1(3) = "FSTP" '- DWORD PTR
FLOATOP1(4) = "FLDENV" '[]
FLOATOP1(5) = "FLDCW" '-WORD PTR
FLOATOP1(6) = "FNSTENV" '[]
FLOATOP1(7) = "FNSTCW" '--WORD PTR

ReDim FLOATOP2(7)
FLOATOP2(0) = "FILD" '- DWORD PTR
'FLOATOP2(1) = "" '-----NO INSTR.
FLOATOP2(2) = "FIST" '- DWORD PTR
FLOATOP2(3) = "FISTP" '- DWORD PTR
'FLOATOP2(4) = "" '-----NO INSTR.
FLOATOP2(5) = "FLD" '-TBYTE PTR
'FLOATOP(6) = "" '----NO INSTR.
FLOATOP2(7) = "FSTP" '--TBYTE PTR

ReDim FLOATOP3(&HC0 To &HFF)
FLOATOP3(&HC0) = "FLD ST(0),ST(0)"
FLOATOP3(&HC1) = "FLD ST(0),ST(1)"
FLOATOP3(&HC2) = "FLD ST(0),ST(2)"
FLOATOP3(&HC3) = "FLD ST(0),ST(3)"
FLOATOP3(&HC4) = "FLD ST(0),ST(4)"
FLOATOP3(&HC5) = "FLD ST(0),ST(5)"
FLOATOP3(&HC6) = "FLD ST(0),ST(6)"
FLOATOP3(&HC7) = "FLD ST(0),ST(7)"
FLOATOP3(&HC8) = "FXCH ST(0),ST(0)"
FLOATOP3(&HC9) = "FXCH ST(0),ST(1)"
FLOATOP3(&HCA) = "FXCH ST(0),ST(2)"
FLOATOP3(&HCB) = "FXCH ST(0),ST(3)"
FLOATOP3(&HCC) = "FXCH ST(0),ST(4)"
FLOATOP3(&HCD) = "FXCH ST(0),ST(5)"
FLOATOP3(&HCE) = "FXCH ST(0),ST(6)"
FLOATOP3(&HCF) = "FXCH ST(0),ST(7)"
FLOATOP3(&HD0) = "FNOP"
'D1-D7 no instr
FLOATOP3(&HD8) = "FSTP1 ST(0),ST(0)"
FLOATOP3(&HD9) = "FSTP1 ST(0),ST(1)"
FLOATOP3(&HDA) = "FSTP1 ST(0),ST(2)"
FLOATOP3(&HDB) = "FSTP1 ST(0),ST(3)"
FLOATOP3(&HDC) = "FSTP1 ST(0),ST(4)"
FLOATOP3(&HDD) = "FSTP1 ST(0),ST(5)"
FLOATOP3(&HDE) = "FSTP1 ST(0),ST(6)"
FLOATOP3(&HDF) = "FSTP1 ST(0),ST(7)"
FLOATOP3(&HE0) = "FCHS"
FLOATOP3(&HE1) = "FABS"
'E2-no inst
'E3-no inst
FLOATOP3(&HE4) = "FTST"
FLOATOP3(&HE5) = "FXAM"
'E6-no inst
'E7-no inst
FLOATOP3(&HE8) = "FLD1"
FLOATOP3(&HE9) = "FLDL2T"
FLOATOP3(&HEA) = "FLDL2E"
FLOATOP3(&HEB) = "FLDPI"
FLOATOP3(&HEC) = "FLDLG2"
FLOATOP3(&HED) = "FLDLN2"
FLOATOP3(&HEE) = "FLDZ"
'EF-no instr
FLOATOP3(&HF0) = "F2XM1"
FLOATOP3(&HF1) = "FYL2X"
FLOATOP3(&HF2) = "FPTAN"
FLOATOP3(&HF3) = "FPATAN"
FLOATOP3(&HF4) = "FEXTRACT"
FLOATOP3(&HF5) = "FPREM1"
FLOATOP3(&HF6) = "FDECSTP"
FLOATOP3(&HF7) = "FINCSTP"
FLOATOP3(&HF8) = "FPREM"
FLOATOP3(&HF9) = "FYL2XP1"
FLOATOP3(&HFA) = "FSQRT"
FLOATOP3(&HFB) = "FSINCOS"
FLOATOP3(&HFC) = "FRNDINT"
FLOATOP3(&HFD) = "FSCALE"
FLOATOP3(&HFE) = "FSIN"
FLOATOP3(&HFF) = "FCOS"

ReDim FLOATOP4(&HC0 To &HFF)
FLOATOP4(&HC0) = "FCMOVNB ST(0),ST(0)"
FLOATOP4(&HC1) = "FCMOVNB ST(0),ST(1)"
FLOATOP4(&HC2) = "FCMOVNB ST(0),ST(2)"
FLOATOP4(&HC3) = "FCMOVNB ST(0),ST(3)"
FLOATOP4(&HC4) = "FCMOVNB ST(0),ST(4)"
FLOATOP4(&HC5) = "FCMOVNB ST(0),ST(5)"
FLOATOP4(&HC6) = "FCMOVNB ST(0),ST(6)"
FLOATOP4(&HC7) = "FCMOVNB ST(0),ST(7)"
FLOATOP4(&HC8) = "FCMOVNE ST(0),ST(0)"
FLOATOP4(&HC9) = "FCMOVNE ST(0),ST(1)"
FLOATOP4(&HCA) = "FCMOVNE ST(0),ST(2)"
FLOATOP4(&HCB) = "FCMOVNE ST(0),ST(3)"
FLOATOP4(&HCC) = "FCMOVNE ST(0),ST(4)"
FLOATOP4(&HCD) = "FCMOVNE ST(0),ST(5)"
FLOATOP4(&HCE) = "FCMOVNE ST(0),ST(6)"
FLOATOP4(&HCF) = "FCMOVNE ST(0),ST(7)"
FLOATOP4(&HD0) = "FCMOVNBE ST(0),ST(0)"
FLOATOP4(&HD1) = "FCMOVNBE ST(0),ST(1)"
FLOATOP4(&HD2) = "FCMOVNBE ST(0),ST(2)"
FLOATOP4(&HD3) = "FCMOVNBE ST(0),ST(3)"
FLOATOP4(&HD4) = "FCMOVNBE ST(0),ST(4)"
FLOATOP4(&HD5) = "FCMOVNBE ST(0),ST(5)"
FLOATOP4(&HD6) = "FCMOVNBE ST(0),ST(6)"
FLOATOP4(&HD7) = "FCMOVNBE ST(0),ST(7)"
FLOATOP4(&HD8) = "FCMOVNU ST(0),ST(0)"
FLOATOP4(&HD9) = "FCMOVNU ST(0),ST(1)"
FLOATOP4(&HDA) = "FCMOVNU ST(0),ST(2)"
FLOATOP4(&HDB) = "FCMOVNU ST(0),ST(3)"
FLOATOP4(&HDC) = "FCMOVNU ST(0),ST(4)"
FLOATOP4(&HDD) = "FCMOVNU ST(0),ST(5)"
FLOATOP4(&HDE) = "FCMOVNU ST(0),ST(6)"
FLOATOP4(&HDF) = "FCMOVNU ST(0),ST(7)"
FLOATOP4(&HE0) = "FENI"
FLOATOP4(&HE1) = "FDISI"
FLOATOP4(&HE2) = "FNCLEX"
FLOATOP4(&HE3) = "FNINIT"
FLOATOP4(&HE4) = "FSETPM"
'E5-E7 no inst
FLOATOP4(&HE8) = "FUCOMI ST(0),ST(0)"
FLOATOP4(&HE9) = "FUCOMI ST(0),ST(1)"
FLOATOP4(&HEA) = "FUCOMI ST(0),ST(2)"
FLOATOP4(&HEB) = "FUCOMI ST(0),ST(3)"
FLOATOP4(&HEC) = "FUCOMI ST(0),ST(4)"
FLOATOP4(&HED) = "FUCOMI ST(0),ST(5)"
FLOATOP4(&HEE) = "FUCOMI ST(0),ST(6)"
FLOATOP4(&HEF) = "FUCOMI ST(0),ST(7)"
FLOATOP4(&HF0) = "FCOMI ST(0),ST(0)"
FLOATOP4(&HF1) = "FCOMI ST(0),ST(1)"
FLOATOP4(&HF2) = "FCOMI ST(0),ST(2)"
FLOATOP4(&HF3) = "FCOMI ST(0),ST(3)"
FLOATOP4(&HF4) = "FCOMI ST(0),ST(4)"
FLOATOP4(&HF5) = "FCOMI ST(0),ST(5)"
FLOATOP4(&HF6) = "FCOMI ST(0),ST(6)"
FLOATOP4(&HF7) = "FCOMI ST(0),ST(7)"
'F8-FF no instr

ReDim FLOATOP5(&HC0 To &HFF)
FLOATOP5(&HC0) = "FCMOVB ST(0),ST(0)"
FLOATOP5(&HC1) = "FCMOVB ST(0),ST(1)"
FLOATOP5(&HC2) = "FCMOVB ST(0),ST(2)"
FLOATOP5(&HC3) = "FCMOVB ST(0),ST(3)"
FLOATOP5(&HC4) = "FCMOVB ST(0),ST(4)"
FLOATOP5(&HC5) = "FCMOVB ST(0),ST(5)"
FLOATOP5(&HC6) = "FCMOVB ST(0),ST(6)"
FLOATOP5(&HC7) = "FCMOVB ST(0),ST(7)"
FLOATOP5(&HC8) = "FCMOVE ST(0),ST(0)"
FLOATOP5(&HC9) = "FCMOVE ST(0),ST(1)"
FLOATOP5(&HCA) = "FCMOVE ST(0),ST(2)"
FLOATOP5(&HCB) = "FCMOVE ST(0),ST(3)"
FLOATOP5(&HCC) = "FCMOVE ST(0),ST(4)"
FLOATOP5(&HCD) = "FCMOVE ST(0),ST(5)"
FLOATOP5(&HCE) = "FCMOVE ST(0),ST(6)"
FLOATOP5(&HCF) = "FCMOVE ST(0),ST(7)"
FLOATOP5(&HD0) = "FCMOVBE ST(0),ST(0)"
FLOATOP5(&HD1) = "FCMOVBE ST(0),ST(1)"
FLOATOP5(&HD2) = "FCMOVBE ST(0),ST(2)"
FLOATOP5(&HD3) = "FCMOVBE ST(0),ST(3)"
FLOATOP5(&HD4) = "FCMOVBE ST(0),ST(4)"
FLOATOP5(&HD5) = "FCMOVBE ST(0),ST(5)"
FLOATOP5(&HD6) = "FCMOVBE ST(0),ST(6)"
FLOATOP5(&HD7) = "FCMOVBE ST(0),ST(7)"
FLOATOP5(&HD8) = "FCMOVU ST(0),ST(0)"
FLOATOP5(&HD9) = "FCMOVU ST(0),ST(1)"
FLOATOP5(&HDA) = "FCMOVU ST(0),ST(2)"
FLOATOP5(&HDB) = "FCMOVU ST(0),ST(3)"
FLOATOP5(&HDC) = "FCMOVU ST(0),ST(4)"
FLOATOP5(&HDD) = "FCMOVU ST(0),ST(5)"
FLOATOP5(&HDE) = "FCMOVU ST(0),ST(6)"
FLOATOP5(&HDF) = "FCMOVU ST(0),ST(7)"
FLOATOP5(&HE9) = "FUCOMPP"

ReDim FLOATOP6(7)
FLOATOP6(0) = "FLD"
'1-no instr
FLOATOP6(2) = "FST"
FLOATOP6(3) = "FSTP"
FLOATOP6(4) = "FRSTOR"
'5-no instr
FLOATOP6(6) = "FNSAVE"
FLOATOP6(7) = "FNSTSW"

ReDim FLOATOP7(7)
FLOATOP7(0) = "FFREE"
FLOATOP7(1) = "FXCH4"
FLOATOP7(2) = "FST"
FLOATOP7(3) = "FSTP"
FLOATOP7(4) = "FUCOM"
FLOATOP7(5) = "FUCOMP"
'6-8 no instr

ReDim FLOATOP8(7)
FLOATOP8(0) = "FADDP"
FLOATOP8(1) = "FMULP"
FLOATOP8(2) = "FCOMP5"
'3 no instr
FLOATOP8(4) = "FSUBRP"
FLOATOP8(5) = "FSUBP"
FLOATOP8(6) = "FDIVRP"
FLOATOP8(7) = "FDIVP"

ReDim FLOATOP9(7)
FLOATOP9(0) = "FILD"
'1 no instr
FLOATOP9(2) = "FIST"
FLOATOP9(3) = "FISTP"
FLOATOP9(4) = "FBLD"
FLOATOP9(5) = "FILD"
FLOATOP9(6) = "FBSTP"
FLOATOP9(7) = "FISTP"

ReDim FLOATOP10(7)
FLOATOP10(0) = "FFREEP"
FLOATOP10(1) = "FXCH7"
FLOATOP10(2) = "FSTP8"
FLOATOP10(3) = "FSTP9"
'4-5 no instr
FLOATOP10(6) = "FUCOMIP"
'7 no instr

ReDim IDCJP(7)
IDCJP(0) = "INC"
IDCJP(1) = "DEC"
IDCJP(2) = "CALL"
IDCJP(3) = "CALL"
IDCJP(4) = "JMP"
IDCJP(5) = "JMP"
IDCJP(6) = "PUSH"
'7-no instr

ReDim MMI(7)
MMI(0) = "SLDT"
MMI(1) = "STR"
MMI(2) = "LLDT"
MMI(3) = "LTR"
MMI(4) = "VERR"
MMI(5) = "VERW"
MMI(6) = "JMPE"
'7 -no instr

ReDim MMI2(7)
MMI2(0) = "SGDT"
MMI2(1) = "SIDT"
MMI2(2) = "LGDT"
MMI2(3) = "LIDT"
MMI2(4) = "SMSW"
'5 -no instr
MMI2(6) = "LMSW"
MMI2(7) = "INVLPG"

ReDim XMM(7)
XMM(0) = "XMM0"
XMM(1) = "XMM1"
XMM(2) = "XMM2"
XMM(3) = "XMM3"
XMM(4) = "XMM4"
XMM(5) = "XMM5"
XMM(6) = "XMM6"
XMM(7) = "XMM7"

ReDim MMX(7)
MMX(0) = "MM0"
MMX(1) = "MM1"
MMX(2) = "MM2"
MMX(3) = "MM3"
MMX(4) = "MM4"
MMX(5) = "MM5"
MMX(6) = "MM6"
MMX(7) = "MM7"

ReDim JXX(&HF)
JXX(&H0) = "JO"
JXX(&H1) = "JNO"
JXX(&H2) = "JB"
JXX(&H3) = "JAE"
JXX(&H4) = "JE"
JXX(&H5) = "JNE"
JXX(&H6) = "JBE"
JXX(&H7) = "JA"
JXX(&H8) = "JS"
JXX(&H9) = "JNS"
JXX(&HA) = "JP"
JXX(&HB) = "JNP"
JXX(&HC) = "JL"
JXX(&HD) = "JGE"
JXX(&HE) = "JLE"
JXX(&HF) = "JG"

ReDim SETXX(&HF)
SETXX(&H0) = "SETO"
SETXX(&H1) = "SETNO"
SETXX(&H2) = "SETB"
SETXX(&H3) = "SETAE"
SETXX(&H4) = "SETE"
SETXX(&H5) = "SETNE"
SETXX(&H6) = "SETBE"
SETXX(&H7) = "SETA"
SETXX(&H8) = "SETS"
SETXX(&H9) = "SETNS"
SETXX(&HA) = "SETP"
SETXX(&HB) = "SETNP"
SETXX(&HC) = "SETL"
SETXX(&HD) = "SETGE"
SETXX(&HE) = "SETLE"
SETXX(&HF) = "SETG"

ReDim CFLX(7)
CFLX(0) = "FXSAVE"
CFLX(1) = "FXRSTOR"
CFLX(2) = "LDMXCSR"
CFLX(3) = "STMXCSR"
'4-6 no instr
CFLX(7) = "CLFLUSH"

ReDim CMPS(7)
CMPS(0) = "CMPEQPS"
CMPS(1) = "CMPLTPS"
CMPS(2) = "CMPLEPS"
CMPS(3) = "CMPUNORDPS"
CMPS(4) = "CMPNEQPS"
CMPS(5) = "CMPNLTPS"
CMPS(6) = "CMPNLEPS"
CMPS(7) = "CMPORDPS"

ReDim MATH2(7)
MATH2(2) = "NOT"
MATH2(3) = "NEG"
MATH2(4) = "MUL"
MATH2(5) = "IMUL"
MATH2(6) = "DIV"
MATH2(7) = "IDIV"

ReDim LOOPX(3)
LOOPX(0) = "LOOPNE"
LOOPX(1) = "LOOPE"
LOOPX(2) = "LOOP"
LOOPX(3) = "JECXZ"

ReDim REG16O1(7)
REG16O1(0) = "BX+SI"
REG16O1(1) = "BX+DI"
REG16O1(2) = "BP+SI"
REG16O1(3) = "BP+DI"
REG16O1(4) = "SI"
REG16O1(5) = "DI"
REG16O1(6) = "BP"
REG16O1(7) = "BX"

End Sub
