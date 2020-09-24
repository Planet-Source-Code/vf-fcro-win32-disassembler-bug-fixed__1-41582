VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Win32 OPCode Disassembler written by:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Disassemble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rt1 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8916
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Load Win32 Executeable File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   $"Form1.frx":007F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   7680
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DISASM As New DISASM
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()
cd1.ShowOpen
If Len(cd1.FileName) = 0 Then Exit Sub
Open cd1.FileName For Binary As #1
ReDim FDATA(LOF(1) - 1)
Get #1, , FDATA
Close #1

If ReadPE(FDATA) = 0 Then MsgBox "File is not Win32 Executeable!", vbCritical, "Error!": Exit Sub

VIRTUALBASEADR = NTHEADER.OptionalHeader.ImageBase
EPOINT = NTHEADER.OptionalHeader.AddressOfEntryPoint



Dim Ln As Long
If FindByVirtual(EPOINT, Ln) = 1 Then
If FindByVirtual(NTHEADER.OptionalHeader.BaseOfCode, Ln) = 1 Then MsgBox "Cannot trace the executive Code!", vbCritical, "Error": Exit Sub
End If


Text1 = "Loaded File " & cd1.FileName & vbCrLf & _
"Executive code beginning at Virtual Address: " & Hex(VIRTUALBASEADR + VIRTUALADR) & "h" & vbCrLf & _
"In Section (OBJECT): " & Left(SECTIONSHEADER(u).nameSec, Ln) & vbCrLf & _
"Size Of Executive code: " & Hex(SIZEOFRAW) & "h" & vbCrLf & _
"Entry Point At: " & Hex(VIRTUALBASEADR + EPOINT) & "h"
End Sub
Private Function FindByVirtual(ByVal ValX As Long, ByRef Ln As Long) As Byte
Dim u As Long
'FIND EXECUTIVE SECTION?!
Dim FA As Long
Dim LA As Long
For u = 0 To UBound(SECTIONSHEADER)
FA = SECTIONSHEADER(u).VirtualAddress
LA = FA + SECTIONSHEADER(u).VirtualSize
If ValX >= FA And ValX <= LA Then GoTo FillIt
Next u
FindByVirtual = 1
Exit Function
FillIt:
Ln = lstrlen(ByVal SECTIONSHEADER(u).nameSec)
POINTERTORAW = SECTIONSHEADER(u).PointerToRawData
SIZEOFRAW = SECTIONSHEADER(u).SizeOfRawData
VIRTUALADR = SECTIONSHEADER(u).VirtualAddress
End Function

Private Sub Command2_Click()
Dim ARD As Long
ArrayDescriptor ARD, FDATA, 4
If ARD = 0 Then MsgBox "File isn't loaded yet!", vbCritical, "Error": Exit Sub

DISASM.BaseAddress = VIRTUALBASEADR + VIRTUALADR - POINTERTORAW

Dim Forward As Byte 'Next Instruction!
Dim CNT As Long
CNT = POINTERTORAW
Dim u As Long
Dim DATAS() As String
ReDim DATAS(SIZEOFRAW)



Dim TC2 As Long
Dim TC As Long

Label1 = "Disassemble..."
DoEvents
TC = GetTickCount
'DISASSEMBLE!!!
Do
DATAS(u) = DISASM.DisAssemble(FDATA, CNT, Forward, 1, 0) & vbCrLf
u = u + 1
CNT = CNT + Forward
Loop While SIZEOFRAW + POINTERTORAW > CNT

TC2 = GetTickCount



Label1 = "Disassembled for " & TC2 - TC & " msec"
DoEvents

ReDim Preserve DATAS(u - 1)


rt1 = Join(DATAS, "")
Erase DATAS

End Sub

Private Sub Command3_Click()
rt1 = ""
Unload Me
End Sub

Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub
