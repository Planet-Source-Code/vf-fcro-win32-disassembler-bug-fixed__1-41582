Attribute VB_Name = "GLOBALDATA"
Declare Sub ArrayDescriptor Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc() As Any, ByVal ByteLen As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Public FDATA() As Byte 'ENTIRE FILE DATA

'************************
'EXECUTEABLE SECTION
Public POINTERTORAW As Long
Public SIZEOFRAW As Long
Public VIRTUALADR As Long
'************************

Public EXESECTION As String

Public VIRTUALBASEADR As Long
Public EPOINT As Long




