Attribute VB_Name = "PEModule"
Public Type IMAGEDOSHEADER
    e_magic As String * 2
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
   
End Type




Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type


Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
   OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public DOSHEADER As IMAGEDOSHEADER
Public NTHEADER As IMAGE_NT_HEADERS
Public SECTIONSHEADER() As IMAGE_SECTION_HEADER


Public Function ReadPE(DATA() As Byte) As Byte
On Error GoTo ErrX
Dim CNT As Long
Dim u As Long
CopyMemory DOSHEADER, DATA(CNT), Len(DOSHEADER)
If DOSHEADER.e_magic <> "MZ" Then Exit Function
CopyMemory NTHEADER, DATA(DOSHEADER.e_lfanew), Len(NTHEADER)
CNT = CNT + DOSHEADER.e_lfanew + Len(NTHEADER)
If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then Exit Function
ReDim SECTIONSHEADER(NTHEADER.FileHeader.NumberOfSections - 1)
For u = 0 To UBound(SECTIONSHEADER)
CopyMemory SECTIONSHEADER(u), DATA(CNT), Len(SECTIONSHEADER(0))
CNT = CNT + Len(SECTIONSHEADER(0))
Next u
ReadPE = 1
Exit Function
ErrX:
On Error GoTo 0
End Function


