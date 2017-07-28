Attribute VB_Name = "Pe_info_bas"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type Section
    SectionName          As String * 8
    VirtualSize          As Long
    RVAOffset            As Long
    RawDataSize          As Long
    PointertoRawData     As Long
    PointertoRelocs      As Long
    PointertoLineNumbers As Long
    NumberofRelocs       As Integer
    NumberofLineNumbers  As Integer
    SectionFlags         As Long
End Type

Public Type PE_Header
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  Reserved                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackCommitSize                As Long
  HeapReserveSize                As Long
  HeapCommitSize                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  
  BoundImportsAddress            As Long
  BoundImportsAddressSize        As Long
  IATAddress                     As Long
  IATAddressSize                 As Long

  DelayImportAddress             As Long
  DelayImportAddressSize         As Long
  COMDescriptorAddress           As Long
  COMDescriptorAddressSize       As Long
  
  ReservedAddress                As Long
  ReservedAddressSize            As Long
  
'  Gap                            As String * &H28&
  Sections(64)                   As Section
End Type




Public Type PE_Header64
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
'  ImageBase64                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  Reserved                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackReservedSize64              As Long
  StackCommitSize                As Long
  StackCommitSize64                As Long
  HeapReserveSize                As Long
  HeapReserveSize64                As Long
  HeapCommitSize                 As Long
  HeapCommitSize64                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  
  BoundImportsAddress            As Long
  BoundImportsAddressSize        As Long
  IATAddress                     As Long
  IATAddressSize                 As Long

  DelayImportAddress             As Long
  DelayImportAddressSize         As Long
  COMDescriptorAddress           As Long
  COMDescriptorAddressSize       As Long
  
  ReservedAddress                As Long
  ReservedAddressSize            As Long
  
'  Gap                            As String * &H28&
  Sections(64)                   As Section
End Type


' ------- Additional API declarations ---------------
Public Const IMAGE_ORDINAL_FLAG = &H80000000

Type IMAGE_IMPORT_BY_NAME
   Hint As Integer
   ImpName As String * 254
End Type

Type IMAGE_IMPORT_DESCRIPTOR
   OriginalFirstThunk As Long
   TimeDateStamp As Long
   ForwarderChain As Long
   pDllName As Long
   FirstThunk As Long
End Type

Type IMAGE_BASE_RELOCATION
   VirtualAddress As Long
   SizeOfBlock As Long
End Type


Public IMAGE_IMPORT_DESCRIPTOR As IMAGE_IMPORT_DESCRIPTOR
Public IMAGE_BASE_RELOCATION As IMAGE_BASE_RELOCATION
Public IMAGE_IMPORT_BY_NAME As IMAGE_IMPORT_BY_NAME


'assumption the .text Sections ist the first and .data Section the second in pe_header
Public Const TEXT_SECTION& = 0
Public Const DATA_SECTION& = 1

'Public PE_info As New PE_info
Public PE_Header As PE_Header
Public PE_Header64 As PE_Header64

Public IsPE64 As Boolean
Public PE_SectionData As Collection
Private PEStart&

'these functions were in a class, but they are used to set these public structs so keeping all in one place...

Public Function VAToRaw(VA As Long) As Long
   Dim i&, RVA&
   RVA = VA - PE_Header.ImageBase
   
   VAToRaw = -1
   
  'find section
   For i = 0 To PE_Header.NumberofSections - 1
      
      With PE_Header.Sections(i)
         If RangeCheck(RVA, .RVAOffset + .VirtualSize, .RVAOffset) Then
            VAToRaw = .PointertoRawData + (RVA - .RVAOffset)
            Exit For
         End If
      End With
   
   Next

End Function

Public Function RVAToRaw(RVA As Long) As Long
   Dim i&
  
   RVAToRaw = -1
   
  'find section
   For i = PE_Header.NumberofSections - 1 To 0 Step -1
      
      With PE_Header.Sections(i)
         If RangeCheck(RVA, .RVAOffset + .VirtualSize, .RVAOffset) Then
            RVAToRaw = .PointertoRawData + (RVA - .RVAOffset)
            Exit For
         End If
      End With
   
   Next

End Function

Public Sub OpenPE()

'     '--- find PE-signature ---
'     'Get First 0x400 Bytes
'      Dim tmpstr$
'      file.Position = 0
'      tmpstr = file.FixedString(&H400)
'
'     'Locate start of PE-header
'      PEStart = InStr(1, tmpstr, "PE" & vbNullChar & vbNullChar, vbBinaryCompare)
'      If PEStart = 0 Then err.Raise vbObjectError Or 1, , "No PE-Header Found"
    
     '--- find PE-signature ---
     'Check DOS Header
      Dim tmpstr$
      File.Position = 0
     
     'to skip the Error in VB-IDE Rightclick Toggle/Break on unhandled errors
     'MZ DOS-Header->e_magic
      If File.int16 <> &H5A4D Then Err.Raise vbObjectError Or 1, , "No ExeFile DOS-Header.e_magic<>""MZ"""

     'Locate & Validate PE-header
      File.Position = &H3C '   DOS-Header->e_lfanew
      PEStart = File.int32
      File.Position = PEStart
      PEStart = PEStart + 1
      
      If File.int32 <> &H4550 Then Err.Raise vbObjectError Or 2, , "No ExeFile 'PE-Header.Signature<>""PE"""
    
    '  --- get PE_Header  ---
      Dim hFile&
      hFile = m_FreeFile("peheader")
      Open File.filename For Binary Access Read As #hFile
      Get hFile, PEStart, PE_Header
      m_Close hFile, "peheader"
      
      
    ' Validate Machine Type
      If PE_Header.Machine <> &H14C Then
         If PE_Header.Machine = &H8664 Then
'            Err.Raise vbObjectError Or 4, , "PE-Header.Signature=HDR64_MAGIC!"
            IsPE64 = True
            
            '  --- get PE_Header64  ---
            hFile = m_FreeFile("machinetype")
            Open File.filename For Binary Access Read As #hFile
            Get hFile, PEStart, PE_Header64
            m_Close hFile, "machinetype"
  
         Else
           Err.Raise vbObjectError Or 3, , "Unsupported PE-Header.Signature 0x" & H16(PE_Header.Machine) & " <>I386(0x14C)."
         End If
      
      Else
         If PE_Header.OptionalHeaderSize <> &HE0 Then
            Err.Raise vbObjectError Or 5, , "PE_Header.OptionalHeaderSize = E0 expected but curvalue is " & H32(PE_Header.OptionalHeaderSize)
         
         End If
   
      End If
    
End Sub


Function AlignForFile(value&)
   AlignForFile = Align(value, PE_Header.FileAlignment)
End Function

Function AlignForSection(value&)
   AlignForSection = Align(value, PE_Header.SectionAlignment)
End Function


Function Align&(value&, alignment&)
   '                  keep if equal  |  round up to next bounary
   Align = IIf(value, alignment * (((value - 1) \ alignment) + 1), 0)
End Function


