VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSystem3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Revision 4
'
'Info:     These are basically macros for VB's built in file processes
'            this should streamline your code quite a bit and hopefully
'            remove alot of redundant coding.
'
'Author:   dzzie@yahoo.com
'Sight:    http://sandsprite.com
    
'Variant version of clsFso for scripting languages, added recurse option to getfolderfiles
' 5.21.12 - bugfix in RandomNum sporotic overflow time of day related...
' 11.14.13- bugfix in FileExists,FolderExists, GetFolderFiles recursive filter, GetParentFolder trailing \ bugfix
'  1.21.14- bugfix deletefile detects readonly attribute and removes
'  7.8.16 - bugfixs synced with current CFileSystem2
'  8.27.16 - !breaking change! -> GetFolderFiles/SubFolders now returns a collection
'  8.28.16 - GetSubFolders added recursive option, buffix to include hidden and hidden system folders..
'  8.28.16 - +GetDrives, DriveType, DriveTypeString, GetSpecialFolder, GetShortName,

Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Enum DriveTypes
    DRIVE_UNKNOWN = 0
    DRIVE_NOTFOUND = 1
    DRIVE_REMOVABLE = 2
    DRIVE_FIXED = 3
    DRIVE_REMOTE = 4
    DRIVE_CDROM = 5
    DRIVE_RAMDISK = 6
End Enum

Public Enum SpecialFolders
    
    sf_DESKTOP = &H0 '<desktop>
    'sf_INTERNET = &H1 'Internet Explorer (icon on desktop)
    sf_PROGRAMS = &H2 'Start Menu\Programs
    'sf_CONTROLS = &H3'My Computer\Control Panel
    'sf_PRINTERS = &H4'My Computer\Printers
    sf_PERSONAL = &H5 'My Documents
    sf_FAVORITES = &H6 '<user name>\Favourites
    sf_STARTUP = &H7 'Start Menu\Programs\Startup
    sf_RECENT = &H8 '<user name>\Recent
    sf_SENDTO = &H9 '<user name>\SendTo
    sf_BITBUCKET = &HA '<desktop>\Recycle Bin
    sf_STARTMENU = &HB '<user name>\Start Menu
'    sf_MYDOCUMENTS = &HC'logical "My Documents" desktop icon
    sf_MYMUSIC = &HD '"My Music" folder
    sf_MYVIDEO = &HE '"My Videos" folder
    sf_DESKTOPDIRECTORY = &H10 '<user name>\Desktop
    'sf_DRIVES = &H11'My Computer
    'sf_NETWORK = &H12'Network Neighborhood (My Network Places)
'    sf_NETHOOD = &H13'<user name>\nethood
    sf_FONTS = &H14 'windows\fonts
'    sf_TEMPLATES = &H15'templates
    sf_COMMON_STARTMENU = &H16 'All Users\Start Menu
'    sf_COMMON_PROGRAMS = &H17 'All Users\Start Menu\Programs
    sf_COMMON_STARTUP = &H18 'All Users\Startup
    sf_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
    sf_APPDATA = &H1A '<user name>\Application Data
'    sf_PRINTHOOD = &H1B'<user name>\PrintHood
    sf_LOCAL_APPDATA = &H1C '<user name>\Local Settings\Application Data (non roaming)
 '   sf_ALTSTARTUP = &H1D'non localized startup
    'non localized common startup
'    sf_COMMON_ALTSTARTUP = &H1E
'    sf_COMMON_FAVORITES = &H1F
'    sf_INTERNET_CACHE = &H20
'    sf_COOKIES = &H21
'    sf_HISTORY = &H22
    'All Users\Application Data
'    sf_COMMON_APPDATA = &H23
    sf_WINDOWS = &H24 'GetWindowsDirectory()
    sf_SYSTEM = &H25 'GetSystemDirectory()
    sf_PROGRAM_FILES = &H26 'C:\Program Files
    sf_MYPICTURES = &H27 'C:\Program Files\My Pictures
    sf_PROFILE = &H28 'USERPROFILE
'    'x86 system directory on RISC
'    sf_SYSTEMX86 = &H29
'    'x86 C:\Program Files on RISC
'    sf_PROGRAM_FILESX86 = &H2A
'    'C:\Program Files\Common
'    sf_PROGRAM_FILES_COMMON = &H2B
'    'x86 Program Files\Common on RISC
'    sf_PROGRAM_FILES_COMMONX86 = &H2C
'     'All Users\Templates
'    sf_COMMON_TEMPLATES = &H2D
'     'All Users\Documents
'    sf_COMMON_DOCUMENTS = &H2E
'    'All Users\Start Menu\Programs\Administrative Tools
'    sf_COMMON_ADMINTOOLS = &H2F
'    '<user name>\Start Menu\Programs\Administrative Tools
'    sf_ADMINTOOLS = &H30
'    'Network and Dial-up Connections
'    sf_CONNECTIONS = &H31
'    'All Users\My Music
'    sf_COMMON_MUSIC = &H35
'    'All Users\My Pictures
    sf_COMMON_PICTURES = &H36
'    'All Users\My Video
    sf_COMMON_VIDEO = &H37
'    'Resource Directory
'    sf_RESOURCES = &H38
'    'Localized Resource Directory
'    sf_RESOURCES_LOCALIZED = &H39
'    'Links to All Users OEM specific apps
'    sf_COMMON_OEM_LINKS = &H3A
'    'USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
'    sf_CDBURN_AREA = &H3B
'    'unused                                      = &H3C
'    'Computers Near Me (computered from Workgroup membership)
'    sf_COMPUTERSNEARME = &H3D
End Enum

'Public dlg As New CCmnDlg

'Function FileProperties(ByVal fPath As String, Optional checkSignature As Boolean = False) As CFileProperties
'    Dim f As New CFileProperties
'    f.LoadFileInfo fPath, checkSignature
'    Set FileProperties = f
'End Function

Function GetFolderFiles(folderPath, Optional Filter = "*", Optional retFullPath = True, Optional recursive = False) As Collection
   Dim fnames As New Collection
   Dim folders As Collection
   Dim tmp As New Collection
   Dim i As Long
   Dim fs, f
   
   Set GetFolderFiles = fnames
   If Not FolderExists(folderPath) Then Exit Function
   folderPath = IIf(Right(folderPath, 1) = "\", folderPath, folderPath & "\")
   
   fs = Dir(folderPath & Filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then fnames.Add IIf(retFullPath, folderPath & fs, fs)
     fs = Dir()
   Wend
   
   If recursive Then
        Set folders = GetSubFolders(folderPath)
        For Each f In folders
            FolderEngine f, tmp, Filter
        Next
        If Not retFullPath Then
            For Each f In tmp
                fnames.Add VBA.Replace(f, folderPath, Empty)    'make relative path from base
            Next
        End If
    End If
    
End Function

Private Sub FolderEngine(fldrpath, c As Collection, Optional Filter = "*")

    Dim files As Collection
    Dim folders As Collection
    Dim i As Long
    Dim f
    
    Set files = GetFolderFiles(fldrpath, Filter)
    Set folders = GetSubFolders(fldrpath)
        
    For Each f In files
        c.Add f
    Next
    
    For Each f In folders
        FolderEngine f, c, Filter
    Next
    
End Sub
 
    
Function GetSubFolders(folderPath, Optional retFullPath = True, Optional recursive = False) As Collection
   Dim fnames As New Collection
   Dim tmp As New Collection
   Dim tmp2 As New Collection
   Dim fd, atr
   Const ATTR_HIDDEN = 2
   On Error Resume Next
   
   Set GetSubFolders = fnames
    
   If Not FolderExists(folderPath) Then Exit Function
   If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

   fd = Dir(folderPath, vbDirectory Or vbHidden Or vbSystem)
   While fd <> ""
        If Left(fd, 1) <> "." Then
            atr = GetAttr(folderPath & fd) 'things like pagefile.sys that are locked can throw an error...
            If (atr And vbDirectory) = vbDirectory Then
               tmp.Add folderPath & fd 'always add full path we filter latter..(needed for recursive option)
            End If
        End If
        fd = Dir()
   Wend
   
   If recursive Then
        For Each fd In tmp
            SubFoldersInternal tmp2, fd
        Next
   End If

   For Each fd In tmp
        If retFullPath = True Then
            fnames.Add fd
        Else
            fnames.Add VBA.Replace(fd, folderPath, Empty) 'make relative to parent dir
        End If
   Next
        
   For Each fd In tmp2
        If retFullPath = True Then
            fnames.Add fd
        Else
            fnames.Add VBA.Replace(fd, folderPath, Empty) 'make relative to parent dir
        End If
   Next
   
End Function

Private Sub SubFoldersInternal(pc As Collection, parentDir)
    Dim c As New Collection, d
    Set c = GetSubFolders(parentDir)
    For Each d In c
        pc.Add d
        SubFoldersInternal pc, d
        DoEvents
    Next
End Sub


Function FolderExists(path) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetParentFolder(path)
    Dim tmp() As String
    Dim my_path
    Dim ub As String
    
    On Error GoTo hell
    If Len(path) = 0 Then Exit Function
    
    my_path = path
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
    
    tmp = Split(CStr(my_path), "\")
    tmp(UBound(tmp)) = Empty
    my_path = Replace(Join(tmp, "\"), "\\", "\")
    If VBA.Right(my_path, 1) = "\" Then my_path = Mid(my_path, 1, Len(my_path) - 1)
    
    GetParentFolder = my_path
    Exit Function
    
hell:
    GetParentFolder = Empty
    
End Function

Function CreateFolder(path)
   On Error GoTo blah
   If FolderExists(path) Then Exit Function
   MkDir path
   If Not FolderExists(path) Then Exit Function
   CreateFolder = True
blah:
End Function

'web or file path
Function FileNameFromPath(fullpath)
    Dim tmp() As String
    Dim marker As String
    marker = "\"
    If InStr(fullpath, "/") > 0 Then marker = "/"
    If InStr(fullpath, marker) > 0 Then
        tmp = Split(CStr(fullpath), marker)
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = CStr(fullpath) 'maybe its just a file name already..
    End If
End Function

Function DeleteFile(fPath)
 On Error GoTo hadErr
    
    Dim attributes As VbFileAttribute

    attributes = GetAttr(fPath)
    If (attributes And vbReadOnly) Then
        attributes = attributes - vbReadOnly
        SetAttr fPath, attributes
    End If
    
    Kill fPath
    DeleteFile = True
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Sub Rename(fullpath, newName)
  Dim pf
  pf = GetParentFolder(fullpath)
  Name fullpath As pf & "\" & newName
End Sub

Sub SetAttribute(fPath, it As VbFileAttribute)
   SetAttr fPath, it
End Sub

Function GetExtension(path)
    Dim tmp() As String
    Dim ub
    tmp = Split(CStr(path), "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       GetExtension = ""
    End If
End Function

Function GetBaseName(path)
    Dim tmp() As String
    Dim ub
    tmp = Split(CStr(path), "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Function ChangeExt(path, ext)
    Dim bn
    ext = IIf(Left(ext, 1) = ".", ext, "." & ext)
    If FileExists(path) Then
        Rename path, GetBaseName(path) & ext
    Else
        'hack to just accept a file name might not be worth supporting
        bn = Mid(path, 1, InStr(1, path, ".") - 1)
        ChangeExt = bn & ext
    End If
End Function

Function SafeFileName(proposed)
  Dim badChars, bad() As String, i As Long
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = VBA.Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function GetFreeFileName(parentDir, Optional extension = ".txt")
    
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String
    Dim folder As String
    
    folder = parentDir
    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Function GetFreeFolderName(ByVal parentFolder, Optional prefix = "")
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(parentFolder) Then Exit Function
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"
        
again:
    Do
      tmp = parentFolder & prefix & RandomNum()
    Loop Until Not FolderExists(tmp)
    
    GetFreeFolderName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Function buildPath(folderPath)
    On Error GoTo oops
    
    If FolderExists(folderPath) Then buildPath = True: Exit Function
    
    Dim tmp() As String, build, i As Long
    
    tmp = Split(CStr(folderPath), "\")
    build = tmp(0)
    For i = 1 To UBound(tmp)
        build = build & "\" & tmp(i)
        If InStr(tmp(i), ".") < 1 Then
            If Not FolderExists(build) Then CreateFolder (build)
        End If
    Next
    buildPath = True
    Exit Function
oops: buildPath = False
End Function


Function ReadFile(filename) As Variant
  On Error GoTo err_readFile
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
  Open filename For Binary As #f        ' Open file.(can be text or image)
  temp = Input(FileLen(filename), #f)   ' Get entire Files data
err_readFile:
  Close #f
  ReadFile = temp
  If Err.Number <> 0 Then Log "ERROR during fso.ReadFile: " & Err.Description
End Function

Sub writeFile(path, it)
    On Error GoTo err_writeFile
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    
err_writeFile:
   Close f
   If Err.Number <> 0 Then Log "ERROR during fso.writeFile: " & Err.Description
End Sub

Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Function Copy(fPath, toFolder)
   Dim baseName, newName
   On Error Resume Next
   If FolderExists(toFolder) Then
       baseName = FileNameFromPath(fPath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       newName = toFolder & baseName
       If FileExists(newName) Then Kill newName
       FileCopy fPath, newName
       Copy = newName
   Else 'assume tofolder is actually new desired file path
       FileCopy fPath, toFolder
       Copy = toFolder
   End If
End Function

Function Move(fPath, toFolder)
    Dim fName As String
    fName = FileNameFromPath(fPath)
    toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
    
    Name fPath As toFolder & fName
    Move = toFolder & fName
    
End Function

Function CreateFile(fPath)
    On Error GoTo hell
    Dim f As Long
    f = FreeFile
    If FileExists(fPath) Then Exit Function
    Open fPath For Binary As f
    Close f
    If FileExists(fPath) Then CreateFile = True
hell:
End Function


Function DeleteFolder(folderPath, Optional force = True)
 On Error GoTo failed
   Call delTree(folderPath, force)
   Call RmDir(folderPath)
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderPath, Optional force = True)
   Dim sfi As Collection, sfo As Collection, i As Integer, f
   
   Set sfi = GetFolderFiles(folderPath)
   Set sfo = GetSubFolders(folderPath)
   
   If sfi.Count > 0 And force = True Then
        For Each f In sfi
            DeleteFile f
        Next
   End If
   
   If sfo.Count > 0 And force = True Then
        For Each f In sfo
            Call DeleteFolder(f, True)
        Next
   End If
   
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function FolderName(folderPath)
    Dim ret, tmp() As String
    tmp = Split(CStr(folderPath), "\")
    If Not AryIsEmpty(tmp) Then
        If Len(tmp(UBound(tmp))) <> 0 Then ret = tmp(UBound(tmp)) _
        Else ret = tmp(UBound(tmp) - 1)
    Else
        ret = CStr(folderPath)
    End If
    FolderName = ret
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 500
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name !!
    If Not FileExists(sFile) Then
        writeFile sFile, ""
        iCreated = True
    End If
    
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then
        GetShortName = sFile
    End If
    
    If iCreated Then Kill sFile

End Function

Function GetDrives() As Collection
    Dim c As New Collection
    Dim strDrives As String
    Dim result As Long
    Dim pos As Long
    Dim drive As String
    
    Set GetDrives = c
    
    'get requireed buffer length and create buffer
    result = GetLogicalDriveStrings(0, strDrives)
    strDrives = String(result, Chr(0))

    result = GetLogicalDriveStrings(Len(strDrives), strDrives)
    If result = 0 Then Exit Function
    
    ' Walk through the logical drive string and display the drive
    ' letters. The logical drive string is a null seperated
    ' double null terminated string.
    pos = 1
    Do While Not Mid$(strDrives, pos, 1) = Chr(0)
        drive = Mid$(strDrives, pos, 3)
        pos = pos + 4
        c.Add UCase(drive) & " " & DriveTypeString(drive)
    Loop
    
End Function

'supports c, c:, c:\, and c:\path
Function DriveType(ByVal drive) As DriveTypes
    Dim a As Long
    
    If Len(drive) = 1 Then drive = drive & ":"
    
    a = InStr(drive, "\")
    If a > 2 And Len(drive) > a Then
        drive = Mid(drive, 1, a)
    End If
    
    DriveType = GetDriveType(drive)
    
End Function

Function DriveTypeString(ByVal drive) As String
   Dim dt As DriveTypes
   
   dt = DriveType(drive)
   
   Select Case dt
      Case 0: DriveTypeString = "cannot be determined"
      Case 1: DriveTypeString = "does not exist"

      Case DRIVE_REMOVABLE:
          Select Case Left$(drive, 1)
              Case "a", "b": DriveTypeString = "Floppy drive"
              Case Else: DriveTypeString = "Removable drive"
          End Select

      Case DRIVE_FIXED:   DriveTypeString = "Hard drive"
      Case DRIVE_REMOTE:  DriveTypeString = "Network drive"
      Case DRIVE_CDROM:   DriveTypeString = "Optical drive"
      Case DRIVE_RAMDISK: DriveTypeString = "RAM disk"
   End Select
  
End Function

Function GetSpecialFolder(sf As SpecialFolders) As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, sf, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      GetSpecialFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
  
End Function



Public Function FileSize(fPath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fPath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function


