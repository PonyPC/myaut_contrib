Attribute VB_Name = "GlobalDefs"
Option Explicit

Public FILE As New FileStream
Public filename As New ClsFilename
Public ExtractedFiles As Collection

Global opts As New CCommandlineOptions
Global fso As New CFileSystem3
Global isAutomationRun As Boolean

Public Enum ErrReturns
    ERR_NO_AUT_EXE& = vbObjectError Or &H10
    ERR_NO_OBFUSCATE_AUT& = vbObjectError Or &H20
    ERR_NO_TEXTFILE& = vbObjectError Or &H30
    NO_AUT_DE_TOKEN_FILE& = &H100
    ERR_CANCEL_ALL& = vbObjectError Or &H1000
    er_SUCCESS = 0
End Enum

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long


'dll missing? called from frmMain.mi_CustomDecrypt
'http://archive.ec/Kiqd1#selection-1345.41-1345.51
Private Declare Function myCryptCall Lib "hwindr.dll" Alias "_deCode" _
   (ByVal Key1 As String, _
    ByVal Key2 As String) _
   As Long
 
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 
 

Public Const StringBody_SingleQuoted As String = "[^']*"
Public Const String_SingleQuoted = "(?:'" & StringBody_SingleQuoted & "')+"
Public Const StringBody_DoubleQuoted As String = "[^""]*"
Public Const String_DoubleQuoted As String = "(?:""" & StringBody_DoubleQuoted & """)+"
Public Const StringPattern As String = String_DoubleQuoted & "|" & String_SingleQuoted


Public Const DE_OBFUSC_TYPE_NOT_OBFUSC& = &H0
Public Const DE_OBFUSC_TYPE_VANZANDE& = &H10000
Public Const DE_OBFUSC_TYPE_ENCODEIT& = &H20000
Public Const DE_OBFUSC_TYPE_CHR_ENCODE& = &H10
Public Const DE_OBFUSC_TYPE_CHR_ENCODE_OLD& = &H8


Public Const DE_OBFUSC_VANZANDE_VER14& = &H10014
Public Const DE_OBFUSC_VANZANDE_VER15& = &H10015
Public Const DE_OBFUSC_VANZANDE_VER15_2& = &H100152
Public Const DE_OBFUSC_VANZANDE_VER24& = &H10024


Private startTime As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private fHandleLog As Long

Private Const DEBUG_FHANDLES As Boolean = False 'i cant find the damn leak...screw it..
Private Const DEBUG_FH_FILT As String = "" '".exe"

'this gives us a place to set a breakpoint to find all native Open calls
Function m_FreeFile(Optional desc As String) As Long
    m_FreeFile = FreeFile
    fHandleDebugLog pad(desc, 20) & "  Open: " & m_FreeFile
End Function

Function m_Close(h As Long, Optional desc As String)
    Close h
    fHandleDebugLog pad(desc, 20) & " Close: " & h
End Function
    

Function fHandleDebugLogOpen()
    
    On Error Resume Next
    
    If Not DEBUG_FHANDLES Then Exit Function
    
    Dim p As String
    p = App.path & "\FileHandleLog.txt"
    If fso.FileExists(p) Then Kill p
    
    fHandleLog = FreeFile
    Open p For Output As fHandleLog
    
    If Err.Number <> 0 Then
        fHandleLog = 0
    Else
        Print #fHandleLog, "Opening new file handle log " & Now
    End If
    
End Function

Function fHandleDebugLog(ByVal x As String)
    On Error Resume Next
    If fHandleLog = 0 Then Exit Function
    If Len(DEBUG_FH_FILT) > 0 Then
        If InStr(1, x, DEBUG_FH_FILT, vbTextCompare) < 1 Then Exit Function
    End If
    Print #fHandleLog, x
End Function

Function fHandleDebugLogClose()
    On Error Resume Next
    If fHandleLog = 0 Then Exit Function
    Print #fHandleLog, "Closing file handle log " & Now
    Close fHandleLog
    fHandleLog = 0
End Function



Sub StartBenchMark(Optional ByRef t As Long)
    If IsMissing(t) Then
        startTime = GetTickCount()
    Else
        t = GetTickCount()
    End If
End Sub

Function EndBenchMark(Optional ByRef t As Long) As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    If IsMissing(t) Then
        loadTime = endTime - startTime
    Else
        loadTime = endTime - t
    End If
    EndBenchMark = loadTime / 1000 & " seconds"
End Function

Sub FL_verbose(Text)
   FrmMain.FL_verbose Text
End Sub

Sub log_verbose(TextLine$)
   FrmMain.log_verbose TextLine
End Sub

Sub FL(Text)
   FrmMain.FL Text
End Sub


''/////////////////////////////////////////////////////////
''// log -Add an entry to the Log
Public Sub Log(TextLine$, Optional LinePrefix$)
   FrmMain.Log TextLine, LinePrefix
End Sub

'/////////////////////////////////////////////////////////
Public Sub Log_Clear()
   FrmMain.ListLog.Clear
End Sub



Public Sub GUIEvent_ProcessBegin(Target&, Optional BarLevel& = 0, Optional Skipable As Boolean = False)
   FrmMain.GUIEvent_ProcessBegin Target, BarLevel, Skipable
End Sub

Public Sub GUIEvent_ProcessUpdate(CurrentValue&, Optional BarLevel& = 0)
   FrmMain.GUIEvent_ProcessUpdate CurrentValue, BarLevel
End Sub
Public Sub GUIEvent_ProcessEnd(Optional BarLevel& = 0)
   FrmMain.GUIEvent_ProcessEnd BarLevel
End Sub

Public Sub GUIEvent_Increase(PerCentToIncrease As Double, Optional BarLevel& = 0)
   FrmMain.GUIEvent_Increase PerCentToIncrease, BarLevel
End Sub

Public Sub GUI_SkipEnable()
   FrmMain.Cmd_Skip.Visible = True
   If FrmMain.bCmd_Skip_HasFocus = False Then
      FrmMain.Cmd_Skip.SetFocus
      FrmMain.bCmd_Skip_HasFocus = True
   End If
End Sub

Public Sub GUI_SkipDisable()
   FrmMain.Cmd_Skip.Visible = False
   FrmMain.bCmd_Skip_HasFocus = False
End Sub



Sub DoEventsSeldom()
   If Rnd < 0.01 Then myDoEvents
End Sub

Sub DoEventsVerySeldom()
   If (GetTickCount() And &H7F) = 1 Then
'   If Rnd < 0.00001 Then
       myDoEvents
   End If
End Sub

Sub ShowScript(ScriptData$)
   
   FrmMain.Txt_Script = Script_RawToText(ScriptData)

End Sub

Function Script_RawToText(ByRef ScriptData$) As String

   If isUTF16(ScriptData) Then
      Script_RawToText = StrConv((Mid(ScriptData, 1 + Len(UTF16_BOM))), vbFromUnicode)
   ElseIf isUTF8(ScriptData) Then
      Script_RawToText = Mid(ScriptData, 1 + Len(UTF8_BOM))
   Else
      Script_RawToText = ScriptData
   End If

End Function

Sub SaveScriptData(ScriptData$, Optional skipTidy As Boolean)

   With FrmMain
      
   ' Not need anymore since Tidy v2.0.24.4 November 30, 2008
'   ' Adding a underscope '_' for lines longer than 2047
'   ' so Tidy will not complain
'      FrmMain.Log "Try to breaks very long lines (about 2000 chars) by adding '_'+<NewLine> ..."
'      ScriptData = AddLineBreakToLongLines(Split(ScriptData, vbCrLf))
      
       ' overwrite script
         If FrmMain.Chk_TmpFile.value = vbChecked Then
            filename.Name = filename.Name & "_restore"
            .Log "Saving script to: " & filename.filename
         Else
   '         FileDelete FileName.Name
            .Log "Save/overwrite script to: " & filename.filename
         End If
   
         fso.writeFile filename.filename, ScriptData
      
      End With
      
      RunTidy ScriptData, skipTidy
End Sub

Public Sub RunTidy(ScriptData$, Optional skipTidy As Boolean)
   
   With FrmMain
        
      ShowScript ScriptData
      .Log ""
     
      If skipTidy Then
         .Log "Skipping to run 'data\Tidy\Tidy.exe' on" & filename.NameWithExt & "' to improve sourcecode readability. (Plz run it manually if you need it.)"
      Else
         
         .Log "Running 'Tidy.exe " & filename.NameWithExt & "' to improve sourcecode readability."
         
         FrmMain.ScriptLines = Split(ScriptData, vbCrLf)
         
         Dim cmdline$, parameters$, Logfile$
         cmdline = App.path & "\" & "data\Tidy\Tidy.exe"
         parameters = """" & filename & """" ' /KeepNVersions=1
         .Log cmdline & " " & parameters
         
         Dim TidyExitCode&
         
         'Dim ConsoleOut$
         'ConsoleOut =
         FrmMain.Console.ShellExConsole cmdline, parameters, TidyExitCode
         
         
         If TidyExitCode = 0 Then
             .Log "=> Okay (ExitCode: " & TidyExitCode & ")."
             Dim TidyBackupFileName As New ClsFilename
             TidyBackupFileName.mvarFileName = filename.mvarFileName
             TidyBackupFileName.Name = TidyBackupFileName.Name & "_old1"
             
           ' Delete Tidy BackupFile
             If FrmMain.Chk_TmpFile.value = vbUnchecked Then
                .Log "Deleting Tidy BackupFile..." ' & TidyBackupFileName.NameWithExt
                FileDelete TidyBackupFileName.filename
             End If
            
            
          ' Readin tidy file
            ScriptData = fso.ReadFile(filename.filename)
          
            ShowScript ScriptData
            
         Else
            .Log "=> Error (ExitCode: " & TidyExitCode & ")" ' TidyOutput >>>"
'            .Log ConsoleOut, "TIDY OUTPUT: "
'            .Log "<<<"
            .Log "Attention: Tidy.exe failed. Deobfucator will probably also fail because scriptfile is not in proper format."
         End If
         
      End If 'skip tidy
      
   End With
End Sub

Public Sub RunTidy2(ScriptFile)
   
   With FrmMain
        
       .Log ""
              
         .Log "Running 'Tidy.exe " & filename.NameWithExt & "' to improve sourcecode readability."
                 
         Dim cmdline$, parameters$, Logfile$
         cmdline = App.path & "\" & "data\Tidy\Tidy.exe"
         parameters = """" & ScriptFile & """" ' /KeepNVersions=1
         .Log cmdline & " " & parameters
         
         Dim TidyExitCode&
         
         'Dim ConsoleOut$
         'ConsoleOut =
         FrmMain.Console.ShellExConsole cmdline, parameters, TidyExitCode
         
         
         If TidyExitCode = 0 Then
             .Log "=> Okay (ExitCode: " & TidyExitCode & ")."
             Dim TidyBackupFileName As New ClsFilename
             TidyBackupFileName.mvarFileName = filename.mvarFileName
             TidyBackupFileName.Name = TidyBackupFileName.Name & "_old1"
             
           ' Delete Tidy BackupFile
             If FrmMain.Chk_TmpFile.value = vbUnchecked Then
                .Log "Deleting Tidy BackupFile..." ' & TidyBackupFileName.NameWithExt
                FileDelete TidyBackupFileName.filename
             End If
            
           'this block dz
           Dim tmp As String
           tmp = fso.ReadFile(filename.filename)
           tmp = Replace(tmp, "EndFunc" & vbCrLf, "EndFunc" & vbCrLf & vbCrLf)
           fso.writeFile filename.filename, tmp
           
          ' Readin tidy file
            ShowScript tmp
            
         Else
            .Log "=> Error (ExitCode: " & TidyExitCode & ")" ' TidyOutput >>>"
'            .Log ConsoleOut, "TIDY OUTPUT: "
'            .Log "<<<"
            .Log "Attention: Tidy.exe failed. Deobfucator will probably also fail because scriptfile is not in proper format."
         End If
         
     
      
   End With
End Sub

Public Function canAttemptUnUPX(exe) As Boolean
    Dim b() As Byte, f As Long, tmp As String
        
    On Error Resume Next
    
    If Not FileExists(exe) Then Exit Function
    
    FrmMain.Log "Testing to see if exe is UPX compressed..."

    ReDim b(&H1000)
    f = FreeFile
    Open exe For Binary As f
    Get f, , b()
    Close f
     
    If b(0) <> Asc("M") And b(1) <> Asc("Z") Then
       FrmMain.Log "canAttemptUnUPX: No MZ header found exiting..."
       Exit Function
    End If
    
    tmp = StrConv(b, vbUnicode, &H409)
    If InStr(1, tmp, "UPX") < 1 Then
       FrmMain.Log "canAttemptUnUPX: No UPX marker found exiting..."
       Exit Function
    End If
         
    canAttemptUnUPX = True
    
End Function

Public Function UnUPX(exe, ByRef outFile As String) As Boolean
   
   Dim cmdline$, parameters$, Logfile$
   Dim b() As Byte, f As Long, tmp As String
   
   With FrmMain
         outFile = exe & ".unupx"
         
         If LCase(Right(exe, 6)) = ".unupx" Then
            UnUPX = True
            Exit Function
         End If
         
         If FileExists(outFile) Then
            .Log "Found previously UPX decompressed file using that..."
            UnUPX = True
            Exit Function
         End If
         
         If Not canAttemptUnUPX(exe) Then Exit Function
         
         .Log "Trying to decompress UPX binary..."

         cmdline = App.path & "\data\upx.exe"
         If Not FileExists(cmdline) Then
            .Log "upx.exe binary not found!"
            .Log cmdline
            Exit Function
         End If
         
         parameters = "-d """ & exe & """ -o """ & outFile & """"
         .Log cmdline & " " & parameters
         
         Dim upxExitCode&
         
         'Dim ConsoleOut$
         'ConsoleOut =
         FrmMain.Console.ShellExConsole cmdline, parameters, upxExitCode
         
         If upxExitCode = 0 And FileExists(outFile) Then
             .Log "=> UPX Decompress Okay!"
             UnUPX = True
         Else
            .Log "=> Error (ExitCode: " & upxExitCode & ")" ' UPXOutput >>>"
'            .Log ConsoleOut, "UPX OUTPUT: "
'            .Log "<<<"
            .Log "Attention: upx decompress failed. May be modified..."
         End If
         
   End With
   
End Function

Public Function ADLER32$(Data As StringReader)
   With Data
'            Dim a
            
            Dim L&, h&
            h = 0: L = 1
'            a = GetTickCount
' taken out for performance reason
'               .EOS = False
'               .DisableAutoMove = False
'               Do Until .EOS
'                 'The largest prime less than 2^16
'                  l = (.int8 + l) Mod 65521 '&HFFF1
'                  H = (H + l) Mod 65521 '&HFFF1
'                  If (l And 8) Then myDoEvents
'               Loop
'
'            Debug.Print "a: ", GetTickCount - a 'Benchmark: 20203

 '           a = GetTickCount
               
               Dim StrCharPos&, tmpBuff$
               tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID)
'               tmpBuff = .mvardata
               For StrCharPos = 1 To Len(.mvardata)
                  'The largest prime less than 2^16
                  L = (AscB(MidB$(tmpBuff, StrCharPos, 1)) + L) Mod 65521 '&HFFF1
                  h = (h + L) Mod 65521 '&HFFF1
                  
                  If 0 = (StrCharPos Mod &H8000) Then myDoEvents

               Next
'            Debug.Print "b: ", GetTickCount - a 'Benchmark: 5969

      ADLER32 = H16(h) & H16(L)
   End With
End Function


Public Function CryptCall(Key1, Key2) As String
    Dim RetStr As Long
    Dim SLen As Long
    Dim Buffer As String
    'Get a pointer to a string, which contains the command line
    RetStr = myCryptCall(Key1, Key2)
    'Get the length of that string
    SLen = lstrlen(RetStr)
    If SLen > 0 Then
        'Create a buffer
        CryptCall = Space$(SLen)
        'Copy to the buffer
        CopyMemory ByVal CryptCall, ByVal RetStr, SLen
    End If
End Function



Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Sub openURL(url$)
   Dim hProc&
   hProc = ShellExecute(0, "open", url, "", "", 1)
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function AnyOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function pad(x, Optional size As Long = 10) As String
    Dim buf As String
    If Len(x) >= size Then
        pad = x
    Else
        pad = Left(x & Space(size), size)
    End If
End Function

'this has been tested to return the same endsAt pointer as the original findBytes did..you reset it if ok..
Private Function ScanForPassword(fpath As String, startPos As Long, ByRef endsAt As Long, ParamArray Bytes()) As Boolean
    
    Dim f As Long, pointer As Long, x As Long
    Dim buf()  As Byte, i As Long, j As Long
    On Error Resume Next
    
    ReDim buf(90)
    endsAt = 0
    
    f = FreeFile
    Open fpath For Binary Access Read As f
    Seek f, startPos
    
    Do While pointer < LOF(f)
        'If abort Then GoTo aborting
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < 9000 Then ReDim buf(x)
        Get f, , buf()
        For i = 0 To UBound(buf)
            If buf(i) = Bytes(0) Then
                For j = 1 To UBound(Bytes)
                    If buf(i + j) <> Bytes(j) Then Exit For
                Next
                If j > UBound(Bytes) Then 'we found it!
                    ScanForPassword = True
                    endsAt = pointer + i + UBound(Bytes)
                    Exit Do
                End If
            End If
        Next
    Loop
    
    Close f
    
End Function
