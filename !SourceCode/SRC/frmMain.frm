VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "myAut2Exe >The Open Source AutoIT/AutoHotKey script decompiler< - dmod"
   ClientHeight    =   10425
   ClientLeft      =   2670
   ClientTop       =   1005
   ClientWidth     =   9330
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10425
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDirMode 
      Height          =   1425
      Left            =   405
      TabIndex        =   19
      Top             =   1035
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   135
      TabIndex        =   14
      Top             =   5175
      Width           =   9150
   End
   Begin VB.TextBox txtFilePath 
      Height          =   330
      Left            =   1170
      OLEDropMode     =   1  'Manual
      TabIndex        =   13
      Text            =   "Drag and Drop"
      Top             =   135
      Width           =   8070
   End
   Begin VB.TextBox Txt_Script 
      Height          =   3945
      Left            =   135
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   900
      Width           =   9135
   End
   Begin VB.ListBox List_Positions 
      Height          =   1425
      Left            =   4770
      TabIndex        =   9
      ToolTipText     =   "Right click for close"
      Top             =   8865
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Skip 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Skip >>>"
      Height          =   260
      Left            =   8325
      TabIndex        =   4
      ToolTipText     =   "Skip current step"
      Top             =   6660
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Fr_Options 
      Height          =   1740
      Left            =   120
      TabIndex        =   1
      Top             =   8640
      Width           =   9135
      Begin VB.TextBox txt_OffAdjust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Text            =   "2C"
         ToolTipText     =   $"frmMain.frx":628A
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_scan 
         Caption         =   "<<"
         Height          =   255
         Left            =   4245
         TabIndex        =   7
         ToolTipText     =   "Finds possible scriptstarts. ( Requires valid options for 'SrcFile_FileInst' and 'CompiledPathName' in options.)"
         Top             =   255
         Width           =   375
      End
      Begin VB.TextBox Txt_Scriptstart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         ToolTipText     =   $"frmMain.frx":631C
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_options 
         Caption         =   "More Options >>"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   225
         Width           =   1455
      End
      Begin VB.CheckBox Chk_TmpFile 
         Caption         =   "Don't delete temp files (for ex. compressed scriptdata)"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Chk_verbose 
         Caption         =   "Verbose Mode"
         Height          =   195
         Left            =   7650
         MaskColor       =   &H8000000F&
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl_Adjustment 
         Caption         =   "Off_Adjust"
         Height          =   255
         Left            =   2550
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "StartOffset"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ListBox ListLog 
      Height          =   1620
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "Double click to see more !"
      Top             =   6930
      Width           =   9135
   End
   Begin VB.Shape Sh_ProgressBar 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   135
      Top             =   495
      Visible         =   0   'False
      Width           =   9090
   End
   Begin VB.Label Label5 
      Caption         =   "File or Folder: "
      Height          =   285
      Left            =   90
      TabIndex        =   18
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Decoded Script results (truncated at MAX_INT chars)"
      Height          =   240
      Left            =   135
      TabIndex        =   17
      Top             =   630
      Width           =   3840
   End
   Begin VB.Label Label3 
      Caption         =   "Automated run log"
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   4950
      Width           =   1500
   End
   Begin VB.Shape Sh_ProgressBar 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   810
      Top             =   6660
      Visible         =   0   'False
      Width           =   7425
   End
   Begin VB.Shape Sh_ProgressBar 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   810
      Top             =   6795
      Visible         =   0   'False
      Width           =   7425
   End
   Begin VB.Label Label2 
      Caption         =   "Run log"
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   6705
      Width           =   1320
   End
   Begin VB.Menu mu_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuUnUPX 
         Caption         =   "UPX Decompress"
      End
      Begin VB.Menu mnuDirModeMoveFails 
         Caption         =   "Dir Mode Move Fails"
      End
      Begin VB.Menu RegExp_Renamer 
         Caption         =   "&RegExp_Renamer"
         Shortcut        =   {F11}
         Visible         =   0   'False
      End
      Begin VB.Menu mi_FunctionRenamer 
         Caption         =   "&FunctionRenamer"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mi_HexToBinTool 
         Caption         =   "&HexToBin_Binary() parser"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mi_CustomDecrypt 
         Caption         =   "&Custom_Decrypt() parser"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mi_GetAutoItVersion 
         Caption         =   "&GetAutoItVersion(Attention this executes the current exe)"
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu mi_SeperateIncludes 
         Caption         =   "&Seperate includes of *.au3"
      End
   End
   Begin VB.Menu mu_BugFix 
      Caption         =   "&BugFix"
      Begin VB.Menu mi_LocalID 
         Caption         =   "SetLocalID"
      End
   End
   Begin VB.Menu mu_Info 
      Caption         =   "&Info"
      Begin VB.Menu mi_About 
         Caption         =   "About"
         Visible         =   0   'False
      End
      Begin VB.Menu mi_Update 
         Caption         =   "&Update"
      End
      Begin VB.Menu mi_Forum 
         Caption         =   "&Forum"
      End
      Begin VB.Menu mi_MD5_pwd_Lookup 
         Caption         =   "Copy Passwordhash"
      End
   End
   Begin VB.Menu mnuScanFile 
      Caption         =   "Scan File"
      Begin VB.Menu mi_Reload 
         Caption         =   "Manual"
      End
      Begin VB.Menu mnuAutomatedRun 
         Caption         =   "Automated"
      End
      Begin VB.Menu mi_cancel 
         Caption         =   "Cancel (Esc)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ScriptLines
Public WithEvents Console As Console
Attribute Console.VB_VarHelpID = -1
Public StartLocations As New Collection
Public bCmd_Skip_HasFocus As Boolean
Public DecompileSuccess As Boolean

Dim GUIEvent_InitialWidth(0 To 2) As Long
Dim GUIEvent_ProcessScale(0 To 2) As Double
Dim GUIEvent_Max(0 To 2) As Long
Dim GUIEvent_Width_before(0 To 2) As Long
Private Form_Initial_Height&
Private Form_Initial_Width&

Enum spStages
    sps_Decompile = 3
    sps_DeTokenization = 2
    sps_DeObfuscation = 1
End Enum

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_GETCOUNT = &H18B
Const LB_SETTOPINDEX = &H197

Private hLog As Long
Private hDirLog As Long
Private curLog As String
Private compactions As Long 'test code to see how often the log compactor actually runs...
Private logLineCnt As Long
Private autoRunIndex As Long

'Private Sub mi_FunctionRenamer_Click()
'  Load FrmFuncRename
''   If FileExists(Combo_Filename) Then
''      FrmFuncRename.Txt_Fn_Org_FileName = Combo_Filename
''   Else
''
''
''   End If
'
'   FrmFuncRename.Show ' vbModal
''   Unload FrmFuncRename
'
'End Sub

Private Sub mi_Update_Click()
   openURL "http://deioncube.in/files/MyAutToExe/index.html"
End Sub

Private Sub mi_Forum_Click()
   openURL "http://board.deioncube.in/showthread.php?tid=29"
End Sub

Private Sub ListLog_DblClick()
   On Error Resume Next
   If Not FileExists(curLog) Then Exit Sub
   Shell "notepad.exe """ & curLog & """", vbNormalFocus
End Sub

Public Sub Log(TextLine$, Optional LinePrefix$ = "", Optional forceScroll As Boolean = False)
    
   On Error Resume Next
   
   'lets keep the ui log compact at less than 1000 entries, always showing the last 50 at least...
   If ListLog.ListCount > 5000 Then
        Dim tmp() As String, i As Long
        compactions = compactions + 1
        For i = ListLog.ListCount - 51 To ListLog.ListCount - 1
            push tmp, ListLog.List(i)
        Next
        ListLog.Clear
        ListLog.AddItem ">>> UI log compacted, double click to view full log (" & logLineCnt & " lines)"
        For i = 0 To UBound(tmp)
            ListLog.AddItem tmp(i)
        Next
    End If
        
   
 ' Output Text /split into line and output it line wise
   Dim Line
   For Each Line In Split(TextLine, vbCrLf)
      ListLog.AddItem LinePrefix & Line
      If hLog <> 0 Then Print #hLog, CStr(LinePrefix & Line)
      logLineCnt = logLineCnt + 1
   Next
 
   'Scroll to last item, ListIndex and ListCount are only Integer types so
   'when there are more than &h7fff items there would be an overflow error
   'we will avoid it with a signed item check (could use api..)
   'If ListLog.ListCount > 0 Then ListLog.ListIndex = ListLog.ListCount - 1
   
   Dim Count As Long
   Count = SendMessage(ListLog.hWnd, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
   If forceScroll Or Count Mod 10 = 0 Then
        SendMessage ListLog.hWnd, LB_SETTOPINDEX, ByVal Count - 1, ByVal 0
   End If
   
   If (Rnd < 0.01) Then myDoEvents
   Err.Clear
   
End Sub

Function cLog(x)
   Dim Count As Long
   List2.AddItem x
   Count = SendMessage(List2.hWnd, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
   SendMessage List2.hWnd, LB_SETTOPINDEX, ByVal Count - 1, ByVal 0
End Function


Private Function LogOpen() As String

    Dim path As String
    Dim fName As String
    Dim ii, i As Long
    Dim data As String
    Dim tmp() As String
    
    On Error Resume Next
    
    'I dont want accidental calls to create partial logs...
    If hLog <> 0 Then Exit Function
    
    Do
        fName = fso.GetBaseName(txtFilePath)
        If Len(fName) > 8 Then fName = VBA.Left(fName, 8)
        fName = Format(Now, "m-d-yy h.n.s") & "_" & fName
        path = fso.GetParentFolder(txtFilePath) & "\" & fName & "_" & ii & ".log"
        ii = ii + 1
    Loop While FileExists(path)
    
    logLineCnt = 0
    compactions = 0
    hLog = FreeFile
    curLog = path
    
    Err.Clear
    Open path For Output As hLog
    
    If Err.Number <> 0 Then
        hLog = 0
    Else
        Print #hLog, "Opening new log " & Now & " isAutomationRun: " & isAutomationRun & IIf(isAutomationRun, " Run: " & autoRunIndex, "")
        Print #hLog, "Analysis run for " & txtFilePath
    End If
    
End Function

Private Function LogClose()
    On Error Resume Next
    Print #hLog, "Lines Logged: " & logLineCnt & " UI Log Compactions: " & compactions
    Print #hLog, "Closing log " & Now
    Close hLog
    hLog = 0
    If compactions > 0 Then
        Me.Caption = "Log Compactions = " & compactions
    End If
End Function

Function StageToStr(s As spStages)
    Select Case s
        Case 1: StageToStr = "DeObfuscation"
        Case 2: StageToStr = "DeTokenization"
        Case 3: StageToStr = "Decompilation"
        Case Else: StageToStr = "Unknown? " & s
    End Select
End Function

Private Sub mi_cancel_Click()
   CancelAll = True
End Sub

Private Sub mi_CustomDecrypt_Click()
   CustomDecrypt
End Sub

Function DetectStageToStartAt(fPath As String) As spStages
    
   Dim ext As String
   
   DetectStageToStartAt = sps_Decompile
   ext = LCase(fso.GetExtension(fPath))
   If Len(ext) = 0 Then Exit Function
   If Left(ext, 3) = ".ex" Then Exit Function
   
   If InStr(DeTokeniser.TokenFile_RequiredInputExtensions, ext) > 0 Then
         DetectStageToStartAt = sps_DeTokenization
         Exit Function
   End If
   
   If ext = ".au3" Or ext = ".ahk" Then
        DetectStageToStartAt = sps_DeObfuscation
        Exit Function
   End If
   
End Function

Private Sub mi_Reload_Click()
      
   Dim stage As spStages
   Dim StartAt_DeTokenize As Boolean
   Dim er As ErrReturns
   Dim fPath As String
   
   List2.Clear
   ListLog.Clear
   List_Positions.Clear
   lstDirMode.Visible = False
   isAutomationRun = False
   fPath = txtFilePath
    
   If fso.FolderExists(fPath) Then
       Log "Manual mode can not handle directories."
   ElseIf fso.FileExists(fPath) Then
        stage = DetectStageToStartAt(fPath)
        er = StartProcessing(fPath, stage)
        DecompileSuccess = (er = er_SUCCESS)
        'we have a leak somewhere? some files end up locked...
        FILE.CloseFile 'not helping...
   Else
       Log "File or folder not found: " & fPath
   End If
   
End Sub


Public Sub GUIEvent_ProcessBegin(Target&, Optional BarLevel& = 0, Optional Skipable As Boolean = False)
On Error GoTo ERR_GUIEvent_ProcessBegin
   
   With Sh_ProgressBar(BarLevel)
      .Visible = True
'      .Tag = .Width
      .Width = 0
      
      GUIEvent_Max(BarLevel) = Target
      
    ' Avoid a division by Zero
      If Target > 0 Then
       ' Get stored length from when created the Form
         GUIEvent_ProcessScale(BarLevel) = GUIEvent_InitialWidth(BarLevel) / Target
      Else
         GUIEvent_ProcessScale(BarLevel) = 1
      End If
   End With
   
   GUIEvent_Width_before(BarLevel) = 0
   
   If BarLevel = 0 Then
      
      If Skipable Then
         GUI_SkipEnable
      Else
         GUI_SkipDisable
      End If
      
   End If
   
'   myDoEvents
   
ERR_GUIEvent_ProcessBegin:
End Sub

Public Sub GUIEvent_ProcessUpdate(CurrentValue&, Optional BarLevel& = 0)
On Error GoTo ERR_GUIEvent_ProcessUpdate
   With Sh_ProgressBar(BarLevel)
      
      .Width = CurrentValue * GUIEvent_ProcessScale(BarLevel)
      
      If (.Width - GUIEvent_Width_before(BarLevel)) > 10 Then
         GUIEvent_Width_before(BarLevel) = .Width
         
         On Error GoTo 0
         myDoEvents
         
      End If
   End With
ERR_GUIEvent_ProcessUpdate:
End Sub

Public Sub GUIEvent_Increase(PerCentToIncrease As Double, Optional BarLevel& = 0)
   
   Dim NewValue&
   NewValue = GUIEvent_ProcessScale(BarLevel) * PerCentToIncrease

   With Sh_ProgressBar(BarLevel)
      .Width = .Width + NewValue
   End With
End Sub

Public Sub GUIEvent_ProcessEnd(Optional BarLevel& = 0)
On Error GoTo ERR_GUIEvent_ProcessEnd
   
   With Sh_ProgressBar(BarLevel)
'      .Width = .Tag
      .Visible = False
   End With
   
   If BarLevel = 0 Then
      Cmd_Skip.Visible = False
   End If
 
ERR_GUIEvent_ProcessEnd:
End Sub

Sub FL_verbose(Text)
   log_verbose H32(FILE.Position) & " -> " & Text
End Sub

Sub log_verbose(TextLine$)
   If Chk_verbose.value = vbChecked Then Log TextLine
End Sub

Sub FL(Text)
   Log H32(FILE.Position) & " -> " & Text
End Sub

Public Sub LogSub(TextLine$)
   Log "  " & TextLine
End Sub


Private Sub Chk_verbose_Click()
   Static value
   Checkbox_TriStateToggle Chk_verbose, value
End Sub

Private Sub ListLogShowCaption()
   Log Me.Caption
   Log String(80, "=")
End Sub

Private Sub ListLogClear()
   ListLog.Clear
End Sub


Private Sub cmd_options_Click()
   Frm_Options.Show
End Sub

Private Sub Cmd_Skip_Click()
   Cmd_Skip.Visible = False
   Skip = True
End Sub

Private Sub Console_OnInit(ProgramName As String)
    On Error Resume Next
'    GUI_SkipEnable
    GUIEvent_ProcessBegin UBound(ScriptLines), 0, True
    GUIEvent_ProcessBegin UBound(ScriptLines), 1
End Sub

Private Function GetCurLineFromTidyOutput(TextLine As String, MatchKeyWord$) As Long

   Dim myRegExp As New RegExp

   With myRegExp
      .Pattern = MatchKeyWord & RE_Group("\d+") & RE_NewLine
      
      Dim Match As Match
      For Each Match In .Execute(TextLine)
         GetCurLineFromTidyOutput = Match.SubMatches(0)
      Next
   End With

End Function

Private Sub Console_OnOutput(TextLine As String, ProgramName As String)
   On Error GoTo Console_OnOutput_err
 ' cut last newline
   Dim NewLinePos&
   NewLinePos = InStrRev(TextLine, vbCrLf)
   If NewLinePos > 0 Then
      Dec NewLinePos
   End If
   
  'TidyOutput - updateProcessBar
   Dim curline&
   curline = GetCurLineFromTidyOutput(TextLine, "Pre-processing record: ")
   If curline Then
      GUIEvent_ProcessUpdate curline, 1
   Else
      curline = GetCurLineFromTidyOutput(TextLine, "Processing record: ")
      If curline Then
         GUIEvent_ProcessUpdate curline, 0
      End If
   End If
 
 ' Show first 100 Lines
   ShowScriptPart ScriptLines, curline
   
 ' Log output
   Log Left(TextLine, NewLinePos), ProgramName & ": "
   
Console_OnOutput_err:
 Exit Sub
Log "ERR: " & Err.Description & "in  FrmMain.Console_OnOutput(TextLine , ProgramName )"
End Sub

 ' Show first 100 Lines (used in console output)
Private Sub ShowScriptPart(ScriptLines, curline&, Optional Lines& = 100)
   Dim ScriptLinesPreview_Start&
   ScriptLinesPreview_Start = Min(curline, UBound(ScriptLines))
   
   Dim ScriptLinesPreview_End&
   ScriptLinesPreview_End = Min(curline + Lines, UBound(ScriptLines))
   
   ReDim ScriptLinesPreview(ScriptLinesPreview_Start To ScriptLinesPreview_End)
   
   Dim i
   For i = ScriptLinesPreview_Start To ScriptLinesPreview_End
      ScriptLinesPreview(i) = ScriptLines(i)
   Next
   ShowScript Join(ScriptLinesPreview, vbCrLf)

End Sub

Private Sub Console_OnDone(ExitCode As Long)
'   GUI_SkipDisable
   GUIEvent_ProcessEnd 0
   GUIEvent_ProcessEnd 1
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
'      Case vbKeyDelete, vbKeyBack
'         ListLogClear
         
      Case vbKeyEscape
         CancelAll = True

   End Select


End Sub



Private Sub Form_Resize()
  
    On Error Resume Next
      
      If WindowState = vbMaximized Then WindowState = vbNormal
      
      If WindowState = vbNormal Then
         If (Me.Height <> Form_Initial_Height) Or _
            (Me.Width <> Form_Initial_Width) Then
               
             Me.Height = Form_Initial_Height
             Me.Width = Form_Initial_Width
         End If
      End If

End Sub



Private Sub List_Positions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = MouseButtonConstants.vbRightButton Then
      List_Positions.Visible = False
      lbl_Adjustment.Visible = False
      txt_OffAdjust.Visible = False
   End If
   
End Sub





Private Sub mi_HexToBinTool_Click()
   HexToBinTool
End Sub

Private Sub mi_LocalID_Click()
InputValue:
    On Error GoTo ERR_mi_LocalID_Click
    
    Dim InputboxTmp$
    InputboxTmp = InputBox( _
        "You will need to adjust that value if your Windows is not a german of english one and you are getting errors(checksum fail;modified JB LZSS Signature) when decompiling ya own freshly compiled files or the included examples." _
        & "See '!SourceCode\languages-ids.txt' for your LCID. '0' will tell VB to use the current LCID. Any invalid value will reset this to the default(German).", _
        "Enter your LocalID(as hex) for handling strings", H16(LocaleID))

    If InputboxTmp = "" Then Exit Sub
    
    Dim InputboxTmpVal&
    
    InputboxTmpVal = HexToInt(InputboxTmp)
    If InputboxTmpVal <> 0 Then
        RangeCheck InputboxTmpVal, &H5000, &H400, "LCID is not inside the valid range"
    End If
    
    LocaleID = InputboxTmpVal
    
ERR_mi_LocalID_Click:
Select Case Err
    Case 0
    Case 13
        LocaleID = LocaleID_GER
    
    Case Else
        MsgBox Err.Description
        Resume InputValue
End Select
End Sub

'Copies hash to clipboard and does an online query.
Private Sub mi_MD5_pwd_Lookup_Click()
   Clipboard.Clear
   Clipboard.SetText MD5PassphraseHashText

   'openURL MD5_CRACKER_URL$ & LCase$(MD5PassphraseHashText)

End Sub

Private Sub HexToBinTool()

   Dim filename As New ClsFilename
   filename.filename = InputBox("FileName:", "", txtFilePath)
   
   If filename.filename = "" Then Exit Sub

   Dim data$
   data = fso.ReadFile(filename.filename)

   Dim myRegExp As New RegExp
   With myRegExp

      
      .Pattern = RE_WSpace(RE_Group("\w*?") & "\(", "[""']0x", _
                            "[0-9A-Fa-f]" & "*?", "[""']", ".*?", "\)")
      Dim matches As MatchCollection
      Set matches = .Execute(data)
      Dim FunctionName$
      If matches.Count < 1 Then
         
         FunctionName = "FnNameOfBinaryToString"
      Else
      
         FunctionName = matches(0).SubMatches(0)
      End If
   

      FunctionName = InputBox("FunctionName:", "", FunctionName)
      

      
      .Global = True
      .Pattern = RE_WSpace(RE_Literal(FunctionName), _
                            "\(", "[""']0x", _
                               RE_Group("[0-9A-Fa-f]" & "*?"), _
                            "[""']", _
                              RE_Group_NonCaptured( _
                                 RE_WSpace( _
                                   ",", _
                                   RE_Group("[1-4]")) _
                                 ) & "?", _
                            "\)")
                            
      Set matches = .Execute(data)
      Dim Match As Match
      For Each Match In matches
         With Match
         
            Dim IsPrintable As Boolean
            Dim BinData$
            BinData = MakeAutoItString( _
               HexStringToString(.SubMatches(0), IsPrintable, .SubMatches(1)))
            
            If IsPrintable Then
               Log "Replacing: " & BinData & " <= " & .value
               ReplaceDo data, .value, EncodeUTF8(BinData), .FirstIndex, 1
            Else
               Log "Skipped replace(not printable): " & MakePrintable(BinData) & " <= " & .value
            End If
            
         End With
      Next
      
      
      
   End With
   
   If matches.Count Then
      filename.Name = filename.Name & "_HexToBin"
       
    ' Save
      fso.writeFile filename.filename, data

       
       Log matches.Count & " replacements done."
       Log "File save to: " & filename.filename
   Else
      Log "Nothing found."
   End If
   
   

End Sub

'this is run manually only...
Private Sub CustomDecrypt()

   Dim filename As New ClsFilename
   filename.filename = InputBox("Note: The CustomDecrypt only makes sense together with the VB6-IDE !" & vbCrLf & _
                     "" & vbCrLf & _
                     "It helps if you encounter stuff like this: 'MsgBox(0, Fn04B6(""dHBKQL LWW~W"", ""FI""),...'" & vbCrLf & _
                     "" & vbCrLf & _
                     "FileName:", "Programmers only!", txtFilePath)
   
   If filename.filename = "" Then Exit Sub
   

   Dim data$
   data = fso.ReadFile(filename.filename)

   
   
   Dim myRegExp As New RegExp
   With myRegExp

      
      .Pattern = RE_WSpace(RE_Group("\w*?") & "\(", "[""']0x", _
                            "[0-9A-Fa-f]" & "*?", "[""']", ".*?", "\)")
      Dim matches As MatchCollection
      Set matches = .Execute(data)
      Dim FunctionName$
      If matches.Count < 1 Then
         
         FunctionName = "FnNameOfBinaryToString"
      Else
      
         FunctionName = matches(0).SubMatches(0)
      End If
   
   
FunctionName = "_deCode"

      FunctionName = InputBox("FunctionName:", "", FunctionName)
      

'_deCode("rATNQ7", "BA")

      .Global = True
      
      
    'We'll just care about "doublequoted" Strings
      Const RE_AU3_QUOTE$ = "[""]"
      
      Const RE_AU3_String$ = _
         RE_AU3_QUOTE & "(" & _
             "[^""]*?" & _
         ")" & RE_AU3_QUOTE
      
      .Pattern = RE_WSpace(RE_Literal(FunctionName), _
                            "\(", _
                              RE_AU3_String$, _
                              ",", _
                              RE_AU3_String$, _
                            "\)")
                            
      Set matches = .Execute(data)
      Dim Match As Match
      For Each Match In matches
         With Match
         
            Dim IsPrintable As Boolean
            Dim BinData$
            

            BinData$ = CryptCall(.SubMatches(0), .SubMatches(1))
            BinData = MakeAutoItString(BinData$)
            
 '           If IsPrintable Then
               Log "Replacing: " & BinData & " <= " & .value
               ReplaceDo data, .value, EncodeUTF8(BinData), .FirstIndex, 1
  '          Else
  '             Log "Skipped replace(not printable): " & MakePrintable(BinData) & " <= " & .value
  '          End If
            
         End With
      Next
      
      
      
   End With
   
   If matches.Count Then
      filename.Name = filename.Name & "_CustomDecrypt"
       
    ' Save
      fso.writeFile filename.filename, data

       
       Log matches.Count & " replacements done."
       Log "File save to: " & filename.filename
   Else
      Log "Nothing found."
   End If
   
   

End Sub

Private Sub Form_Load()
  ' Command line support is handled in CCommandlineOptions in Class_Initilize (opts global object)

  'CamoGet
 
   Set Console = New Console
  
   Dim hRandRot As Long
   Dim path As String
   Dim i As Long
   
   Form_Initial_Height = Me.Height
   Form_Initial_Width = Me.Width
   
   path = App.path & "\data\RanRot_MT.dll"
   hRandRot = LoadLibrary(path)
   
   If hRandRot = 0 Then
        Log "RanRot_MT.dll not found: " & path
        mnuScanFile.Enabled = False
        txtFilePath.Enabled = False
        Exit Sub
   End If

   For i = 0 To Sh_ProgressBar.Count - 1
        GUIEvent_InitialWidth(i) = Sh_ProgressBar(i).Width
   Next
   
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " build(" & App.Revision & ")"
   
   LocaleID = LocaleID_GER
   FormSettings_Load Me, "txt_OffAdjust Txt_Scriptstart"
   Load Frm_Options 'this sets all the defaults from the textboxes..
   Listbox_SetHorizontalExtent ListLog, 6000
   ListLogClear
   ListLogShowCaption
   fHandleDebugLogOpen
   FILE.id = "global"
   
  'Show Form if SilentMode is not Enable
   If Not opts.SilentMode Then Me.Show
  
   If False And opts.IsCommandlineMode Then
   'If opts.IsCommandlineMode Then
      If opts.DirMode Then
            HandleDirectoryMode opts.path
      Else
            txtFilePath = opts.path
            DoAutomationRun
      End If
   Else
        Log "Can decompiled AutoHotKey/AutoIT binaries or detokenize .tok or .mem files"
        Log "Automated mode can unpack UPX files if necessary. You may have to extract SFX or memdump other packers"
   End If

End Sub
   
Function DirLog(msg)
    On Error Resume Next
    Dim c As Long
    Print #hDirLog, msg
    lstDirMode.AddItem msg
    c = SendMessage(lstDirMode.hWnd, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
    SendMessage lstDirMode.hWnd, LB_SETTOPINDEX, ByVal c - 1, ByVal 0
End Function

Function HandleDirectoryMode(dirPath As String)
    
    Dim f As Collection, Count As Long
    Dim ff, bn, pd As String, sample
    Dim success As Long, processed As Long
    Dim st As Long, fDirLog As String
    
    On Error Resume Next
    
    With lstDirMode
        .Top = Label4.Top
        .Left = Label4.Left
        .Height = Txt_Script.Height + Label4.Height
        .Width = Txt_Script.Width
        .Clear
        .Visible = True
    End With
    
    Set f = fso.GetFolderFiles(dirPath)
    
    fDirLog = dirPath & "\dirlog.txt"
    If fso.FileExists(fDirLog) Then
        hDirLog = FreeFile
        Open fDirLog For Append As hDirLog
    Else
        hDirLog = FreeFile
        Open fDirLog For Output As hDirLog
    End If
    
    DirLog "Started Directory mode " & Now & " found " & f.Count & " files"
    GUIEvent_ProcessBegin f.Count, 2
    
    For Each ff In f
        Err.Clear
        If fso.GetExtension(ff) <> ".txt" Then
            processed = processed + 1
            pd = fso.GetParentFolder(ff) & "\"
            bn = fso.GetBaseName(ff) & "_files"
            'If fso.FolderExists(pd & bn) Then fso.DeleteFolder pd & bn
            If fso.FolderExists(pd & bn) Then
                DirLog "Skipping: " & bn & " Directory already found resuming previous run?"
            Else
                fso.CreateFolder pd & bn
                fso.Copy ff, pd & bn
                sample = pd & bn & "\" & fso.FileNameFromPath(ff)
                DirLog "Handling Sample: " & sample
                If Err.Number <> 0 Then DirLog "  Error copying sample? " & Err.Description & " Line: " & Erl
                If fso.FileExists(sample) Then
                    txtFilePath = sample
                    StartBenchMark st
                    DoAutomationRun
                    DirLog IIf(DecompileSuccess, "     SUCCESS", "           FAILED") & " Time: " & EndBenchMark(st) & " Offsets: " & List_Positions.ListCount
                    If Not DecompileSuccess Then
                        Name pd & bn As pd & "fail_" & bn
                    Else
                        success = success + 1
                    End If
                Else
                    DirLog "  copied sample file does not exist skipping :("
                End If
            End If
        End If
        DoEvents
        GUIEvent_ProcessUpdate processed, 2
    Next
    
    GUIEvent_ProcessEnd 2
    DirLog "Directory mode completed " & Now & " Success " & success & " / " & processed
    Close hDirLog
            
End Function



Private Sub Form_Unload(Cancel As Integer)
   
   fHandleDebugLogClose
   FormSettings_Save Me
  
 'Close might be clicked 'inside' some myDoEvents so
 'in case it was do a hard END
 '   End
 
   Dim form_i As Form
   For Each form_i In Forms
     Unload form_i
   Next
 
   APP_REQUEST_UNLOAD = True
   
   'WH_close
End Sub

Private Sub mi_SeperateIncludes_Click()
   Dim FILE$
   FILE = InputBox("Normally seperating includes is done automatically after you decompiled some au3.exe(of old none tokend format)." & vbCrLf & _
          "However that tool is useful in the case you have some decompiled *.au3 with these '; <AUT2EXE INCLUDE-START: C:\ ...' comments you like to process." & vbCrLf & vbCrLf & _
          "Please enter(/paste) full path of the file: (Or drag it into the myAutToExe filebox and then run me again)", "Manually run 'seperate au3 includes' on file", txtFilePath)
   If FILE <> "" Then
      filename.filename = FILE
      SeperateIncludes
   End If
End Sub


Sub ResetUI(Optional startUp As Boolean = True)
    
    If startUp Then
        CancelAll = False
        Set ExtractedFiles = New Collection
        
        ' Block any new files during DoEvents
        txtFilePath.Enabled = False
        mi_Reload.Enabled = False
        mi_cancel.Enabled = True
        
        ' Reset ProgressBars
        GUIEvent_ProcessEnd 0
        GUIEvent_ProcessEnd 1
        
        ' Clear Log (expect when run via commandline)
        If opts.IsCommandlineMode = False Then
           ListLogClear
           ListLogShowCaption
        End If
        
        Txt_Script = ""
    Else
        ' Allow Reload / Block Cancel
        txtFilePath.Enabled = True
        mi_Reload.Enabled = True
        mi_cancel.Enabled = False
        opts.IsCommandlineMode = False
    End If
    
  
End Sub


Function DetectFakeScript() As Boolean
    
    If InStr(1, Txt_Script, "Hacker. Nice try, but Wrong :)", vbTextCompare) > 0 Then
        If Len(Txt_Script) < 100 Then
            cLog ">>> FAKE - Detected w0uter protected fake script..ignoring. Script length = " & Len(Txt_Script)
            DetectFakeScript = True
        Else
            cLog "I found the w0uter fake script marker but script is to long to ignore? run manual if you need more."
        End If
    End If
    
    If Len(Txt_Script) > 100 And Not AnyOfTheseInstr(Txt_Script, "Local,Dim,Global") Then
        cLog ">>> FAKE - We decrypted something, but it doesnt seem to be a valid script..we will keep exploring.. Script length = " & Len(Txt_Script)
        DetectFakeScript = True
    End If
    
End Function

Function AutomatedController() As Boolean
    
    Dim fPath As String
    Dim er As ErrReturns
    Dim tmp As String
    Dim i As Long, j As Long
    Dim stage As spStages
    Dim StartAt_DeTokenize As Boolean
    
    List2.Clear
    List_Positions.Clear
    fPath = txtFilePath
    
    If Not FileExists(fPath) Then
        cLog "File not found: " & fPath
        Exit Function
    End If
       
    stage = DetectStageToStartAt(fPath)
    Txt_Scriptstart = Empty
    
    For i = 4 To 1 Step -1
        DeCompiler.AutoChooseVersion = i
        cLog "Trying run with default: " & AutoVerToStr(i)
        er = StartProcessing(fPath, stage)
        If er = ERR_NO_AUT_EXE Then Exit For
        If er = ERR_CANCEL_ALL Then Exit Function
        If er = er_SUCCESS Then
            If DetectFakeScript() Then Exit For
            AutomatedController = True
            Exit Function
        End If
    Next

    cLog "Doing an offset scan for possible script start addresses"
    cmd_scan_Click
    
    If List_Positions.ListCount = 0 Then
        cLog "None found"
    Else
        For j = 0 To List_Positions.ListCount - 1
            Txt_Scriptstart = List_Positions.List(j)
            For i = 4 To 1 Step -1
                DeCompiler.AutoChooseVersion = i
                cLog "Starting at script offset " & j & " = " & Txt_Scriptstart & " default: " & AutoVerToStr(i)
                er = StartProcessing(fPath, stage)
    '            If er = ERR_NO_AUT_EXE Then Exit For
                If er = ERR_CANCEL_ALL Then Exit Function
                If er = er_SUCCESS Then
                    If DetectFakeScript() Then
                        'its bullshit so we ignore this start offset and keep processing...
                        Exit For
                    Else
                        AutomatedController = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
    
    List_Positions.Clear
    Txt_Scriptstart = Empty
    cLog "Ok, lets see if we can unUPX the sample..."
    
    If Not canAttemptUnUPX(fPath) Then
        cLog "This sample does not appear upx compressed, aborting"
        Exit Function
    End If
    
    If Not UnUPX(fPath, tmp) Then
        cLog "upx decompress failed..."
        Exit Function
    End If
    
    For i = 4 To 1 Step -1
        DeCompiler.AutoChooseVersion = i
        cLog "Trying run with default: " & AutoVerToStr(i)
        er = StartProcessing(fPath, stage)
        If er = ERR_NO_AUT_EXE Then Exit For
        If er = ERR_CANCEL_ALL Then Exit Function
        If er = er_SUCCESS Then
            If DetectFakeScript() Then Exit For
            AutomatedController = True
            Exit Function
        End If
    Next
    
    cLog "Doing an offset scan in decompressed sample for possible script start addresses"
    cmd_scan_Click

    If List_Positions.ListCount = 0 Then
        cLog "None found"
    Else
        For j = 0 To List_Positions.ListCount - 1
            Txt_Scriptstart = List_Positions.List(j)
            For i = 4 To 1 Step -1
                DeCompiler.AutoChooseVersion = i
                cLog "Starting at script offset " & j & " = " & Txt_Scriptstart & " default: " & AutoVerToStr(i)
                er = StartProcessing(fPath, stage)
    '            If er = ERR_NO_AUT_EXE Then Exit For
                If er = ERR_CANCEL_ALL Then Exit Function
                If er = er_SUCCESS Then
                    If DetectFakeScript() Then
                        'its bullshit so we ignore this start offset and keep processing...
                        Exit For
                    Else
                        AutomatedController = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
    
    cLog "Out of automated options, try processing it manually in verbose mode!"

End Function


Function StartProcessing(fPath As String, Optional startProcessingAtStage As spStages = sps_Decompile) As ErrReturns

    'note no exit function calls, must cleanup at end of function and give retval
    On Error Resume Next

    Dim startTime As Long
    
    autoRunIndex = autoRunIndex + 1 'inc before logopen, only shown if isAutomationRun= true so always increment ok..
    
    LogOpen
    ResetUI
    filename = fPath 'global clsFileName object used and abused throughout sub functions...
    StartBenchMark startTime
    
    Log "Starting processing at stage: " & StageToStr(startProcessingAtStage)
    
    If startProcessingAtStage >= sps_Decompile Then
1        Decompile
         If DeCompiler.UsedPasswordScan Then
            Log ">>> USED PASSWORD SCAN"
            If lstDirMode.Visible Then DirLog ">>> USED PASSWORD SCAN" 'trying to catch a sample that uses this...
         End If
'        If Err = ERR_NO_AUT_EXE Then 'instead of cascading down through the various stages automatically..we make it testable from above
'            Log Err.Description
'            Err.Clear
'        End If
        If Err.Number <> 0 Then GoTo final
        Log "Decompiled ok!"
    Else
        ExtractedFiles.Add fPath, "MainScript"
    End If
    
    filename = ExtractedFiles("MainScript")

    If startProcessingAtStage >= sps_DeTokenization Then
2        DeToken
        If Err = NO_AUT_DE_TOKEN_FILE Then 'not all exe contain scripts which will need to be detokenized...this is ok..
            Log Err.Description
            Err.Clear
        End If
        If Err.Number <> 0 Then GoTo final
    End If
        
    Log String(79, "=")
    
    If startProcessingAtStage >= sps_DeTokenization Then
3        DeObfuscate
        If Err = ERR_NO_OBFUSCATE_AUT Then
            Log Err.Description
            Err.Clear
        End If
        If Err.Number = ERR_CANCEL_ALL Then GoTo final
        If Err.Number <> 0 Then
            Log "DeObfuscation failed with error: " & Err.Description
            Err.Clear 'this is to complex/fragile to abort everything if it chokes...
        End If
    End If
    
    'these two will not throw errors...
4    If Frm_Options.Chk_RestoreIncludes.value = vbChecked Then SeperateIncludes
5    If IsTextFile Then CheckScriptFor_COMPILED_Macro

final:
  
  Dim routines()
  routines = Array("Unknown", "Decompile", "Detoken", "DeObfuscate", "SeperateIncludes", "CheckScriptFor_COMPILED_Macro")
  Dim er As Long
  Dim isScript As Boolean
  Dim details As String
  Dim errText As String
  
  If Err.Number <> 0 Then
        If Len(Txt_Script) > 40 Then
            isScript = AnyOfTheseInstr(Txt_Script, "Local,Dim,Global")
            details = " TxtScriptLen: " & Len(Txt_Script) & " isScript: " & isScript
        End If
        er = Err.Number
        errText = "ERR: " & Err.Description & " Line: " & Erl
        If Erl > 0 And Erl < UBound(routines) Then errText = errText & " (" & routines(Erl) & ") "
        Log errText & details
        If IsIde() And isScript Then Stop
  End If
  
  'this would reset the err object because it uses on error resume next..there is no scope its global object..
  'this is why return values are better than using the err object globally across functions...pfft
  If KeyExistsInCollection(ExtractedFiles, "MainScript") Then
        filename = ExtractedFiles("MainScript").filename
        If filename.ext = ".au3" Then RunTidy2 filename    'dz: tidy may not have been run...
  End If
  
  Log "Time = " & EndBenchMark(startTime) & " Success = " & (er = er_SUCCESS), , True
  
  LogClose
  ResetUI False
  StartProcessing = er
  
  If Not isAutomationRun Then
        If APP_REQUEST_UNLOAD Then End
        If opts.QuitWhenFinish Then Unload Me
  End If
  
End Function

Sub SaveControlLog(lst As ListBox, Optional suffix = "_auto", Optional logDir As String = Empty)
    
    Dim path As String
    Dim fName As String
    Dim ii, i As Long
    Dim data As String
    Dim tmp() As String
    Dim pd As String
    Dim bn As String
    
    On Error Resume Next
    
    Do
        fName = fso.GetBaseName(txtFilePath)
        If Len(fName) > 8 Then fName = VBA.Left(fName, 8)
        fName = Format(Now, "m-d-yy h.n.s") & "_" & fName
        pd = fso.GetParentFolder(txtFilePath)
        If Len(logDir) > 0 Then pd = logDir
        If Not fso.FolderExists(pd) Then pd = App.path
        path = pd & "\" & fName & suffix & ii & ".log"
        ii = ii + 1
    Loop While fso.FileExists(path)
        
    lst.AddItem "Saving Logdata to : " & path
    'If list.ListCount = 0 Then Exit Sub 'nothing to do
    For i = 0 To lst.ListCount - 1
        push tmp, lst.List(i)
    Next
    data = Join(tmp, vbCrLf)
    
    fso.writeFile path, data
  
End Sub

Private Sub mnuAutomatedRun_Click()

    lstDirMode.Visible = False
    
    If fso.FileExists(txtFilePath) Then
        DoAutomationRun
    ElseIf fso.FolderExists(txtFilePath) Then
        HandleDirectoryMode txtFilePath
    Else
        Log "Invalid file or folder path: " & txtFilePath
    End If
    
End Sub

Private Function DoAutomationRun()

    Dim success As Boolean
    Dim startTime As Long
    
    autoRunIndex = 0
    StartBenchMark startTime
    isAutomationRun = True
    DecompileSuccess = AutomatedController()
    isAutomationRun = False
    
    If Len(MD5PassphraseHashText) > 0 Then cLog "MD5 Password hash: " & MD5PassphraseHashText
    cLog "Automation run complete Success = " & DecompileSuccess
    cLog "Time: " & EndBenchMark(startTime)
    SaveControlLog List2
    
    'we have a leak somewhere.
    FILE.CloseFile 'not helping...
    
    If APP_REQUEST_UNLOAD Then End
    If Not opts.DirMode And opts.QuitWhenFinish Then Unload Me

End Function

Private Sub mnuDeobsfuscate_Click()
    Dim er As ErrReturns
    
    ListLog.Clear
    If fso.FileExists(txtFilePath) Then
        Log "Trying to deobsfuscate script " & fso.GetBaseName(txtFilePath)
        er = StartProcessing(txtFilePath, sps_DeObfuscation)
        If er = er_SUCCESS Then
            Log "Success!"
        Else
            Log "Error return: " & er
        End If
    Else
        Log "File not found: " & txtFilePath
    End If
        
End Sub

Private Sub mnuDirModeMoveFails_Click()
    
    On Error Resume Next
    
    List2.Clear
    
    If Not fso.FolderExists(txtFilePath) Then
        MsgBox "Folder not found: " & txtFilePath, vbInformation
        Exit Sub
    End If
    
    Dim c As Collection
    Set c = fso.GetSubFolders(txtFilePath)
    
    If c.Count = 0 Then
        MsgBox "No sub folders found", vbInformation
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    Dim f, bn
    Dim failDir As String
    
    failDir = txtFilePath & "\fails"
    
    If Not fso.FolderExists(failDir) Then MkDir failDir
    
    If Not fso.FolderExists(failDir) Then
        MsgBox "failed to create fail dir?:" & failDir
        Exit Sub
    End If
    
    GUIEvent_ProcessBegin c.Count, 2
    
    For Each f In c
        bn = fso.GetBaseName(f)
        If VBA.Left(bn, 5) = "fail_" Then
            Name f As failDir & "\" & bn
            If Err.Number <> 0 Then
                List2.AddItem "Failed to move " & f
            Else
                bn = Replace(Replace(bn, "fail_", Empty), "_files", Empty)
                If fso.FileExists(txtFilePath & "\" & bn) Then
                    Name txtFilePath & "\" & bn As failDir & "\" & bn
                    If Err.Number <> 0 Then
                        List2.AddItem "Failed to move file " & bn
                    Else
                        List2.AddItem "Moved file and folder to /fails " & bn
                        i = i + 1
                    End If
                Else
                    List2.AddItem "Sample file does not exist? " & bn
                End If
            End If
        End If
        j = j + 1
        GUIEvent_ProcessUpdate j, 2
    Next
    
    GUIEvent_ProcessEnd 2
    
    List2.AddItem "Moved " & i & " fails / " & c.Count & " sub folders"
    
End Sub

Private Sub mnuScanFile_Click()
    mi_Reload.Enabled = Not fso.FolderExists(txtFilePath)
End Sub

Private Sub mnuUnUPX_Click()
    Dim tmp As String
    If UnUPX(txtFilePath, tmp) Then
        Log "UnUPX unpack Success!"
        txtFilePath = tmp
    Else
        Log "UnUPX failed"
    End If
End Sub

'Private Sub RegExp_Renamer_Click()
'    FrmRegExp_Renamer.Show ' vbModal
''   Unload FrmRegExp_Renamer
'End Sub

Private Sub txtFilePath_KeyDown(KeyCode As Integer, Shift As Integer)
   Form_KeyDown KeyCode, Shift
End Sub

Private Sub txt_OffAdjust_Change()
   updateStartLocations_List
End Sub

Private Sub txtFilePath_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   txtFilePath.Text = data.files(1)
   'mi_Reload_Click
End Sub
 
Private Sub Txt_Scriptstart_Change()
   On Error Resume Next
   Dim scriptstart&
   scriptstart = HexToInt(Txt_Scriptstart)
   
   Frm_Options.Chk_NormalSigScan.Enabled = (Err.Number <> 0)
   
End Sub

Private Sub cmd_scan_Click()
   
   'LongValScan_Init
   Set StartLocations = New Collection
   
   Dim b As New CBinaryReader
   If Not b.Load(txtFilePath) Then
        Log "Scan for offsets: Failed to load " & txtFilePath
        Exit Sub
   End If
   
   'StartBenchMark
   'now step through the file byte by byte reading long values and xoring them to try to find the value..
 ' New Script 0xADBC / 0F820 WideChar_Unicode
   LongValScan2 b, Xorkey_SrcFile_FileInstNEW_Len, _
               Xorkey_CompiledPathNameNEW_Len, _
               2
   
   Log "Testing for old AU3-scripttype"
 ' Old Script 0x29BC / 29AC ACCII
   LongValScan2 b, Xorkey_SrcFile_FileInst_Len, _
               Xorkey_CompiledPathName_Len, _
               1

   'Log "Scan for markers took: " & EndBenchMark '22 seconds, down to 3 seconds with progress bar update updated on mod 100..

End Sub

Private Sub List_Positions_DblClick()
   Txt_Scriptstart = List_Positions.Text
   Txt_Script = "Ok now scan file to use this offset"
End Sub


Private Sub List_Positions_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then List_Positions_DblClick
End Sub


Public Sub updateStartLocations_List()
   On Error Resume Next 'GoTo updateStartLocations_List_err

   With FrmMain.List_Positions
      
      Dim adjustment&
      adjustment = HexToInt(FrmMain.txt_OffAdjust)
      
      .Clear
      
      lbl_Adjustment.Visible = True
      txt_OffAdjust.Visible = True
      .Visible = True
'      .SetFocus
      
      Dim Location
      For Each Location In StartLocations
         Dec Location, adjustment
         .AddItem H32(Location)
      Next
   End With
updateStartLocations_List_err:
End Sub

 
'Sub ORGINAL_StartProcessing()
'  On Error GoTo StartProcessing_err
'
'  CancelAll = False
'
'' Block any new files during DoEvents
'  txtFilePath.Enabled = False
'  mi_Reload.Enabled = False
'  mi_cancel.Enabled = True
'
'' Reset ProgressBars
'  GUIEvent_ProcessEnd 0
'  GUIEvent_ProcessEnd 1
'
'' Clear Log (expect when run via commandline)
'  If IsCommandlineMode = False Then
'     ListLogClear
'     ListLogShowCaption
'  End If
'  Txt_Script = ""
'  'txtFilePath = UnUPX(txtFilePath) 'dz: can not do this automatically breaks some...
'  filename = txtFilePath
'
'' Log String(80, "=")
'' log "           -=  " & Me.Caption & "  =-"
'
'  On Error Resume Next
'
'  Decompile
'
'  If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
'  If Err Then
'     Log "ERR: " & Err.Description
'  End If
'
'  filename = ExtractedFiles("MainScript")
'
'  DeToken
'     If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
'     If Err Then Log "ERR: " & Err.Description
'
'     Log String(79, "=")
'     On Error Resume Next
'
'     DeObfuscate
'     If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
'     Select Case Err
'     Case 0, ERR_NO_OBFUSCATE_AUT
'        If Frm_Options.Chk_RestoreIncludes.value = vbChecked Then _
'           SeperateIncludes
'
'
'     Case Else
'        Log Err.Description
'
'     End Select
'
'
'    If IsTextFile Then CheckScriptFor_COMPILED_Macro
'
'
'Err.Clear
'GoTo StartProcessing_err
'
'' ErrorHandle for resume from Errors
'DeToken:
'   Log String(79, "=")
'   DeToken
'
'DeObfuscate:
'   Log String(79, "=")
'   DeObfuscate
'
'StartProcessing_err:
'
'' Add some fileName if it weren't done during decompile()
'  If IsAlreadyInCollection(ExtractedFiles, "MainScript") = False Then
'     ExtractedFiles.Add File.filename, "MainScript"
'  End If
'
'
'' Note: Resume is necessary to reenable Errorhandler
''       Else the VB-standard Handler will catch the error -> Exit Programm
'  Select Case Err
'  Case 0
'
'  Case ERR_NO_AUT_EXE
'     Log Err.Description
'     Resume DeToken
'
'  Case NO_AUT_DE_TOKEN_FILE
'     Log Err.Description
'     Resume DeObfuscate
'
'  Case ERR_NO_OBFUSCATE_AUT
''    Log Err.Description
'     Resume StartProcessing_err
'
'  Case ERR_CANCEL_ALL
'     Log "Processing CANCELED!  " & Err.Description
'     Resume Finally
'
'  Case Else
'     Log Err.Description
'     Resume StartProcessing_err
'  End Select
''-----------------------------------------------
'
'Finally:
'' Save Log Data
'  On Error Resume Next
'  Resume Finally
'
'
'  filename = ExtractedFiles("MainScript").filename
'
'  If filename.ext = ".au3" Then RunTidy2 filename   'dz: tidy may not have been run...
'
'  filename.NameWithExt = filename.Name & "_myExeToAut.log"
'
'  Log ""
'  Log "Saving Logdata to : " & filename.filename
'  'fso.writeFile filename.filename, Log_GetData
'  Log "--- COMPLETE ---"
'
'' process Quit
'  If APP_REQUEST_UNLOAD Then End
'
'' Allow Reload / Block Cancel
'  txtFilePath.Enabled = True
'  mi_Reload.Enabled = True
'  mi_cancel.Enabled = False
'
'
'
'  IsCommandlineMode = False
'  If IsOpt_QuitWhenFinish Then Unload Me
'
'End Sub

'Private Function OpenFile(Target_FileName As ClsFilename) As Boolean
'
'   On Error GoTo Scanfile_err
'   Log "------------------------------------------------"
'
'   Log Space(4) & Target_FileName.NameWithExt
'
'   File.Create Target_FileName.mvarFileName, Readonly:=True
'
'   Me.Show
'
'Err.Clear
'Scanfile_err:
'Select Case Err
'   Case 0
'
'   Case Else
'      Log "-->ERR: " & Err.Description
'
'End Select
'
'End Function
