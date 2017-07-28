VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommandlineOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CommandLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public QuitWhenFinish As Boolean
Public SilentMode As Boolean
Public DirMode As Boolean
Public path As String
Public IsCommandlineMode As Boolean

Public NumberOfCommandLineArgs
Public CommandLine As String

Private argsWithQuotes$()
Private Args$()
Private EmptyArray$()

Private Sub Class_Initialize()
  Dim tmp As String
  
  CommandLine = Command
  If Len(CommandLine) = 0 Then Exit Sub
  
  'so the split function requires quotes around paths with spaces in them,
  'lets first just do a basic check to see if only a path was specified..
  tmp = Trim(VBA.Replace(Command, """", Empty))
  If fso.FileExists(tmp) Or fso.FolderExists(tmp) Then
        IsCommandlineMode = True
        NumberOfCommandLineArgs = 1
        path = tmp
        DirMode = fso.FolderExists(tmp)
        Log "Command Line Mode: Processing: " & IIf(DirMode, "Directory", "File") & ": " & path
  Else
        Call commandLineSplitRek
        ProcessCommandline
  End If
  
  
End Sub

Private Sub ProcessCommandline()

   With Me
   
      If .NumberOfCommandLineArgs Then
      
         Log "Cmdline Args: " & .CommandLine
         
         Dim arg
         For Each arg In .getArgs
            
           'Check for options
            If arg Like "[/-]*" Then

               If arg Like "?[qQ]" Then
                  QuitWhenFinish = True
                  Log "Option 'QuitWhenFinish' enabled."
                  
               ElseIf arg Like "?[sS]" Then
                  SilentMode = True
                  Log "Option 'RunSilent' enabled."
                  
               ElseIf arg Like "?[hH]" Then
                  MsgBox Replace(Replace("Command line option support: [-s -q] <file or dir path>\n\t-s   SilentMode (no UI)\n\t-q   Quit When Finish", "\n", vbCrLf), "\t", vbTab), vbInformation
                  End
               Else
                  Log "ERR_Unknow option: '" & arg & "'"
                  
               End If
               
          ' Check if CommandArg is a FileName
            Else
           
               If IsCommandlineMode Then
                  Log "ERR_Invalid Argument ('" & arg & "') filename already set."
                  
               Else
                  If fso.FileExists(arg) Then
                     Log "Command line file mode ok: " & arg
                     IsCommandlineMode = True
                     path = arg
                     Log "FileName : " & arg
                  ElseIf fso.FolderExists(arg) Then
                    Log "Command line directory mode ok: " & arg
                    IsCommandlineMode = True
                    path = arg
                    DirMode = True
                    SilentMode = False
                  Else
                     Log "ERR_Invalid Argument. Can't open file '" & arg & "'"
                  End If
               End If
               
            End If
         Next
      End If
   End With

   'Verify
   If SilentMode And Not QuitWhenFinish Then
      Log "ERR 'SilentMode' only makes sense together with 'QuitWhenFinish'. As long as you don't also enable 'QuitWhenFinish' 'SilentMode' is ignored "
      SilentMode = False
   End If
   
   If DirMode And SilentMode Then SilentMode = False

End Sub


'///////////////////////////////////////////////////
'// getArgs - Get CommandLineArguments with Quotes
Public Function getArgs() As String()
   getArgs = Args
End Function

Public Function getArgsWithQuotes()
   getArgsWithQuotes = argsWithQuotes
End Function





'////////////////////////////////////////////////////////////////////////////////////////////////
'// commandLineSplitRek -  Split the commandLine and store chunks in args() and argsWithQuotes()
'//                        should only used by the Constructor
Private Sub commandLineSplitRek(Optional ByVal startPos = 1, Optional levelCounter = 0)
  'On Error GoTo commandLineSplitRek_err
   
   Dim endPos%                ' Endposition for cut (...e.exe"<-)
'   Dim startPos%           ' Startposition for "new" commandline (->"C:\p...")
   
   endPos = Len(CommandLine) + 1
  
  'Filter Leading Space
   For startPos = startPos To endPos
     If Mid(CommandLine, startPos, 1) > " " Then Exit For
   Next
   
 ' Trivial case - Empty String
   If startPos >= endPos Then
    ' create array for saving commandline arguments
      NumberOfCommandLineArgs = levelCounter
      If NumberOfCommandLineArgs > 0 Then
         ReDim Args(NumberOfCommandLineArgs - 1)
      Else
       ' Clear Arrays
         Args = EmptyArray
      End If
 
 ' Recursion case
   Else
 
   ' does Commandline starts with "C:\P...
     If Mid(CommandLine, startPos, 1) = """" Then
           
      ' seek over quote
        Inc startPos
        
      ' find next "(DoubleQuote)
        endPos = InStr(startPos, CommandLine, """")
     Else
      ' find next ' '(Space)
        endPos = InStr(startPos, CommandLine, " ")
     End If
     
   ' we are already at the end
     If (endPos <= 0) Then
        endPos = Len(CommandLine) + 1
     End If

   ' Call commandLineSplitRek recursiv with "new" commandline 6 increase levelCounter
     commandLineSplitRek endPos + 1, levelCounter + 1
     
   ' Save Data in Array
     Args(levelCounter) = Mid(CommandLine, startPos, endPos - startPos)
   
   End If
   
'commandLineSplitRek_err:
End Sub
 

