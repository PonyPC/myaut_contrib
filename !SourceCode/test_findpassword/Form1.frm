VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   510
      Left            =   1710
      TabIndex        =   1
      Top             =   2520
      Width           =   1185
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   4380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3mb file...embedded markers at:
'001097C0   AA AA AA
'0020EC50   BB BB BB
'002FFFA0   FF FF FF

'output:
'----------------------------
'Starting new ScanForPassword()
'found! ends at20EC53
'0.172 seconds
'
'Starting FileStream.FindBytes scan
'found!  pointer is:20EC53
'25.547 seconds

'to integrate this new code:
'Decompiler.bas - Function ReadPassword
'old:
'  'FILE.FindBytes &HF9, 0, 0
   'If FILE.EOS = False Then
'
'new:
'   dim newPointer as long
'   If ScanForPassword(FILE.filename, FILE.Position, newPointer, &HF9, 0, 0) Then
'        FILE.Position = newPointer

Private Sub Form_Load()
    
    Dim endsAt As Long
    Dim path As String
    
    Me.Show
    Me.Refresh
    path = App.path & "\noname.dat"
    
    If Not FileExists(path) Then
        List1.AddItem "test file not found..."
        Exit Sub
    End If
    
    List1.AddItem "Starting new ScanForPassword()"
    
    StartBenchMark
    
    If ScanForPassword(path, &H1000, endsAt, &HBB, &HBB, &HBB) Then
        List1.AddItem "found! ends at" & Hex(endsAt)
    Else
        List1.AddItem "Not found!"
    End If
    
    List1.AddItem EndBenchMark
    List1.AddItem ""
    List1.AddItem "Starting FileStream.FindBytes scan"
    Me.Refresh
    
    Dim f As New FileStream
    StartBenchMark
    
    f.Create path, False, False, True
    f.Position = &H1000
    f.FindBytes &HBB, &HBB, &HBB
    
    If f.EOS = False Then
         List1.AddItem "found!  pointer is:" & Hex(f.Position)
    Else
        List1.AddItem "Not found!"
    End If
    
    f.CloseFile
    List1.AddItem EndBenchMark
       
End Sub

'this has been tested to return the same endsAt pointer as the original findBytes did..you reset it if ok..
Private Function ScanForPassword(fpath As String, startPos As Long, ByRef endsAt As Long, ParamArray Bytes()) As Boolean
    
    Dim f As Long, pointer As Long, x As Long
    Dim buf()  As Byte, i As Long, j As Long

    On Error Resume Next
    
    endsAt = 0
    Const bufSz As Long = 9000
    ReDim buf(bufSz)
    
    f = FreeFile
    Open fpath For Binary Access Read As f
    Seek f, startPos
    
    Do While pointer < LOF(f)
        'If abort Then GoTo aborting
        pointer = Seek(f)
        x = LOF(f) - pointer
        If x < 1 Then Exit Do
        If x < bufSz Then ReDim buf(x)
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


Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText lbCopy(List1)
End Sub


