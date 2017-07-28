VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public LoadedFile As String
Public Position As Long
Private m_Length As Long

Private b() As Byte
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal src As Long, ByVal length&)

Property Get length() As Long
    length = m_Length
End Property

Property Get EOS() As Boolean
    If Position >= m_Length Then EOS = True
End Property

Function Load(fPath) As Boolean
    
    On Error GoTo hell
    
    LoadedFile = fPath
    
    Dim f As Long
    f = m_FreeFile("BinaryReader.Load")
    Open fPath For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    m_Close f, "BinaryReader.Load"
    
    Position = 0
    m_Length = UBound(b) + 1
    Load = True

hell:
End Function

Sub Move(ByVal size As Long)
    Position = Position + size
End Sub

'should not be a property because ide will query properties
'at runtime if you mouse over them and then it would move pointer!!
Function int32() As Long
       If Position + 4 <= m_Length Then
            MoveMemory int32, VarPtr(b(Position)), 4
       End If
       Position = Position + 4
End Function

Function int16() As Integer
       If Position + 2 <= m_Length Then
            MoveMemory int16, VarPtr(b(Position)), 2
       End If
       Position = Position + 2
End Function

Function int8() As Byte
       If Position + 1 <= m_Length Then
            MoveMemory int8, VarPtr(b(Position)), 1
       End If
       Position = Position + 1
End Function



