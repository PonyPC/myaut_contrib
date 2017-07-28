VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'redone

Private buf() As String
Private sz As Long

Public Property Let value(vData As String)
    Erase buf
    sz = Len(vData)
    push buf, vData
End Property

Public Property Get value() As String
    value = Join(buf, "")
End Property

Public Sub Clear()
  Erase buf
  sz = 0
End Sub

Public Sub Concat(ByRef value As String)
  push buf, value
  sz = sz + Len(value)
End Sub

'Public Sub ConcatVariant(ByRef value As Variant)
'
'End Sub

'not used
Public Property Get Length() As Long
    Length = sz
End Property



