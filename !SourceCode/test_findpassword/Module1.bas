Attribute VB_Name = "Module1"
'misc fx to support FileStreamClass plus some debug code i added
Global Const ERR_OPENFILE = 0
Global Const DEBUG_FH = False
Global Const isAutomationRun = False

Private startTime As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function fHandleDebugLog(x)

End Function

Function pad(x, y)

End Function

Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function lbCopy(lstBox As Object) As String
    
    Dim i As Long
    Dim tmp() As String
    
    For i = 0 To lstBox.ListCount
        push tmp, lstBox.List(i)
    Next
    
    lbCopy = Join(tmp, vbCrLf)
    
End Function


Sub StartBenchMark()
    startTime = GetTickCount()
End Sub

Function EndBenchMark() As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    loadTime = endTime - startTime
    EndBenchMark = loadTime / 1000 & " seconds"
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function
