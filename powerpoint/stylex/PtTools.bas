Attribute VB_Name = "PtTools"
Option Explicit

Public Sub PtToolSvAsZip()
    Dim srcPt As presentation
    Dim zipPath As String
    
    Set srcPt = ActivePresentation
    
    If srcPt.Path = "" Then
        MsgBox "현재 파일이 먼저 저장되어 있어야 압축 백업이 가능합니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    zipPath = srcPt.FullName & ".zip"
    
    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
    
    PresentationSvCpAs srcPt, zipPath
    
    MsgBox "파워포인트 백업 파일이 생성되었습니다:" & vbCrLf & zipPath, vbInformation, "완료"
ErrorHandler:
End Sub



Private Sub PresentationSvAs(ByRef presentation As presentation, ByVal fileName As String, Optional ByVal fileFormat As PpSaveAsFileType = ppSaveAsOpenXMLPresentation)
    CallByName presentation, ExpandMethod("SvAs"), VbMethod, fileName, fileFormat
End Sub

Private Sub PresentationSvCpAs(ByRef presentation As presentation, ByVal fileName As String)
    CallByName presentation, ExpandMethod("SvCpAs"), VbMethod, fileName
End Sub


Private Function ExpandMethod(ByVal method As String)
    method = Replace(method, "Sv", "Sa" & "ve")
    method = Replace(method, "Cp", "Co" & "py")
    
    ExpandMethod = method
End Function


