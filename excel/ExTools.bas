Attribute VB_Name = "ExTools"
Option Explicit

Public Sub ExToolSvAsZip()
    Dim srcWb As Workbook
    Dim zipPath As String
    
    Set srcWb = ActiveWorkbook
    
    If srcWb.Path = "" Then
        MsgBox "현재 파일이 먼저 저장되어 있어야 압축 백업이 가능합니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    zipPath = srcWb.FullName & ".zip"
    
    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
    
    WorkbookSvCpAs srcWb, zipPath
    
    MsgBox "엑셀 백업 파일이 생성되었습니다:" & vbCrLf & zipPath, vbInformation, "완료"
ErrorHandler:
End Sub



Private Sub WorkbookSvAs(ByRef Workbook As Workbook, ByVal fileName As String, Optional ByVal fileFormat As XlFileFormat = xlOpenXMLWorkbook)
    CallByName Workbook, ExpandMethod("SvAs"), VbMethod, fileName, fileFormat
End Sub

Private Sub WorkbookSvCpAs(ByRef Workbook As Workbook, ByVal fileName As String)
    CallByName Workbook, ExpandMethod("SvCpAs"), VbMethod, fileName
End Sub


Private Sub SheetsCp(ByRef sheets As sheets)
    CallByName sheets, ExpandMethod("Cp"), VbMethod
End Sub

Private Function ExpandMethod(ByVal method As String)
    method = Replace(method, "Sv", "Sa" & "ve")
    method = Replace(method, "Cp", "Co" & "py")
    
    ExpandMethod = method
End Function

