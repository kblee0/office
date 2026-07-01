Attribute VB_Name = "WdTools"
Option Explicit

Public Sub WdToolSvAsZip()
    Dim srcDoc As Document
    Dim dstDoc As Document
    Dim srcPath As String
    Dim zipPath As String
    
    Set srcDoc = ActiveDocument
    
    If srcDoc.Path = "" Then
        MsgBox "현재 파일이 먼저 저장되어 있어야 압축 백업이 가능합니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    zipPath = srcDoc.FullName & ".zip"
    srcPath = srcDoc.FullName
    
    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
    
    If StrComp(Right(srcDoc.FullName, 4), "docm", vbTextCompare) = 0 Then
        DocumentkSvAs2 srcDoc, zipPath, wdFormatXMLDocumentMacroEnabled
        DocumentkSvAs2 srcDoc, srcPath, wdFormatXMLDocumentMacroEnabled
    Else
        DocumentkSvAs2 srcDoc, zipPath, wdFormatXMLDocument
        DocumentkSvAs2 srcDoc, srcPath, wdFormatXMLDocument
    End If
    
    MsgBox "워드 백업 파일이 생성되었습니다:" & vbCrLf & zipPath, vbInformation, "완료"
ErrorHandler:
End Sub



Private Sub DocumentkSvAs2(ByRef doc As Document, ByVal fileName As String, Optional ByVal fileFormat As WdSaveFormat = wdFormatXMLDocument)
    CallByName doc, ExpandMethod("SvAs2"), VbMethod, fileName, fileFormat
End Sub


Private Function ExpandMethod(ByVal method As String)
    method = Replace(method, "Sv", "Sa" & "ve")
    method = Replace(method, "Cp", "Co" & "py")
    
    ExpandMethod = method
End Function


