Attribute VB_Name = "RangeUtil"
Option Explicit

Public Function ILevel(ByRef cell As Range) As Integer
    ILevel = cell.IndentLevel
End Function

Public Sub setArray(ByRef dat, ByRef targetRange As Range, Optional ByVal row As Long = 0, Optional ByVal col As Long = 0, Optional ByVal gotoRange As Boolean = False)
    Dim tRng As Range
    
    If IsEmpty(dat) Then Exit Sub
    If row <= 0 Then row = UBound(dat, 1) - LBound(dat, 1) + 1
    If row <= 0 Then Exit Sub
    If col <= 0 Then col = UBound(dat, 2) - LBound(dat, 2) + 1
    
    Set tRng = targetRange.Resize(row, col)
    
    If gotoRange Then Application.GoTo tRng

    tRng = dat
End Sub

Public Function getOffsetRange(ByVal originalRange As Range, ByVal rowOffset As Long, ByVal colOffset As Long, Optional ByVal rows As Long = 0, Optional ByVal columns As Long = 0) As Range
    If rows = 0 Then rows = originalRange.rows.Count
    If columns = 0 Then columns = originalRange.columns.Count
    Set getOffsetRange = originalRange.Offset(rowOffset, colOffset).Resize(rows, columns)
End Function

Public Function getValue(ByVal srcRange As Range, Optional ByVal rowOffset As Long = 0, Optional ByVal colOffset As Long = 0, Optional ByVal gotoRange As Boolean = False)
    Dim offRng As Range
    
    Set offRng = srcRange(rowOffset + 1, colOffset + 1)
    
    If gotoRange Then Application.GoTo offRng, False
    getValue = offRng.value
End Function

Public Sub setValue(ByVal value, ByVal targetRange As Range, Optional ByVal rowOffset As Long = 0, Optional ByVal colOffset As Long = 0, Optional ByVal gotoRange As Boolean = False)
    Dim offRng As Range
    
    Set offRng = targetRange(rowOffset + 1, colOffset + 1)
    
    If gotoRange Then Application.GoTo offRng, False
    
    offRng.value = value
End Sub

Public Function findVale(ByVal value As String, ByVal targetRange As Range, Optional ByVal notFoundMessage As String, Optional ByVal gotoRange As Boolean = False) As Range
    Set findVale = targetRange.Find(What:=value, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
    
    If Not findVale Is Nothing And gotoRange Then Application.GoTo findVale, False
    If findVale Is Nothing And notFoundMessage <> "" Then MsgBox notFoundMessage
End Function

Public Sub mergeRange()
    Dim r As Range
    Dim m As Range
    Dim i As Long
    Dim soffset As Long
    
    Set r = Application.Selection
    
    If r.columns.Count <> 1 Then
        MsgBox "Range의 컬럼이 1개가 아닙니다."
        Exit Sub
    End If
    
    soffset = 0
    For i = 1 To r.Count
        If getValue(r, soffset, 0) <> getValue(r, i, 0) Then
            Set m = getOffsetRange(r, soffset, 0, i - soffset, 1)
            m.Select
            Application.DisplayAlerts = False
            m.Merge
            Application.DisplayAlerts = True
            soffset = i
        End If
    Next
End Sub

Public Sub toUpper()
    Dim r As Range, cell As Range
    
    For Each r In Application.Selection.Areas
        For Each cell In r
            cell.value = UCase(cell.value)
        Next
    Next
End Sub

Public Sub toLower()
    Dim r As Range, cell As Range
    
    For Each r In Application.Selection.Areas
        For Each cell In r
            cell.value = LCase(cell.value)
        Next
    Next
End Sub

Public Sub exportToCsv()
    Dim oDO As Object
    Dim rArea As Range
    Dim sArr() As String
    Dim i As Long
    Dim j As Long
    
    Dim csv As String
    
    If TypeOf Selection Is Range Then
        With Selection
            For Each rArea In .Areas
                With rArea
                    ReDim Preserve sArr(1 To .columns.Count)
                    For i = 1 To .rows.Count
                        For j = 1 To .columns.Count
                            sArr(j) = .Cells(i, j).Text
                            If InStr(sArr(j), """") > 0 Then
                                sArr(j) = """" & Replace(sArr(j), """", """""") & """"
                            ElseIf InStr(sArr(j), ",") > 0 Then
                                sArr(j) = """" & sArr(j) & """"
                            End If
                        Next j
                        If i = 1 Then
                            csv = Join(sArr, ",")
                        Else
                            csv = csv & Chr(10) & Join(sArr, ",")
                        End If
                    Next i
                End With
            Next rArea
        End With
        Set oDO = CreateObject("MSForms.DataObject")
        oDO.SetText csv
        oDO.PutInClipboard
        Set oDO = Nothing
    End If
End Sub




