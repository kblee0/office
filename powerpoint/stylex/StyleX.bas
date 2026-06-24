Attribute VB_Name = "StyleX"
'==================================================
' StyleX RibbonX PowerPoint 매크로
' SPEC: 스타일 도구모음
'==================================================

Option Explicit

Const PT_PER_CM As Double = 28.34645652771
Public gBorderOutsideMode As Boolean

Sub OnToggleAction(control As IRibbonControl, pressed As Boolean)
    If control.id = "tglBorderOutside" Then gBorderOutsideMode = pressed
End Sub

Sub OnGetPressed(control As IRibbonControl, ByRef returnedVal)
    If control.id = "tglBorderOutside" Then returnedVal = gBorderOutsideMode
End Sub

'==================================================
' 그룹1: 채우기 (배경색 변경 - 다중 선택 지원)
'==================================================
Public Sub OnFillColorClick(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    
    Dim fill As FillFormat
    Dim colorHex As String
    Dim colorRGB As Long
    
    On Error GoTo ErrorHandler
    
    If Right(control.id, 4) = "NONE" Then
        colorRGB = -1
    Else
        colorHex = Right(control.id, 6)
        colorRGB = HexToRGB(colorHex)
    End If
    
    ' 선택된 모든 도형에 배경색 적용
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        
        ' 테이블 셀인 경우
        If shape.HasTable Then
            ApplyFillColorToTable shape, colorRGB
        ' 일반 도형인 경우
        Else ' If shape.fill.Type <> msoNoFill Then
            Set fill = shape.fill
            If colorRGB = -1 Then
                fill.Transparency = 1#
                fill.Visible = msoFalse
            Else
                fill.ForeColor.RGB = colorRGB
                fill.Transparency = 0
            End If
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

' 테이블의 모든 셀에 배경색 적용
Private Sub ApplyFillColorToTable(tableShape As shape, colorRGB As Long)
    Dim table As table
    Dim cell As cell
    Dim row As Long
    Dim col As Long
    
    On Error Resume Next
    
    Set table = tableShape.table
    
    ' 모든 셀의 배경색 설정
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Set cell = table.cell(row, col)
            If cell.Selected Then
                cell.shape.fill.ForeColor.RGB = colorRGB
                cell.shape.fill.Transparency = 0
            End If
        Next col
    Next row
End Sub



'==================================================
' 그룹2: 선색 (테두리선 색상 변경 - 다중 선택 지원)
'==================================================
Public Sub OnLineColorClick(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    
    Dim line As LineFormat
    Dim colorHex As String
    Dim colorRGB As Long
    
    On Error GoTo ErrorHandler
    
    ' 버튼 ID에서 16진수 추출 (예: L2F2F2F -> 2F2F2F)
    If Right(control.id, 4) = "NONE" Then
        colorRGB = -1
    Else
        colorHex = Right(control.id, 6)
        colorRGB = HexToRGB(colorHex)
    End If
    
    ' 선택된 모든 도형의 테두리선 색상 변경
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        
        ' 테이블 셀인 경우
        If shape.HasTable Then
            ApplyLineColorToTable shape, colorRGB
        ' 일반 도형인 경우
        Else
            Set line = shape.line
            line.ForeColor.RGB = colorRGB
            If Not line.Visible Then line.Visible = msoTrue
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

' 테이블의 모든 셀 테두리에 색상 적용
Private Sub ApplyLineColorToTable(tableShape As shape, colorRGB As Long)
    Dim table As table
    Dim cell As cell
    Dim row As Long
    Dim col As Long
    Dim minRow As Long, maxRow As Long
    Dim minCol As Long, maxCol As Long
    
    On Error Resume Next
    
    Set table = tableShape.table
    
    minRow = 999
    minCol = 999
    maxRow = 0
    maxCol = 0
    
    If gBorderOutsideMode Then
        For row = 1 To table.Rows.Count
            For col = 1 To table.Columns.Count
                Set cell = table.cell(row, col)
                If cell.Selected Then
                    If row < minRow Then minRow = row
                    If col < minCol Then minCol = col
                    If row > maxRow Then maxRow = row
                    If col > maxCol Then maxCol = col
                End If
            Next col
        Next row
    End If
    
    ' 모든 셀의 테두리 색상 설정
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Set cell = table.cell(row, col)
            
            If cell.Selected Then
                With cell.borders
                    If Not gBorderOutsideMode Or row = minRow Then
                        .item(ppBorderTop).Visible = msoFalse
                        .item(ppBorderTop).Transparency = 0#
                        .item(ppBorderTop).ForeColor.RGB = colorRGB
                        If .item(ppBorderTop).weight <= 0 Then .item(ppBorderTop).weight = 1
                    End If
                    If Not gBorderOutsideMode Or row = maxRow Then
                        .item(ppBorderBottom).Visible = msoFalse
                        .item(ppBorderBottom).Transparency = 0#
                        .item(ppBorderBottom).ForeColor.RGB = colorRGB
                        If .item(ppBorderBottom).weight <= 0 Then .item(ppBorderBottom).weight = 1
                    End If
                    If Not gBorderOutsideMode Or col = minCol Then
                        .item(ppBorderLeft).Visible = msoFalse
                        .item(ppBorderLeft).Transparency = 0#
                        .item(ppBorderLeft).ForeColor.RGB = colorRGB
                        If .item(ppBorderLeft).weight <= 0 Then .item(ppBorderLeft).weight = 1
                    End If
                    If Not gBorderOutsideMode Or col = maxCol Then
                        .item(ppBorderRight).Visible = msoFalse
                        .item(ppBorderRight).Transparency = 0#
                        .item(ppBorderRight).ForeColor.RGB = colorRGB
                        If .item(ppBorderRight).weight <= 0 Then .item(ppBorderRight).weight = 1
                    End If
                End With
            End If
        Next col
    Next row
End Sub



Public Sub OnShapeTypeAdjustments(control As IRibbonControl)
    Dim parts() As String
    parts = Split(control.tag, ",")
    
    ApplyShapeTypeAdjustments CInt(parts(0)), CDbl(parts(1)), CDbl(parts(2))
End Sub

Private Sub ApplyShapeTypeAdjustments(shapeType As MsoAutoShapeType, Adjustment1 As Double, Adjustment2 As Double)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    For i = shapeRange.Count To 1 Step -1
        Set shape = shapeRange(i)
        
        shape.AutoShapeType = shapeType
        
        ' 둥근 정도 설정 (0 ~ 1 범위)
        If Adjustment1 >= 0 Then shape.Adjustments(1) = CmToAdjustment(shape.Width, shape.Height, Adjustment1)
        If Adjustment2 >= 0 Then shape.Adjustments(2) = CmToAdjustment(shape.Width, shape.Height, Adjustment2)
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다.", vbExclamation
End Sub


Public Sub OnLineWeight(control As IRibbonControl)
    ApplyLineWeight CDbl(control.tag)
End Sub


Private Sub ApplyLineWeight(ByVal lineWeight As Double)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        
        ' 테이블 셀인 경우
        If shape.HasTable Then
            ApplyLineWeightToTable shape, lineWeight
        ' 일반 도형인 경우
        Else ' If shape.fill.Type <> msoNoFill Then
            shape.line.weight = lineWeight
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

Private Sub ApplyLineWeightToTable(tableShape As shape, ByVal weight As Single)
    Dim table As table
    Dim cell As cell
    Dim row As Long
    Dim col As Long
    Dim minRow As Long, maxRow As Long
    Dim minCol As Long, maxCol As Long
    
    On Error Resume Next
    
    Set table = tableShape.table
    
    minRow = 999
    minCol = 999
    maxRow = 0
    maxCol = 0
    
    If gBorderOutsideMode Then
        For row = 1 To table.Rows.Count
            For col = 1 To table.Columns.Count
                Set cell = table.cell(row, col)
                If cell.Selected Then
                    If row < minRow Then minRow = row
                    If col < minCol Then minCol = col
                    If row > maxRow Then maxRow = row
                    If col > maxCol Then maxCol = col
                End If
            Next col
        Next row
    End If
    
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Set cell = table.cell(row, col)
            
            If cell.Selected Then
                With cell.borders
                    If Not gBorderOutsideMode Or row = minRow Then .item(ppBorderTop).weight = weight
                    If Not gBorderOutsideMode Or row = maxRow Then .item(ppBorderBottom).weight = weight
                    If Not gBorderOutsideMode Or col = minCol Then .item(ppBorderLeft).weight = weight
                    If Not gBorderOutsideMode Or col = maxCol Then .item(ppBorderRight).weight = weight
                End With
            End If
        Next col
    Next row
End Sub

Public Sub OnTableBordersLRRemove(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long

    Dim table As table
    Dim row As Long
    
    On Error GoTo ErrorHandler
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        If shape.HasTable Then
            Set table = shape.table
            For row = 1 To table.Rows.Count
                table.cell(row, 1).borders.item(ppBorderLeft).Visible = msoFalse
                table.cell(row, 1).borders.item(ppBorderLeft).Transparency = 1#
                table.cell(row, table.Columns.Count).borders.item(ppBorderRight).Visible = msoFalse
                table.cell(row, table.Columns.Count).borders.item(ppBorderRight).Transparency = 1#
            Next row
        End If
    Next
    Exit Sub
ErrorHandler:
    MsgBox "선택된 표가 없습니다. 표를 선택 후 시도해주세요.", vbExclamation
End Sub


Public Sub OnLineDashStyle(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long

    Dim table As table
    Dim cell As cell
    Dim dashStyle As MsoLineDashStyle
    Dim row, col As Long
    
    On Error GoTo ErrorHandler
    
    dashStyle = CInt(control.tag)
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
   
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        
        ' 테이블 셀인 경우
        If shape.HasTable Then
            ApplyLineDashStyleToTable shape, dashStyle
        ' 일반 도형인 경우
        Else ' If shape.fill.Type <> msoNoFill Then
            shape.line.dashStyle = dashStyle
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

Private Sub ApplyLineDashStyleToTable(tableShape As shape, ByVal dashStyle As MsoLineDashStyle)
    Dim table As table
    Dim cell As cell
    Dim row As Long
    Dim col As Long
    Dim minRow As Long, maxRow As Long
    Dim minCol As Long, maxCol As Long
    
    On Error Resume Next
    
    Set table = tableShape.table
    
    minRow = 999
    minCol = 999
    maxRow = 0
    maxCol = 0
    
    If gBorderOutsideMode Then
        For row = 1 To table.Rows.Count
            For col = 1 To table.Columns.Count
                Set cell = table.cell(row, col)
                If cell.Selected Then
                    If row < minRow Then minRow = row
                    If col < minCol Then minCol = col
                    If row > maxRow Then maxRow = row
                    If col > maxCol Then maxCol = col
                End If
            Next col
        Next row
    End If

    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Set cell = table.cell(row, col)
            If cell.Selected Then
                With cell.borders
                    If Not gBorderOutsideMode Or row = minRow Then .item(ppBorderTop).dashStyle = dashStyle
                    If Not gBorderOutsideMode Or row = maxRow Then .item(ppBorderBottom).dashStyle = dashStyle
                    If Not gBorderOutsideMode Or col = minCol Then .item(ppBorderLeft).dashStyle = dashStyle
                    If Not gBorderOutsideMode Or col = maxCol Then .item(ppBorderRight).dashStyle = dashStyle
                End With
            End If
        Next col
    Next row

End Sub


'==================================================
' 그룹5: 폰트 도구 (폰트 속성 변경)
'==================================================
Public Sub OnFontChange(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long

    Dim table As table
    Dim cell As cell
    Dim row, col As Long
    
    Dim textRange As textRange
    Dim fontParts() As String
    Dim fontName As String
    Dim fontColor As String
    Dim colorRGB As Long
    
    On Error GoTo ErrorHandler
    
    ' 태그에서 폰트 정보 파싱 (형식: 폰트명,색상16진수)
    fontParts = Split(control.tag, ",")
    
    If UBound(fontParts) < 1 Then
        MsgBox "폰트 정보가 올바르지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    fontName = Trim(fontParts(0))
    fontColor = Trim(fontParts(1))
    
    colorRGB = HexToRGB(fontColor)
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        Set textRange = ActiveWindow.Selection.textRange
        With textRange.Font
            .Name = fontName
            .NameFarEast = fontName
            .Color.RGB = colorRGB
        End With
    Else
        Set shapeRange = ActiveWindow.Selection.shapeRange
   
        For i = 1 To shapeRange.Count
            Set shape = shapeRange(i)
            
            If shape.HasTable Then
                Set table = shape.table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        Set cell = table.cell(row, col)
                        If cell.Selected Then
                            Set textRange = cell.shape.TextFrame.textRange
                            With textRange.Font
                                .Name = fontName
                                .NameFarEast = fontName
                                .Color.RGB = colorRGB
                            End With
                        End If
                    Next
                Next
            Else
                Set textRange = shape.TextFrame.textRange
                With textRange.Font
                    .Name = fontName
                    .NameFarEast = fontName
                    .Color.RGB = colorRGB
                End With
            End If
        Next
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "텍스트를 포함한 도형을 선택해주세요.", vbExclamation
End Sub

Public Sub OnParagraphSpaceWithin(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    
    Dim space As Double

    Dim table As table
    Dim cell As cell
    Dim row, col As Long
    
    Dim textRange As textRange
    
    On Error GoTo ErrorHandler
    
    space = CDbl(control.tag)
    
    If ActiveWindow.Selection.Type = ppSelectionText Then
        Set textRange = ActiveWindow.Selection.textRange
        textRange.ParagraphFormat.SpaceWithin = space
    Else
        Set shapeRange = ActiveWindow.Selection.shapeRange
   
        For i = 1 To shapeRange.Count
            Set shape = shapeRange(i)
            
            If shape.HasTable Then
                Set table = shape.table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        Set cell = table.cell(row, col)
                        If cell.Selected Then
                            Set textRange = cell.shape.TextFrame.textRange
                            textRange.ParagraphFormat.SpaceWithin = space
                        End If
                    Next
                Next
            Else
                Set textRange = shape.TextFrame.textRange
                textRange.ParagraphFormat.SpaceWithin = space
            End If
        Next
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "텍스트를 포함한 도형을 선택해주세요.", vbExclamation
End Sub


Public Sub OnCopyHexColor(control As IRibbonControl, id As String, index As Integer)
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "Text", "#" & Right(id, 6)
End Sub


Public Sub OnhapeSizeNormalize(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long
    Dim baseCm As Double
    
    On Error GoTo ErrorHandler
    
    baseCm = CDbl(control.tag)

    Set shapeRange = ActiveWindow.Selection.shapeRange
   
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        shape.Width = RoundPtToCmMultiple(shape.Width, baseCm)
        shape.Height = RoundPtToCmMultiple(shape.Height, baseCm)
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

Public Sub OnShapeLocNormalize(control As IRibbonControl)
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim i As Long

    Dim baseCm As Double
    
    On Error GoTo ErrorHandler
    
    baseCm = CDbl(control.tag)

    Set shapeRange = ActiveWindow.Selection.shapeRange
   
    For i = 1 To shapeRange.Count
        Set shape = shapeRange(i)
        shape.Left = RoundPtToCmMultiple(shape.Left, baseCm, 1)
        shape.Top = RoundPtToCmMultiple(shape.Top, baseCm, 2)
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

'==================================================
' 유틸리티 함수
'==================================================

' 16진수 문자열을 RGB 색상 값으로 변환
Private Function HexToRGB(hexColor As String) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    
    ' 16진수 문자열 정리
    hexColor = Replace(hexColor, "#", "")
    hexColor = Replace(hexColor, "0x", "")
    
    ' RGB 값 추출
    r = Val("&H" & Mid(hexColor, 1, 2))
    g = Val("&H" & Mid(hexColor, 3, 2))
    b = Val("&H" & Mid(hexColor, 5, 2))
    
    ' RGB 값을 Long으로 변환
    HexToRGB = RGB(r, g, b)
End Function

Private Function CmToAdjustment( _
        ByVal widthPt As Double, _
        ByVal heightPt As Double, _
        ByVal radiusCm As Double) As Double

    Dim radiusPt As Double
    Dim shortSide As Double

    radiusPt = radiusCm * PT_PER_CM

    shortSide = IIf(widthPt < heightPt, widthPt, heightPt)

    CmToAdjustment = radiusPt / shortSide

    If CmToAdjustment > 1 Then CmToAdjustment = 1
    If CmToAdjustment < 0 Then CmToAdjustment = 0

End Function


Function RoundPtToCmMultiple(ByVal pt As Double, ByVal cmMultiple As Double, Optional ByVal baseType As Integer = 0) As Double
    Dim cm As Double
    Dim base As Double
    
    base = 0
    
    If baseType = 1 Then base = ActiveWindow.presentation.PageSetup.SlideWidth / 2
    If baseType = 2 Then base = ActiveWindow.presentation.PageSetup.SlideHeight / 2

    If cmMultiple = 0 Then
        RoundPtToCmMultiple = 0
        Exit Function
    End If
    
    If base = 0 Then
        RoundPtToCmMultiple = Round(((pt - base) / PT_PER_CM + cmMultiple / 10#) / cmMultiple, 0) * cmMultiple * PT_PER_CM + base
    Else
        RoundPtToCmMultiple = Round(((pt - base) / PT_PER_CM - cmMultiple / 10#) / cmMultiple, 0) * cmMultiple * PT_PER_CM + base
    End If
End Function


