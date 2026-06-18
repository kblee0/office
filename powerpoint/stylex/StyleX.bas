Attribute VB_Name = "StyleX"
'==================================================
' StyleX RibbonX PowerPoint 매크로
' SPEC: 스타일 도구모음
'==================================================

Option Explicit

'==================================================
' StyleX RibbonX PowerPoint 매크로
' SPEC: 스타일 도구모음
'==================================================

Option Explicit

'==================================================
' 그룹1: 채우기 (배경색 변경 - 다중 선택 지원)
'==================================================
Public Sub OnFillColorClick(control As IRibbonControl)
    Dim colorHex As String
    Dim colorRGB As Long
    Dim shapeRange As shapeRange
    Dim shape As shape
    Dim fill As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 버튼 ID에서 16진수 추출 (예: C2F2F2F -> 2F2F2F)
    colorHex = Right(control.Id, 6)
    colorRGB = HexToRGB(colorHex)
    
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
            fill.ForeColor.RGB = colorRGB
            fill.Transparency = 0
        End If
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다. 도형을 선택 후 시도해주세요.", vbExclamation
End Sub

'==================================================
' 그룹2: 선색 (테두리선 색상 변경 - 다중 선택 지원)
'==================================================
Public Sub OnLineColorClick(control As IRibbonControl)
    Dim colorHex As String
    Dim colorRGB As Long
    Dim shapeRange As Object
    Dim shape As Object
    Dim line As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 버튼 ID에서 16진수 추출 (예: L2F2F2F -> 2F2F2F)
    colorHex = Right(control.Id, 6)
    colorRGB = HexToRGB(colorHex)
    
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

'==================================================
' 그룹3: 사각형도구 (다중 선택 지원)
'==================================================

' 직사각형으로 변경
Public Sub OnRectangleShape(control As IRibbonControl)
    Dim shapeRange As Object
    Dim shape As shape
    Dim slide As Object
    Dim newShape As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    Set slide = ActiveWindow.Selection.shapeRange(1).Parent
    
    ' 선택된 모든 도형을 직사각형으로 변환
    For i = shapeRange.Count To 1 Step -1
        Set shape = shapeRange(i)
        
        ' 새로운 직사각형 생성 msoShapeRectangle ' AutoShapeType msoShapeRound2SameRectangle
        shape.AutoShapeType = msoShapeRectangle
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다.", vbExclamation
End Sub

' 0.1cm 둥근 모서리 직사각형
Public Sub OnRoundedRectangle01(control As IRibbonControl)
    ChangeShapeToRoundedRectangle msoShapeRoundedRectangle, 0.1, -1
End Sub

' 0.2cm 둥근 모서리 직사각형
Public Sub OnRoundedRectangle02(control As IRibbonControl)
    ChangeShapeToRoundedRectangle msoShapeRoundedRectangle, 0.2, -1
End Sub

' 위쪽 0.1cm 둥근 모서리 직사각형
Public Sub OnTopRoundedRectangle01(control As IRibbonControl)
    ' 위쪽만 둥근 모서리 (근사치 구현)
    ChangeShapeToRoundedRectangle msoShapeRound2SameRectangle, 0.1, 0
End Sub

' 위쪽 0.2cm 둥근 모서리 직사각형
Public Sub OnTopRoundedRectangle02(control As IRibbonControl)
    ChangeShapeToRoundedRectangle msoShapeRound2SameRectangle, 0.2, 0
End Sub

' 아래쪽 0.1cm 둥근 모서리 직사각형
Public Sub OnBottomRoundedRectangle01(control As IRibbonControl)
    ChangeShapeToRoundedRectangle msoShapeRound2SameRectangle, 0, 0.1
End Sub

' 아래쪽 0.2cm 둥근 모서리 직사각형
Public Sub OnBottomRoundedRectangle02(control As IRibbonControl)
    ChangeShapeToRoundedRectangle msoShapeRound2SameRectangle, 0, 0.2
End Sub

' 둥근 직사각형으로 변경 헬퍼 함수 (다중 선택 지원)
Private Sub ChangeShapeToRoundedRectangle(shapeType As MsoAutoShapeType, roundValue1 As Double, roundValue2 As Double)
    Dim shapeRange As Object
    Dim shape As shape
    Dim slide As Object
    Dim newShape As Object
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    Set slide = shapeRange(1).Parent
    
    ' 선택된 모든 도형을 둥근 직사각형으로 변환
    For i = shapeRange.Count To 1 Step -1
        Set shape = shapeRange(i)
        
        shape.AutoShapeType = shapeType
        
        ' 둥근 정도 설정 (0 ~ 1 범위)
        shape.Adjustments(1) = roundValue1
        
        If roundValue2 > 0 Then shape.Adjustments(2) = roundValue2
        
    Next i
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 도형이 없습니다.", vbExclamation
End Sub

'==================================================
' 그룹4: 표 도구 (테이블 테두리 설정)
'==================================================

' 모든 테두리 0.5pt
Public Sub OnTableBorder05(control As IRibbonControl)
    ApplyTableBorders 0.5, True, True
End Sub

' 모든 테두리 1.0pt
Public Sub OnTableBorder10(control As IRibbonControl)
    ApplyTableBorders 1, True, True
End Sub

' 모든 테두리 1.5pt
Public Sub OnTableBorder15(control As IRibbonControl)
    ApplyTableBorders 1.5, True, True
End Sub

' 좌우 제외 0.5pt
Public Sub OnTableBorder05NoLR(control As IRibbonControl)
    ApplyTableBorders 0.5, False, True
End Sub

' 좌우 제외 1.0pt
Public Sub OnTableBorder10NoLR(control As IRibbonControl)
    ApplyTableBorders 1, False, True
End Sub

' 좌우 제외 1.5pt
Public Sub OnTableBorder15NoLR(control As IRibbonControl)
    ApplyTableBorders 1.5, False, True
End Sub

' 테이블 테두리 적용 헬퍼 함수
Private Sub ApplyTableBorders(thickness As Double, includeLR As Boolean, includeAll As Boolean)
    Dim table As Object
    Dim shape As Object
    Dim lineWeight As Double
    Dim borders As borders
    Dim row As Long
    Dim col As Long
    
    On Error GoTo ErrorHandler
    
    ' 선택된 표 가져오기
    Set shape = ActiveWindow.Selection.shapeRange(1)
    Set table = shape.table
    
    ' 선 두께 변환 (포인트 -> 포인트)
    lineWeight = thickness * 1.333  ' 1포인트 = 1.333
    
    ' 모든 셀의 테두리 설정
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            
            Set borders = table.cell(row, col).borders
            With borders
                .item(ppBorderTop).Weight = lineWeight
                .item(ppBorderBottom).Weight = lineWeight
                .item(ppBorderLeft).Weight = lineWeight
                .item(ppBorderRight).Weight = lineWeight
                .item(ppBorderTop).Visible = msoTrue
                .item(ppBorderBottom).Visible = msoTrue
                .item(ppBorderLeft).Visible = msoTrue
                .item(ppBorderRight).Visible = msoTrue
                .item(ppBorderTop).Transparency = 0#
                .item(ppBorderBottom).Transparency = 0#
                .item(ppBorderLeft).Transparency = 0#
                .item(ppBorderRight).Transparency = 0#
                
                If Not includeLR Then
                    If col = 1 Then
                        .item(ppBorderLeft).Visible = msoFalse
                        .item(ppBorderLeft).Transparency = 1#       ' 투명도 100%
                    End If
                    If col = table.Columns.Count Then
                        .item(ppBorderRight).Visible = msoFalse
                        .item(ppBorderRight).Transparency = 1#        ' 투명도 100%
                    End If
                End If
                
            End With
        Next col
    Next row
    
    Exit Sub
ErrorHandler:
    MsgBox "선택된 표가 없습니다. 표를 선택 후 시도해주세요.", vbExclamation
End Sub

'==================================================
' 그룹5: 폰트 도구 (폰트 속성 변경)
'==================================================
Public Sub OnFontChange(control As IRibbonControl)
    Dim fontInfo As String
    Dim fontParts() As String
    Dim fontName As String
    Dim fontColor As String
    Dim fontBold As Boolean
    Dim colorRGB As Long
    Dim shape As Object
    
    On Error GoTo ErrorHandler
    
    ' 태그에서 폰트 정보 파싱 (형식: 폰트명,색상16진수,굵은체여부)
    fontInfo = control.Tag
    fontParts = Split(fontInfo, ",")
    
    If UBound(fontParts) < 2 Then
        MsgBox "폰트 정보가 올바르지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    fontName = Trim(fontParts(0))
    fontColor = Trim(fontParts(1))
    fontBold = (Trim(fontParts(2)) = "true")
    
    colorRGB = HexToRGB(fontColor)
    
    ' 선택된 텍스트 또는 도형 텍스트에 폰트 적용
    Set shape = ActiveWindow.Selection.shapeRange(1)
    
    With shape.TextFrame.TextRange.Font
        .Name = fontName
        .Color.RGB = colorRGB
        .Bold = fontBold
    End With
    
    Exit Sub
ErrorHandler:
    MsgBox "텍스트를 포함한 도형을 선택해주세요.", vbExclamation
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

' 테이블의 모든 셀에 배경색 적용
Private Sub ApplyFillColorToTable(tableShape As Object, colorRGB As Long)
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

' 테이블의 모든 셀 테두리에 색상 적용
Private Sub ApplyLineColorToTable(tableShape As Object, colorRGB As Long)
    Dim table As table
    Dim cell As cell
    Dim row As Long
    Dim col As Long
    
    On Error Resume Next
    
    Set table = tableShape.table
    
    ' 모든 셀의 테두리 색상 설정
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Set cell = table.cell(row, col)
            
            If cell.Selected Then
                With cell.borders
                    .item(ppBorderTop).ForeColor.RGB = colorRGB
                    .item(ppBorderBottom).ForeColor.RGB = colorRGB
                    .item(ppBorderLeft).ForeColor.RGB = colorRGB
                    .item(ppBorderRight).ForeColor.RGB = colorRGB
                End With
            End If
        Next col
    Next row
End Sub

' 도형 포맷팅 복사
Private Sub CopyShapeFormatting(sourceShape As Object, targetShape As Object)
    On Error Resume Next
    
    With targetShape
        ' 채우기 복사
        If sourceShape.fill.Type <> msoNoFill Then
            .fill.Copy
            .fill.Paste
        End If
        
        ' 선 복사
        If sourceShape.line.Visible Then
            .line.Copy
            .line.Paste
        End If
        
        ' 그림자 복사
        .Shadow.Copy
        .Shadow.Paste
        
        ' 텍스트 복사
        If sourceShape.HasTextFrame Then
            .TextFrame.TextRange.Copy
            .TextFrame.TextRange.Paste
        End If
    End With
End Sub

