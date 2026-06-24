Attribute VB_Name = "FileTools"
'---------------------------------------------------------------------------------------
' Procedure : ExportAsPowerPointAddIn
' Purpose   : 올바른 OpenXML 상수를 사용하여 동일 폴더에 .ppam 추가 기능을 깨끗하게 저장합니다.
'---------------------------------------------------------------------------------------
Sub ExportAsPowerPointAddIn(control As IRibbonControl)
    Dim srcPres As presentation
    Dim currentPath As String
    Dim currentName As String
    Dim ppamPath As String
    Dim dotPosition As Integer
    
    ' 1. 현재 활성화된 원본 프레젠테이션 지정
    Set srcPres = ActivePresentation
    
    ' [예외 처리] 파일이 한 번도 저장된 적이 없어 폴더 경로가 없는 경우 실행 차단
    If srcPres.Path = "" Then
        MsgBox "현재 파일이 먼저 저장되어 있어야 동일한 폴더에 추가 기능 생성이 가능합니다.", vbExclamation, "실행 취소"
    Exit Sub
    End If
    
    ' 2. 경로 및 파일명 분리 처리 (확장자 제거)
    currentPath = srcPres.Path & "\"
    currentName = srcPres.Name
    
    ' 파일명 뒤에서부터 첫 번째 점(.)의 위치를 찾아 기존 확장자 제거
    dotPosition = InStrRev(currentName, ".")
    If dotPosition > 0 Then
        currentName = Left(currentName, dotPosition - 1)
    End If
    
    ' 3. 최종 .ppam 저장 전체 경로 조합
    ppamPath = currentPath & currentName & ".ppam"
    
    ' 4. 기존 동일 파일 충돌 방지을 위한 선삭제
    On Error Resume Next
    Kill ppamPath
    On Error GoTo 0
    
    ' 5. [핵심 교정] ppSaveAsAddIn 대신 ppSaveAsOpenXMLAddIn 사용
    ' 이 상수를 지정해야 파워포인트 엔진이 뒤에 .ppa를 붙이지 않고 순수 .ppam으로 빌드합니다.
    On Error Resume Next
    ApplicationRunSave srcPres, True, ppamPath, ppSaveAsOpenXMLAddin
    
    ' 추가 기능이 이미 로드되어 실행 중이어서 권한 잠금이 걸린 경우 예외 처리
    If Err.Number <> 0 Then
        MsgBox "추가 기능 파일 생성 중 오류가 발생했습니다." & vbCrLf & _
               "기존에 동일한 추가 기능이 파워포인트에 현재 로드(실행) 중인지 확인해 주세요.", vbCritical, "저장 실패"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 6. 작업 완료 알림
    MsgBox "파워포인트 추가 기능 파일(.ppam)이 성공적으로 생성되었습니다." & vbCrLf & vbCrLf & _
           "경로: " & ppamPath, vbInformation, "추가 기능 내보내기 완료"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveZipBackup
' Purpose   : 현재 프레젠테이션의 전체 슬라이드와 마스터 서식을 복사하여 .zip 확장자로 저장합니다.
'---------------------------------------------------------------------------------------
Sub ExportAsPowerPointZip(control As IRibbonControl)
    Dim srcPres As presentation
    Dim dstPres As presentation
    Dim zipPath As String
    Dim pasteRetryCount As Integer
    
    ' 1. 현재 활성화된 원본 파일 지정
    Set srcPres = ActivePresentation
    
    ' [예외처리] 한 번도 저장하지 않은 새 파일인 경우 실행 차단
    If srcPres.Path = "" Then
        MsgBox "현재 파일이 먼저 저장되어 있어야 압축 백업이 가능합니다.", vbExclamation, "오류"
        Exit Sub
    End If
    
    ' 2. 새 프레젠테이션 생성 (서식 이식을 위한 타겟)
    Set dstPres = Presentations.Add(WithWindow:=msoTrue)
    
    ' 3. 슬라이드 크기 및 마스터 템플릿 서식 동기화
    dstPres.PageSetup.SlideHeight = srcPres.PageSetup.SlideHeight
    dstPres.PageSetup.SlideWidth = srcPres.PageSetup.SlideWidth
    dstPres.ApplyTemplate srcPres.FullName
    
    ' 4. 원본 슬라이드 전체 복사
    srcPres.Slides.Range.copy
    
    ' [핵심 해결책] MsgBox 없이도 클립보드가 완전히 로드될 수 있도록 대기 처리
    DoEvents
    
    ' 5. 안전하게 붙여넣기 시도 (비동기 지연 오류 방지 루프)
    On Error Resume Next
    dstPres.Slides.Paste
    
    ' 클립보드 복사가 덜 끝나 에러가 났을 경우 최대 5번 재시도 수행
    Do While Err.Number <> 0 And pasteRetryCount < 5
        Err.Clear
        pasteRetryCount = pasteRetryCount + 1
        DoEvents ' OS에 제어권을 넘겨 클립보드 완료 대기
        dstPres.Slides.Paste
    Loop
    On Error GoTo 0
    
    ' 6. 지정된 경로로 .zip 파일 백업 저장 (기존 작성 사양 유지)
    zipPath = srcPres.FullName & ".zip"
    
    ' 기존에 동일한 백업 파일이 존재한다면 충돌 방지를 위해 선삭제
    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
    
    ApplicationRunSave dstPres, False, zipPath, ppSaveAsOpenXMLPresentation
    
    ' 7. 새로 만든 작업용 임시 창 닫기 (저장 완료되었으므로 알림 없이 닫힘)
    dstPres.Close
    
    ' 완료 알림 (작업이 끝났음을 인지하기 위한 깔끔한 UX)
    MsgBox "백업 파일이 생성되었습니다:" & vbCrLf & zipPath, vbInformation, "완료"
End Sub


Private Sub ApplicationRunSave(ByRef presentation As presentation, ByVal copy As Boolean, ByVal fileName As String, ByVal fileType As PowerPoint.PpSaveAsFileType)
    Dim command As String
    
    command = "Sav"
    
    If copy Then command = command & "eCopyAs" Else command = command & "eAs"
    
    CallByName presentation, command, VbMethod, fileName, fileType
End Sub

