Attribute VB_Name = "Globals"
Option Explicit

Public Const APP_GUID As String = "{8C1B8E75-4B35-48D1-9D8A-4C56A2A41F63}"

Public Type Config
    Initialized As Boolean
    FullName As String
    Path As String
    IniFile As String
    Icon As Object
    Value As Object
End Type

Public Type ToggleState
    BorderOutside As Boolean
    ParagraphSpace As Boolean
End Type

' --- Windows API 선언 (32비트 / 64비트 모두 호환) ---
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    Size As Long
    Type As Long
#If Win64 Then
    hPic As LongPtr
    hPal As LongPtr
#Else
    hPic As Long
    hPal As Long
#End If
    Reserved As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
#If Win64 Then
    DebugEventCallback As LongPtr
#Else
    DebugEventCallback As Long
#End If
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef InputBuf As GdiplusStartupInput, ByVal OutputBuf As LongPtr) As Long
    Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
    Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef image As LongPtr) As Long
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As Long) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPictureDisp) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef InputBuf As GdiplusStartupInput, ByVal OutputBuf As Long) As Long
    Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPictureDisp) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If


' --- Windows API 선언 (32비트 / 64비트 호환) ---
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#End If


Public gConfig As Config
Public gToggleState As ToggleState

Public Function AppGuid() As String
    AppGuid = APP_GUID
End Function

Private Sub initConfig()
    On Error Resume Next
    
    Dim addIn As addIn
    Dim GUID As String
    
    GUID = Application.run("AppGuid")
    If GUID = APP_GUID Then
        gConfig.FullName = ActivePresentation.FullName
        gConfig.Path = ActivePresentation.Path
        gConfig.IniFile = Left(gConfig.FullName, InStrRev(gConfig.FullName, ".")) & "ini"
        gConfig.Initialized = True
    Else
        For Each addIn In Application.AddIns
            If addIn.Loaded Then
                GUID = Application.run("'" & addIn.Name & "'!AppGuid")
                If GUID = APP_GUID Then
                    gConfig.FullName = addIn.FullName
                    gConfig.Path = addIn.Path
                    gConfig.IniFile = Left(gConfig.FullName, InStrRev(gConfig.FullName, ".")) & "ini"
                    gConfig.Initialized = True
                    Exit For
                End If
            End If
        Next
    End If
    
    On Error GoTo 0
    
    If gConfig.Icon Is Nothing Then Set gConfig.Icon = CreateObject("Scripting.Dictionary")
    If gConfig.Value Is Nothing Then Set gConfig.Value = CreateObject("Scripting.Dictionary")
End Sub

Public Function GetResFullName(ByVal filename As String) As String
    GetResFullName = gConfig.Path & "\res\" & filename
End Function

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Dim theme As String
    Dim key As String
    Dim val As String
    Dim i As Integer
    Dim fontTag() As String
    Dim image As IPictureDisp
    
    initConfig

    theme = ReadIni("common", "theme", gConfig.IniFile)

    For i = 1 To 13
        key = "color" & Format(i, "00")
        val = Trim(ReadIni(theme, key, gConfig.IniFile))
        Set image = LoadPNG(GetResFullName(val & ".png"))
        Set gConfig.Icon(key) = image
        gConfig.Value(key) = val
    Next

    For i = 1 To 7
        If i <= 2 Then key = "title" & Format(i, "00") Else key = "text" & Format(i - 2, "00")
        val = ReadIni(theme, key, gConfig.IniFile)

        fontTag = Split(val, ",")

        gConfig.Value(key) = fontTag
        Set image = LoadPNG(GetResFullName(Trim(fontTag(1)) & ".png"))
        Set gConfig.Icon(key) = image
    Next
End Sub


' --- 외부에서 호출할 INI 읽기 함수 ---
' Section : [세션명]
' Key     : 키이름
' FilePath: INI 파일의 전체 경로
' DefaultValue: 값이 없거나 에러 발생 시 반환할 기본값 (생략 가능)
Public Function ReadIni(ByVal Section As String, ByVal key As String, ByVal FilePath As String, Optional ByVal DefaultValue As String = "") As String
    Dim Buffer As String
    Dim BufferSize As Long
    Dim ResultCode As Long
    
    ' 읽어올 텍스트를 담을 버퍼 공간을 넉넉하게 확보 (255자)
    Buffer = space$(1024)
    BufferSize = Len(Buffer)
    
    ' API 호출
    ResultCode = GetPrivateProfileString(Section, key, DefaultValue, Buffer, BufferSize, FilePath)
    
    ' 호출 성공 시 버퍼에서 실제 읽어온 문자열만큼만 잘라서 리턴
    If ResultCode > 0 Then
        ReadIni = Left$(Buffer, ResultCode)
    Else
        ReadIni = DefaultValue
    End If
End Function


' --- 외부 호출용 PNG 로드 함수 ---
Public Function LoadPNG(ByVal FilePath As String) As IPictureDisp
#If Win64 Then
    Dim gdiToken As LongPtr
    Dim gdiBitmap As LongPtr
    Dim hBitmap As LongPtr
#Else
    Dim gdiToken As Long
    Dim gdiBitmap As Long
    Dim hBitmap As Long
#End If
    Dim gdiInput As GdiplusStartupInput
    Dim pAsg As GUID
    Dim pDesc As PICTDESC
    
    ' GDI+ 초기화
    gdiInput.GdiplusVersion = 1
    If GdiplusStartup(gdiToken, gdiInput, 0) <> 0 Then Exit Function
    
    ' 파일에서 PNG 이미지 로드
    If GdipLoadImageFromFile(StrPtr(FilePath), gdiBitmap) = 0 Then
        ' 리본의 투명 배경(알파채널) 유지를 위해 배경색은 0으로 지정하여 래핑
        If GdipCreateHBITMAPFromBitmap(gdiBitmap, hBitmap, 0) = 0 Then
            
            ' OLE IPictureDisp 객체 구조체 바인딩
            With pDesc
                .Size = Len(pDesc)
                .Type = 1 ' PICTYPE_BITMAP
                .hPic = hBitmap
                .hPal = 0
            End With
            
            ' IPictureDisp 가이드 ID (IID_IPictureDisp)
            With pAsg
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
                .Data4(0) = &H8B
                .Data4(1) = &HBB
                .Data4(2) = &H0
                .Data4(3) = &HAA
                .Data4(4) = &H0
                .Data4(5) = &H30
                .Data4(6) = &HC
                .Data4(7) = &HAB
            End With
            
            ' 핸들을 연결하여 최종 IPictureDisp 객체 생성
            Call OleCreatePictureIndirect(pDesc, pAsg, 1, LoadPNG)
        End If
        ' GDI+ 리소스 해제
        Call GdipDisposeImage(gdiBitmap)
    End If
    
    ' GDI+ 종료
    Call GdiplusShutdown(gdiToken)
End Function



