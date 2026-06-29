Attribute VB_Name = "Globals"

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

Public gConfig As Config
Public gToggleState As ToggleState

Public Function AppGuid() As String
    AppGuid = APP_GUID
End Function

Public Sub initConfig()
    On Error Resume Next
    
    Dim addIn As addIn
    Dim guid As String
    
    guid = Application.Run("AppGuid")
    If guid = APP_GUID Then
        gConfig.FullName = ActivePresentation.FullName
        gConfig.Path = ActivePresentation.Path
        gConfig.IniFile = Left(gConfig.FullName, InStrRev(gConfig.FullName, ".")) & "ini"
        Initialized = True
    Else
        For Each addIn In Application.AddIns
            If addIn.Loaded Then
                guid = Application.Run("'" & addIn.Name & "'!AppGuid")
                If guid = APP_GUID Then
                    gConfig.FullName = addIn.FullName
                    gConfig.Path = addIn.Path
                    gConfig.IniFile = Left(gConfig.FullName, InStrRev(gConfig.FullName, ".")) & "ini"
                    Initialized = True
                    Exit For
                End If
            End If
        Next
    End If
    
    On Error GoTo 0
    
    If gConfig.Icon Is Nothing Then Set gConfig.Icon = CreateObject("Scripting.Dictionary")
    If gConfig.Value Is Nothing Then Set gConfig.Value = CreateObject("Scripting.Dictionary")
End Sub

Public Function GetResFullName(ByVal fileName As String) As String
    GetResFullName = gConfig.Path & "\res\" & fileName
End Function

Public Function ReadIni(ByVal section As String, ByVal key As String, Optional ByVal sp As String = "") As Variant

End Function

Public Sub loadIcon(ByVal id As String, ByVal fileName As String)

End Sub

