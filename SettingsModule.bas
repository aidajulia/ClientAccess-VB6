Attribute VB_Name = "SettingsModule"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpDefault As String, _
                        ByVal lpReturnedString As String, _
                        ByVal nSize As Long, _
                        ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" _
                        (ByVal lpApplicationName As String, _
                        ByVal lpKeyName As Any, _
                        ByVal lpString As Any, _
                        ByVal lpFileName As String) As Long

Public Function ReadSettingString(sSection As String, sKeyName As String) As String
    Dim sRet As String
    Dim IniFileName As String
    IniFileName = GetIniFileName()
    If IniFileName = "" Then Exit Function
    If Dir(IniFileName, vbNormal) = "" Then Exit Function

  sRet = String(255, Chr(0))
  ReadSetting = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), IniFileName))
End Function
Public Function ReadSettingNumeric(sSection As String, sKeyName As String) As Double
    Dim sRet As String
    Dim IniFileName As String
    Dim TempValue As String
    IniFileName = GetIniFileName()
    If IniFileName = "" Then Exit Function
    If Dir(IniFileName, vbNormal) = "" Then Exit Function

  sRet = String(255, Chr(0))
  TempValue = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), IniFileName))
  ReadSettingNumeric = Val(TempValue)
End Function

Public Function WriteSetting(sSection As String, sKeyName As String, sValue As String) As Boolean
    Dim IniFileName As String
    IniFileName = GetIniFileName()
    If IniFileName = "" Then Exit Function
 
    Call WritePrivateProfileString(sSection, sKeyName, sValue, IniFileName)
    WriteSetting = (Err.Number = 0)
End Function

Public Function GetIniFileName() As String
    Dim IniFileName As String
    IniFileName = App.Path
    If Right(IniFileName, 1) <> "\" Then IniFileName = IniFileName & "\"
    IniFileName = IniFileName & App.EXEName & ".ini"
    GetIniFileName = IniFileName
End Function
