Attribute VB_Name = "AppSettings"
Option Explicit


'Settings Array Fields
    
    '1.  SettingsType - can be EML, DIR, or FIL (Email, Directory, or File Name)
    '2.  Next broadest descriptive category
    '3.  Narrowest descriptive category
    '4.  The settings value
    'For example, FIL STRAT INPUT are the descriptors for a holdings file name
    'for a strategy.
    
    'The settings data is loaded into an array that has public scope and is
    'thus available throughout the application.
    
Sub LoadSettings(ByVal strFileName As String)
    'The purpose of this procedure is to load default settings into
    'application.  The settings are for email addresses, preparers, and
    'file locations.
    Dim SettingsFile As New FileName
    Dim intFile As Integer
    Dim strData As String
    Dim i As Integer
    
    
    'Check for file existence and re-size the settings array.
    With SettingsFile
        .LongName = strFileName
        If .FileCheck = False Then
            ErrorStatus.ErrorExists = True
            ErrorStatus.Description = "Settings File Not Found"
            Exit Sub
        Else
            ReDim CurrentSettings(1 To .Lines, 1 To 4)
        End If
    End With
    
    'Load data into array
    intFile = FreeFile
    Open SettingsFile.LongName For Input As intFile
        For i = 1 To UBound(CurrentSettings)
            Line Input #intFile, strData
                CurrentSettings(i, 1) = Left(strData, 3)
                CurrentSettings(i, 2) = Trim(Mid(strData, 5, 9))
                CurrentSettings(i, 3) = Trim(Mid(strData, 15, 9))
                CurrentSettings(i, 4) = Trim(Mid(strData, 25))
        Next i
    Close intFile
End Sub

Public Function GetSetting(ByVal SettingsType As String, ByVal Category As String, ByVal Subcategory As String) As String
    'This function is useful for settings that have unique descriptors.
    'If more than one setting matches the descriptors provided by the caller, the first
    'setting is returned.  If there are no matches, "NOTFOUND" is returned.
    Dim i As Integer
    
    GetSetting = "NOTFOUND"
    For i = 1 To UBound(CurrentSettings)
        If CurrentSettings(i, 1) = SettingsType Then
            If CurrentSettings(i, 2) = Category Then
                If CurrentSettings(i, 3) = Subcategory Then
                    GetSetting = CurrentSettings(i, 4)
                End If
            End If
        End If
    Next i
End Function

