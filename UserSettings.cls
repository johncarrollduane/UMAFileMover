VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The purpose of this class is to create an object variable
'that will make it easy to identify the user through the use of a central
'repository of user data.  This class works only if accompanied by
'the module called GetUser.  This module contains a Windows API function
'to get the network ID of the user directly from the operating system.


Private pUsers() As String
'1.  Network ID
'2.  Advent user folder ID
'3.  Telephone extension (4 digits)
'5.  First and last name

Private pNetworkID As String


Private Sub Class_Initialize()
    
    'This procedure runs when the class is initialized.
    'It loads an array with user data from a text file
    'called UserSettings.txt.  This data is used throughout
    'this class.
    
    Dim strFileName As String
    Dim strData As String
    Dim i As Integer
    Dim intFile As Integer
    
    'Get user settings file name.  See function definition in AppSettings.
    'strFileName = "\\mdtafile\MDT_Share\MDT-OPS\Boston Ops\Admin\Systems - General\UserSettings.txt"
    strFileName = GetSetting("DIR", "SYSTEMS", "GENERAL") & GetSetting("FIL", "SYSTEMS", "USERS")
    
    'First count how many rows are in the file.
    'Then resize the pUsers() array to fit the number
    'of rows in the file.
    intFile = FreeFile
    i = 0
    Open strFileName For Input As intFile
        Do Until EOF(intFile)
            Line Input #intFile, strData
            i = i + 1
        Loop
    Close intFile
    ReDim pUsers(1 To i, 1 To 4)
    
    'Load the contents of UserSettings.txt into pUsers().
    intFile = FreeFile
    i = 0
    Open strFileName For Input As intFile
        Do Until EOF(intFile)
            Line Input #intFile, strData
            i = i + 1
            pUsers(i, 1) = Left(strData, 7)
            pUsers(i, 2) = Mid(strData, 9, 3)
            pUsers(i, 3) = Mid(strData, 13, 4)
            pUsers(i, 4) = Trim(Mid(strData, 18, 25))
        Loop
    Close intFile
    
    
End Sub



Public Property Get AdventFolder() As String
    Dim i As Integer
    
    For i = 1 To UBound(pUsers)
        If NetworkID = pUsers(i, 1) Then
            AdventFolder = pUsers(i, 2)
            Exit For
        End If
    Next i
End Property



Public Property Get LongName() As String
    Dim i As Integer
    
    For i = 1 To UBound(pUsers)
        If NetworkID = pUsers(i, 1) Then
            LongName = pUsers(i, 4)
            Exit For
        End If
    Next i
End Property



Public Property Get Telephone() As String
    Dim i As Integer
    
    For i = 1 To UBound(pUsers)
        If NetworkID = pUsers(i, 1) Then
            Telephone = "(617) 235-" & pUsers(i, 3)
            Exit For
        End If
    Next i
End Property



Public Property Get NetworkID() As String
    NetworkID = pNetworkID
End Property


Public Property Let NetworkID(ByVal strNewValue As String)
    pNetworkID = strNewValue
End Property
