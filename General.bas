Attribute VB_Name = "General"
Option Explicit

'public scope variables
Public CurrentSettings() As Variant
Public ErrorStatus As ErrorNotice
Public booCheckPerformed As Boolean
Public myUser As UserSettings
Public strScenario As String

Public UMAFiles() As Variant
Public FinalUMAFiles() As Variant
    'Single dimension array: filenames only.

Public CUSIPS_C() As Variant    'sorted by cusip
Public CUSIPS_T() As Variant    'sorted by ticker
Public PrelimCUSIPS() As Variant
    '1.  Ticker
    '2.  CUSIP

'File handling variables
Public strCopyToFolder As String
Public SettingsFile As FileName
Public CUSIPFile As FileName
Public LogFile As String


'Constant to hold name of Settings File and CUSIP file and Log file.
Public Const conSettingsFile As String = _
    "\\mdtafile\MDT_Share\MDT-OPS\Boston & Pitt Ops\Daily UMA Holdings\Resources\Settings.txt"
'Public Const conCUSIPFile As String = "\\mdtafile\MDT_Share\QuantOps\Ops\Trading\CUSIP9.txt"
'Public Const conLogFile As String = _
    '"\\mdtafile\MDT_Share\MDT-OPS\Boston & Pitt Ops\Daily UMA Holdings\Output\UMALogFile.txt"



Sub ReportProgress(ByVal strMsg As String)
    'Procedure to print processing status on Main Form for user to see.
    frmMain.lblStatus.Caption = strMsg
End Sub


Sub CreateUMAFileNames(ByVal CurrentSettings As Variant, ByRef UMAFiles As Variant)
    'The purpose of this procedure is to create the complete
    'path and name for each file.  This is done by combining
    'various components from the CurrentSettings array and the date
    'from the Main Form.
    
    Dim i As Integer
    Dim j As Integer
    Dim intFileCounter As Integer
    Dim strDate As String
    Dim strPath As String
    
    'Establish date that will be used in file name
    strDate = Format(frmMain.tbDate.Text, "YYYYMMDD")
    
    'Lookup the name of the folder holding the UMA holdings files.
    'See function definition in AppSettings.
    strPath = GetSetting("DIR", "UMAAPP", "INPUT")

    
    'Count how many files there will be and size the array
    intFileCounter = 0
    For i = 1 To UBound(CurrentSettings)
        If CurrentSettings(i, 1) = "FIL" And CurrentSettings(i, 2) = "STRAT" Then
            intFileCounter = intFileCounter + 1
        End If
    Next i
    ReDim UMAFiles(1 To intFileCounter)
    
    'Create file names and write them to array.
    j = 0
    For i = 1 To UBound(CurrentSettings)
        If CurrentSettings(i, 1) = "FIL" And CurrentSettings(i, 2) = "STRAT" Then
            j = j + 1
            UMAFiles(j) = strPath & CurrentSettings(i, 4) & "_" & strDate & ".csv"
        End If
    Next i
End Sub

Sub WriteLogFile(ByVal strMsg As String)

    'The purpose of this procedure is to record an entry
    'in the log file.  The entries represent key actions
    'taken by the macro.  Each entry holds the name of the preparer,
    'the date and the time.  The variable passed to this
    'procedure is the text describing the action being logged.
    
    Dim intFile As Integer
    Dim strTime As String
    Dim strDate As String
    
    intFile = FreeFile
    strTime = Format(CStr(Time), "HH:MM:SS")
    strDate = Format(CStr(Date), "MM/DD/YYYY")

    Open LogFile For Append As intFile
        
        Print #intFile, strDate; "  "; strTime; "  "; myUser.LongName; "  "; strMsg

    Close intFile
End Sub


Sub RefineFileArray()
    'The purpose of this procedure is to eliminate from the UMAFiles array any file
    'names that are not to be processed in this session.  The CreateUMAFileNames
    'procedure runs when the program starts up and it loads into an array all the
    'possible file names that could be processed.  Using the option buttons on the main
    'form, the user can select to run only a particular program or all programs.
    
    'Note: If the user has selected All programs, this procedure does not eliminate
    'any file names from the array.
    

    Dim myFile As New FileName
    Dim intFileCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    intFileCount = 0
    If frmMain.obSB.Value = True Then
        
        'Count the number of Smith Barney files
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "sb" Then
                intFileCount = intFileCount + 1
            End If
        Next i
        
        'Set the size of the new array
        ReDim FinalUMAFiles(1 To intFileCount)
        
        'Load only the files needed by this session
        j = 0
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "sb" Then
                j = j + 1
                FinalUMAFiles(j) = myFile.LongName
            End If
        Next i

        
    ElseIf frmMain.obMS.Value = True Then
    
        'Count the number of Morgan Stanley files
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "ms" Then
                intFileCount = intFileCount + 1
            End If
        Next i
        
        'Set the size of the new array
        ReDim FinalUMAFiles(1 To intFileCount)
        
        'Load only the files needed by this session
        j = 0
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "ms" Then
                j = j + 1
                FinalUMAFiles(j) = myFile.LongName
            End If
        Next i
        
        
    ElseIf frmMain.obCiti.Value = True Then
        
        'Count the number of Citi DAP files
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "pe" Then
                intFileCount = intFileCount + 1
            End If
        Next i
        
        'Set the size of the new array
        ReDim FinalUMAFiles(1 To intFileCount)
        
        'Load only the files needed by this session
        j = 0
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            If Mid(myFile.Name, 8, 2) = "pe" Then
                j = j + 1
                FinalUMAFiles(j) = myFile.LongName
            End If
        Next i
        
    Else
    
        'This is for the All programs option.  The FinalUMAFiles array should be identical
        'to the UMAFiles array.
        
        'Set the size of the new array
        ReDim FinalUMAFiles(1 To UBound(UMAFiles))
        
        'Load all the files.
        j = 0
        For i = 1 To UBound(UMAFiles)
            myFile.LongName = UMAFiles(i)
            j = j + 1
            FinalUMAFiles(j) = myFile.LongName
        Next i
        
        
    End If
    
    Erase UMAFiles
    Set myFile = Nothing

End Sub
