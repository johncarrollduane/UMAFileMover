VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MDT Advisers"
   ClientHeight    =   7695
   ClientLeft      =   5775
   ClientTop       =   2865
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   7200
   Begin VB.Frame Frame1 
      Caption         =   "Select UMA Programs to Process"
      Height          =   2055
      Left            =   1800
      TabIndex        =   11
      Top             =   3000
      Width           =   3735
      Begin VB.OptionButton obCiti 
         Caption         =   "Citigroup Dynamic Allocation Portfolios"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   3495
      End
      Begin VB.OptionButton obMS 
         Caption         =   "Morgan Stanley PPA"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton obSB 
         Caption         =   "Smith Barney Select UMA"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton obAll 
         Caption         =   "All Programs"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.CommandButton btnFix 
      Caption         =   "Repair Data"
      Height          =   615
      Left            =   2396
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox tbDate 
      Height          =   375
      Left            =   3319
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton btnCheck 
      Caption         =   "Check Files"
      Height          =   615
      Left            =   1069
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5036
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Move Files"
      Height          =   615
      Left            =   3716
      TabIndex        =   0
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblUser 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2933
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Effective Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "UMA Holdings File Mover (Varonis)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   5535
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuProcedure 
         Caption         =   "&Procedure"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCheck_Click()
    'The purpose of this procedure is to check the UMA holdings files
    'one by one for any ticker symbols or CUSIP's that are not found
    'in APL.  It also checks to see that all files that need to be
    'checked to indeed exist.  This provides the benefit of checking
    'the spelling of the file name (since the files were created
    'manually.  The list of files to check is taken from the
    'UMAFiles() array which was filled earlier with filenames from
    'the Settings file.
    
    Dim i As Integer
    Dim j As Integer
    Dim intFile As Integer
    Dim strFound As String
    Dim strNotFound As String
    Dim strMsg As String
    Dim strData As String
    Dim Sorter As New SortOp
    Dim TempFile As New FileName
    Dim intNotSecurities As Integer
    
    ReportProgress ("Checking Files")
    WriteLogFile ("Checking Files: Begin Processing")
    
    'Eliminate any unneeded files from the UMAFiles array.
    RefineFileArray
    
    'Look up in the UMAFiles() array the path of the UMA files are to be copied to.
    'Using the UMAFile object, check for the existence of each file.
    'Load the file names of those found and not found in the
    'appropriate string variable.  These variables will be used
    'later to report file existence to the user.
    For i = 1 To UBound(FinalUMAFiles)
        Dim UMAFile As New FileName
        UMAFile.LongName = FinalUMAFiles(i)
        If UMAFile.FileCheck = True Then
            strFound = strFound & UMAFile.Name & vbCrLf
        Else
            strNotFound = strNotFound & UMAFile.Name & vbCrLf
        End If
    Next i
    
    'If variables holding the lists are empty, fill the variable with an
    'appropriate message.
    If strFound = Empty Then
        strFound = "No Files Found"
        MsgBox "No files were found.", vbOKOnly, "File Check"
        Exit Sub
    End If
    
    If strNotFound = Empty Then
        strNotFound = "All Files Were Found"
    End If
    
    'Construct the message to be displyed to the user.
    strMsg = "FILES FOUND THAT CAN BE MOVED:" & vbCrLf
    strMsg = strMsg & strFound & vbCrLf & vbCrLf
    strMsg = strMsg & "FILES NOT FOUND:" & vbCrLf
    strMsg = strMsg & strNotFound
    
    'Scan each file's contents and validate CUSIP's.
    
    'First open the CUSIP file and read the CUSIP data into an array.
    'File format:
    '1.  CUSIP
    '2.  Ticker
    '3.  Security name
    '4.  Price
    '5.  CRSP #
    '6.  Is current ticker or not?
    
    'Read the entire contents of the CUSIP file into the PrelimCUSIPs array.
    'PrelimCUSIPS() fields:
    '1.  Ticker
    '2.  CUSIP
    '3.  Is current?
    
    intFile = FreeFile
    i = 0
    Open CUSIPFile.LongName For Input As intFile
        For i = 1 To UBound(PrelimCUSIPS)
            Line Input #intFile, strData
            PrelimCUSIPS(i, 1) = Trim(Mid(strData, 11, 6))
            PrelimCUSIPS(i, 2) = Left(strData, 9)
            PrelimCUSIPS(i, 3) = Trim(Mid(strData, 72, 3))
            
            If PrelimCUSIPS(i, 3) = "no" Then
                intNotSecurities = intNotSecurities + 1
            End If
        Next i
    Close intFile
    
    
    
    'Fill the final CUSIPS() array with all the items in the PrelimCUSIPS
    'array that do not have security type 2.
    'CUSIPS() fields:
    '1.  Ticker
    '2.  CUSIP
    '3.  Is current?
    ReDim CUSIPS_C(1 To UBound(PrelimCUSIPS) - intNotSecurities, 1 To 2)
    ReDim CUSIPS_T(1 To UBound(PrelimCUSIPS) - intNotSecurities, 1 To 2)
    j = 0
    For i = 1 To UBound(PrelimCUSIPS)
        If PrelimCUSIPS(i, 3) <> "no" Then
            j = j + 1
            CUSIPS_C(j, 1) = PrelimCUSIPS(i, 1)
            CUSIPS_C(j, 2) = PrelimCUSIPS(i, 2)
        End If
    Next i
    Erase PrelimCUSIPS
    CUSIPS_T = CUSIPS_C

    
    'Next sort the CUSIPS() array by Ticker to make searching easier.
    With Sorter
        .DataType = dtText
        .Descending = False
        .SortColumn = 1
        .Sort CUSIPS_T
    End With
    
    With Sorter
        .DataType = dtText
        .Descending = False
        .SortColumn = 2
        .Sort CUSIPS_C
    End With
    Set Sorter = Nothing
    
    'Finally, loop through each file and scan using the procedure in Module
    'CUSIP_Check.  Only call the CheckCUSIPs procedure if the file
    'exists.
    ErrorStatus.CusipProblemNumber = 0
    ErrorStatus.CusipProblemExists = False
    
    For i = 1 To UBound(FinalUMAFiles)
        TempFile.LongName = FinalUMAFiles(i)
        If TempFile.FileCheck = True Then
            
            'If the file contains a "pe", the file was obtained
            'from APL and thus has a unique format that must be
            'changed to the Smith Barney format.  The ConvertAPLFile
            'function does this.  See the APL_Files module.
            If Mid(TempFile.Name, 8, 2) = "pe" Then
                ConvertAPLFile (TempFile.LongName)
            End If
            
            CheckCUSIPs TempFile.LongName, Left(TempFile.Name, 2), _
                Mid(TempFile.Name, 10, 1)
        End If
    Next i

    
    'Report processing results.
    MsgBox strMsg, vbOKOnly, "Search Results: UMA Files To Move"
    If ErrorStatus.CusipProblemExists = True Then
        If ErrorStatus.CusipProblemNumber = 1 Then
            MsgBox "There is " & ErrorStatus.CusipProblemNumber & _
                " CUSIP problem." & vbCrLf & "Check Log File for Details: " & _
                LogFile, vbOKOnly, "CUSIP Check"
        Else
            MsgBox "There are " & ErrorStatus.CusipProblemNumber & _
                " CUSIP problems." & vbCrLf & "Check Log File for Details: " & _
                LogFile, vbOKOnly, "CUSIP Check"
        End If
    Else
        MsgBox "There are no CUSIP Problems.", vbOKOnly, "CUSIP Check"
    End If
    
    If ErrorStatus.TickerProblemExists = True Then
        If ErrorStatus.TickerProblemNumber = 1 Then
            MsgBox "There is " & ErrorStatus.TickerProblemNumber & _
                " ticker symbol problem." & vbCrLf & "Check Log File for Details: " & _
                LogFile, vbOKOnly, "Ticker Check"
        Else
            MsgBox "There are " & ErrorStatus.TickerProblemNumber & _
                " ticker symbol problems." & vbCrLf & "Check Log File for Details: " & _
                LogFile, vbOKOnly, "Ticker Check"
        End If
    Else
        MsgBox "There are no ticker symbol problems.", vbOKOnly, "Ticker Check"
    End If
    
    'Fill "Check Performed?" variable with true to indicate that the check has been
    'performed.
    booCheckPerformed = True
    
    WriteLogFile ("Checking Files: End Processing")
    ReportProgress ("Checking Completed")
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnFix_Click()
    'The purpose of this procedure is to scan the holdings
    'files and check to see if the CUSIP's and Tickers in the file are
    'recognized, that is, agree to the CUSIP in the APL CUSIP
    'file.  If the CUSIP does not agree, replace it with the
    'proper CUSIP.  Another data fix is to eliminate rows that
    'represent zero quantity holdings.
    
    'The reason some CUSIP's need to be repaired is because the
    'download from the web site of the UMA sponsors may cause
    'leading zeroes on the CUSIP's to be dropped and CUSIP's
    'containing the letter E to be converted to scientific notation.
    'Another problem is sometimes the CUSIP will have a leading character
    'that is not really part of the CUSIP, most often a lower case "a"
    'with an accent mark (ASCII character 256).
    
    Dim i As Integer
    Dim intFile As Integer
    Dim strFound As String
    Dim strNotFound As String
    Dim strMsg As String
    Dim strData As String
    Dim Sorter As New SortOp
    Dim TempFile As New FileName
    
    WriteLogFile ("Fixing Data: Begin Processing")
    ReportProgress ("Fixing Data")
    
    'Do not go forward unless the check has been performed.
    If booCheckPerformed = False Then
        MsgBox "Please perform Check before attempting to fix CUSIP's.", vbOKOnly, _
            "Check Not Yet Done"
        Exit Sub
    End If
    
    'Loop through each file and scan using the procedure in Module
    'Data_Repair.
    For i = 1 To UBound(FinalUMAFiles)
        TempFile.LongName = FinalUMAFiles(i)
        If TempFile.FileCheck = True Then
            RepairData TempFile.LongName, Left(TempFile.Name, 2)
        Else
            WriteLogFile ("Did not fix: " & TempFile.Name)
        End If
    Next i
    
    
    WriteLogFile ("Fixing Data: End Processing")
    ReportProgress ("Done Fixing Data")
End Sub

Private Sub btnGo_Click()
    'The purpose of this procedure is to make copies of the
    'UMA holdings files and place one set in the Archives files
    'and another set in the MDT folder on the Momentum server.
    
    
    Dim i As Integer
    Dim strMsg As String
    Dim strDestinationPath As String
    Dim strArchivePath As String
    
    'Variables used to keep track of which files exist and which ones do not.
    'strNA contains the files that are "N/A" because there are no assets in
    'that strategy on the platform.
    Dim strFound As String
    Dim strNotFound As String
    Dim strNA As String
    
    
    WriteLogFile ("Moving Files: Begin Processing")
    ReportProgress ("Moving Files")
    
    'Do not go forward unless the check has been performed.
    If booCheckPerformed = False Then
        MsgBox "Please perform Check before attempting to move files.", vbOKOnly, _
            "Check Not Yet Done"
        Exit Sub
    End If

    
    'Look up in the Settings the path the UMA files are to be copied to.
    'Note that the files will be copied to two locations: an archive folder and the
    'folder on the Momentum server.
    'Store the paths in the variables strDestinationPath and strArchivePath.
    strDestinationPath = GetSetting("DIR", "UMAAPP", "MOM")
    strArchivePath = GetSetting("DIR", "UMAAPP", "ARC")
    
    'Loop through all the file names listed in the settings.  Make sure the
    'file exists.  If it does, copy it to the destination and archive folders.
    'Then erase the file in the original folder.  This will reduce the likelihood
    'of copying the same file more than once.
    'Keep lists of files found and copies as well as files not found in separate variables.
    For i = 1 To UBound(FinalUMAFiles)
        Dim UMAFile As New FileName
        UMAFile.LongName = FinalUMAFiles(i)
        If UMAFile.FileCheck = True Then
            strFound = strFound & UMAFile.Name & vbCrLf
            
            FileCopy UMAFile.LongName, strDestinationPath & UMAFile.Name
            WriteLogFile ("Wrote " & UMAFile.Name & " to: " & strDestinationPath)
            
            FileCopy UMAFile.LongName, strArchivePath & UMAFile.Name
            WriteLogFile ("Wrote " & UMAFile.Name & " to: " & strArchivePath)
            
            Kill UMAFile.LongName
        Else
            
            If MsgBox("Is " & UMAFile.Name & " missing because this strategy has no assets on the platform?", _
                vbYesNo, "Why is file missing?") = vbYes Then
                
                strNA = strNA & UMAFile.Name & vbCrLf
                WriteLogFile ("Did NOT write " & UMAFile.Name & " to: " & strArchivePath) & _
                ".  No assets."
            
            Else
            
                strNotFound = strNotFound & UMAFile.Name & vbCrLf
                WriteLogFile ("Did NOT write " & UMAFile.Name & " to: " & strArchivePath) & _
                    ".  Not found."
            End If
            
        End If
    Next i
    WriteLogFile ("Moving Files: End Processing")
    
    'If variables holding the lists are empty, fill the variable with an
    'appropriate message.
    If strFound = Empty Then
        strFound = "No Files Were Found."
    End If
    
    If strNotFound = Empty Then
        strNotFound = "All Files Were Found and Moved."
    End If
    
    If strNA = Empty Then
        strNA = "All strategies have assets.  Files were found and moved."
    End If

    'Send email notification to interested parties.
    SendEmail strFound, strNA, strNotFound
    
    WriteLogFile ("Notification emails sent.")
    ReportProgress ("Moving Files Completed")
End Sub

Private Sub Form_Load()
    'Procedure that runs when Main form is loaded.
    Dim i As Integer
    Dim myDate As New DateTool
    Set ErrorStatus = New ErrorNotice
    
    'Display version on Form
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor
    
    'Load default value into variable that indicates whether or not the Check has been done.
    booCheckPerformed = False
    
    'Load settings data into memory from settings file.
    LoadSettings (conSettingsFile)
    
    If ErrorStatus.ErrorExists = True Then
        MsgBox ErrorStatus.Description, vbOKOnly, conSettingsFile
        Exit Sub
    End If
    
    'Load the current date into the date object.
    myDate.TheDate = Date
    
    'Load user ID into user object
    Set myUser = New UserSettings
    myUser.NetworkID = GetNameOfUser
    
    'Establish instances of file handling objects.
    Set SettingsFile = New FileName
    Set CUSIPFile = New FileName


    'Display User name on Form
    lblUser.Caption = "User: " & myUser.LongName
    
    'Create error object to hold error status and set default value
    Set ErrorStatus = New ErrorNotice
    ErrorStatus.ErrorExists = False
    
    
    
    'Establish file name for the CUSIP file.  Make sure it exists. Size the CUSIP array.
    'Get the CUSIP9 file name.  See function definition in AppSettings.
    CUSIPFile.LongName = GetSetting("DIR", "SPOT", "TRADING") & GetSetting("FIL", "SPOT", "CUSIPS")
    If CUSIPFile.FileCheck = False Then
        MsgBox "File Not Found", vbOKOnly, CUSIPFile.LongName
        Exit Sub
    Else
        ReDim PrelimCUSIPS(1 To CUSIPFile.Lines, 1 To 3)
    End If
    
    'Establish file name for the Log file.
    'Get the log file name.  See function definition in AppSettings.
    LogFile = GetSetting("DIR", "UMAAPP", "OUTPUT") & GetSetting("FIL", "UMAAPP", "LOG")
    
    'Load the default date into the date box
    frmMain.tbDate.Text = myDate.PrevBusDay
    
    'Create an array of file names to be used during processing.
    CreateUMAFileNames CurrentSettings, FinalUMAFiles
    
    WriteLogFile ("Application loaded and ready for processing")
    
End Sub



Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuProcedure_Click()
    frmProcedure.Show
End Sub

Private Sub mnuSettings_Click()
    frmSettings.Show
End Sub

Private Sub tbDate_Change()
    'Clear data from existing file nazme array
    Erase UMAFiles
    
    'Create an array of file names to be used during processing.
    CreateUMAFileNames CurrentSettings, UMAFiles
End Sub
