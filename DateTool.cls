VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The purpose of this class is to assemble in one place the various date
'functions that are used widely in MDT applications that are not easily
'called using VB functions.

Private pHolidays() As Date
Public TheDate As Date



Private Sub Class_Initialize()
    
    'This procedure runs when the class is initialized.
    'It loads an array with holiday dates from a test file
    'called Holidays.txt.  This data is used in the IsHoliday
    'property below.
    
    Dim strFileName As String
    Dim strData As String
    Dim i As Integer
    Dim intFile As Integer
    
    'Get holidays file name.  See function definition in AppSettings.
    strFileName = GetSetting("DIR", "SYSTEMS", "GENERAL") & GetSetting("FIL", "SYSTEMS", "HOLIDAYS")
    'strFileName = "\\mdtafile\MDT_Share\MDT-OPS\Boston Ops\Admin\Systems - General\Holidays.txt"
    
    'First count how many rows are in the Holiday file.
    'Then resize the Holidays() array to fit the number
    'of rows in the file.
    intFile = FreeFile
    i = 0
    Open strFileName For Input As intFile
        Do Until EOF(intFile)
            Line Input #intFile, strData
            i = i + 1
        Loop
    Close intFile
    ReDim pHolidays(1 To i)
    
    'Load the contents of Holidays.txt into Holidays().
    intFile = FreeFile
    i = 0
    Open strFileName For Input As intFile
        Do Until EOF(intFile)
            Line Input #intFile, strData
            i = i + 1
            pHolidays(i) = CDate(strData)
        Loop
    Close intFile
    
    
End Sub

Private Property Get pIsHoliday(ByVal dtDate As Date) As Boolean
    'This procedure checks to see if a given date is a holiday
    'by checking the data in Holidays().
    'This property is private.  It is used by the public property
    'IsHoliday and also by the Next BusDay and PrevBusDay
    'properties.
    'It returns True if it is a holiday.
    
    Dim i As Integer
    
    If IsDate(dtDate) = True Then
        pIsHoliday = False
        For i = 1 To UBound(pHolidays)
            If dtDate = CDate(pHolidays(i)) Then
                pIsHoliday = True
                Exit For
            End If
        Next i
    Else
        MsgBox "Have not assigned a valid date to this object.", vbOKOnly, _
            CStr(TheDate)
        Exit Sub
    End If
End Property


Public Property Get IsHoliday() As Boolean
    'Checks to see if the public property TheDate is a holiday.
    'It returns True if it is a holiday.
    
    If IsDate(TheDate) = True Then
        If pIsHoliday(TheDate) Then
            IsHoliday = True
        Else
            IsHoliday = False
        End If
    Else
        MsgBox "Have not assigned a valid date to this object.", vbOKOnly, _
            CStr(TheDate)
        Exit Sub
    End If
End Property

Public Property Get NextBusDay() As Date
    'Returns the next business day.  This is defined as the next day following
    'TheDate as long as that date is not a holiday, is not a Saturday,
    'and is not a Sunday.  If it is, it checks the next date, and so on.
    
    If IsDate(TheDate) = True Then
        NextBusDay = TheDate + 1
        Do Until pIsHoliday(NextBusDay) = False And _
            DatePart("w", NextBusDay) <> 7 And _
            DatePart("w", NextBusDay) <> 1
        
            NextBusDay = NextBusDay + 1
        Loop
    Else
        MsgBox "Have not assigned a valid date to this object.", vbOKOnly, _
            CStr(TheDate)
        Exit Sub
    End If
End Property

Public Property Get PrevBusDay() As Date
    'Returns the previous business day.  This is defined as the day immediately
    'before TheDate as long as that date is not a holiday, is not a Saturday,
    'and is not a Sunday.  If it is, it checks the date one day older, and so on.
    
    If IsDate(TheDate) = True Then
        PrevBusDay = TheDate - 1
        Do Until pIsHoliday(PrevBusDay) = False And _
            DatePart("w", PrevBusDay) <> 7 And _
            DatePart("w", PrevBusDay) <> 1
        
            PrevBusDay = PrevBusDay - 1
        Loop
    Else
        MsgBox "Have not assigned a valid date to this object.", vbOKOnly, _
            CStr(TheDate)
        Exit Sub
    End If
End Property

Public Property Get DayName() As String
    'Returns the spelled out name of the day of the week.
    
    If IsDate(TheDate) = True Then
    
        Select Case DatePart("w", TheDate, vbSunday, vbFirstJan1)
            Case 1
                DayName = "Sunday"
            Case 2
                DayName = "Monday"
            Case 3
                DayName = "Tuesday"
            Case 4
                DayName = "Wednesday"
            Case 5
                DayName = "Thursday"
            Case 6
                DayName = "Friday"
            Case 7
                DayName = "Saturday"
        End Select
    
    Else
    
        MsgBox "Have not assigned a valid date to this object.", vbOKOnly, _
            CStr(TheDate)
        Exit Sub
    End If
End Property


