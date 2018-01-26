Attribute VB_Name = "APL_Files"
Option Explicit


Sub ConvertAPLFile(ByVal strFile As String)
    'The purpose of this procedure is to transform the APL format
    'holdings file into a more readable format.
    
    Dim intFile As Integer
    Dim intFile2 As Integer
    Dim lngFound As Long
    Dim strRow As String
    Dim strTempFile As String
    Dim strTempTicker As String
    Dim strTempQty As String
    Dim strTempMV As String
    Dim strFiller As String
    Dim TempFile As New FileName
    Dim Sorter As New SortOp
    
    TempFile.LongName = strFile
    
    'Make a copy of the original file.  Replace the first
    'two characters of the name with "te".
    strTempFile = TempFile.Path & "te" & Right(TempFile.Name, 24)
    FileCopy TempFile.LongName, strTempFile
    
    intFile = FreeFile
    intFile2 = intFile + 1
    
    Open strTempFile For Input As intFile
    Open TempFile.LongName For Output As intFile2
        
        'Create five rows that will not be read by the application
        'in a later procedure.  This will make the output file look
        'like the Smith Barney files.  Note that on row 2 there is a
        'strategy identifier.
        Print #intFile2, "Analyze by :,,,,,,,,,,,,,,,,,,,,,,,,,,,,Holding Analytics"
        Print #intFile2, "Model id:,,,,,,,,,,,,,,,,,,,,,,,,,,,," & Mid(TempFile.Name, 8, 6)
        Print #intFile2, "Security,Cusip,Sedol,Target%,Underweights(% from Target),,,,,W/in,Overweights(% from Target),,,,,Statistics,,,,,,,,Summary,,,,,"
        Print #intFile2, ",,,,<-5,<-4,<-3,<-2,<-1,1,>1,>2,>3,>4,>5,Min,Max,Avg,Std Dev,Cur Prc,Hi Cst,Lw Cst,Avg Cst,Hldrs,Miss Sec,Total,Holdings#,Mkt Val,Hldrs (%)"
        Print #intFile2, "$,$$,$$,2.04,7,16,36,57,68,-,-,-,-,-,-,0.03,2.41,0.62,0.56,-,-,-,-,68,0,68,N/A,$231,618.01 ,100"
        
        'This variable will replace all the data that is on the typical
        'Smith Barney file that is not used.
        strFiller = "-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-,-"
        
        Do Until EOF(intFile)
            Line Input #intFile, strRow
            
            'Because the input file is an electronic copy of a report which lists the
            'holdings twice, the second time the holdings is read, it should be ignored
            'for future processing.  Use the existence of an X in column 80 to indicate
            'the first instance of this holding.
            If Mid(strRow, 80, 1) = "X" Then
            
                'Only use data from rows that have COMMON or FOREIGN in the second
                'column. Place data in strRow variable.
                If Mid(strRow, 24, 6) = "COMMON" Or _
                    Mid(strRow, 24, 7) = "FOREIGN" Then
                
                    'Look up the ticker symbol in the CUSIPS() array using the
                    'CUSIP from the input file.  Then replace the long name
                    'of the issuer with the ticker symbol.
                
                    'Fields in CUSIP array are:
                    '1.  Ticker
                    '2.  CUSIP
                    '3.  Issue Type
                    lngFound = Searcher(CUSIPS_C, Mid(strRow, 33, 9), 2)
                    If lngFound <> -1 Then
                        strTempTicker = CUSIPS_C(lngFound, 1)
                    Else
                        strTempTicker = Left(strRow, 5)
                    End If
                
                    'Write the useful data, with the ticker symbol, to the output file.
                    'Note that the market value field is surrounded with quotes and
                    'that commas have not been removed.
                    Print #intFile2, strTempTicker & "," & _
                        Mid(strRow, 33, 9) & "," & _
                        strFiller & "," & _
                        RemoveCommas(Trim(Mid(strRow, 59, 12))) & "," & Chr(34) & _
                        Trim(Mid(strRow, 44, 14)) & ".00" & Chr(34) & ",-"
                
                End If
            End If
        Loop
    Close intFile2
    Close intFile
    
End Sub

Function RemoveCommas(ByVal strOldValue) As String
    
    'Rebuild the value without commas or quotes.
    
    Dim i As Integer
    Dim strChar As String

    For i = 1 To Len(strOldValue)
        strChar = Mid(strOldValue, i, 1)
        If strChar <> "," And Asc(strChar) <> 34 Then
            RemoveCommas = RemoveCommas & strChar
        End If
    Next i


End Function
