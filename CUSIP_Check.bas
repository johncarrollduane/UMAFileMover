Attribute VB_Name = "CUSIP_Check"
Option Explicit

Sub CheckCUSIPs(ByVal strFile As String, ByVal strFirm As String, _
    ByVal strEntity As String)
    
    'The purpose of this procedure is to check each CUSIP and Ticker in the file for validity.
    'Because the file format differs depending on whether the file is from Morgan Stanley
    'or Smith Barney, the procedure accomodates each one.
    
    'Arguments:
    'strFile - the file name of the Morgan Stanley or Smith Barney File to be evaluated.
    'strFirm - a two character variable (either ms or sb) that indicates which firm produced
    'the file in strFile.
    'strEntity - a one character field indicating which money manager firm the file is for
    '(either m, f, or c).  The character is in position 10 of the file name.
    
    'Procedural Flow:
    '1.  The procedure reads each row of the file.
    '2.  If there are extraneous commas in the firm name they are deleted.  This is because
    'subsequent processing of this file, which is comma delimited, uses a comma counting technique
    'to determine which field is which.
    '3.  The procedure evaluates the ticker symbol in the row by reference to an array of Tickers
    'and CUSIPS loaded from a file downloaded from the APL system.
    '4  If it finds the ticker symbol in the array it continues on to compare the CUSIP in the file
    'to the one in the CUSIPS_T() array.  Because of this logic, only CUSIP's associated with valid
    'ticker symbols are evaluated for validity.  So, there is a possibility that a bad CUSIP will
    'pass through the procedure undetected if the ticker was not found.
    '5.  Exceptions are written to the log.  CUSIP problems detected are reported to the user
    'via a message box.
    
    'VARIABLES USED BY BOTH SECTIONS
    Dim intFile As Integer
    Dim strTicker As String
    Dim strCUSIP As String
    Dim strRowSkip As String
    Dim strRowUse As String
    Dim intFound As Integer
    Dim strMsg As String
    Dim Extracted(1 To 2) As Variant
    '1.  Ticker
    '2.  CUSIP
    
    ReportProgress ("Checking CUSIP's")

        
        'MORGAN STANLEY FILE LAYOUT
        'File is comma-delimited.
        '1st row are labels.  Data begins in 2nd row.
        '1.  Ticker
        '2.  Firm Name
        '3.  Strategy
        '4.  CUSIP
        '5.  Security Name
        '6.  Qty
        '7.  Price
        '8.  Total MV of position
    
        'SMITH BARNEY FILE LAYOUT
        'File is comma-delimited.
        '1st five rows are labels.  Data begins in 6th row.
        '1.  Ticker
        '2.  CUSIP
        '3. to 28 Not used
        
        intFile = FreeFile
        Open strFile For Input As intFile
        
            If strFirm = "ms" Then
                'Skip row 1
                Line Input #intFile, strRowSkip
            Else
                'Skip rows 1 to 5
                Line Input #intFile, strRowSkip
                Line Input #intFile, strRowSkip
                Line Input #intFile, strRowSkip
                Line Input #intFile, strRowSkip
                Line Input #intFile, strRowSkip
            End If
            
        
            'Read the rest of the rows one by one.
            Do Until EOF(intFile)
                Line Input #intFile, strRowUse
                
                'The Morgan Stanley file contains the name of the money management
                'firm in field 2.  In some instances the firm name will contain a comma
                'which would disrupt subsequent processing of the comma delimited file.
                'ReplaceFirm effectively removes the comma from the firm name.
                'This is not a problem in the Smith Barney file.
                If strEntity = "m" Then
                    strRowUse = ReplaceFirm(strRowUse, 2, "MDT ADVISERS INC.")
                End If
                
                'This function takes a row from the file, parses the fields, and extracts
                'the CUSIP and ticker symbol.  It returns a two element array containing
                'those two data items.
                If strFirm = "ms" Then
                    
                    'Morgan Stanley
                    GetCUSIPandTicker strRowUse, 4, 1, Extracted
                    
                Else
                    
                    'Smith Barney
                    GetCUSIPandTicker strRowUse, 2, 1, Extracted
                    
                End If
                
                
                'Check ticker in Extracted(1) using binary search function.
                'If function returns -1, the item could not be found.
                'Look for the Ticker in the CUSIPS_T() array.
                'For items not found, write the details to the log file.
                intFound = Searcher(CUSIPS_T, Extracted(1), 1)
                If intFound = -1 Then
            
                    'Ticker not found in CUSIP array
                    With ErrorStatus
                        .TickerProblemExists = True
                        .TickerProblemNumber = .TickerProblemNumber + 1
                    End With
                    
                    strMsg = Extracted(1) & " " & "Unknown Ticker " & "in " & strFile
                    WriteLogFile (strMsg)
                
                Else
                    If Extracted(2) <> CUSIPS_T(intFound, 2) Then
                
                        'Ticker found, but CUSIP's do not agree.
                        'Count the number of CUSIP problems.
                        With ErrorStatus
                            .CusipProblemExists = True
                            .CusipProblemNumber = .CusipProblemNumber + 1
                        End With
                        
                        'Create a string message to report the problem in the log.
                        strMsg = Extracted(1) & " " & "CUSIP does not agree " & "in " & _
                            strFile & ": " & CUSIPS_T(intFound, 2) & " vs. " & Extracted(2)
                        WriteLogFile (strMsg)
                    
                    End If
                End If
            Loop
        Close intFile
    
End Sub

Sub GetCUSIPandTicker(ByVal strRow As String, ByVal intPosCUSIP As Integer, _
    ByVal intPosTicker As Integer, ByRef Extracted As Variant)
    
    'The purpose of this procedure is to locate the CUSIP and Ticker symbol in
    'the row of data read from the UMA holdings file.  Depending on whether the
    'Smith Barney or Morgan Stanley files are being read, the CUSIP and Ticker are
    'in different locations.  Both files are comma delimited.
    'Arguments:
    'strRow - variable that contains the entire row of data.
    'intPosCUSIP - indicates the position of the CUSIP in the row.  For example, 2
    'would indicate the field is in second position, that is, it follows the first
    'comma delimiter.
    'intPosTicker - indicates the position of the Ticker in the row.
    'Extracted - a single dimension array holding two values (Ticker and CUSIP).
    
    'This procedure returns the Extracted array to the calling procedure.
    
    Dim i As Integer
    Dim intCommaCounter As Integer
    Dim strChar As String
    
    
    intCommaCounter = 0
    i = 0
    
    'Ticker
    Extracted(1) = Empty
    'CUSIP
    Extracted(2) = Empty
    
    'This part of the procedure loops through strRow character by character.
    'Count the number of commas.  The comma count will indicate what field
    'in the row is currently being read.  Note that the comma count will be
    'one less than the field position variables intPosTicker and intPosCUSIP.
    'Once the required number of commas has been passed, begin building the
    'CUSIP or Ticker one character at a time.  Stop when the comma counter
    'equals the larger of the positions for CUSIP or Ticker.
    Do Until intCommaCounter = TheLarger(intPosCUSIP, intPosTicker)
        i = i + 1
        
        If i > Len(strRow) Then Exit Do
        
        strChar = Mid(strRow, i, 1)
        If strChar = "," Then
            intCommaCounter = intCommaCounter + 1
        Else
            If intCommaCounter = intPosTicker - 1 Then
                'Ticker
                Extracted(1) = Extracted(1) & strChar
            ElseIf intCommaCounter = intPosCUSIP - 1 Then
                'CUSIP
                Extracted(2) = Extracted(2) & strChar
            End If
        End If

    Loop
    
    
End Sub

Function IsZero(ByVal strRow As String, ByVal intPosQty As Integer) As Boolean
    
    'The purpose of this function is to return a True/False indicator of whether
    'the row of data read from the UMA holdings file contains a zero quantity for
    'the position.  This is important because the final file sent to Federated should
    'not have any zero positions.
    
    'Smith Barney or Morgan Stanley files are being read. The number of shares is
    'in different locations.  Both files are comma delimited.
    'Arguments:
    'strRow - variable that contains the entire row of data.
    'intPosQty - indicates the position of the quantity in the row.  For example, 2
    'would indicate the field is in second position, that is, it follows the first
    'comma delimiter.
    
    'This function returns a FALSE to the calling procedure if the row has a quantity
    'greater than zero.  Note these are long-only portfolios.
    
    Dim i As Integer
    Dim intCommaCounter As Integer
    Dim strChar As String
    Dim strQty As String
    
    'Default values
    intCommaCounter = 0
    i = 0
    IsZero = True
    strQty = Empty
    
    'This part of the procedure loops through strRow character by character.
    'Count the number of commas.  The comma count will indicate what field
    'in the row is currently being read.  Note that the comma count will be
    'one less than the field position variable intPosQty .
    'Once the required number of commas has been passed, begin building the
    'QTY one character at a time.  Stop when the comma counter
    'equals intPosQty.
    Do Until intCommaCounter = intPosQty
        i = i + 1
        
        If i > Len(strRow) Then Exit Do
        
        strChar = Mid(strRow, i, 1)
        If strChar = "," Then
            intCommaCounter = intCommaCounter + 1
        Else
            If intCommaCounter = intPosQty - 1 Then
                'QTY
                strQty = strQty & strChar
            End If
        End If

    Loop
    
    'In the Smith Barney file a zero quantity is indicated with a hyphen
    'in the QTY field.  Morgan Stanley denotes a zero quantity with a zero.
    'Test for these values.  If present, set IsZero to True.  Also set to
    'zero any QTY that is a negative number or is not a number at all.
    If IsNumeric(strQty) = False Then
        IsZero = True
    ElseIf CDbl(strQty) <= 0 Then
        IsZero = True
    Else
        IsZero = False
    End If
    
End Function


Function TheLarger(ByVal varItem1 As Variant, ByVal varItem2 As Variant) _
    As Variant
    
    If varItem1 > varItem2 Then
        TheLarger = varItem1
    Else
        TheLarger = varItem2
    End If
End Function
