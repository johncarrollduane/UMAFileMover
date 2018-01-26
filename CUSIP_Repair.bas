Attribute VB_Name = "Data_Repair"
Option Explicit

Sub RepairData(ByVal strFile As String, ByVal strFirm As String)
    'The purpose of this procedure is to "repair" the data in the input file.  Repairing
    'includes removing extraneous commas that get in the way of efficient parsing, removing
    'positions with zero quantities, and replacing CUSIP's that do not match their ticker
    'symbols.

    Dim intFile1 As Integer
    Dim intFile2 As Integer
    Dim InputFile As New FileName
    Dim OutputFile As New FileName
    
    
    ReportProgress ("Repairing Data")
    
    'Create file objects.  Assign name to temporary file.
    'The temporary file is used to hold the "post-processing version
    'of the data file.  Once processing is complete, the temporary
    'file overwrites the input file.
    InputFile.LongName = strFile
    OutputFile.LongName = GetSetting("DIR", "UMAAPP", "RESOURCES") & _
        GetSetting("FIL", "UMAAPP", "INPROCESS")
    
    'Remove unwanted commas from input file.
    CleanUpCommas InputFile.LongName, OutputFile.LongName, strFirm

    
    'This process cleans up the CUSIP problems and eliminates rows with
    'zero positions.  This is for both firms.
    CleanUpCUSIPs InputFile.LongName, OutputFile.LongName, strFirm


End Sub



Sub CleanUpCommas(ByVal InputFile As String, ByVal OutputFile As String, _
    ByVal strFirm As String)
    
    'The purpose of this procedure is to "clean up" the input file
    'by removing extraneous commas that will make parsing the
    'comma delimited rows easier.
    
    'Arguments:
    'InputFile - the file produced by the wirehouse
    'OutputFile - the cleaned up file
    'intRowSkip Count - the number of rows at the top of the
    'file that do not include usable data (column headings)
    
    'Processing:
    '1.  Proceed down the input file row by row.  For each row read
    '    write a row to the output file.
    '2.  The rows that are skipped are simply re-written to the
    '    output file.
    '3.  The remaining rows are fed one by one into the ReplaceFirm and
    '    ReplaceCompany functions.  The functions return a row of data
    '    that does not have a comma in the Firm name or in the Company
    '    name.
    '4.  Finally, copy the output file to the input file's memory location.
    '    Thus the input file will now contain clean data to be used by
    '    subsequent processes.
    
    Dim intFile1 As Integer
    Dim intFile2 As Integer
    Dim strRowSkip As String
    Dim strRowUse As String
    Dim intRowSkipCount As Integer
    
    intFile1 = 1
    intFile2 = 2
    
    'Establish value of variable holding number of rows to skip.
    If strFirm = "ms" Then
        intRowSkipCount = 1
    ElseIf strFirm = "sb" Then
        intRowSkipCount = 5
    Else
        intRowSkipCount = 0
    End If
        
    
    Open InputFile For Input As intFile1
    Open OutputFile For Output As intFile2
        
        'Skip rows without data
        
        Do While intRowSkipCount > 0
            Line Input #intFile1, strRowSkip
            Print #intFile2, strRowSkip
            intRowSkipCount = intRowSkipCount - 1
        Loop
        
        'Read the rest of the lines one by one.  Send the lines through
        'functions that remove unwanted commas from the Firm name and the
        'Company name and the Quantity.  Then  re-write the line to the output file.
        'Then write a new file for subsequent processing.
        
        If strFirm = "ms" Then
            Do Until EOF(intFile1)
                Line Input #intFile1, strRowUse
                    
                strRowUse = ReplaceFirm(strRowUse, 2, "MDT ADVISERS INC.")
                strRowUse = ReplaceCompany(strRowUse, 5, 7)

                Print #intFile2, strRowUse

            Loop
            
        ElseIf strFirm = "sb" Then
            Do Until EOF(intFile1)
                Line Input #intFile1, strRowUse

                strRowUse = ReplaceQty(strRowUse, 27)

                Print #intFile2, strRowUse

            Loop
        End If
        
    Close intFile1
    
    
    
    
    
    
    Close intFile2
        
    'Overwrite input file with output file.  Input file is now clean.
    FileCopy OutputFile, InputFile
    
End Sub


Sub CleanUpCUSIPs(ByVal InputFile As String, ByVal OutputFile As String, ByVal strFirm As String)
    
    'The purpose of this procedure is to "clean up" the input file's
    'CUSIP's.  CUSIP's are checked for length (they should have be
    '(9 characters long) and for correspondence to their ticker symbol.
    'The procedure will reference the array of CUSIP data when
    'evaluating the quality of the CUSIP's in the file.
    
    'Arguments:
    'InputFile - the file produced by the wirehouse (after comma clean up)
    'OutputFile - the file after CUSIP clean up
    'intRowSkip Count - the number of rows at the top of the
    'file that do not include usable data (column headings)
    
    'Processing:
    '1.  Proceed down the input file row by row.  For each row read
    '    write a row to the output file, unless the row has a zero position.
    '2.  The rows that are skipped are simply re-written to the
    '    output file.
    '3.  The remaining rows are evaluated one by one for length.  If the CUSIP
    '    is 10 characters long, truncate the first character.
    '4.  The reamining rows are evaluated one by one for a legitimate ticker
    '    symbol by checking the CUSIPs array.
    '5.  For those rows with a valid ticker symbol, check to see that the CUSIP
    '    and the ticker symbol correspond by reference to the CUSIPs array.
    '    If the CUSIP does not match the ticker, replace the CUSIP in the file
    '    with a good CUSIP.
    '6.  Finally, copy the output file to the input file's memory location.
    '    Thus the input file will now contain clean data to be used by
    '    subsequent processes.
    
    Dim intFile1 As Integer
    Dim intFile2 As Integer
    Dim intFound As Integer
    Dim intRowSkipCount As Integer
    Dim intPosCUSIP As Integer
    Dim intPosTicker As Integer
    Dim intPosQty As Integer
    Dim strRowSkip As String
    Dim strRowUse As String
    Dim strOldCUSIP As String
    Dim strMsg As String
    Dim Extracted(1 To 2) As Variant
    '1.  Ticker
    '2.  CUSIP
    
    intFile1 = 1
    intFile2 = 2
    
    'MORGAN STANLEY FILE LAYOUT
        'File is comma-delimited.
        '1st row are labels.  Data begins in 2nd row.
        '1.  Ticker
        '2.  Not used
        '3.  Not used
        '4.  CUSIP
        '5.  Security Name
        '6.  Qty
        '7.  Not used
        '8.  Not used
    
    'SMITH BARNEY FILE LAYOUT
        'File is comma-delimited.
        '1st five rows are labels.  Data begins in 6th row.
        '1.  Ticker
        '2.  CUSIP
        '3. to 28 Not used (except 10 is Qty)
        
        
    'Establish value of firm specific variables.
    If strFirm = "ms" Then
        intRowSkipCount = 1
        intPosCUSIP = 4
        intPosTicker = 1
        intPosQty = 6
    ElseIf strFirm = "sb" Then
        intRowSkipCount = 5
        intPosCUSIP = 2
        intPosTicker = 1
        intPosQty = 27
    Else
        MsgBox ("Error in Clean Up CUSIP's")
        Exit Sub
    End If
    
    Open InputFile For Input As intFile1
    Open OutputFile For Output As intFile2
        
        'Skip rows without data
        Do While intRowSkipCount > 0
            Line Input #intFile1, strRowSkip
            Print #intFile2, strRowSkip
            intRowSkipCount = intRowSkipCount - 1
        Loop
        
        'Read the rest of the lines one by one.
        Do Until EOF(intFile1)
            Line Input #intFile1, strRowUse
                If IsZero(strRowUse, intPosQty) = False Then
                    GetCUSIPandTicker strRowUse, intPosCUSIP, intPosTicker, Extracted
                    
                    'Store old CUSIP in variable
                    strOldCUSIP = Extracted(2)
                    If Len(strOldCUSIP) = 0 Then strOldCUSIP = "BLANK"
                    
                    'Check if CUSIP is equal to 10 characters.  If it is,
                    'truncate the first character.
                    If Len(strOldCUSIP) = 10 Then
                        strRowUse = ReplaceCUSIP(strRowUse, intPosCUSIP, Right(strOldCUSIP, 9))
                        strMsg = Extracted(1) & " " & "CUSIP replaced " & "in " & InputFile & _
                            ": " & Right(strOldCUSIP, 9) & " for " & strOldCUSIP
                        WriteLogFile (strMsg)
                    End If
                    
                    'Check ticker in Extracted(1) using binary search function.
                    'If function returns -1, the item could not be found in the
                    'CUSIPS_T() array.  Field 1 in the array is ticker and field 2 is CUSIP.
                    'For items not found, write the details to the log file.
                    intFound = Searcher(CUSIPS_T, Extracted(1), 1)
                    If intFound = -1 Then
            
                        'Ticker not found in CUSIP array.  Write an entry to the log noting
                        'that we have an unknown ticker symbol.  Write the row to the
                        'output file anyway given that CUSIP is the field that Federated
                        'uses as its security identifier.
                        strMsg = Extracted(1) & " " & "Unknown Ticker " & "in " & InputFile
                        WriteLogFile (strMsg)
                        Print #intFile2, strRowUse
                
                    Else
                        If Extracted(2) <> CUSIPS_T(intFound, 2) Then
                
                            'Ticker found, but CUSIP's do not agree.
                            strRowUse = ReplaceCUSIP(strRowUse, intPosCUSIP, CUSIPS_T(intFound, 2))
                            Print #intFile2, strRowUse
                            strMsg = Extracted(1) & " " & "CUSIP replaced " & "in " & InputFile & _
                                ": " & CUSIPS_T(intFound, 2) & " for " & strOldCUSIP
                            WriteLogFile (strMsg)
                    
                        Else
                    
                            Print #intFile2, strRowUse
                    
                        End If
                    End If
                Else
                    WriteLogFile ("Zero Qty Row: " & InputFile & ", " & Left(strRowUse, 20))
                End If
            Loop
        Close intFile1
        Close intFile2
        
        FileCopy OutputFile, InputFile
End Sub


Function ReplaceCUSIP(ByRef strRow As String, ByVal intPosCUSIP As Integer, _
    ByVal strNewCUSIP As String) As String
    
    'The general idea behind this function is that it iterates
    'through the row character by character building a new row
    'by simply appending each character from the original row
    'to the new one.  While iterating through the row, it counts
    'the number of commas it has encountered.  The comma count
    'is how the function knows where the CUSIP is or where it should
    'be.
    
    'There are two situations where this function is called.  The
    'first situation is where a CUSIP is missing.  The other is where
    'the CUSIP is wrong.  The function will either insert the CUSIP
    'where there is none and continue building the row character by
    'character.  Or it will insert the correct CUSIP after the appropriate
    'comma and ignore the characters making up the bad CUSIP while
    'continuing to build the row.

    Dim i As Integer
    Dim intCommaCounter As Integer
    Dim strTicker As String
    Dim strCUSIP As String
    Dim strChar As String
    Dim booCUSIPReplaced As Boolean
    
    'Beginning values
    intCommaCounter = 0
    i = 0
    booCUSIPReplaced = False
    ReplaceCUSIP = Empty
    
    
    For i = 1 To Len(strRow)
        
        'Loop through the row character by character.
        'Each time a comma is encountered, increment the
        'intCommaCounter variable by 1.
        
        strChar = Mid(strRow, i, 1)
        If strChar = "," Then
            intCommaCounter = intCommaCounter + 1
            
            'Check to see if the CUSIP has already been replaced.
            If booCUSIPReplaced = False Then
                
                'It has not been replaced yet.
                'Check to see if this comma represents the beginning
                'of the CUSIP field.  If so append the good CUSIP to the
                'new row and toggle the ReplaceCUSIP variable to TRUE.
                If intCommaCounter = intPosCUSIP - 1 Then
                    ReplaceCUSIP = ReplaceCUSIP & strChar & strNewCUSIP
                    booCUSIPReplaced = True
                    
                'Because we have not yet reached the CUSIP field,
                'continue building the row by appending this character.
                Else
                    ReplaceCUSIP = ReplaceCUSIP & strChar
                End If
                
            Else
                'The CUSIP has been replaced AND this character is a comma.
                'Simply append it to the new row.
                ReplaceCUSIP = ReplaceCUSIP & strChar
                
            End If
        Else
            
            'This character is not a comma.
            'Check to see if we are before or have already passed the CUSIP
            'field.  If so, we are safe in appending whatever the character
            'is.  If we are still in the CUSIP field, we DO NOT append,
            'since this character is part of the bad CUSIP.
            
            If intCommaCounter > intPosCUSIP - 1 Then
                
                'We are past the CUSIP field so append.
                ReplaceCUSIP = ReplaceCUSIP & strChar
                
            ElseIf intCommaCounter < intPosCUSIP - 1 Then
                
                'We are before the CUSIP field so append.
                ReplaceCUSIP = ReplaceCUSIP & strChar

            End If
        End If
            
            
    Next i
    
    
End Function



Function ReplaceFirm(ByRef strRow As String, ByVal intPosFirm As Integer, _
    ByVal strNewFirm As String) As String
    
    'The purpose of this function is to remove any commas from the Firm name
    'such as "MDT Advisers, Inc."  This is necessary to make parsing the comma-
    'delimited row easier.
    
    'Arguments:
    'strRow - This is the row of data from the file before any changes are
    '         made to it.
    'intPosFirm - This is the field position of the Firm name in the row.
    'strNewFirmName - The Firm name to be used in place of the one with a comma.
    
    'Processing:
    '1.  Evaluate each row to see if a replacement has to be made.  This is done
    '    by searching for an 18 character string that matches "MDT ADVISERS, INC." _
    '    or "MDT ADVISORS, INC."
    '2.  If a replacement has to be made, rewrite the row one character at a time
    '    using the new Firm name.  Use the intPosFirm variable to locate the
    '    beginning of the Firm field in the row.
    '3.  Once the row has been rewritten, return the new value to the caller.

    Dim i As Integer
    Dim j As Integer
    Dim strToken As String
    Dim intCommaCounter As Integer
    Dim strTicker As String
    Dim strCUSIP As String
    Dim strChar As String
    Dim booFirmNeedsReplacing As Boolean
    Dim booFirmReplaced As Boolean
    
    'Beginning values
    intCommaCounter = 0
    i = 0
    booFirmNeedsReplacing = False
    booFirmReplaced = False
    ReplaceFirm = Empty
    
    'Check to see if the firm name has a comma in it.
    'Scan chunks of the row 18 characters at a time looking
    'for the firm name and then the comma.
    'If the firm has already been replaced
    For i = 1 To Len(strRow)
        strToken = Mid(strRow, i, 18)
        If strToken = "MDT ADVISERS, INC." Or strToken = "MDT ADVISORS, INC." Then
            booFirmNeedsReplacing = True
            Exit For
        End If
    Next i
    
    'If the firm name has a comma in it, replace it with
    'a firm name that does not have a comma in it.
    'Otherwise, return the original value of the row, unchanged.
    If booFirmNeedsReplacing = True Then
        For i = 1 To Len(strRow)
            strChar = Mid(strRow, i, 1)

            If strChar = "," Then
                intCommaCounter = intCommaCounter + 1
            End If
        
        
            If booFirmReplaced = True Then
                If intCommaCounter >= intPosFirm + 1 Then
                    ReplaceFirm = ReplaceFirm & strChar
                End If
            Else
                If intCommaCounter < intPosFirm - 1 Then
                    ReplaceFirm = ReplaceFirm & strChar
                ElseIf intCommaCounter = intPosFirm - 1 Then
                    ReplaceFirm = ReplaceFirm & strChar & strNewFirm
                    booFirmReplaced = True
                End If
            End If
        Next i
    Else
        ReplaceFirm = strRow
    End If
    
    
End Function


Function ReplaceCompany(ByRef strRow As String, ByVal intPosCo As Integer, _
    ByVal intCommaCountNormal As Integer) As String
    'The purpose of this function is to check the company name to see if it has
    'a comma in it (such as IBM, Inc.).  If so, remove the comma and return a new row.
    
    'Arguments:
    'strRow - This is the row of data from the file before any changes are
    '         made to it.
    'intPosCo - This is the field position of the company name in the row.
    'intCommaCountNormal - Integer representing the number of commas in a row where
    '                      there are no extra commas.  In other words, this is the
    '                      number of commas in a row where the commas are just field
    '                      delimiters.
    
    'Processing:
    '1.  Evaluate each row to see if a company name needs to have a comma removed.
    '    This is done by counting the number of commas in each row and comparing
    '    that total to the "normal" number of commas in a row.  It is safe to assume
    '    that any extra commas will be in the company name because any other field
    '    that could hold an extra comma has already been cleaned up.
    '2.  If a company name has to be changed, rebuild the company name without the
    '    comma.  Then, rewrite the row one character at a time
    '    inserting the new company name.  Use the intPosCo variable to locate the
    '    beginning of the company field in the row.
    '3.  Once the row has been rewritten, return the new value to the caller.
    
    Dim i As Integer
    Dim j As Integer
    Dim strToken As String
    Dim intCommaCounter As Integer
    Dim strOldName As String
    Dim strNewName As String
    Dim strChar As String
    Dim booCoNeedsReplacing As Boolean
    Dim booCoReplaced As Boolean
    
    'Beginning values
    intCommaCounter = 0
    i = 0
    booCoNeedsReplacing = False
    booCoReplaced = False
    ReplaceCompany = Empty
    strOldName = Empty
    strNewName = Empty
    
    'Check to see if the company name has a comma in it.
    'Count the number of commas in the entire row.  The
    'number of commas will be a fixed number unless there
    'is a comma in the company name, then there will be one more.
    
    'Count the commas.
    For i = 1 To Len(strRow)
        strChar = Mid(strRow, i, 1)
        If strChar = "," Then
            intCommaCounter = intCommaCounter + 1
        End If
    Next i
    
    'If there are too many commas, set the boolean indicator to true.
    'If not, simply exit the function after returning the original
    'value of the row.
    If intCommaCounter > intCommaCountNormal Then
        booCoNeedsReplacing = True
    Else
        ReplaceCompany = strRow
        Exit Function
    End If
    
    'If the company name has a comma in it, replace it with
    'a company name that does not have a comma in it.
    'Otherwise, return the original value of the row, unchanged.
    
    If booCoNeedsReplacing = True Then
        
        'Extract the company name from the row, including the unwanted comma.
        'The name will go into the variable called strOldName.
        intCommaCounter = 0
        For i = 1 To Len(strRow)
            strChar = Mid(strRow, i, 1)

            If strChar = "," Then
                intCommaCounter = intCommaCounter + 1
            End If
        
            If intCommaCounter = intPosCo - 1 Or intCommaCounter = intPosCo Then
                strOldName = strOldName & strChar
            End If

        Next i
        
        'Rebuild the company name without the comma.
        For i = 1 To Len(strOldName)
            strChar = Mid(strOldName, i, 1)
            If strChar <> "," Then
                strNewName = strNewName & strChar
            Else
                strNewName = strNewName & " "
            End If
        Next i
        
        'Replace the old name with the new name
        intCommaCounter = 0
        For i = 1 To Len(strRow)
            strChar = Mid(strRow, i, 1)

            If strChar = "," Then
                intCommaCounter = intCommaCounter + 1
            End If
        
        
            If booCoReplaced = True Then
                If intCommaCounter >= intPosCo + 1 Then
                    ReplaceCompany = ReplaceCompany & strChar
                End If
            Else
                If intCommaCounter < intPosCo - 1 Then
                    ReplaceCompany = ReplaceCompany & strChar
                ElseIf intCommaCounter = intPosCo - 1 Then
                    ReplaceCompany = ReplaceCompany & strChar & strNewName
                    booCoReplaced = True
                End If
            End If
        Next i
        
    End If
    
    
End Function

Function ReplaceQty(ByRef strRow As String, ByVal intPosQty As Integer) As String
    'The purpose of this function is to check the quantity field to see if it has
    'a comma in it (such as 2,345).  If so, remove the comma and return a new row.
    
    'Arguments:
    'strRow - This is the row of data from the file before any changes are
    '         made to it.
    'intPosQty - This is the field position of the quantity in the row.
    
    'Processing:
    '1.  Evaluate each row to see if the quantity needs to have a comma removed.
    '    This is done by iterating through the characters in the row until reaching
    '    the quantity field.  If the quantity has a comma in it, the number will be
    '    surrounded by quotes.
    '2.  If a quantity has to be changed, rebuild the quantity without the quotes and
    '    comma.  Then, rewrite the row one character at a time
    '    inserting the new quantity.  Use the intQtyCo variable to locate the
    '    beginning of the quantity in the row.
    '3.  Once the row has been rewritten, return the new value to the caller.
    
    Dim i As Integer
    Dim j As Integer
    Dim strToken As String
    Dim intCommaCounter As Integer
    Dim strOldQty As String
    Dim strNewQty As String
    Dim strChar As String
    Dim booQtyNeedsReplacing As Boolean
    Dim booQtyReplaced As Boolean
    
    'Beginning values
    i = 0
    booQtyNeedsReplacing = False
    booQtyReplaced = False
    ReplaceQty = Empty
    strOldQty = Empty
    strNewQty = Empty
    
    'Check to see if the quantity field begins with a quote.
    'If so, the field will contain a comma.
    
    'Move to the quantity field and check the first character.
    For i = 1 To Len(strRow)
        strChar = Mid(strRow, i, 1)
        If strChar = "," Then
            intCommaCounter = intCommaCounter + 1
        End If
        
        'Check the comma counter to see if you have reached the quantity field.
        'If so, check to see if the first character after the comma is a quote.
        'if so, toggle the booQtyNeeds Replacing variable to True.
        'Note that a quote is ASCII character 34.
        If intCommaCounter = intPosQty - 1 Then
            If Asc(Mid(strRow, i + 1, 1)) = 34 Then
                booQtyNeedsReplacing = True
            End If
            Exit For
        End If
    Next i
    
    
    'If the quantity has a comma in it, replace it with
    'a quantity that does not have a comma in it.
    'Otherwise, return the original value of the row, unchanged.
    
    If booQtyNeedsReplacing = True Then
        
        'Extract the quantity from the row, including the unwanted comma.
        'The name will go into the variable called strOldQty.
        intCommaCounter = 0
        For i = 1 To Len(strRow)
            strChar = Mid(strRow, i, 1)

            If strChar = "," Then
                intCommaCounter = intCommaCounter + 1
            End If
        
            If intCommaCounter = intPosQty - 1 Or intCommaCounter = intPosQty Then
                strOldQty = strOldQty & strChar
            End If

        Next i
        
        'Rebuild the quantity without the comma or the quotes.
        For i = 1 To Len(strOldQty)
            strChar = Mid(strOldQty, i, 1)
            If strChar <> "," And Asc(strChar) <> 34 Then
                    strNewQty = strNewQty & strChar
            End If
        Next i
        
        'Replace the old qty with the new qty
        intCommaCounter = 0
        For i = 1 To Len(strRow)
            strChar = Mid(strRow, i, 1)

            If strChar = "," Then
                intCommaCounter = intCommaCounter + 1
            End If
        
        
            If booQtyReplaced = True Then
                If intCommaCounter >= intPosQty + 1 Then
                    ReplaceQty = ReplaceQty & strChar
                    'MsgBox (ReplaceQty)
                End If
            Else
                If intCommaCounter < intPosQty - 1 Then
                    ReplaceQty = ReplaceQty & strChar
                ElseIf intCommaCounter = intPosQty - 1 Then
                    ReplaceQty = ReplaceQty & strChar & strNewQty
                    
                    'MsgBox (ReplaceQty)
                    booQtyReplaced = True
                End If
            End If
        Next i
    Else
        'There is no comma in the quantity field.
        ReplaceQty = strRow
    
    End If
    
    
End Function

