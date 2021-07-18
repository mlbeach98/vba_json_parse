Attribute VB_Name = "Module2"
Sub main_format()
    'This subprocedure will be the main procedure run to format the json data
    'This will take the list of dictionaries, and for each dictionary it will
    ' use two columns on the formatted_result tab to show the data. This will
    ' have the person's name as the header for the data, then the key and value
    ' pairs below that
    
    'To first clean up the data we'll break up each of the key value pairs into
    ' their own cells (still one row per dictionary) to clean it up a bit
    
    Sheets("raw_data").Activate
    
    'Find the bottom row of the data to know how far we need our for loop to go
    Range("A1").Select
    Selection.End(xlDown).Select
    bottomRow = Selection.Row
    
    'Loop over each row
    For i = 1 To bottomRow
    
        'We will split the data on "," as this seperates the key/value pairs,
        ' but since most lines end with a "," we will remove this so we don't
        ' have an extra element in our array
        'Note here that we are saving this changed version in a variable and
        ' not overwriting the cell value so we can keep the original version
        ' of the JSON if it is later needed
        If Right(Range("A" & i), 1) = "," Then
            newDict = Left(Range("A" & i).Value, Len(Range("A" & i).Value) - 1)
        Else
            newDict = Range("A" & i).Value
        End If
        
        'We will also remove all quotations, brackets, and braces to clean our
        ' raw data since we know how the data is structured
        newDict = Replace(newDict, "[", "")
        newDict = Replace(newDict, "{", "")
        newDict = Replace(newDict, "}", "")
        newDict = Replace(newDict, "]", "")
        newDict = Replace(newDict, """", "")
        
    
        'Use the split function to create an array with each key/value pair
        Dim tempArr() As String
        tempArr = Split(newDict, ",")
        
        'We will loop over the elements in the array and put them into the
        ' spreadsheet starting in column A on the step_1 sheet
        
        'We will also use a number to letter conversion function (not my own code
        ' but very useful function I've found) to determine the column letter
        startColNum = 1
        
        For Each j In tempArr
            'Before we put the key/value pairs into their own cells, we will clean
            ' up the pairs by trimming any leading or trailing spaces
            j = Trim(j)
            Sheets("step_1").Range(Col_Letter(CLng(startColNum)) & i).Value = j
            
            'Remember to increment startColNum to not overwrite the data we have
            startColNum = startColNum + 1
            
        Next j
        
    Next i
    
    'Now that we've done our first cleaning step to the data and broken
    ' the key/value pairs up, we will turn this into our formatted version
    
    Sheets("step_1").Select
    
    'Again, we'll find the bottom row of the data for looping purposes; this should
    ' be the same as our first sheet, but if we changed something above to limit
    ' the records that make it in, we want to make sure we have an accurate stopping
    ' point
    Range("A1").Select
    Selection.End(xlDown).Select
    bottomRow2 = Selection.Row
    
    For i = 1 To bottomRow2
    
        'For this data, we know the first name is in column B and the last
        ' name is in column C, but lets assume that this won't always be the case
        ' as order within a dictionary does not matter (in terms of uniqueness)
        'We will loop over the columns in each row and see which one starts with
        ' the key value we are looking for
        'We can also generalize this example and not assume that all dictionaries
        ' have the same number of elements, so we will check the length for each
        ' one individually
        Range("A" & i).Select
        Selection.End(xlToRight).Select
        tempEndCol = Selection.Column
        
        'Just a note, when thinking about how I wanted the formatted data to look,
        ' I figured making cards of sort would make the most sense. It will be set
        ' up to be 3 cards wide and then however many rows are needed down
        'To determine the starting location, we will use mod 3 to find the column
        ' and then use integer division to find the row
        formCol = (((i - 1) Mod 3) * 3) + 1
        formRow = (((i - 1) \ 3) * 8) + 1
        
        nameFlag = False
        firstName = ""
        lastName = ""
        For j = 1 To tempEndCol
            'Check if key matches "last_name"
            If Left(Range(Col_Letter(CLng(j)) & i).Value, 9) = "last_name" Then
                lastName = Trim(Split(Range(Col_Letter(CLng(j)) & i).Value, ":")(1))
            
            'Check if key matches "first_name"
            ElseIf Left(Range(Col_Letter(CLng(j)) & i).Value, 10) = "first_name" Then
                firstName = Trim(Split(Range(Col_Letter(CLng(j)) & i).Value, ":")(1))
            
            End If
            
            'Even though we are using the first and last name in the header for each
            ' person, we will still keep their key/value pairs in the overall dictionary
            ' we create for them
            Sheets("formatted_result").Range(Col_Letter(CLng(formCol)) & formRow + j).Value = Trim(Split(Range(Col_Letter(CLng(j)) & i).Value, ":")(0))
            Sheets("formatted_result").Range(Col_Letter(CLng(formCol + 1)) & formRow + j).Value = Trim(Split(Range(Col_Letter(CLng(j)) & i).Value, ":")(1))
            
            'Once we have the first and last name, we will add this to the header
            If nameFlag = False And firstName <> "" And lastName <> "" Then
                'merge the header cells together
                Sheets("formatted_result").Activate
                merge_cells Col_Letter(CLng(formCol)) & formRow, Col_Letter(CLng(formCol + 1)) & formRow
                Sheets("step_1").Activate
                
                Sheets("formatted_result").Range(Col_Letter(CLng(formCol)) & formRow).Value = lastName & ", " & firstName
                nameFlag = True
            
            End If
            
        Next j
    
    Next i
    
    
    'Now that we have the main data on the formatted_result tab, we will make
    ' everything look nice with excel formatting
    
    'First we will resize the columns to fit all the data in them
    'Since we went 3 cards wide and each one takes 3 spots (including a buffer)
    ' we will need to work on columns A through I
    Sheets("formatted_result").Select
    Columns("A:I").EntireColumn.AutoFit
    
    'We will also add thick borders around each card and set the background to
    ' white so the excel gridlines aren't seen
    
    'We will use the bottomRow2 variable to recall how many total cards there
    ' are and use that to limit our looping
    endLimit = bottomRow2 \ 3 + 1
    'Including header and buffer row, each card uses 8 rows
    rowSkipFactor = 8
    For i = 1 To endLimit
        startRow = ((i - 1) * rowSkipFactor) + 1
        
        For j = 1 To 3 'total number of cards in a row
            range1 = Col_Letter(CLng(((j - 1) * 3) + 1)) & startRow
            range2 = Col_Letter(CLng(((j - 1) * 3) + 1)) & startRow + 1
            range3 = Col_Letter(CLng(((j - 1) * 3) + 2)) & startRow + 6
        
            'Using a macro I recorded on one test card, we can generalize this and
            ' use dynamic cell references to apply the formatting to all cards
            
            If Range(range1).Value <> "" Then
                card_border range1, range2, range3
            End If
        
        Next j
        
    Next i
    
    'We will also set the background color of the whole sheet to white
    ' so the cards stand out
    Cells.Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Sub merge_cells(range1, range2)
    
    Range(range1 & ":" & range2).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge

End Sub

Sub card_border(range1, range2, range3)
'
' Macro3 Macro
'

'
    Range(range1 & ":" & range3).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range(range2 & ":" & range3).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlHairline
    End With
End Sub


Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
