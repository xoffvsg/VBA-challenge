Attribute VB_Name = "tickerSummary"
Sub tickerSummary()


'******* WARNING ********
'   This Subroutine relies on a dictionnary having multiple Items per Key
'   A new Class Module needs to be created with the following entries:
'       ' tickerList Class Module Code
'       Public jan01Open As Double
'       Public dec31Close As Double
'       Public volume As Double
'
'
'   In addition, the box in the VBA menu "Tools->References" for "Microsoft Scripting Runtime" must be checked
'
'************************

'----------------------------------------------------------------
    Dim WS_Count, i, j, k, m As Integer, wsheetNames() As String
    
    
    WS_Count = ActiveWorkbook.Worksheets.Count                  'Counts the number of worksheets in the workbook
    ReDim wsheetNames(WS_Count)                                 'Creates an array for the names of all the worksheets
    
    For i = 1 To WS_Count

        wsheetNames(i - 1) = ActiveWorkbook.Worksheets(i).Name  'Fill the array with the names of all the worksheets

    Next i
    
'----------------------------------------------------------------
    ' Create and initialize a dictionary to track the open value on Jan 2nd,
    ' the close value and Dec 31st, and the cumulated volume through the year
    
    
    
    Dim dict As New Scripting.Dictionary    'Early binding: Need to check the box in Tools->References for Microsoft Scripting Runtime
    

    Dim ticker As tickerList                 'Declare the new object as the Class. This must be outside of the loop

    
'----------------------------------------------------------------
'   Start looping the first worksheet to populare the dictionary

   

    
    Dim entry As String
    Dim sh As Worksheet
    
    For j = 0 To WS_Count - 1
    dict.RemoveAll                              'Reset the dictionary for the current spreadsheet
        k = 2   'Start the While loop at row 2.
        Set sh = ThisWorkbook.Worksheets(wsheetNames(j))
        
    
        While sh.Cells(k, "A").Value <> ""
            entry = sh.Cells(k, "A").Value
        
            If Not dict.Exists(entry) Then
                Set ticker = New tickerList 'Instantiate a new object. It must be inside the loop or the last entry to the dictionnary
                                            'will override the Items of the previous entries!
                                            'https://stackoverflow.com/questions/18332166/vba-dictionary-adding-an-item-overwrites-all-items

                dict.Add entry, ticker
            
            End If
        
            If Right(sh.Cells(k, "B").Value, 4) = "0102" Then
            dict(entry).jan01Open = sh.Cells(k, "C").Value
            End If
        
            If Right(sh.Cells(k, "B").Value, 4) = "1231" Then
            dict(entry).dec31Close = sh.Cells(k, "F").Value
            End If
            
            dict(entry).volume = dict(entry).volume + sh.Cells(k, "G").Value
    
            k = k + 1
        Wend
        

        ' Display the results
        With sh              'Format new Headers
            .Cells(1, "H").ColumnWidth = 20
            .Cells(1, "I").Value = "Ticker"
            .Cells(1, "I").HorizontalAlignment = xlLeft
            .Cells(1, "I").ColumnWidth = 10
            .Cells(1, "J").Value = "Yearly Change"
            .Cells(1, "J").HorizontalAlignment = xlRight
            .Cells(1, "J").ColumnWidth = 13
            .Cells(1, "K").Value = "Percent Change"
            .Cells(1, "K").HorizontalAlignment = xlRight
            .Cells(1, "K").ColumnWidth = 15
            .Cells(1, "L").Value = "Total Stock Volume"
            .Cells(1, "L").HorizontalAlignment = xlRight
            .Cells(1, "L").ColumnWidth = 20
            .Cells(1, "M").ColumnWidth = 20
            .Cells(2, "N").Value = "Greatest % Increase"
            .Cells(3, "N").Value = "Greatest % Decrease"
            .Cells(4, "N").Value = "Greatest Total Volume"
            .Cells(1, "N").ColumnWidth = 20
            .Cells(1, "O").Value = "Ticker"
            .Cells(1, "O").ColumnWidth = 10
            .Cells(1, "P").Value = "Value"
            .Cells(1, "P").ColumnWidth = 20
            .Range("O1:P4").HorizontalAlignment = xlLeft
            .Range("P2:P4").HorizontalAlignment = xlRight
            .Cells.VerticalAlignment = xlCenter
        End With



    
        Dim index As Variant
        i = 2
        For Each index In dict.Keys()
            sh.Cells(i, "I").Value = index
            Dim a, b, c As Double
            a = dict(index).dec31Close
            b = dict(index).jan01Open
            c = (a - b) / b
            sh.Cells(i, "J").Value = a - b
            sh.Cells(i, "K").Value = c
            sh.Cells(i, "K").NumberFormat = "0.00%"
            sh.Cells(i, "L").Value = dict(index).volume
            sh.Cells(i, "L").NumberFormat = "#,##0"
            If sh.Cells(i, "J").Value < 0 Then
            sh.Cells(i, "J").Interior.Color = RGB(255, 0, 0)
            Else
                sh.Cells(i, "J").Interior.Color = RGB(0, 255, 0)
            End If
            i = i + 1
        Next
    

        'Identifies the Max and Min items in the list
        Dim ticker1, ticker2, ticker3 As String, val_1, val_2, val_3 As Double
        val_1 = 0
        val_2 = 0
        val_3 = 0
    
        For i = 2 To dict.Count + 1
            If sh.Cells(i, "K").Value >= val_1 Then
                val_1 = sh.Cells(i, "K").Value
                ticker1 = sh.Cells(i, "I").Value
            ElseIf sh.Cells(i, "K").Value < val_2 Then
                val_2 = sh.Cells(i, "K").Value
                ticker2 = sh.Cells(i, "I").Value
            ElseIf sh.Cells(i, "L").Value > val_3 Then
                val_3 = sh.Cells(i, "L").Value
                ticker3 = sh.Cells(i, "I").Value
            End If
        Next i
    
        sh.Cells(2, "O").Value = ticker1
        sh.Cells(2, "P").Value = val_1
        sh.Cells(2, "P").NumberFormat = "0.00%"
        sh.Cells(3, "O").Value = ticker2
        sh.Cells(3, "P").Value = val_2
        sh.Cells(3, "P").NumberFormat = "0.00%"
        sh.Cells(4, "O").Value = ticker3
        sh.Cells(4, "P").Value = val_3
        sh.Cells(4, "P").NumberFormat = "#,##0"
            
    Next j


End Sub
