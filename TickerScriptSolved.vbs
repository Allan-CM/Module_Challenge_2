Attribute VB_Name = "Module1"
Sub Ticker()

'Defining Ticker Symbol as a string variable
Dim Ticker As String
'Defining LastRow a long integer to store # of rows
Dim LastRow As Long
'Defining Output table row
Dim OutputTableRow As Long

'Setting excel formula to Determine number of rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Setting the starting value of the Output table
OutputTableRow = 2

'Defining a new column head for Ticker Symbols
Range("I1").Value = "Ticker Symbol"

'Fit the Ticker Column Header within the cell
Columns("I:I").EntireColumn.AutoFit

'Creating a for loop to scan through each cell in Ticker COlumn
For i = 2 To LastRow

    ' Check cells to see if we are the symbol ticker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Defines the ticket variable as the last symbol before it detects a change
        Ticker = Cells(i, 1).Value
        
        'Storing Ticket names in the new Ticker Symbol Column
        Cells(OutputTableRow, 9).Value = Ticker
        
        'Telling OutputTable Row to increase by 1 to store the next symbol in a new row within the same column
        OutputTableRow = OutputTableRow + 1
        
    End If
    
Next i

End Sub
