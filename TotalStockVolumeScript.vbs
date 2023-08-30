Attribute VB_Name = "Module1"
Sub Ticker()

'Defining Ticker Symbol as a string variable
Dim Ticker As String
'Defining LastRow a long integer to store # of rows
Dim LastRow As Long
'Defining YearlyChangeRow a long integer to store # of rows in that column
Dim YearlyChangeRow As Long
'Defining Output table row
Dim OutputTableRow As Long
'Closing Value for a given ticker
Dim ClosingValue As Double
'Opening Value for a given ticker
Dim OpeningValue As Double
'Defining Variable Count to determine the number of rows for a given ticker
Dim Count As Double
'Defining the variable to store stock volume for each ticker
Dim StockVolume As Double


'Setting excel formula to Determine number of rows in original dataset
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Setting the starting value of the Output table
OutputTableRow = 2

'Defining a new column head for Ticker Symbols
Range("I1").Value = "Ticker Symbol"

'Defining a new column head for Yearly Change
Range("J1").Value = "Yearly Change"

'Defining a new column head for Percentage Change
Range("K1").Value = "Percentage Change"

'Defining a new column head for Total Stock Volume
Range("L1").Value = "Total Stock Volume"

'Fit the Column Header within the defined range
Columns("I:L").EntireColumn.AutoFit

'Creating a for loop to scan through each cell in Ticker COlumn
For i = 2 To LastRow

    ' Check cells to see if we there is a change in ticker symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Defines the ticket variable as the last symbol before it detects a change
        Ticker = Cells(i, 1).Value
        
        'Determines the Closing Value at the end of the year
        ClosingValue = Cells(i, 6).Value
        
        'Determine the Opening Value at the start of the year
        OpeningValue = Cells(i - Count, 3).Value
        
        'Storing Ticket names in the new Ticker Symbol Column
        Cells(OutputTableRow, 9).Value = Ticker
        
        'Storing the Change in Closing vs opening Price at the end of the year to the
        Cells(OutputTableRow, 10).Value = ClosingValue - OpeningValue
        
        'Storing the Percentage Change in Price in the Percentage Change Column
        Cells(OutputTableRow, 11).Value = (Cells(OutputTableRow, 10).Value) / OpeningValue
        
        
        'calculating the total sum stock volume for a given ticker
        StockVolume = StockVolume + Cells(i, 7).Value
        'Storing Total Sotck Value for each Ticker in the Total Stock Value Column
        Cells(OutputTableRow, 12).Value = StockVolume
        
        'Telling OutputTable Row to increase by 1 to store the next symbol in a new row within the same column
        OutputTableRow = OutputTableRow + 1
        
        'Reseting Data for the next ticker
        ClosingValue = 0
        OpeningValue = 0
        Count = 0
        StockVolume = 0
        
    Else
    
        Count = Count + 1
        'calculating the total sum stock volume for a given ticker
        StockVolume = StockVolume + Cells(i, 7).Value
        
    End If
   
Next i

'Setting excel formula to Determine number of rows in Yearly Change Column
YearlyChangeRow = Cells(Rows.Count, 10).End(xlUp).Row

'Creating a for loop to scan through each cell in Yearly Change Column
For j = 2 To YearlyChangeRow
    'If the value in the cell is positive change then
    If Cells(j, 10).Value > 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(j, 10).Value < 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

'Formats the Percentage Change Column into a percentage and rounds it to 2 decimal places
Columns("K").NumberFormat = "0.00%"

End Sub
