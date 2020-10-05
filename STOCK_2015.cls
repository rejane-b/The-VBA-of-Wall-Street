

Sub STOCK_2015()

'set an initial variable for holding the total volume
Dim volume_total As Double
volume_total = 0

'Variable that stores the ticker for year end analysis
Dim ticker_name As String
ticker_name = ""

'Row to store the stock ticker
Dim ticker_row As Long
ticker_row = 2

Dim opening_price As Double
opening_price = 0

Dim percent_change As Double
percent_change = 0

Dim closing_price As Double
closing_price = 0

'Boolean to evaluate whether a new ticker has been found
Dim found_first_occurrence As Boolean

found_first_occurrence = True


Dim max_change As Double
Dim max_change_stock As String

max_change = 0


Dim min_change As Double
Dim min_change_stock As String

min_change = 0

Dim max_volume As Double
Dim max_volume_stock As String

max_volume = 0


Dim i As Long
i = 2


'loop through all stock market data

Do While Not IsEmpty(Cells(i, 1))

If found_first_occurrence = True Then
    opening_price = Cells(i, 3)
    found_first_occurrence = False
End If

'check if we are still analyzing data from the same stock, if not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


'set the ticker
ticker_name = Cells(i, 1).Value

closing_price = Cells(i, 6).Value

volume_total = volume_total + Cells(i, 7).Value

If volume_total > max_volume Then
max_volume = volume_total
max_volume_stock = ticker_name

End If



'Fill out the summary table
Range("i" & ticker_row).Value = ticker_name

Range("j" & ticker_row).Value = closing_price - opening_price

If opening_price = 0 Then

    Range("k" & ticker_row).Value = "N/A"
Else
    percent_change = (closing_price - opening_price) / opening_price
    
    If percent_change < min_change Then
            min_change = percent_change
            min_change_stock = ticker_name
    ElseIf percent_change > max_change Then
        max_change = percent_change
        max_change_stock = ticker_name
    End If
    
    Range("k" & ticker_row).Value = percent_change
    
    If percent_change < 0 Then
        Range("k" & ticker_row).Interior.ColorIndex = 3
    ElseIf percent_change > 0 Then
        Range("k" & ticker_row).Interior.ColorIndex = 4
    End If
End If


Range("l" & ticker_row).Value = volume_total

ticker_row = ticker_row + 1

'reset looping variable
volume_total = 0
found_first_occurrence = True


'if the cell immediately following a row is the same ticker...
Else

volume_total = volume_total + Cells(i, 7).Value

If volume_total > max_volume Then
max_volume = volume_total
max_volume_stock = ticker_name

End If


End If

i = i + 1

Loop


'Print Challenge assignment
Range("o2").Value = max_change_stock
Range("p2").Value = max_change

Range("o3").Value = min_change_stock
Range("p3").Value = min_change

Range("o4").Value = max_volume_stock
Range("p4").Value = max_volume



End Sub

