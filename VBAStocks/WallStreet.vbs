Sub WallStreet()

Dim i As Integer
Dim j As Long
Dim m As Long
Dim n As Long
Dim lRowNew As Integer
Dim k As Integer
Dim lRow As Long
Dim FirstValue(2) As Double
Dim LowestValue(1) As Double
Dim HighestValue(1) As Double
Dim HighestVol(1) As Double
Dim TopVolume(1) As Double
Dim LastValue As Double
Dim LineOne As Long
i = 1
'Scaning through all the worksheet

For i = 1 To Worksheets.Count
    k = 2
    j = 2
    o = 2
    lRow = 0
    lRowNew = 0
'Activating current sheet
    Worksheets(i).Activate
'Finding last row in the sheet
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
'Saving first value for finding Volume,Percentage
    LowestValue(0) = Cells(2, 11)
    HighestValue(0) = Cells(2, 11)
    HighestVol(0) = Cells(2, 12)
'Looping to finding the different values of output
    For j = 2 To lRow
'Finding ticker
    If Cells(j, 1) <> Cells(j + 1, 1) Then
        Cells(k, 9) = Cells(j, 1)
        LastValue = Cells(j, 6)
'Processing first tickers value seperately
        If k = 2 Then
        Cells(k, 10) = LastValue - Cells(2, 3)
'Checking if the finding percentage it is not failing and replacing the denominator with 1 from zero
        If Cells(2, 3) = 0 Then
            Cells(2, 3) = 1
        End If
' Finding percentage for first ticker
        Cells(k, 11) = FormatPercent(Cells(k, 10) / Cells(2, 3))
        Cells(k, 12) = Application.Sum(Range("G2:G" & j))
        Else
'Finding Yearly Change
        Cells(k, 10) = LastValue - FirstValue(0)
        If FirstValue(0) = 0 Then
        FirstValue(0) = 1
        End If
'Finding percentage
        Cells(k, 11) = FormatPercent(Cells(k, 10) / FirstValue(0))
        For n = m To j
'Finding total Stock Volume
        Cells(k, 12) = Cells(n, 7) + Cells(k, 12)
        Next n
         End If
            m = j + 1
         FirstValue(0) = Cells(j + 1, 3)
         k = k + 1
    End If
    Next j
'Challenge section
    lRowNew = Cells(Rows.Count, 9).End(xlUp).Row
    For o = 2 To lRowNew
'Finding lowest value for Percentage Change
    If LowestValue(0) > Cells(o, 11) Then
        LowestValue(0) = Cells(o, 11)
        Cells(3, 16) = Cells(o, 9)
        End If
'Finding Highest value for percentage change
    If HighestValue(0) < Cells(o, 11) Then
        HighestValue(0) = Cells(o, 11)
        Cells(2, 16) = Cells(o, 9)
        End If
'Finding Highest volumne
    If HighestVol(0) < Cells(o, 12) Then
        HighestVol(0) = Cells(o, 12)
        Cells(4, 16) = Cells(o, 9)
        End If
'Colouring the cell
    If Cells(o, 10) < 0 Then
        Cells(o, 10).Interior.ColorIndex = 3
        Else
        Cells(o, 10).Interior.ColorIndex = 4
        End If
        
        
    Next o
'Pasting in the found value

    Cells(3, 17) = FormatPercent(LowestValue(0))
    Cells(2, 17) = FormatPercent(HighestValue(0))
    Cells(4, 17) = HighestVol(0)
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volumne"
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percentage Change"
    Cells(1, 12) = "Total Stock Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    
Next i


End Sub