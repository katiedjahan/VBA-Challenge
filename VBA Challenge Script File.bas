Attribute VB_Name = "Module1"
Sub MYSDMacro()

Dim Current As Worksheet
' Loop through all of the worksheets in the active workbook.
For Each Current In Worksheets

Dim counter As Double
counter = 2
Dim tickercounter As Double
tickercounter = 2
Dim yearcounter As Integer
yearcounter = 0

Dim Column1 As Long
Column1 = Current.Cells(Current.Rows.Count, 1).End(xlUp).Row



Current.Range("I1").Value = "Ticker"
Current.Cells(2, 9).Value = Current.Cells(2, 1).Value
Current.Range("J1").Value = "Yearly Change"
Current.Range("K1").Value = "Percent Change"
Current.Range("L1").Value = "Total Stock Volume"


For i = 2 To Column1
    
    If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
        
        ' Ticker
        tickercounter = tickercounter + 1
        Current.Cells(tickercounter, 9).Value = Current.Cells(i + 1, 1)
        
        ' Yearly Change
        Current.Cells(counter, 10).Value = Current.Cells(i, 6).Value - Current.Cells(i - yearcounter, 3).Value
        
        'Percent Change
        Current.Cells(counter, 11).Value = (Current.Cells(i, 6).Value / Current.Cells(i - yearcounter, 3).Value) - 1
        
        ' Total Stock Volume
        Current.Cells(counter, 12).Value = WorksheetFunction.Sum(Current.Range(Current.Cells(i - yearcounter, 7), Current.Cells(i, 7)))
        
        ' Adjust row
        counter = counter + 1
        
        yearcounter = -1
    
    End If

yearcounter = yearcounter + 1

Next i

Dim Column10 As Integer
Column10 = Current.Cells(Current.Rows.Count, 10).End(xlUp).Row

' Format Color
For i = 2 To Column10
        
        If Current.Cells(i, 10).Value > 0 Then
        
            Current.Cells(i, 10).Interior.Color = vbGreen
        Else
            Current.Cells(i, 10).Interior.Color = vbRed
            
        End If
    
Next i

Dim Column11 As Integer
Column11 = Current.Cells(Current.Rows.Count, 11).End(xlUp).Row

' Percent Change Format
For i = 2 To Column11

    Current.Cells(i, 11).NumberFormat = "0.00%"
    
Next i


Dim Kmax As Double
Kmax = 0
Dim Kmin As Double
Kmin = 0
Dim Lmax As Double
Lmax = 0

Current.Range("O2").Value = "Greatest % Increase"
Current.Range("O3").Value = "Greatest % Decrease"
Current.Range("O4").Value = "Greatest Total Volume"
Current.Range("P1").Value = "Ticker"
Current.Range("Q1").Value = "Value"

For i = 2 To Column11

    If Current.Cells(i, 11).Value > Kmax Then
    Kmax = Current.Cells(i, 11).Value
    Current.Range("P2").Value = Current.Cells(i, 9).Value
    
    End If
    
    If Current.Cells(i, 11).Value < Kmin Then
    Kmin = Current.Cells(i, 11).Value
    Current.Range("P3").Value = Current.Cells(i, 9).Value
    
    End If
    
    If Current.Cells(i, 12).Value > Lmax Then
    Lmax = Current.Cells(i, 12).Value
    Current.Range("P4").Value = Current.Cells(i, 9).Value
    
    End If
    
Next i

Current.Range("Q2:Q3").NumberFormat = "0.00%"
Current.Range("Q2").Value = Kmax
Current.Range("Q3").Value = Kmin
Current.Range("Q4").Value = Lmax

Next

End Sub

Sub FitColumnsMYSD()

Dim Current As Worksheet
For Each Current In Worksheets

Current.Range("J:Q").EntireColumn.AutoFit

Next

End Sub


Sub ResetMYSD()

Dim Current As Worksheet
For Each Current In Worksheets

Current.Range("I:Q").Delete

Next

End Sub
