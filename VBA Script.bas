Attribute VB_Name = "Module2"
Sub VBA_Challenge()
    'Identify variables
    Dim openp As Double
    Dim newclose As Double
    Dim yc As Double
    Dim pc As Double
    Dim tsv As LongLong
    Dim ftsv As LongLong
    Dim r As LongLong
    Dim r2 As Long
    Dim Max_inc As Double
    Dim Max_dec As Double
    Dim Max_vol As LongLong
    Dim Tic_inc As Variant
    Dim Tic_dec As Variant
    Dim Tic_vol As Variant
    
    'Insert labels
    Range("I1,P1").Value = "Ticker"
    Range("J11").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest total volume"
    
    'Collect the data
    'Establish the start
    openp = Cells(2, 3).Value
    tsv = 0
    r2 = 2
    
    'Run a Loop
    For r = 2 To 753001
        'Split the tickers
        If Cells(r, 1).Value = Cells(r + 1, 1).Value Then
            'Set the open
            tsv = tsv + Cells(r, 7).Value
        Else
            'Identify the close
            newclose = Cells(r, 6).Value
            'List results
            yc = newclose - openp
            pc = -1 * (1 - ((openp + yc) / openp))
            ftsv = tsv + Cells(r, 7).Value
            Cells(r2, 9).Value = Cells(r, 1).Value
            Cells(r2, 10).Value = yc
            Cells(r2, 11).Value = FormatPercent(pc)
            Cells(r2, 12).Value = ftsv
            'reset the start
            openp = Cells(r + 1, 3).Value
            r2 = r2 + 1
            tsv = 0
        End If
    Next r
    
    'Highlight the positive and negative yearly changes
    'Start loop
    For r = 2 To 3001
        'Highlight green if yc is positive
        If Cells(r, 10).Value > 0 Then
            Cells(r, 10).Interior.ColorIndex = 4
        'Highlight red if yc is negative
        ElseIf Cells(r, 10).Value < 0 Then
            Cells(r, 10).Interior.ColorIndex = 3
        'Highlight yellow if yc is 0
        ElseIf Cells(r, 10).Value = 0 Then
            Cells(r, 10).Interior.ColorIndex = 6
        End If
    Next r
    
    'Analyze the data
    Max_inc = Application.WorksheetFunction.Max(Range("K2:K3001"))
    Max_dec = Application.WorksheetFunction.Min(Range("K2:K3001"))
    Max_vol = Application.WorksheetFunction.Max(Range("L2:L3001"))
    
    'Insert the data results
    Range("Q2").Value = FormatPercent(Max_inc)
    Range("Q3").Value = FormatPercent(Max_dec)
    Range("Q4").Value = Max_vol
    
    'Identify the tickers associated with the data
    Tic_inc = Application.WorksheetFunction.Index(Range("I2:I3001"), Application.WorksheetFunction.Match(Max_inc, Range("K2:K3001"), 0))
    Tic_dec = Application.WorksheetFunction.Index(Range("I2:I3001"), Application.WorksheetFunction.Match(Max_dec, Range("K2:K3001"), 0))
    Tic_vol = Application.WorksheetFunction.Index(Range("I2:I3001"), Application.WorksheetFunction.Match(Max_vol, Range("L2:L3001"), 0))
    
    'Insert the tickers
    Range("P2").Value = Tic_inc
    Range("P3").Value = Tic_dec
    Range("P4").Value = Tic_vol
End Sub
