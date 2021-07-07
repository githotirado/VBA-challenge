Sub Main()
' VBA Challenge main subroutine.  START WITH THIS SUBROUTINE!
    ' Run through each worksheet
    Dim Current_Sheet As Worksheet

    For Each Current_Sheet In Worksheets
        ' Go to sheet
        Current_Sheet.Activate
        ' Run yearly change and greatest change subroutines
        YearlyChange
        GreatestChanges
    Next
End Sub

Sub YearlyChange()
    ' Define variables saved open, saved close, yearly change, saved volume
    Dim Saved_Open, Saved_Close, Yearly_Change As Double
    Dim Saved_Volume As LongLong
    Saved_Volume = 0

    ' Initialize lowest saved close date, highest saved open date
    Dim Saved_Close_Date, Saved_Open_Date As Long
    Saved_Open_Date  = 99999999
    Saved_Close_Date = 0

    ' Define each row element type when reading line by line
    Dim Ticker As String
    Dim Ticker_Date, Volume As Long
    Dim Open_Amt, High_Amt, Low_Amt, Close_Amt As Double

    ' Define variables for percentage calculation
    Dim pct_change As Double
    Dim Percent_Change As String

    ' Initialize first row number where results are printed
    Dim Result_Row As String
    Result_Row = 2

    ' Write out result headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Identify the sheet max row
    Dim Max_Row As Long
    Max_Row = Range("A1").End(xlDown).row

    ' loop through each row in the original data
    For My_Row = 2 To Max_Row
        ' Read Ticker, Date, Open, High, Low, Close, Volume
        Ticker      = Cells(My_Row, 1).Value
        Ticker_Date = Cells(My_Row, 2).Value
        Open_Amt    = Cells(My_Row, 3).Value
        High_Amt    = Cells(My_Row, 4).Value
        Low_Amt     = Cells(My_Row, 5).Value
        Close_Amt   = Cells(My_Row, 6).Value
        Volume      = Cells(My_Row, 7).Value

        ' If Date < saved open date
        If Ticker_Date < Saved_Open_Date Then
            ' New saved open date = Date
            Saved_Open_Date = Ticker_Date
            ' earliest open amount is Open
            Saved_Open = Open_Amt
        End If

        ' If Date > saved close date
        If Ticker_Date > Saved_Close_Date Then
            ' New saved close date = Date
            Saved_Close_Date = Ticker_Date
            ' latest close amount is Close
            Saved_Close = Close_Amt
        End If
        
        ' Increment my saved volume by current volume amount
        Saved_Volume = Saved_Volume + Volume

        ' If next row ticker is not same as current row ticker
        If Cells(My_Row, 1).Offset(1, 0).Value <> Ticker Then
            ' Calculate yearly channge from saved open and close prices
            Yearly_Change = Saved_Close - Saved_Open

            ' Calculate percent change from saved prices. Avoid divide by zero
            If Saved_Open = 0 Then
                pct_change = 0
            Else
                pct_change = Yearly_Change / Saved_Open
            End If
            Percent_Change = FormatPercent(pct_change)

            ' Display results at result row with any formatting
            Cells(Result_Row,  9).Value = Ticker
            Cells(Result_Row, 10).Value = Yearly_Change
            Cells(Result_Row, 11).Value = Percent_Change
            Cells(Result_Row, 12).Value = Saved_Volume

            ' Yearly change is green if positive, red if negative
            If Yearly_Change >= 0 Then
                Cells(Result_Row, 10).Interior.Color = RGB(0,255,0)
            Else
                Cells(Result_Row, 10).Interior.Color = RGB(255,0,0)
            End If

            ' Increment result row for next ticker result
            Result_Row = Result_Row + 1
            ' Reset saved close date, saved open date, saved volume
            Saved_Open_Date  = 99999999
            Saved_Close_Date = 0
            Saved_Volume = 0
            ' Reset variables saved open, saved close
            Saved_Open  = 0
            Saved_Close = 0
        End If
    ' Next row
    Next My_Row

End Sub

Sub GreatestChanges()

    ' Set up table headers
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Initialize counters and output variables
    Dim Current_ticker, GI_ticker, GD_ticker, G_ticker As String
    Dim GI_percent, GD_percent As String
    Dim Current_pct, GI_pct, GD_pct As Double
    Dim Current_vol, G_vol As LongLong
    GI_pct = 0
    GD_pct = 0
    G_vol  = 0

    ' Identify max record row number
    Dim Max_Result_row As Long
    Max_Result_row = Range("K1").End(xlDown).row

    ' Loop through ticker aggregates
    For Current_Row = 2 To Max_Result_row
        ' Read in result values
        Current_ticker = Cells(Current_Row,  9).Value
        Current_pct    = Cells(Current_Row, 11).Value
        Current_vol    = Cells(Current_Row, 12).Value

        ' if value encountered is greater than stored
            ' overwrite previous value
        If Current_pct > GI_pct Then
            GI_pct = Current_pct
            GI_ticker = Current_ticker
        End If
        If Current_pct < GD_pct Then
            GD_pct = Current_pct
            GD_ticker = Current_ticker
        End If
        If Current_vol > G_vol Then
            G_vol = Current_vol
            G_ticker = Current_ticker
        End If
    Next Current_row

    ' Format percentages
    GI_percent = FormatPercent(GI_pct)
    GD_percent = FormatPercent(GD_pct)

    ' Display results
    Range("P2").Value = GI_ticker
    Range("Q2").Value = GI_percent
    Range("P3").Value = GD_ticker
    Range("Q3").Value = GD_percent
    Range("P4").Value = G_ticker
    Range("Q4").Value = G_vol

End Sub