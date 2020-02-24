Attribute VB_Name = "Module1"
Sub Ticker_Volume():

    Dim ws As Worksheet
    Dim Ticker_Name As String
    Dim Total_Volume As Double
    Dim Current_Ticker_Row As Integer
        Current_Ticker_Row = 2
   
   
   'Looping through worksheets to add headers
   
    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        Current_Ticker_Row = 2
        
            Cells(1, 10) = "Ticker"
            Cells(1, 13) = "Total_Volume"
            
        lastrow = ActiveSheet.UsedRange.Rows.Count
       
   'Looping through cells, sheets to calculate Total Volume by Ticker Name and output into appropriate columns
   
        For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker_Name = Cells(i, 1).Value
        
            Total_Volume = Total_Volume + Cells(i, 7).Value
        
            Range("J" & Current_Ticker_Row).Value = Ticker_Name
            
            Range("M" & Current_Ticker_Row).Value = Total_Volume
        
            Current_Ticker_Row = Current_Ticker_Row + 1
        
            Total_Volume = 0
        
        Else
        
            Total_Volume = Total_Volume + Cells(i, 7).Value
        
        End If
        
        Next i
            
    Next ws
            
End Sub

Sub Change_Metrics()
    
    Dim ws As Worksheet
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Open_Date As Double
    Dim Close_Date As Double
    Dim Close_Price As Double
        Close_Price = 0
    Dim Open_Price As Double
        Open_Price = 0
    Dim Current_Ticker_Row As Integer
        Current_Ticker_Row = 2

    'Looping through worksheets to add headers and set non-static Open & Close Dates

    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
            Current_Ticker_Row = 2
    
            Cells(1, 11) = "Yearly_Change"
            Cells(1, 12) = "Percent_Change"
            
            lastrow = ActiveSheet.UsedRange.Rows.Count
            Open_Date = WorksheetFunction.Min(ws.Range("B2:B" & lastrow).Value)
            Close_Date = WorksheetFunction.Max(ws.Range("B2:B" & lastrow).Value)
    
    'Looping through cells, sheets to determine Open Price and Close Price
    
    For i = 2 To lastrow
    
            If Cells(i, 2).Value = Close_Date Then
            
            Close_Price = Cells(i, 6).Value
            
            End If
    
            If Cells(i, 2).Value = Open_Date Then
            
            Open_Price = Cells(i, 3).Value
            
            End If
                
    'Calculating Yearly Change and Percent Change w/formatting
                
        If Cells(i, 2).Value = Close_Date Then
        
            Yearly_Change = Close_Price - Open_Price
            
            Percent_Change = (Close_Price / Open_Price) - 1
            
            Range("K" & Current_Ticker_Row).Value = Yearly_Change
            
            Range("L" & Current_Ticker_Row).Value = Percent_Change
                
            Current_Ticker_Row = Current_Ticker_Row + 1
            
            Range("L2:L" & lastrow).NumberFormat = "0.00%"
            
        End If
        
   'Conditional Formatting of the Yearly Change column by worksheet
        
        If Cells(i, 11).Value > "0" Then
            
           Cells(i, 11).Interior.ColorIndex = 4
           
           End If
        
        If Cells(i, 11).Value < "0" Then
        
            Cells(i, 11).Interior.ColorIndex = 3
            
        End If
        
        If IsEmpty(Cells(i, 11).Value) Then
        
            Cells(i, 11).Interior.ColorIndex = x1None
        End If
        
        Next i
        
    Next ws
    
End Sub

Sub Summary_Metrics():

    Dim ws As Worksheet
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Ticker_Name As String

    'Looping through sheets to add headers and determine summary metrics

    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
            
            Cells(2, 15) = "Greatest % Increase"
            Cells(3, 15) = "Greatest % Decrease"
            Cells(4, 15) = "Greatest Total Volume"
            Cells(1, 16) = "Ticker"
            Cells(1, 17) = "Value"

            lastrow = ActiveSheet.UsedRange.Rows.Count
            Greatest_Increase = WorksheetFunction.Max(ws.Range("L2:L" & lastrow).Value)
            Greatest_Decrease = WorksheetFunction.Min(ws.Range("L2:L" & lastrow).Value)
            Greatest_Volume = WorksheetFunction.Max(ws.Range("M2:M" & lastrow).Value)
            
   'Looping through to determine and output the summary metrics
        
        For i = 2 To lastrow
        
            If Cells(i, 12).Value = Greatest_Increase Then
            
            Ticker_Name = Cells(i, 10).Value
            
            Cells(2, 17).Value = Greatest_Increase
            
            Cells(2, 16).Value = Ticker_Name
            
            End If
            
            If Cells(i, 12).Value = Greatest_Decrease Then
            
            Ticker_Name = Cells(i, 10).Value
            
            Cells(3, 17).Value = Greatest_Decrease
            
            Cells(3, 16).Value = Ticker_Name
            
            Range("Q2:Q3").NumberFormat = "0.00%"
            
            End If
            
            If Cells(i, 13).Value = Greatest_Volume Then
            
            Ticker_Name = Cells(i, 10).Value
            
            Cells(4, 17).Value = Greatest_Volume
            
            Cells(4, 16).Value = Ticker_Name
        
        End If
            
        Next i
        
    Next ws

End Sub
