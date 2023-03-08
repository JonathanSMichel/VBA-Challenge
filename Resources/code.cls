VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub code()

Dim ws As Worksheet

For Each ws In Worksheets

Dim Ticker_Symbol As String
Dim Start_Row As Integer
Dim last_row As Double
Dim Year_Close As Double
Dim Year_Open As Double
Dim Yearly_Change As Double
Dim Volume As Double

    ' printing table headers
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

Start_Row = 2

' finding the last row
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To last_row
    
    If ws.Cells(i, 3).Value = 0 Then
    
        If ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Symbol = ws.Cells(i, 1).Value
            
        End If

    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

        Volume = Volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Year_Open = ws.Cells(i, 3).Value
        End If
        
    Else
        
        Ticker_Symbol = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        Year_Close = ws.Cells(i, 6).Value
        
        ws.Range("I" & Start_Row).Value = Ticker_Symbol
        ws.Range("L" & Start_Row).Value = Volume
        
        If Volume > 0 Then
        
            ws.Range("J" & Start_Row).Value = Year_Close - Year_Open
        
        ' changing color
            If ws.Range("J" & Start_Row).Value > 0 Then
                ws.Range("J" & Start_Row).Interior.ColorIndex = 4
                
            Else
                ws.Range("J" & Start_Row).Interior.ColorIndex = 3
            
            End If
            
        ws.Range("K" & Start_Row).Value = ws.Range("J" & Start_Row) / Year_Open
        
        Else
        
        ws.Range("J" & Start_Row) = 0
        ws.Range("K" & Start_Row) = 0
        
    
    End If
    
        ws.Range("K" & Start_Row).Style = "percent"
        
        Volume = 0
       Start_Row = Start_Row + 1
        
End If
    
Next i
        
        

Next ws

End Sub

