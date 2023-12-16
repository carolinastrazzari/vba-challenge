Attribute VB_Name = "Module1"
Sub Stock()
    For Each ws In Worksheets
        ws.Activate
        Dim Yearly_Change As Double
        Dim Percentage_change As Double
        Dim Total_Stock As Double
        'MsgBox (ws.Name)
        'For Each ws In worsheets
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Row_count = Cells(Rows.Count, "A").End(xlUp).Row
        
        Total_vol = 0
        Open_Price = Cells(2, "C").Value
        Summary_Pointer = 2
            
        For i = 2 To Row_count
            
            
            
            If Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
                Total_vol = Total_vol + Cells(i, "G").Value
                Close_Price = Cells(i, "F").Value
                
                Yearly_Change = Close_Price - Open_Price
                
                Percentage_change = Yearly_Change / Open_Price * 100
                
                Cells(Summary_Pointer, "I").Value = Cells(i, "A").Value
                Cells(Summary_Pointer, "J").Value = Yearly_Change
                Cells(Summary_Pointer, "K").Value = "%" & Percentage_change
                Cells(Summary_Pointer, "L").Value = Total_vol
                
                If Yearly_Change > 0 Then
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 3
                Else
                    Cells(Summary_Pointer, "J").Interior.ColorIndex = 2
                End If
                
                Total_vol = 0
                
                Open_Price = Cells(i + 1, "C").Value
                Summary_Pointer = Summary_Pointer + 1
            Else
        
                Total_vol = Total_vol + Cells(i, "G").Value

            End If

        Next i
    Next ws
    MsgBox ("Complete")


End Sub
