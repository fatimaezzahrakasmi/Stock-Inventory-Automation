Attribute VB_Name = "Module3"
Sub transfertData()
    ' Sheets setup
    Dim wsData As Worksheet, wsDetail As Worksheet
    Set wsData = ThisWorkbook.Sheets(" stock")
    Set wsDetail = ThisWorkbook.Sheets("details")
    
    ' Last rows in both sheets
    Dim lastRowData As Long, lastRowDetail As Long
    lastRowData = wsData.Cells(wsData.Rows.Count, "E").End(xlUp).Row
    lastRowDetail = wsDetail.Cells(wsDetail.Rows.Count, "E").End(xlUp).Row
    
    If lastRowDetail >= 6 Then
        Dim i As Long, found As Boolean
        found = False
        
        ' Loop through the stock data
        For i = 6 To lastRowData
            If wsDetail.Cells(lastRowDetail, 6).Value = wsData.Cells(i, 5).Value Then
                wsData.Cells(i, 8).Value = wsData.Cells(i, 8).Value + wsDetail.Cells(lastRowDetail, 8).Value
                wsData.Cells(i, 9).Value = wsData.Cells(i, 9).Value + wsDetail.Cells(lastRowDetail, 9).Value
                wsData.Cells(i, 10).Value = wsData.Cells(i, 7).Value + wsData.Cells(i, 8).Value - wsData.Cells(i, 9).Value
                
                ' Prevent negative stock
                If wsData.Cells(i, 10).Value < 0 Then
                    wsData.Cells(i, 10).Value = 0
                End If
                
                ' Determine stock status
                Dim status As String
                If wsData.Cells(i, 10).Value = 0 Then
                    status = "Rupture de stock"
                ElseIf wsData.Cells(i, 10).Value <= 10 Then
                    status = "faible stock"
                Else
                    status = "En stock"
                End If
                wsData.Cells(i, 12).Value = status ' Update status
                
                ' Color the Status Cell
                Select Case status
                    Case "Rupture de stock"
                        wsData.Cells(i, 12).Interior.Color = RGB(255, 0, 0) ' Red
                    Case "faible stock"
                        wsData.Cells(i, 12).Interior.Color = RGB(255, 255, 0) ' Yellow
                    Case "En stock"
                        wsData.Cells(i, 12).Interior.Color = RGB(0, 255, 0) ' Green
                End Select
                
                found = True
                Exit For
            End If
        Next i
        
        ' If no match is found, add a new row
        If Not found Then
            lastRowData = lastRowData + 1
            wsData.Cells(lastRowData, 5).Value = wsDetail.Cells(lastRowDetail, 6).Value ' Reference
            wsData.Cells(lastRowData, 6).Value = wsDetail.Cells(lastRowDetail, 7).Value ' Product name
            wsData.Cells(lastRowData, 7).Value = 0 ' Initial stock
            wsData.Cells(lastRowData, 8).Value = wsDetail.Cells(lastRowDetail, 8).Value ' Entries
            wsData.Cells(lastRowData, 9).Value = wsDetail.Cells(lastRowDetail, 9).Value ' Exits
            wsData.Cells(lastRowData, 10).Value = wsData.Cells(lastRowData, 7).Value + wsData.Cells(lastRowData, 8).Value - wsData.Cells(lastRowData, 9).Value
            
            ' Determine and update stock status
            Dim newStatus As String
            If wsData.Cells(lastRowData, 10).Value = 0 Then
                newStatus = "Rupture de stock"
            ElseIf wsData.Cells(lastRowData, 10).Value <= 10 Then
                newStatus = "faible stock"
            Else
                newStatus = "En stock"
            End If
            wsData.Cells(lastRowData, 12).Value = newStatus ' Update status
            
            ' Color the status cell
            Select Case newStatus
                Case "Rupture de stock"
                    wsData.Cells(lastRowData, 12).Interior.Color = RGB(255, 0, 0) ' Red
                Case "faible stock"
                    wsData.Cells(lastRowData, 12).Interior.Color = RGB(255, 255, 0) ' Yellow
                Case "En stock"
                    wsData.Cells(lastRowData, 12).Interior.Color = RGB(0, 255, 0) ' Green
            End Select
        End If
    End If
End Sub

