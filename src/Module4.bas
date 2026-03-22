Attribute VB_Name = "Module4"
Sub AjouteClient()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim clientID As Long
    
    'definir la feuille
    Set ws = ThisWorkbook.Sheets("client")
    
    'trouver la dernier ligne dans le tableau des clients
    If ws.Cells(14, 12).Value = "" Then
        lastRow = 13
    Else
        lastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    End If
    
    'generer un nouvelle ID Client
    clientID = lastRow - 12
    
    'ajouter les informations du client dans la prochaine ligne
    ws.Cells(lastRow + 1, 12).Value = clientID 'ID client
    ws.Cells(lastRow + 1, 11).Value = ws.Range("G6").Value 'nom
    ws.Cells(lastRow + 1, 10).Value = ws.Range("G7").Value 'adresse
    ws.Cells(lastRow + 1, 9).Value = ws.Range("G8").Value 'tele
    ws.Cells(lastRow + 1, 8).Value = ws.Range("G9").Value 'ville
    ws.Cells(lastRow + 1, 7).Value = ws.Range("G10").Value 'email
    ws.Cells(lastRow + 1, 6).Value = Date 'date d'inscription
    
    ws.Range("G6:G10").ClearContents
    
    MsgBox "cleint ajoute avec succes! "

End Sub
