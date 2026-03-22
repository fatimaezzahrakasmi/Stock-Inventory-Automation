Attribute VB_Name = "Module5"
Sub GenererSauvgarderFacture()
    Dim wsFacture As Worksheet
    Dim wsclient As Worksheet
    Dim lastRow As Long
    Dim ClientName As String
    Dim invoiceNumber As String
    Dim folderPath As String
    Dim savePath As String
    Dim currentYear As String
    Dim currentMonth As String
    Dim invoiceCounter As Long
    Dim lastInvoiceNumber As String
    Dim clientRow As Long
    Dim foundClient As Range
    
    ' feuille de la facture
    Set wsFacture = ThisWorkbook.Sheets("facture")
    ' feuille des clients
    Set wsclient = ThisWorkbook.Sheets("client")
    
    ' générer le numéro de la facture
    ' déterminer l'année et le mois actuels
    currentYear = Format(Date, "yyyy")
    currentMonth = Format(Date, "mm")
    
    lastRow = wsFacture.Cells(wsFacture.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 1 And wsFacture.Cells(lastRow, 1).Value <> "" Then
        lastInvoiceNumber = wsFacture.Cells(lastRow, 1).Value
        invoiceCounter = CLng(Right(lastInvoiceNumber, 4)) + 1
    Else
        invoiceCounter = 1
    End If
    
    ' générer le nouveau numéro de facture
    invoiceNumber = "F-" & currentYear & "-" & currentMonth & "-" & Format(invoiceCounter, "0000")
    
    ' assigner le numéro de facture
    wsFacture.Range("A1").Value = invoiceNumber
    
    ' ---- sauvegarder la facture en PDF
    ' récupérer le nom du client dans la facture
    ClientName = wsFacture.Range("B4").Value
    
    ' définir le chemin du dossier pour le client
    folderPath = "C:\Users\pc\Desktop\Factures\" & ClientName & "\"
    
    ' créer le dossier s'il n'existe pas
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' chemin du fichier PDF
    savePath = folderPath & invoiceNumber & ".pdf"
    
    ' sauvegarder en PDF
    wsFacture.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath
    
    ' chercher le client dans la feuille des clients
    Set foundClient = wsclient.Columns(11).Find(ClientName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundClient Is Nothing Then
        clientRow = foundClient.Row
        ' ajouter un lien vers le dossier client dans la feuille des clients
        wsclient.Hyperlinks.Add Anchor:=wsclient.Cells(clientRow, 5), Address:=folderPath, TextToDisplay:="Dossier Factures"
    Else
        MsgBox "Client non trouvé dans la feuille des clients."
    End If
    
    ' effacer le contenu de la facture pour la prochaine entrée
    ' wsFacture.Range("B2").ClearContents
    ' wsFacture.Range("B4:B6").ClearContents
    ' wsFacture.Range("B11:F31").ClearContents
    ' wsFacture.Range("B32").ClearContents
    
    MsgBox "Facture générée, sauvegardée et lien ajouté avec succès !"
    
End Sub

