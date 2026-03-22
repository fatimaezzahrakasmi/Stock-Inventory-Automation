Attribute VB_Name = "Module1"
Sub AddToTable()
    ' Declare variables for the inputs
    Dim DateValue As String
    Dim ClientName As String
    Dim Classification As String
    Dim Reference As String
    Dim Product As String
    Dim Price As String
    Dim Total As String
    Dim TableObject As ListObject
    Dim NewRow As ListRow

    ' Get the values from the cells
    DateValue = Range("F5").Value
    ClientName = Range("F6").Value
    Classification = Range("F7").Value
    Reference = Range("F8").Value
    Product = Range("F9").Value
    Price = Range("F10").Value
    Total = Range("F11").Value

    ' Reference the table by its name (replace "YourTableName" with the actual table name)
    Set TableObject = ActiveSheet.ListObjects("tableInfo")

    ' Add a new row at the end of the table
    Set NewRow = TableObject.ListRows.Add

    ' Populate the new row with data
    NewRow.Range(1, 7).Value = DateValue         ' seventh column (Date)
    NewRow.Range(1, 6).Value = ClientName        ' sixth column (Client Name)
    NewRow.Range(1, 5).Value = Classification    ' fifth column (Classification)
    NewRow.Range(1, 4).Value = Reference         ' Fourth column (Reference)
    NewRow.Range(1, 3).Value = Product           ' third column (Product)
    NewRow.Range(1, 2).Value = Price             ' second column (Price)
    NewRow.Range(1, 1).Value = Total             ' first column (Total)

    ' Confirmation message
    MsgBox "The data has been added to the table successfully!", vbInformation
End Sub




