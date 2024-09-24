Sub RemoveColumnsByName()
    Dim columnsToKeep As String
    Dim lastColumn As Integer
    Dim i As Integer
    Dim keepColumn As Boolean
    Dim colName As String

    ' Define which columns to keep (list column headers here)
    columnsToKeep = "PRODUCT_SKU_NO,PRODUCT_NAME,LONG_DESC,SUBCATEGORY,BRAND_NAME,BC_INDIGENOUS_PRODUCT,SPECIES,STRAIN,EXTRACTION_PROCESS,PACKAGING_MATERIAL,GROWING_METHOD,TERPENE_1_TYPE,TERPENE_2_TYPE,TERPENE_3_TYPE"

    ' Find the last used column in the first row
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

    ' Loop through columns from the last to the first
    For i = lastColumn To 1 Step -1
        ' Reset keep flag
        keepColumn = False
        ' Get column name from the first row
        colName = Cells(1, i).Value

        ' Check if the column name is in the list of columns to keep
        If InStr(1, columnsToKeep, colName) > 0 Then
            keepColumn = True
        End If
        
        ' If the column is not in the list, delete it
        If Not keepColumn Then
            Columns(i).Delete
        End If
    Next i
End Sub

