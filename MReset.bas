Attribute VB_Name = "MReset"
Option Explicit


Public Sub ResetWorkbook()
    ResetRawCards
End Sub
Public Sub HardReset()
    ResetWorkbook
    ResetCardDetails
End Sub


Private Sub ClearListObject( _
    ByRef pToClear As ListObject _
)
    Dim lrw_firstRow As ListRow
    Dim var_firstRowFormulas As Variant
    Dim lCol As Long
    
    With pToClear
        Set lrw_firstRow = .ListRows(1)
    End With
    
    With lrw_firstRow
        var_firstRowFormulas = lrw_firstRow.Range.Formula
    End With
    
    For lCol = LBound(var_firstRowFormulas, 2) To UBound(var_firstRowFormulas, 2)
        If Left(var_firstRowFormulas(1, lCol), 1) <> "=" Then
            var_firstRowFormulas(1, lCol) = vbNullString
        End If
    Next lCol
    
    With pToClear
        .DataBodyRange.Clear
        .Resize Union(.HeaderRowRange, lrw_firstRow.Range)
        
        With .ListRows.Add
            .Range.Formula = var_firstRowFormulas
        End With
    End With
    
    Set lrw_firstRow = Nothing
End Sub


Private Sub ResetRawCards()
    ClearListObject SRawCards.LObj_RawCards
End Sub

Private Sub ResetCardDetails()
    ClearListObject SCardDetails.LObj_CardDetails
End Sub
