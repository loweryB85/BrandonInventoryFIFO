Attribute VB_Name = "Nonserialized"
Sub IfNotSerialized(ByVal inventoryRow As Double, ByVal scanCount As Double)

    Dim inventory, moves, results As Worksheet
    Set inventory = Worksheets("Inventory")
    Set moves = Worksheets("Pickface Moves")
    Set results = Worksheets("Results")
    
    Dim oldestSerial, serialRow
    
    
    'do any pallets for current part with qty>0 exist in inventory?
    oldestSerial = Utility.GetOldest(inventoryRow, scanCount)
    
    'If at least one pallet was found with Qty>0
    If (oldestSerial <> -1) Then
    
        'if current scan is oldest serial
        If (moves.Range("D" & scanCount).Value & moves.Range("E" & scanCount).Value = inventory.Range("M" & oldestSerial)) Then
        
            results.Range("N" & scanCount) = "FIFO"
        
        Else
            
            results.Range("N" & scanCount) = "NOT FIFO"
            
        End If
        
        'now adjust qty - get inventory row of serial
        serialRow = Application.Match(moves.Range("D" & scanCount).Value & moves.Range("E" & scanCount).Value, inventory.Range("M:M"), 0)
        
        'if the serial was found adjust quantity
        If (IsError(serialRow) = False) Then
            
            inventory.Range("F" & serialRow).Value = inventory.Range("F" & serialRow).Value - moves.Range("I" & scanCount).Value
        
        End If
        
    Else
    'all pallets have been depleted or no inventory existed so we will assume FIFO
    
        results.Range("N" & scanCount) = "FIFO"
    
    End If

End Sub


