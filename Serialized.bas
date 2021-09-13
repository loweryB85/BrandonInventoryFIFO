Attribute VB_Name = "Serialized"
Sub IfSerialized(ByVal inventoryRow As Double, ByVal scanCount As Double)

    Dim inventory, moves, results As Worksheet
    Set inventory = Worksheets("Inventory")
    Set moves = Worksheets("Pickface Moves")
    Set results = Worksheets("Results")
    
    Dim oldestSerial, serialRow
    
    'If it's a tote
    If (Utility.IsMaster(scanCount) = False) Then
    
        results.Range("N" & scanCount) = "TOTE"
        
        'now we have to look for this tote's master in order to adjust Qty
        'get inventory row of serial
        serialRow = Application.Match(moves.Range("D" & scanCount).Value & moves.Range("J" & scanCount).Value, inventory.Range("M:M"), 0)
        
        'if the serial was found adjust quantity
        If (IsError(serialRow) = False) Then
            
            inventory.Range("F" & serialRow) = inventory.Range("F" & serialRow) - moves.Range("I" & scanCount)
        
        End If
        
    Else
    'if it's a master
        
        'check for existence of pallets with Qty > 0
        oldestSerial = Utility.GetOldest(inventoryRow, scanCount)
        
        If (oldestSerial <> -1) Then
        
            'check FIFO
            If (moves.Range("D" & scanCount).Value & moves.Range("E" & scanCount).Value = inventory.Range("M" & oldestSerial)) Then
            
                results.Range("N" & scanCount) = "FIFO"
            
            Else
                
                results.Range("N" & scanCount) = "NOT FIFO"
            
            End If
            
            'we do not adjust quantities with serialized masters - only with serialized totes
            
        Else
        'No pallets with Qty>0 Existed so we assume FIFO
            
            results.Range("N" & scanCount) = "FIFO"
        
        End If
        
    End If
    

End Sub
