Attribute VB_Name = "Utility"
'********************** UTILITY MODULE ************************
'
'Contains helper subroutines/functions not directly associated with
'the main module.
'
'**************************************************************

'concatenate part number and serial

Function Concatenate()

    
    Dim inventory As Worksheet 'declare worksheet variable
    
    Dim counter As Double ' declare counter to be used in for loop
    
    Dim inventoryCount As Double 'declare variable to hold number of instances in inventory data
    
    Set inventory = Worksheets("Inventory") ' assign worksheet variable to worksheet containing inventory data
    
    
    
    'with inventory data
    
    With inventory
    
    inventoryCount = .Range("A1", .Range("A1").End(xlDown)).Rows.count 'count number of instances in inventory
    
    'for each instance in inventory, concatonate part number and serial number to create a more unique serial
    For counter = 2 To inventoryCount
    
    .Range("M" & counter).Value = .Range("A" & counter).Value & .Range("E" & counter).Value
    
    
    'move to next scan
    Next counter
    
    
    
    End With


End Function


'Get number of lines in scan data pull
Function GetScanLines()

    Dim numLines As Double     'number of scans in scan data pull

    Dim moves As Worksheet       'declare a worksheet variable
    
    'assign worksheet variable to worksheet containing pickface move data pull
    Set moves = Worksheets("Pickface Moves")
    
    'with the pickface move data
    With moves
    
    'count the number of rows in the data pull that are not blank
    numLines = .Range("A1", .Range("A1").End(xlDown)).Rows.count
    
    'Debug Code - comment this code out for Release
    'MsgBox "Result: " & numScans
    
    End With
    
    'when returning the value, add 1 to compensate for header row.
    GetScanLines = numLines
    
End Function

'Returns the oldest serial of a part number
Function GetOldest(ByVal pointer As Double, ByVal scanCount As Double)

    'worksheet variables
    Dim inventory, moves As Worksheet
    
    'set target worksheets
    Set inventory = Worksheets("Inventory")
    Set moves = Worksheets("Pickface Moves")

    'Start at the beginning of the subset of Inventory and find the first pallet
    'that has QTY > 0. Since the pallets have already been sorted by creation date, this will give us
    'the oldest pallet that has not been depleted.
    Do While ((inventory.Range("F" & pointer).Value <= 0) And (inventory.Range("A" & pointer).Value = moves.Range("D" & scanCount).Value))
        pointer = pointer + 1
    Loop

    'need to perform a quick check to see if all pallets were depleted
    If (inventory.Range("A" & pointer).Value <> moves.Range("D" & scanCount)) Then
        
        GetOldest = -1
        
    Else
    
        'return the row number of the oldest serial
        GetOldest = pointer
    
    End If
    
       
End Function

'check a scan log to determine if master tag was scanned
Function IsMaster(ByVal pointer As Double)

    'declare worksheet variable and set target
    Dim moves As Worksheet
    Set moves = Worksheets("Pickface Moves")
    
    'variable for return value
    Dim test As Boolean
    
    'if "PFUSERID" is found within the scan location
    If InStr(moves.Range("F" & pointer).Value, "PFUSERID") > 0 Then
        
        'this means we are dealing with a master label
        test = True
        
    Else
    
        'absence of "PFUSERID" substring indicates tote label
        test = False
    
    End If
    
    'return result
    IsMaster = test

End Function

'checks a scan log to determine if the scan was done on a serialized part
Function IsSerialized(ByVal pointer As Double)

    'declare worksheet variable and set target
    Dim moves As Worksheet
    Set moves = Worksheets("Pickface Moves")
    
    'variable for return value
    Dim test As Boolean
    
    'if "Serialized" is found within the scan LogInfo
    If InStr(moves.Range("H" & pointer).Value, "Serialized") > 0 Then
        
        'this means we are dealing with a serialized part
        test = True
        
    Else
    
        'absence of "Serialized" substring indicates a non-serialized part
        test = False
    
    End If
    
    'return result
    IsSerialized = test

End Function

'adjust the quantity of inventory on row "pointer" by the amount provided
Sub AdjustQuantity(ByVal pointer As Double, ByVal scanCount As Double)

   'declare worksheet variable and set target
    Dim inventory, moves As Worksheet
    Set moves = Worksheets("Pickface Moves")
    Set inventory = Worksheets("Inventory")

    'reduce inventory quantity for this row by scan quantity
    inventory.Range("F" & pointer) = inventory.Range("F" & pointer).Value - moves.Range("I" & scanCount).Value

End Sub


