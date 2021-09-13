Attribute VB_Name = "Main"
'------------------------------------------------------------------------------------
'Process Scans Routine
'
'This routine will move down the rows of scan data and look for corresponding entries
'within the "Inventory In Location" to determine whether FIFO process was followed.
'
'This routine assumes that Inventory has been filtered down to solely CMA master labels
'
'------------------------------------------------------------------------------------
Sub ProcessScans()

    'Begin variable declarations
    
    'Define worksheet variables for use in processing
    Dim inventory, moves, results, shifts As Worksheet
    
    'number of lines in scan data
    Dim scanLines As Double
    
    'iteration counter for traversing scan data by line #
    Dim scanCount As Double
    
    'variable for row # of master pallet in Inventory data pull
    Dim inventoryRow
    
    'variable for serial number of master pallet from Inventory data pull
    Dim toFindSerial As String
    
    'variable to store oldest serial during searches
    Dim oldestSerial, testTime As String
    
    'Boolean variables
    Dim fifo As Boolean
    
    'min/max for results page
    Dim earliestScan, latestScan  As Double
    
    'range for earliest/latest scans
    Dim dateRange As Range
    
    'Begin routine
    
    'Set worksheet variables as they will be referenced throughout the sub
    Set inventory = Worksheets("Inventory")
    Set moves = Worksheets("Pickface Moves")
    Set results = Worksheets("Results")
    Set dateRange = moves.Range("G:G")
    Set shifts = Worksheets("Shift Times")
    
    
    'Start by obtaining the number of lines in the scan data pull from WMS
    scanLines = Utility.GetScanLines
    
    'determine earliest and latest scan
    earliestScan = WorksheetFunction.Min(dateRange)
    latestScan = WorksheetFunction.Max(dateRange)
        
    'create search key for determining oldest serial later
    Call Concatenate
    
    'set headers for results page and clear any existing data from previous runs
    
    results.Cells.ClearContents
    
    results.Range("A1:K1").Value = moves.Range("A1:K1").Value
    results.Range("L1").Value = "DLOC"
    results.Range("M1").Value = "Row"
    results.Range("N1").Value = "FIFO"
    results.Range("O1").Value = "Accurate?"
    results.Range("P1").Value = "Earliest Scan"
    results.Range("Q1").Value = "Latest Scan"
    results.Range("R1").Value = "Scan Shift"
    
    '- The main driver of this sub will be the following For loop. -
    'The For loop will be responsible for iterating through the scan data so that
    'each entry can be compared to Inventory to determine FIFO status.
    
    'scanCount will be initialized at 2 so that the header row in the data is skipped
    'and the loop will continue until each line has been examined.
    For scanCount = 2 To scanLines
    
        'check to see if the part number from the current scan exists in inventory
        inventoryRow = Application.Match(moves.Range("D" & scanCount).Value, inventory.Range("A:A"), 0)
            
        'copy data from inventory and moves to results page
        results.Range("A" & scanCount & ":K" & scanCount).Value = moves.Range("A" & scanCount & ":K" & scanCount).Value
        results.Range("P" & scanCount).Value = earliestScan
        results.Range("Q" & scanCount).Value = latestScan
                     
            'check if part is in part lookup, if not set dloc to blank
            If (Not IsError(Application.Index(Sheets("Part Lookup").Range("G:G"), Application.Match(Sheets("Pickface Moves").Range("D" & scanCount), Sheets("Part Lookup").Range("B:B"), 0)))) Then
                    
            'record DLOC of part number in result page - index/match from Part Lookup and check against scan location
            results.Range("L" & scanCount).Value = Application.Index(Sheets("Part Lookup").Range("G:G"), Application.Match(Sheets("Pickface Moves").Range("D" & scanCount), Sheets("Part Lookup").Range("B:B"), 0))
            Else
            results.Range("L" & scanCount).Value = ""
            End If
            
            'check for part number's existence - if IsError returns false, then part was found
            If (IsError(inventoryRow) = False) Then
            
            
            'copy row and dloc from inventory when inventory row exists
            'results.Range("L" & scanCount).Value = inventory.Range("I" & inventoryRow).Value
            results.Range("M" & scanCount).Value = Left(inventory.Range("I" & inventoryRow).Value, 2) & "00"
            
                'part # was found. moving to next step
                If (Utility.IsSerialized(scanCount)) Then
                
                    Call Serialized.IfSerialized(inventoryRow, scanCount)
                    
                Else
                    
                    Call Nonserialized.IfNotSerialized(inventoryRow, scanCount)
                              
                End If
        
            Else
            'part was not found in inventory - need to mark appropriately
            
                'filter out serialized totes
                If (Not (Utility.IsSerialized(scanCount) And (Utility.IsMaster(scanCount) = False))) Then
                       
                    'assume FIFO
                    results.Range("N" & scanCount) = "FIFO"
                
                Else
                
                    results.Range("N" & scanCount) = "TOTE"
                    
                End If
                
            End If
            
            
            
            'compare DLOC to stocked DLOC - use an If statement to skip over serialized master entries
            If (InStr(results.Range("F" & scanCount).Value, "PFUSER") = 0) Then
                
                'if the part is not present in the part lookup table, set the dloc to blank otherwise check if it's the same
                If (results.Range("L" & scanCount) = "") Then
                
                    results.Range("O" & scanCount).Value = ""
                
                Else
                               
       
                'this record is not a serialized master - perform comparison
                If (results.Range("L" & scanCount) = results.Range("F" & scanCount)) Then results.Range("O" & scanCount).Value = "TRUE" Else results.Range("O" & scanCount).Value = "FALSE"
            
                End If
                
            End If
            
        
            'fill shift column based on which day of the week the scan is made. It accounts for long/regular shifts.
            
            Select Case Weekday(results.Range("G" & scanCount))
            
                Case 1 'long sunday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("K2:K9"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("J2:J9"), 0))
                
                Case 2 'regular monday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("B2:B25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("A2:A25"), 0))
                        
                Case 3 'regular tuesday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("B2:B25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("A2:A25"), 0))
                
                Case 4 'regular wednesday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("B2:B25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("A2:A25"), 0))
                
                Case 5 'regular thurday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("B2:B25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("A2:A25"), 0))
                
                Case 6 'long friday
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("E2:E25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("D2:D25"), 0))
                
                Case 7 'long friday( after midnight into saturday)
                        results.Range("R" & scanCount).Value = Application.Index(shifts.Range("H2:H5"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("G2:G5"), 0))
                           
            End Select
            ' this is the original statement for a general case of each shift.
            'results.Range("R" & scanCount).Value = Application.Index(shifts.Range("B2:B25"), Application.Match(CDbl(Format(results.Range("G" & scanCount), "hh")), shifts.Range("A2:A25"), 0))
            
           
            
    
           
        
    'move to next scan
    Next scanCount

End Sub

