Imports ClosedXML.Excel
Imports System.IO

Module SawCodeOutput1
    '********STRATEGY*******    
    'identify parts that have angles that could nest together and would benefit from being output together in the same code line
    'identify parts that can only get 1 part1 per line of code (parts that don't have the same left and right angle)
    'put square cuts to the end of the list so if the operator has to manully cut these.
    'identify if a part1 has any warnings and if so, exclude it from being output in the code and write the warning to the excel sheet for the operator to reference.
    'generate a collection of parts that can share a line of code based on the angles of the parts, this will be used to determine which parts can be output together in the same line of code.
    'convert angles form +- to L or R (- angles = L, + angles = R)
    'find the length values that need to be input into the saw code based on the bottom of the part1
    'output code to excel sheet referencing the individuals stick "OutputStartRow" property to determine where to write the code for each stick.
    '***********************
    Public Sub PrintSawCode1(nestCollection As Collection, stockList As List(Of stockObject), Optional ws As IXLWorksheet = Nothing)
        For Each stock In stockList
            Dim stickIndex As Integer = 1
            For Each stick As StickObject In nestCollection.Item(stock.Name)("Orientation 1")
                'find the parts that can share a line of code
                'Dim sharedLineParts As List(Of Tuple(Of partObject, partObject)) = FindSharedLineParts1(stick.PartList)
                Dim sharedLinePartsCollection As Collection = FindSharedLineParts2(stick.PartList)

                'sort the parts so that square cuts are at the end of the list
                Dim sortedParts As List(Of partObject) = MoveSquareCutsToEnd(stick.PartList)

                'set the quantity of parts that are to be added to the cutlist for the current stick
                For Each part As partObject In sortedParts
                    part.SetQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
                Next

                'output the saw code
                WriteSawCode(sortedParts, sharedLinePartsCollection, stick.OutputStartRow, stick.Orientation)
            Next

            For Each stick As StickObject In nestCollection.Item(stock.Name)("Orientation 2")
                'find the parts that can share a line of code
                'Dim sharedLineParts As List(Of Tuple(Of partObject, partObject)) = FindSharedLineParts1(stick.PartList)
                Dim sharedLinePartsCollection As Collection = FindSharedLineParts2(stick.PartList)

                'sort the parts so that square cuts are at the end of the list
                Dim sortedParts As List(Of partObject) = MoveSquareCutsToEnd(stick.PartList)

                'set the quantity of parts that are to be added to the cutlist for the current stick
                For Each part As partObject In sortedParts
                    part.SetQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
                Next

                'output the saw code
                WriteSawCode(sortedParts, sharedLinePartsCollection, stick.OutputStartRow, stick.Orientation)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Used for writing individual stick saw code to excel
    ''' </summary>
    ''' <param name="nestCollection"></param>
    ''' <param name="stockList"></param>
    ''' <param name="ws"></param>
    Public Sub StickSawCode1ToXL(nestCollection As Collection, stockList As List(Of stockObject), stickStartRow As Integer, stick As StickObject, ws As IXLWorksheet)
        'find the parts that can share a line of code
        'Dim sharedLineParts As List(Of Tuple(Of partObject, partObject)) = FindSharedLineParts1(stick.PartList)
        Dim sharedLinePartsCollection As Collection = FindSharedLineParts2(stick.PartList)

        'sort the parts so that square cuts are at the end of the list
        Dim sortedParts As List(Of partObject) = MoveSquareCutsToEnd(stick.PartList)

        'set the quantity of parts that are to be added to the cutlist for the current stick
        For Each part As partObject In sortedParts
            part.SetQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
        Next

        'output the saw code
        WriteSawCode(sortedParts, sharedLinePartsCollection, stickStartRow, stick.Orientation, ws)
    End Sub

    ''' <summary>
    ''' Find parts that would benefit from being output together in the same line of code because they have opposite angles that would allow them to nest together.
    ''' In some instances parts can nest with multiple other parts
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <returns>List(Of Tuple(Of partObject,partObject) List of tuples of parts that would benifit from sharing a line of code</returns>
    Private Function FindSharedLineParts1(partList As List(Of partObject)) As List(Of Tuple(Of partObject, partObject))
        'only parts with at least one non-0 (or non-90 in the case of code output) angle can benefit from being nested together. 
        'the only time nesting 2 parts together would benefit the operator is part1 1 has the same left angle as part1 2's right angle and part1 1's right angle is the same as part1 2's left angle. 

        Dim sharedLinePartList As New List(Of Tuple(Of partObject, partObject))

        'loop through all the parts and find pairs of parts that have opposite angles and the same material type
        For partIndex As Integer = 0 To partList.Count - 1
            Dim part1 As partObject = partList(partIndex)
            If part1.LeftAngle <> 0 Or part1.RightAngle <> 0 Then 'only consider parts that have at least one angle that is not 0
                For secondPartIndex As Integer = partIndex To partList.Count - 1 'don't compare parts that have already been compared, so start the second loop at the index of the first loop
                    Dim part2 As partObject = partList(secondPartIndex)
                    If part2.PartNumber <> part1.PartNumber Then 'don't compare the same part1 to itself, only compare parts with the same material type
                        If part1.LeftAngle = part2.RightAngle And part1.RightAngle = part2.LeftAngle Then
                            'found a pair of parts that can nest together, add them to the list as a tuple
                            sharedLinePartList.Add(New Tuple(Of partObject, partObject)(part1, part2))
                        End If
                    End If
                Next
            End If
            partIndex += 1
        Next
        Return sharedLinePartList
    End Function

    ''' <summary>
    ''' Get a collection that contains a list of parts that can share a line of code based on the angles of the parts.
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <returns>Collection of a lists(Of partObjects) that can share the same line of code with the main part1 (indexed by main part1 partnumber)</returns>
    Private Function FindSharedLineParts2(partList As List(Of partObject)) As Collection
        Dim sharedLinePartCollection As New Collection
        For Each part1 As partObject In partList
            If part1.LeftAngle <> 0 Or part1.RightAngle <> 0 Then 'only consider parts that have at least one angle that is not 0
                For Each part2 As partObject In partList
                    If part2.PartNumber <> part1.PartNumber Then 'don't compare the same part1 to itself, only compare parts with the same material type
                        If part1.LeftAngle = part2.RightAngle And part1.RightAngle = part2.LeftAngle Then
                            'found a pair of parts that can nest together, add them to the collection
                            sharedLinePartCollection.Add(part2, part1.PartNumber)
                        End If
                    End If
                Next
            End If
        Next

        Return sharedLinePartCollection
    End Function

    Private Function MoveSquareCutsToEnd(partsList As List(Of partObject)) As List(Of partObject)
        'move square cuts to the end of the list so if the operator has to manually cut these, they are all together at the end of the cutlist. 
        Dim sortedPartsList As New List(Of partObject)
        For Each part As partObject In partsList
            If part.LeftAngle <> 0 Or part.RightAngle <> 0 Then
                sortedPartsList.Add(part)
            End If
        Next
        For Each part As partObject In partsList
            If part.LeftAngle = 0 And part.RightAngle = 0 Then
                sortedPartsList.Add(part)
            End If
        Next
        Return sortedPartsList
    End Function

    ''' <summary>
    ''' Generates the cutlist for code output based on the parts list and the shared line parts list.
    ''' parts that have warnings will be excluded from the cutlist and the warning will be written to the excel sheet for the operator to reference.
    ''' </summary>
    ''' <param name="partsList"></param>
    ''' <param name="sharedLineParts"></param>
    ''' <returns>List(Of Tuple(Of partObject,partObject)) The represents parts on the same line of code, if a part1 can be cut multiple times on a single line it will be assigned (part1, part1), if it can only be cut once per line it will be (part1, Nothing)
    ''' if the line contains 2 parts that can share a line of code it will be (part1, part2) where part1 and part2 are the parts that can share a line of code.
    ''' </returns>
    Private Sub WriteSawCode(sortedPartsList As List(Of partObject), sharedLinePartsCol As Collection, startRow As Integer, cutOrientation As Integer, Optional ws As IXLWorksheet = Nothing, Optional stick As StickObject = Nothing)
        Dim currentRow As Integer = startRow
        If Globals.enableOutput Then
            'set the first row to the stick output start row from the nesting output
            currentRow = startRow

            Dim a1Range As IXLRange = ws.Range(ws.Cell(currentRow, 4), ws.Cell(currentRow, 5)).Merge
            a1Range.Value = "A1"
            a1Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
            a1Range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
            a1Range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
            a1Range.Style.Fill.BackgroundColor = XLColor.LightGray

            Dim l1Cell As IXLCell = ws.Cell(currentRow, 6)
            l1Cell.Value = "L1"
            FormatNestXLHeaderCell(ws, l1Cell)

            Dim a2Range As IXLRange = ws.Range(ws.Cell(currentRow, 7), ws.Cell(currentRow, 8)).Merge
            a2Range.Value = "A2"
            a2Range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
            a2Range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
            a2Range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
            a2Range.Style.Fill.BackgroundColor = XLColor.LightGray

            Dim l2Cell As IXLCell = ws.Cell(currentRow, 9)
            l2Cell.Value = "L2"
            FormatNestXLHeaderCell(ws, l2Cell)

            Dim qtyCutCell As IXLCell = ws.Cell(currentRow, 10)
            qtyCutCell.Value = "qty cut"
            FormatNestXLHeaderCell(ws, qtyCutCell)

            Dim qtyRemCell As IXLCell = ws.Cell(currentRow, 11)
            qtyRemCell.Value = "qty rem"
            FormatNestXLHeaderCell(ws, qtyRemCell)

            currentRow += 1
        Else
            cw(" a1 | l1 | a2 | l2 | qty cut | qty rem.", 1, 0)

        End If


        For Each part1 As partObject In sortedPartsList
            'check if the part1 still has a quantity that needs to be added to the cutlist, if it doesn't, skip it and move on to the next part1 in the list
            If part1.RemainingQtyOnStick = 0 Then
                Continue For
            End If

            'list of parts that are excluded from the code output because they have warnings
            Dim excludedPartList As New List(Of String)
            'see if the part1 has any warnings, if it does, skip it and write the warning to the excel sheet for the operator to reference            
            If part1.WarningList.Count > 0 Then
                excludedPartList.Add(part1.PartNumber)
                Continue For 'skip the part1 and move on to the next part1 in the list, this part1 will not be added to the cutlist
            End If

            'see if the part1 can share a line of code with another part1.
            Dim hasSharedParts As Boolean = True
            Dim listOfSharedParts As New List(Of partObject)
            Try
                listOfSharedParts = sharedLinePartsCol.Item(part1.PartNumber) 'try to get the part1 from the shared line parts collection, if it exists, this part1 can share a line of code with another part1
            Catch ex As Exception
                hasSharedParts = False 'if the part1 is not in the shared line parts collection, it cannot share a line of code with another part1
            End Try

            'write the code for part1 (and any parts that it can share a line with) until all of part1's that need to be on the stick are done
            While part1.RemainingQtyOnStick > 0
                If hasSharedParts Then
                    Dim sharedPartsRemaining As Boolean = False
                    For Each part2 As partObject In listOfSharedParts
                        'if we have added all of the part1's that are needed before we added all of the available shared line parts, exit the loop and move on to a new part1
                        If part1.RemainingQtyOnStick > 0 Then
                            Exit For
                        End If

                        'add the part1 and all of its shared line parts to the same line of code in the excel sheet, make sure to update the quantity of each part1 as you add it to the cutlist
                        If part2.RemainingQtyOnStick > 0 Then
                            sharedPartsRemaining = True
                            'The part1 that has the highest remaining quantity should go first because you can get an extra cut
                            If part1.RemainingQtyOnStick = part2.RemainingQtyOnStick Then
                                'if they have the same remaining quantity, it doesn't matter which one goes first in the code output
                                'put all of the remaining of each part1 on the current line
                                OutputCodeLine(cutOrientation, part1, part1.RemainingQtyOnStick, part2, part2.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
                                currentRow += 1
                                'reduce the quantity remaining on the stick for both parts by the remaining quantity
                                part1.ReduceQtyOnStick(part1.RemainingQtyOnStick)
                                part2.ReduceQtyOnStick(part2.RemainingQtyOnStick)

                            ElseIf part1.RemainingQtyOnStick > part2.RemainingQtyOnStick Then
                                'part1 1 has more remaining so it should be cut first
                                'put the remaining quantity of part2 + 1 of part1 on the current line along with the remaining quantity of part2 of part2 
                                OutputCodeLine(cutOrientation, part1, part2.RemainingQtyOnStick + 1, part2, part2.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
                                currentRow += 1

                                part1.ReduceQtyOnStick(part2.RemainingQtyOnStick + 1) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                                part2.ReduceQtyOnStick(part2.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part2 by the quantity that was just added to the cutlist
                            Else
                                'part1 2 has more remaining than part1, so part2 should go first in the code output because you can get an extra cut of part2
                                'put the remaining quantity of part1 + 1 of part2 on the current line along with the remaining quantity of part1 of part1
                                OutputCodeLine(cutOrientation, part2, part1.RemainingQtyOnStick + 1, part1, part1.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
                                currentRow += 1

                                part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                                part2.ReduceQtyOnStick(part1.RemainingQtyOnStick + 1) 'reduce the quantity remaining on the stick for part2 by the quantity that was just added to the cutlist
                            End If
                        End If
                    Next
                Else
                    'the part1 cannot share a line of code with another part1

                    'if the part1 has the same angle on each side, then it can be cut multiple times on the same line of code.
                    If part1.LeftAngle = part1.RightAngle Then
                        'the part1 can be cut multiple times on the same line of code.
                        If part1.RemainingQty Mod 2 = 0 Then
                            'there are an even number of parts remaining
                            OutputCodeLine(cutOrientation, part1, part1.RemainingQtyOnStick / 2, part1, part1.RemainingQtyOnStick / 2, ws:=ws, curRow:=currentRow) 'put 2 of the same part1 on the current line of code and reduce the quantity remaining on the stick by 2
                            currentRow += 1

                            part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                        Else
                            'there are an odd number of parts remaining, put 1 part1 on the current line and then the rest can be put on lines with 2 parts per line
                            OutputCodeLine(cutOrientation, part1, Math.Ceiling(part1.RemainingQtyOnStick / 2), part1, Math.Floor(part1.RemainingQtyOnStick / 2), ws:=ws, curRow:=currentRow) 'put 1 of the part1 on the current line of code and reduce the quantity remaining on the stick by 1))
                            currentRow += 1
                            part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                        End If
                    Else
                        OutputCodeLine(cutOrientation, part1, part1.RemainingQtyOnStick, ws:=ws, curRow:=currentRow) 'put 1 of the part1 on the current line of code and reduce the quantity remaining on the stick by 1
                        currentRow += 1
                        part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                    End If
                End If
            End While

        Next
    End Sub

    ''' <summary>
    ''' Writes a line of saw code to the excel sheet (console for testing) 
    ''' If a second part1 isn't provided, it indicates that only one part1 can be on each line of code.
    ''' </summary>
    ''' <param name="firstPartOnLine"></param>
    ''' <param name="qty1"></param>
    ''' <param name="secondPartOnLine"></param>
    ''' <param name="qty2"></param>
    Private Sub OutputCodeLine(cutOrientation As Integer, firstPartOnLine As partObject, qty1 As Integer, Optional secondPartOnLine As partObject = Nothing, Optional qty2 As Integer = 0,
                               Optional ws As IXLWorksheet = Nothing, Optional curRow As Integer = Nothing)
        Dim a1 As String
        Dim a1Dir As String
        Dim l1 As String
        Dim a2 As String
        Dim a2Dir As String
        Dim l2 As String
        Dim qtyRemaining

        'prepare the angle data
        'angle data can be determined off of first part only because for the parts to both be on the same line, the left angle of second part needs to be the right angle of the first part
        'and the left angle of the first part needs to be the right angle of the second part
        Dim angle1Data As (String, String) = ConvertAngleToCode(firstPartOnLine.LeftAngle)
        a1 = angle1Data.Item1
        a1Dir = angle1Data.Item2
        Dim angle2Data As (String, String) = ConvertAngleToCode(firstPartOnLine.RightAngle)
        a2 = angle2Data.Item1
        a2Dir = angle2Data.Item2

        'check if there are more than one part1 being cut per line
        If secondPartOnLine Is Nothing Then
            'there are only one part1 per line
            If qty1 = 1 Then
                'there is only one part on this line because there only needs to be one cut
                l1 = FindL1_L2(cutOrientation, firstPartOnLine).Item1
                'go ahead and assign it a value even though we won't use it
                l2 = 0.1
                qtyRemaining = 1
            Else
                'there is only one part on this line because it can't nest with other parts or itself
                'we need to figure out the drop cut dimension
                l1 = FindL1_L2(cutOrientation, firstPartOnLine).Item1
                If cutOrientation = 1 Then
                    l2 = Str(Math.Round(0.1 + firstPartOnLine.Stock.Height / Math.Tan(((90 - firstPartOnLine.RightAngle) * Math.PI) / 180), 3))
                ElseIf cutOrientation = 2 Then
                    l2 = Str(Math.Round(0.1 + firstPartOnLine.Stock.Width / Math.Tan(((90 - firstPartOnLine.RightAngle) * Math.PI) / 180), 3) + 0.1)
                Else
                    Throw New Exception("Cut Orientation " & cutOrientation & " sent but only cut orientations 1 and 2 allowed")
                End If

                'the qty that needs to be programmed is 2 times the total quantity but 1 can be subtracted because we don't need the last drop cut
                qtyRemaining = (qty1 * 2) - 1
            End If
        Else
            'there are multiple parts per line
            l1 = firstPartOnLine.GetBottomLength(cutOrientation)
            l2 = secondPartOnLine.GetBottomLength(cutOrientation)

            qtyRemaining = qty1 + qty2


        End If

        If Globals.enableOutput Then
            'print the line to the excel sheet
            Dim a1ValCell As IXLCell = ws.Cell(curRow, 4)
            a1ValCell.Value = a1
            Dim a1DirCell As IXLCell = ws.Cell(curRow, 5)
            a1DirCell.Value = a1Dir
            Dim l1Cell As IXLCell = ws.Cell(curRow, 6)
            l1Cell.Value = l1
            Dim a2ValCell As IXLCell = ws.Cell(curRow, 7)
            a2ValCell.Value = a2
            Dim a2DirCell As IXLCell = ws.Cell(curRow, 8)
            a2DirCell.Value = a2Dir
            Dim l2Cell As IXLCell = ws.Cell(curRow, 9)
            l2Cell.Value = l2
            Dim qtyCutCell As IXLCell = ws.Cell(curRow, 10)
            qtyCutCell.Value = "0"
            Dim qtyRemCell As IXLCell = ws.Cell(curRow, 11)
            qtyRemCell.Value = Str(qtyRemaining)


        Else
            'print the code output to the console
            If secondPartOnLine Is Nothing Then
                cw(a1 & " | " & a1Dir & " | " & l1 & " | " & a2 & " | " & a2Dir & " | " & l2 & " | 0 | " & qtyRemaining & " | " & firstPartOnLine.PartNumber & "x" & qty1)
            Else
                cw(a1 & " | " & a1Dir & " | " & l1 & " | " & a2 & " | " & a2Dir & " | " & l2 & " | 0 | " & qtyRemaining & " | " & firstPartOnLine.PartNumber & "x" & qty1 & " | " & secondPartOnLine.PartNumber & "x" & qty2)
            End If

        End If

    End Sub

    ''' <summary>
    ''' Converts the angle data in the partslist which is based on a +- angle from vertical
    ''' to the angle that the saw uses which is the angle from horizontal and the direction (L or R) which indicates which way the saw blade needs to be tilted to make the cut.
    ''' </summary>
    ''' <param name="angle"></param>
    ''' <returns>(angleValue,angleDirection)</returns>
    Private Function ConvertAngleToCode(angle As Integer) As (String, String)
        Dim angleValue As String
        Dim angleDir As String

        If angle < 0 Then
            angleValue = Str(angle + 90)
            angleDir = "L"
        Else
            angleValue = Str(90 - angle)
            angleDir = "R"
        End If

        Return (angleValue, angleDir)
    End Function

    ''' <summary>
    ''' Finds the length values that need to be input into the saw code. The saw uses the length on the bottom of the part1, but the cutsheet has the maximum length of the part1. 
    ''' </summary>
    ''' <param name="part1"></param>
    ''' <returns></returns>
    Private Function FindL1_L2(cutOrientation As Integer, part1 As partObject, Optional part2 As partObject = Nothing) As (String, String)
        Dim l1 As String
        Dim l2 As String
        If part2 Is Nothing Then
            'there is only one part on this line
            l1 = Str(part1.GetBottomLength(cutOrientation))
            l2 = l1
        Else
            l1 = Str(part1.GetBottomLength(cutOrientation))
            l2 = Str(part2.GetBottomLength(cutOrientation))
        End If

        Return (l1, l2)
    End Function



End Module
