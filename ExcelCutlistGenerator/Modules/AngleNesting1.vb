Imports System.IO
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Spreadsheet

Module AngleNesting1
    ''' <summary>
    ''' Performs the angle nesting algorithm #1
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <param name="stockList"></param>
    ''' <returns>Collection indexed by stock name of Collections indexed by "Orientation 1" and "Orientation 2 of List(of stickObjects)"</returns>
    Public Function CreateAngleNest1(ByRef partList As List(Of partObject), ByRef stockList As List(Of stockObject)) As Collection
        ''''''''''''''''
        'strategy:
        'sort parts into lists by stock and cut orientation in order from longest to shortest
        'nest parts with orientation 1 by using the same strategy used for legacy nesting
        'nest parts with orientation 2 by using the same strategy used for legacy nesting
        'nest parts with orientation 0 by trying to fill in the drops at the end of the orientation 1 and 2 nests first, then nest the remaining parts using the same strategy used for legacy nesting
        '''''''''''''''''

        'sort parts by stock type, cut orientation, and length (longest to shortest)
        Dim sortedParts As Collection = SortParts(partList, stockList)

        'PrintSummaryOfSortedParts(sortedPartsList, stockList)


        'Dim nestList As New List(Of List(Of StickObject)) 'list of nests, each nest is a list of sticks

        'collection of nests, indexed by stock name, each nest is a collection of sticks indexed by orientation "orientation 1" and "orientation 2"
        Dim nestCollection As New Collection

        For Each stock As stockObject In stockList

            Dim stickListOrient1 As New List(Of StickObject)
            Dim stickListOrient2 As New List(Of StickObject)

            'nest parts with orienation 1
            stickListOrient1 = LegacyNest1_2(sortedParts(stock.Name)(1), stickListOrient1)

            'nest parts with orientation 2
            stickListOrient2 = LegacyNest1_2(sortedParts(stock.Name)(2), stickListOrient2)

            'collection of list of stick objects indexed by "orienation 1" and "orientation 2"
            Dim stickCollection As New Collection
            stickCollection = LegacyNest0(sortedParts(stock.Name)(0), stickListOrient1, stickListOrient2)

            nestCollection.Add(stickCollection, stock.Name) 'add the collection of stick lists to the nest collection indexed by stock name
        Next

        PrintSummaryOfNests(nestCollection, stockList)

        Return nestCollection
    End Function

    ''' <summary>
    ''' Returns a collection of lists of lists, sorted first by stock type, then by cut orientation, then by length (longest to shortest). 
    ''' The collection is indexed by stock name, and the lists of lists are indexed by cut orientation.
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <param name="stockList"></param>
    ''' <returns></returns>
    Private Function SortParts(ByRef partList As List(Of partObject), ByRef stockList As List(Of stockObject)) As Collection
        'sort parts into lists by stock and cut orientation in order from longest to shortest
        Dim sortedParts As New Collection
        'first sort parts by stock type
        For Each stock As stockObject In stockList
            Dim stockParts As New List(Of partObject)
            'then sort each stock type into lists by cut orientation
            For Each part As partObject In partList
                If part.Stock.Name = stock.Name Then
                    stockParts.Add(part) 'this is a list of all parts that match the current stock type.
                End If
            Next

            'then sort each stock type by cut orientation and length
            Dim orientedList As New List(Of List(Of partObject))
            For orientation As Integer = 0 To 2
                Dim matchingOrientationParts As New List(Of partObject)
                For Each part As partObject In stockParts
                    If part.CutOrientation = orientation Then
                        matchingOrientationParts.Add(part)
                    End If
                Next

                matchingOrientationParts = SortLongestToShortest(matchingOrientationParts) 'sort the orientation parts by length (longest to shortest)

                orientedList.Add(matchingOrientationParts) 'this is a list of lists, each list contains parts that match the current stock type and cut orientation.)                
            Next
            sortedParts.Add(orientedList, stock.Name) 'this is a collection of lists of lists, each list of lists contains parts that match the current stock type and are sorted by cut orientation and length.
        Next

        Return sortedParts
    End Function

    Private Function SortLongestToShortest(ByRef partList As List(Of partObject)) As List(Of partObject)
        Dim sortedList As New List(Of partObject)
        For Each part As partObject In partList
            If sortedList.Count = 0 Then
                sortedList.Add(part)
            Else
                For index As Integer = 0 To sortedList.Count - 1
                    If part.Length > sortedList(index).Length Then
                        sortedList.Insert(index, part)
                        Exit For
                    ElseIf index = sortedList.Count - 1 Then
                        sortedList.Add(part)
                        Exit For
                    End If
                Next
            End If
        Next
        Return sortedList
    End Function

    ''' <summary>
    ''' For nesting parts with orientation 1 and 2,
    ''' Places parts on sticks starting from the longest part and working down to the shortest, 
    ''' The part list being sent needs to already be sorted by stock type and cut orientation,
    ''' New Sticks are created based on the cut orientation of the part,
    ''' new sticks created with part orientation 0 will be assigned orientation 1
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <param name="stickList"></param>
    ''' <returns></returns>
    Private Function LegacyNest1_2(ByRef partList As List(Of partObject), ByRef stickList As List(Of StickObject)) As List(Of StickObject)
        For Each part As partObject In partList
            If stickList.Count = 0 Then
                'create the first stick with the first part stock and cut orientation
                'if the part has orientation 0 then we make a new stick with orientation 1
                If part.CutOrientation = 0 Then
                    'Dim firstStick As New StickObject(part.Stock, 1)
                    'stickListOrientation1.Add(firstStick)
                    Throw New Exception("Parts with orientation 0 should not be nested in this function.")
                Else
                    Dim firstStick As New StickObject(part.Stock, part.CutOrientation)
                    stickList.Add(firstStick)
                End If
            End If
            While part.RemainingQty > 0
                'used for creating a new stick if needed
                Dim partNested As Boolean = False

                For Each stick As StickObject In stickList
                    'add the part to any stick with enough remaining length
                    If part.Length <= stick.RemainingStockLengthInches + Globals.bladeWidth Then
                        stick.AddPart(part)
                        part.ReduceRemainingQty(1)
                        partNested = True
                        Exit For
                    End If
                Next

                If partNested = False Then
                    'if the part was not nested, create a new stick and add it to the stick list
                    Dim newStick As New StickObject(part.Stock, part.CutOrientation)
                    newStick.AddPart(part)
                    part.ReduceRemainingQty(1)
                    stickList.Add(newStick)
                End If

            End While

        Next
        Return stickList
    End Function

    ''' <summary>
    ''' Used for nesting the parts with orientation 0, it first tries to fill in the drops from the orientation 1 and 2 nests, then nests the remaining parts using the same strategy used for legacy nesting.
    ''' 
    ''' </summary>
    ''' <param name="partList"></param>
    ''' <param name="stickCollection"></param>
    ''' <returns>Collection of list(oF stickObjects) indexed by "Orientation 1" or "Orientation 2"</returns>
    Private Function LegacyNest0(ByRef partList As List(Of partObject), ByRef stickList1 As List(Of StickObject), stickList2 As List(Of StickObject)) As Collection

        'there aren't any sticks created yet so we need to create the first stick
        If stickList1.Count = 0 And stickList2.Count = 0 Then
            Dim firstStick As New StickObject(partList(0).Stock, 1)
            stickList1.Add(firstStick)
        End If

        For Each part As partObject In partList
            While part.RemainingQty > 0
                Dim partNested As Boolean = False
                'try to fit the part in the drops from the orientation 1 sticks first
                For Each stick As StickObject In stickList1
                    If part.Length <= stick.RemainingStockLengthInches + Globals.bladeWidth Then
                        stick.AddPart(part)
                        part.ReduceRemainingQty(1)
                        partNested = True
                        Exit For
                    End If
                Next
                'if the part was not nested in the orientation 1 sticks, try to fit it in the drops from the orientation 2 sticks
                If partNested = False Then
                    For Each stick As StickObject In stickList2
                        If part.Length <= stick.RemainingStockLengthInches + Globals.bladeWidth Then
                            stick.AddPart(part)
                            part.ReduceRemainingQty(1)
                            partNested = True
                            Exit For
                        End If
                    Next
                End If
                'if the part was not nested, create a new stick with orientation 1 and add it to the orientation 1 stick list (we could also add it to the orientation 2 stick list but we have to choose one for organizational purposes)
                If partNested = False Then
                    Dim newStick As New StickObject(part.Stock, 1)
                    newStick.AddPart(part)
                    stickList1.Add(newStick)
                End If
            End While
        Next

        Dim stickCollection As New Collection
        stickCollection.Add(stickList1, "Orientation 1")
        stickCollection.Add(stickList2, "Orientation 2")

        Return stickCollection
    End Function

    Private Sub PrintSummaryOfNests(ByRef nestCollection As Collection, ByRef stockList As List(Of stockObject))

        For Each stock As stockObject In stockList
            Dim stickListOrientation1 As List(Of StickObject) = nestCollection(stock.Name)("Orientation 1")
            Dim stickCount As Integer = 1
            For Each stick As StickObject In stickListOrientation1
                cw(stock.Name & " Stick " & stickCount & ":", 1, 0)
                For Each part As partObject In stick.PartList
                    cw("Part " & part.PartNumber & " Length: " & part.Length & " Orientation: " & part.CutOrientation, 0, 1)
                Next
                cw("Remaining Length: " & stick.RemainingStockLengthInches, 0, 0)
                stickCount += 1
            Next
            Dim stickListOrientation2 As List(Of StickObject) = nestCollection(stock.Name)("Orientation 2")
            For Each stick As StickObject In stickListOrientation2
                cw(stock.Name & " Stick " & stickCount & ":", 1, 0)
                For Each part As partObject In stick.PartList
                    cw("Part " & part.PartNumber & " Length: " & part.Length & " Orientation: " & part.CutOrientation, 0, 1)
                Next
                cw("Remaining Length: " & stick.RemainingStockLengthInches, 0, 0)
                stickCount += 1
            Next

        Next

    End Sub

    Private Sub PrintSummaryOfSortedParts(ByRef sortedParts As Collection, ByRef stockList As List(Of stockObject))
        For Each stock As stockObject In stockList
            Dim orientedList As List(Of List(Of partObject)) = sortedParts(stock.Name)
            For orientation As Integer = 0 To 2
                Dim parts As List(Of partObject) = orientedList(orientation)
                cw(stock.Name & " Orientation " & orientation & ":", 1, 0)
                For Each part As partObject In parts
                    cw(part.PartNumber, 0, 1)
                Next
            Next
        Next
    End Sub

#Region "Data Output"

    Sub WriteOutputAngleNest1(ByRef nestCollection As Collection, ByRef stockList As List(Of stockObject))
        If File.Exists(Globals.outPath) Then
            File.Delete(Globals.outPath)
        End If

        Using wb As New XLWorkbook()
            Dim ws = wb.Worksheets.Add("CutSheet")
            ws.PageSetup.VerticalDpi = 300 'set to 300 dpi
            ws.Rows().Height = 20 ' set the height of all rows to 20 points
            Dim lastPageBreak As Integer = 0
            Dim currentRow As Integer = Nothing
            Dim stickStartRow As Integer = Nothing
            For Each stock In stockList
                Dim stickCount As Integer = 1
                For Each stick As StickObject In nestCollection(stock.Name)("Orientation 1")
                    If ws.LastRowUsed() Is Nothing Then
                        'first entry
                        currentRow = 1
                        stickStartRow = currentRow
                    Else
                        currentRow = ws.LastRowUsed().RowNumber() + 2 'leave a blank row between sticks
                        stickStartRow = currentRow
                        lastPageBreak = CheckForPrintBreak(ws, currentRow, lastPageBreak, stick.PartList.Count) ' check if the next list will fit on the current page
                    End If
                    stick.OutputStartRow = currentRow
                    WriteStickOutput_AngleNest1(ws, stick, currentRow, stickCount)
                    WriteSawCode(stick, ws)
                    stickCount += 1
                Next
                For Each stick As StickObject In nestCollection(stock.Name)("Orientation 2")
                    If ws.LastRowUsed() Is Nothing Then
                        'first entry
                        currentRow = 1
                        stickStartRow = currentRow
                    Else
                        currentRow = ws.LastRowUsed().RowNumber() + 2 'leave a blank row between sticks
                        stickStartRow = currentRow
                        lastPageBreak = CheckForPrintBreak(ws, currentRow, lastPageBreak, stick.PartList.Count) ' check if the next list will fit on the current page
                    End If
                    WriteStickOutput_AngleNest1(ws, stick, currentRow, stickCount)
                    WriteSawCode(stick, ws)
                    stickCount += 1
                Next
            Next

            ws.Columns("A:Z").AdjustToContents()
            wb.SaveAs(Globals.outPath)

        End Using

        'open the new output file
        Process.Start(New ProcessStartInfo With {
            .FileName = Globals.outPath,
            .UseShellExecute = True
        })
    End Sub

    Private Sub WriteStickOutput_AngleNest1(ByRef ws As IXLWorksheet, ByRef stick As StickObject, ByRef currentRow As Integer, ByRef stickCount As Integer)
        'assign the first row of this stick to the stick object so that the code can start on the same line
        stick.OutputStartRow = currentRow
        'write the header for the stick        
        ws.Cell(currentRow, 1).Value = stick.StockName
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 1))
        ws.Cell(currentRow, 2).Value = "Stick " & stickCount
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 2))

        currentRow += 1
        ws.Cell(currentRow, 1).Value = "Orientation: "
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 1))
        ws.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right
        ws.Cell(currentRow, 2).Value = stick.Orientation
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 2))
        ws.Cell(currentRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left

        'move to the next row for data
        currentRow += 1
        'write the part data
        For Each part As partObject In stick.PartList
            InsertPartRow(ws, currentRow, part)
            currentRow += 1
        Next

        ws.Cell(currentRow, 1).Value = "Drop(in):"
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 1))
        ws.Cell(currentRow, 2).Value = stick.RemainingStockLengthInches
        FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 2))
    End Sub

    '********SAW CODE STRATEGY*******    
    'identify parts that have angles that could nest together and would benefit from being output together in the same code line
    'identify parts that can only get 1 part1 per line of code (parts that don't have the same left and right angle)
    'put square cuts to the end of the list so if the operator has to manully cut these.
    'identify if a part1 has any warnings and if so, exclude it from being output in the code and write the warning to the excel sheet for the operator to reference.
    'generate a collection of parts that can share a line of code based on the angles of the parts, this will be used to determine which parts can be output together in the same line of code.
    'convert angles form +- to L or R (- angles = L, + angles = R)
    'find the length values that need to be input into the saw code based on the bottom of the part1
    'output code to excel sheet referencing the individuals stick "OutputStartRow" property to determine where to write the code for each stick.
    '***********************

    ''' <summary>
    ''' FOR DEBUGGING
    ''' Prints the Saw Code to the console
    ''' </summary>
    ''' <param name="nestCollection"></param>
    ''' <param name="stockList"></param>
    ''' <param name="ws"></param>
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
                    part.AddQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
                Next

                'output the saw code
                WriteSawCode(stick)
            Next

            For Each stick As StickObject In nestCollection.Item(stock.Name)("Orientation 2")
                'find the parts that can share a line of code
                'Dim sharedLineParts As List(Of Tuple(Of partObject, partObject)) = FindSharedLineParts1(stick.PartList)
                Dim sharedLinePartsCollection As Collection = FindSharedLineParts2(stick.PartList)

                'sort the parts so that square cuts are at the end of the list
                Dim sortedParts As List(Of partObject) = MoveSquareCutsToEnd(stick.PartList)

                'set the quantity of parts that are to be added to the cutlist for the current stick
                For Each part As partObject In sortedParts
                    part.AddQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
                Next

                'output the saw code
                WriteSawCode(stick)
            Next
        Next
    End Sub


    ''' <summary>
    ''' Generates the cutlist for code output based on the parts list and the shared line parts list.
    ''' parts that have warnings will be excluded from the cutlist and the warning will be written to the excel sheet for the operator to reference.
    ''' </summary>
    ''' <param name="partsList"></param>
    ''' <param name="sharedLineParts"></param>
    ''' <returns>List(Of Tuple(Of partObject,partObject)) The represents parts on the same line of code, if a part1 can be cut multiple times on a single line it will be assigned (part1, part1), if it can only be cut once per line it will be (part1, Nothing)
    ''' if the line contains 2 parts that can share a line of code it will be (part1, part2) where part1 and part2 are the parts that can share a line of code.
    ''' </returns>
    Public Sub WriteSawCode(stick As StickObject, Optional ws As IXLWorksheet = Nothing)
        'find the parts that can share a line of code
        'Dim sharedLineParts As List(Of Tuple(Of partObject, partObject)) = FindSharedLineParts1(stick.PartList)
        Dim sharedLinePartsCol As Collection = FindSharedLineParts2(stick.PartList)

        'sort the parts so that square cuts are at the end of the list
        Dim sortedPartsList As List(Of partObject) = MoveSquareCutsToEnd(stick.PartList)

        'set the quantity of parts that are to be added to the cutlist for the current stick
        For Each part As partObject In sortedPartsList
            part.AddQtyOnStick(1) 'set the quantity that needs to be added to the cutlist for this part1 based on how many times it appears on the stick, this is used for outputting the correct quantity in the code output
        Next


        'set the starting row for the Saw code to the same starting row for the nested part output on the spreadsheet
        Dim currentRow As Integer = stick.OutputStartRow


        If Globals.enableOutput Then
            'set the first row to the stick output start row from the nesting output            

            WriteSawCodeAngleHeader(ws, currentRow, 4, 5, "A1")

            Dim l1Cell As IXLCell = ws.Cell(currentRow, 6)
            l1Cell.Value = "L1"
            FormatNestXLHeaderCell(ws, l1Cell, True)

            WriteSawCodeAngleHeader(ws, currentRow, 7, 8, "A2")

            Dim l2Cell As IXLCell = ws.Cell(currentRow, 9)
            l2Cell.Value = "L2"
            FormatNestXLHeaderCell(ws, l2Cell, True)

            Dim qtyCutCell As IXLCell = ws.Cell(currentRow, 10)
            qtyCutCell.Value = "qty cut"
            FormatNestXLHeaderCell(ws, qtyCutCell, True)

            Dim qtyRemCell As IXLCell = ws.Cell(currentRow, 11)
            qtyRemCell.Value = "qty rem"
            FormatNestXLHeaderCell(ws, qtyRemCell, True)

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
                        'step through each part on the shared parts list
                        'if the current "part2" has 0 qty left, move on to the next one.
                        'once all items on the shared parts list have a qyt of 0, the "sharedPartsRemaining" flag will not be set to true on the next while loop iteration
                        'so the part1 will be treated like a part that has no shared parts.

                        'if we have added all of the part1's that are needed before we added all of the available shared line parts, exit the loop and move on to a new part1
                        If part1.RemainingQtyOnStick > 0 Then
                            Exit For
                        End If

                        'add the part1 and all of its shared line parts to the same line of code in the excel sheet, make sure to update the quantity of each part1 and part 2 as you add it to the cutlist                        

                        If part2.RemainingQtyOnStick > 0 Then
                            sharedPartsRemaining = True
                            'The part1 that has the highest remaining quantity should go first because you can get an extra cut
                            If part1.RemainingQtyOnStick = part2.RemainingQtyOnStick Then
                                'if they have the same remaining quantity, it doesn't matter which one goes first in the code output
                                'put all of the remaining of each part1 on the current line
                                OutputCodeLine(stick.Orientation, part1, part1.RemainingQtyOnStick, part2, part2.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
                                currentRow += 1
                                'reduce the quantity remaining on the stick for both parts by the remaining quantity
                                part1.ReduceQtyOnStick(part1.RemainingQtyOnStick)
                                part2.ReduceQtyOnStick(part2.RemainingQtyOnStick)

                            ElseIf part1.RemainingQtyOnStick > part2.RemainingQtyOnStick Then
                                'part1 1 has more remaining so it should be cut first
                                'put the remaining quantity of part2 + 1 of part1 on the current line along with the remaining quantity of part2 of part2 
                                OutputCodeLine(stick.Orientation, part1, part2.RemainingQtyOnStick + 1, part2, part2.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
                                currentRow += 1

                                part1.ReduceQtyOnStick(part2.RemainingQtyOnStick + 1) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                                part2.ReduceQtyOnStick(part2.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part2 by the quantity that was just added to the cutlist
                            Else
                                'part1 2 has more remaining than part1, so part2 should go first in the code output because you can get an extra cut of part2
                                'put the remaining quantity of part1 + 1 of part2 on the current line along with the remaining quantity of part1 of part1
                                OutputCodeLine(stick.Orientation, part2, part1.RemainingQtyOnStick + 1, part1, part1.RemainingQtyOnStick, ws:=ws, curRow:=currentRow)
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
                            OutputCodeLine(stick.Orientation, part1, part1.RemainingQtyOnStick / 2, part1, part1.RemainingQtyOnStick / 2, ws:=ws, curRow:=currentRow) 'put 2 of the same part1 on the current line of code and reduce the quantity remaining on the stick by 2
                            currentRow += 1

                            part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                        Else
                            'there are an odd number of parts remaining, put 1 part1 on the current line and then the rest can be put on lines with 2 parts per line
                            OutputCodeLine(stick.Orientation, part1, Math.Ceiling(part1.RemainingQtyOnStick / 2), part1, Math.Floor(part1.RemainingQtyOnStick / 2), ws:=ws, curRow:=currentRow) 'put 1 of the part1 on the current line of code and reduce the quantity remaining on the stick by 1))
                            currentRow += 1
                            part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                        End If
                    Else
                        OutputCodeLine(stick.Orientation, part1, part1.RemainingQtyOnStick, ws:=ws, curRow:=currentRow) 'put 1 of the part1 on the current line of code and reduce the quantity remaining on the stick by 1
                        currentRow += 1
                        part1.ReduceQtyOnStick(part1.RemainingQtyOnStick) 'reduce the quantity remaining on the stick for part1 by the quantity that was just added to the cutlist
                    End If
                End If
            End While
        Next

    End Sub

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
                            Try
                                sharedLinePartCollection.Add(part2, part1.PartNumber)
                            Catch ex As Exception
                                Debug.WriteLine("Part " & part1.PartNumber & " already has a shared line part, adding part " & part2.PartNumber & " to the list of shared line parts for part " & part1.PartNumber)
                            End Try
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
                l1 = firstPartOnLine.GetBottomLength(cutOrientation).ToString("0.000")
                'go ahead and assign it a value even though we won't use it
                l2 = 0.1
                qtyRemaining = 1
            Else
                'there is only one part on this line because it can't nest with other parts or itself
                'we need to figure out the drop cut dimension
                'l1 = FindL1_L2(cutOrientation, firstPartOnLine).Item1
                l1 = firstPartOnLine.GetBottomLength(cutOrientation).ToString("0.000")
                l2 = firstPartOnLine.GetDropCutLength(cutOrientation, additionalLength:=0.1).ToString("0.000")
                'If cutOrientation = 1 Then
                '    l2 = Str(Math.Round(0.1 + firstPartOnLine.Stock.Height / Math.Tan(((90 - firstPartOnLine.RightAngle) * Math.PI) / 180), 3))
                'ElseIf cutOrientation = 2 Then
                '    l2 = Str(Math.Round(0.1 + firstPartOnLine.Stock.Width / Math.Tan(((90 - firstPartOnLine.RightAngle) * Math.PI) / 180), 3))
                'Else
                '    Throw New Exception("Cut Orientation " & cutOrientation & " sent but only cut orientations 1 and 2 allowed")
                'End If

                'the qty that needs to be programmed is 2 times the total quantity but 1 can be subtracted because we don't need the last drop cut
                qtyRemaining = (qty1 * 2) - 1
            End If
        Else
            'there are multiple parts per line
            l1 = firstPartOnLine.GetBottomLength(cutOrientation).ToString("0.000")
            l2 = secondPartOnLine.GetBottomLength(cutOrientation).ToString("0.000")

            qtyRemaining = qty1 + qty2


        End If

        If Globals.enableOutput Then
            'print the line to the excel sheet

            WriteCodeCell(ws, curRow, 4, a1)
            WriteCodeCell(ws, curRow, 5, a1Dir)
            WriteCodeCell(ws, curRow, 6, l1)
            WriteCodeCell(ws, curRow, 7, a2)
            WriteCodeCell(ws, curRow, 8, a2Dir)
            WriteCodeCell(ws, curRow, 9, l2)
            WriteCodeCell(ws, curRow, 10, "0")
            WriteCodeCell(ws, curRow, 11, qtyRemaining)

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

    Private Sub WriteSawCodeAngleHeader(ws As IXLWorksheet, currentRow As Integer, column1 As Integer, column2 As Integer, value As String)
        Dim angleRange As IXLRange = ws.Range(ws.Cell(currentRow, column1), ws.Cell(currentRow, column2)).Merge
        angleRange.Value = value
        angleRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        angleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
        angleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
        angleRange.Style.Fill.BackgroundColor = XLColor.LightGray
        angleRange.Style.Font.Bold = True
    End Sub

    Private Sub WriteCodeCell(ws As IXLWorksheet, row As Integer, col As Integer, value As String)
        Dim cell As IXLCell = ws.Cell(row, col)
        cell.Value = value
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
    End Sub
#End Region
End Module
