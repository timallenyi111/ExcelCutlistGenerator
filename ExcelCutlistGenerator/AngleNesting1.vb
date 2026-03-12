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

        'PrintSummaryOfSortedParts(sortedParts, stockList)


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
End Module
