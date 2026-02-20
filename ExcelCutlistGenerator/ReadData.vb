Imports System.IO
Imports ClosedXML.Excel
Imports Excel = Microsoft.Office.Interop.Excel
''' <summary>
''' All functions responsible for reading in data and creating parts, part list, and stock list
''' </summary>
Module ReadData

    ''' <summary>
    ''' Reads the parts and available stock from the Excel sheet  
    ''' </summary>
    ''' <returns>
    ''' <param name="partList"></param>Item 1: a list of all parts from excel sheet |
    ''' <param name="uniqueMaterialList"></param>Item 2: a list of every different material used |
    ''' <param name="stockList"></param>Item 3: a list of available stock 
    ''' </returns>
    Function ReadExcelData() As (List(Of partObject), List(Of String), List(Of stockObject))
        Dim partList As New List(Of partObject)
        Dim uniqueMaterialList As New List(Of String)
        Dim stockList As New List(Of stockObject)

        'read data from spreadsheet
        Using fs = New FileStream(inPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using workbook = New XLWorkbook(fs)
                '''Setup Stock List'''
                stockList = ReadStockList(workbook)

                ''setup the part list
                Dim partData = ReadPartData(workbook, stockList)

                partList = partData.Item1
                uniqueMaterialList = partData.Item2

            End Using
        End Using

        Console.WriteLine(vbCrLf & "Checking for Duplicate Parts...")
        partList = CheckForDuplicateParts(partList)

        Return (partList, uniqueMaterialList, stockList)
    End Function

    Private Function ReadPartData(ByRef workbook As XLWorkbook, ByRef stockList As List(Of stockObject)) As (List(Of partObject), List(Of String))
        Dim partList As New List(Of partObject)
        Dim uniqueMaterials As New List(Of String)
        'read data from spreadsheet

        Dim partWS = workbook.Worksheet(1) ' Assuming the first partWS
        Dim partLastRow = partWS.LastRowUsed().RowNumber()

        'these values have been modified to work with the AS partlist directly
        Dim partNumCol = 1
        Dim materialCol = 2
        Dim qtyCol = 3
        Dim lenCol = 4

        For row As Integer = 2 To partLastRow ' Assuming the first row is headers
            Dim partNumber = partWS.Cell(row, partNumCol).GetString().Trim()
            Dim lengthFrac = partWS.Cell(row, lenCol).GetString().Trim()
            Dim lengthDecimal = InchFracToDecimal(lengthFrac)
            Dim qty = partWS.Cell(row, qtyCol).GetValue(Of Integer)() ' Assuming quantities are in the third column
            Dim material = partWS.Cell(row, materialCol).GetString().Trim().ToUpper() ' Assuming material types are in the fourth column
            If uniqueMaterials.Contains(material) = False Then 'add to unique materials list
                uniqueMaterials.Add(material)
            End If

            Dim partItem As New partObject(partNumber, qty, lengthDecimal, material)

            cw(partNumber, 1, 0)
            cw(material & "*", 0, 1)

            'find the stock that this part will be cut from 
            Dim stockUsed As stockObject = Nothing
            cw("Stock Check:", 0, 1)
            For Each stock As stockObject In stockList
                cw(stock.Name.ToUpper & "*", 0, 1)
                If stock.Name = material Then
                    stockUsed = stock
                    Exit For
                End If
            Next

            If stockUsed Is Nothing Then 'throw an exception if the material isn't in the stocklist
                Throw New Exception("The material used for this part was not fouond in the list of available stock")
            End If

            'read in angle date and assign it to the part instance
            cw(stockUsed.Name, 0, 1)
            Dim angleRange As IXLRange = partWS.Range(partWS.Cell(row, 5), partWS.Cell(row, 8)) 'the range of cells containing angle information
            partItem = GetPartAngles(angleRange, stockUsed, partItem)

            'add the part to the part list and move on to the next row
            partList.Add(partItem)
        Next

        Return (partList, uniqueMaterials)
    End Function

    Private Function ReadStockList(ByRef workbook As XLWorkbook) As List(Of stockObject)
        Dim stockList As New List(Of stockObject)
        Dim stockworksheet = workbook.Worksheet(2) ' Assuming the second worksheet is for material
        'Dim stockLastRow = stockworksheet.LastRowUsed().RowNumber()
        Dim stockLastRow = stockworksheet.LastRowUsed(options:=XLCellsUsedOptions.AllFormats).RowNumber()
        For row As Integer = 2 To stockLastRow
            Dim stockName = stockworksheet.Cell(row, 1).GetString().ToUpper().Trim()
            If stockName = "" Then
                'the list is over but the equations in the cells give a false number of used rows
                Exit For
            End If
            Dim stockLength = stockworksheet.Cell(row, 2).GetValue(Of Integer)()
            Dim type = stockworksheet.Cell(row, 3).GetString().Trim()
            Dim height = stockworksheet.Cell(row, 4).GetValue(Of Double)()
            Dim width = stockworksheet.Cell(row, 5).GetValue(Of Double)()

            Dim newStock As New stockObject(stockName, stockLength, type, height, width)

            stockList.Add(newStock)

            cw("Stock: " & stockName, 1)
            cw("Height: " & height, 0, 1)
            cw("Width: " & width, 0, 1)
            cw(type, 0, 1)
            If newStock.SubType IsNot Nothing Then
                cw("Sub-type: " & newStock.SubType, 0, 2)
            End If
        Next

        Return stockList
    End Function

    Private Function CheckForDuplicateParts(ByRef partList As List(Of partObject)) As List(Of partObject)
        Dim uniquePartList As New List(Of (Integer, partObject))
        Dim partIndex As Integer = 0 'used for keeping track of unique parts in the part list so we can go back and add the warning to the original part
        For Each part As partObject In partList
            Dim isDuplicate As Boolean = False
            'search for a duplicate part in the unique part list
            For Each uniquePart As (Integer, partObject) In uniquePartList
                If part.PartNumber = uniquePart.Item2.PartNumber And part.Length = uniquePart.Item2.Length And part.Stock = uniquePart.Item2.Stock Then
                    'identical parts found with same length and material (probably a part used in different assemblies)
                    'assign the duplicate flag to the current part
                    part.DuplicateIdentical = True
                    partList(uniquePart.Item1).DuplicateIdentical = True 'assign the duplicate flag to the original part
                    isDuplicate = True
                    Exit For
                ElseIf part.PartNumber = uniquePart.Item2.PartNumber And part.Length <> uniquePart.Item2.Length Then
                    'an identical part number with a different length was found (this is probably user error)
                    part.DuplicateDifLength = True
                    partList(uniquePart.Item1).DuplicateDifLength = True 'assign the duplicate flag to the original part
                    isDuplicate = True
                    Exit For
                ElseIf part.PartNumber = uniquePart.Item2.PartNumber And part.Stock <> uniquePart.Item2.Stock Then
                    'an identical part number with a different material was found (this is probably user error)
                    part.DuplicateDifMaterial = True
                    partList(uniquePart.Item1).DuplicateDifMaterial = True 'assign the duplicate flag to the original part
                    isDuplicate = True
                    Exit For
                End If
            Next
            If isDuplicate = False Then
                uniquePartList.Add((partIndex, part))
            Else
                'this is a duplicate part, output the warning to the console
                Console.WriteLine(vbCrLf & "Duplicate part found: " & part.PartNumber)
                Console.WriteLine(vbTab & "Identical Length and Material: " & part.DuplicateIdentical)
                Console.WriteLine(vbTab & "Different Length: " & part.DuplicateDifLength)
                Console.WriteLine(vbTab & "Different Material: " & part.DuplicateDifMaterial)
            End If
            partIndex += 1
        Next

        'return part list with adjusted duplicate flags
        Return partList
    End Function

    Private Function GetPartAngles(ByRef angleRange As IXLRange, ByRef stockUsed As stockObject, ByRef partObj As partObject) As partObject
        Dim part As partObject = partObj
        'round angles to integers because we don't need decimal degrees (and because it's easier to work with)
        Dim LWebAngle As Integer = CInt(Math.Round(angleRange.Cell(1, 1).GetValue(Of Double)()))
        Dim RWebAngle As Integer = CInt(Math.Round(angleRange.Cell(1, 2).GetValue(Of Double)()))
        Dim LFlangeAngle As Integer = CInt(Math.Round(angleRange.Cell(1, 3).GetValue(Of Double)()))
        Dim RFlangeAngle As Integer = CInt(Math.Round(angleRange.Cell(1, 4).GetValue(Of Double)()))

        'all angles are 0 so no need to run the res
        If LWebAngle = 0 And RWebAngle = 0 And LFlangeAngle = 0 And RFlangeAngle = 0 Then
            cw("No Angles, Cut Orientation = 0", 0, 1)
            Return part
        End If

#Region "determine cut orientation and angle"
        'we need to determine which way the material needs to be place in the saw
        'this depends on if the angles are on the web or flange
        'angles in the flange columns would reflect an angle cut on the width plane (sitting on height side in saw)
        'angles in the web columns would reflect an angle cut on the height plane (sitting on width side in saw)

        'to address this the part is given a cut orientation
        '0 = either height or width side down 
        '1 = width side down in the saw
        '2 = height side down in the saw
        '0 is the default but won't be used here (except for ST) because we know the part has an angle 

        'example: a wide flange beam has an angle in the flange column, the stock would be sitting on it's "side" in the saw
        '         this beam would have a cut orienation of 2
        '  ***         ***
        '  * *         * *
        '  * *********** *
        '  * *********** *
        '  * *         * *
        '  ***         ***
        '---------------------


        If (LWebAngle <> 0 Or RWebAngle <> 0) And (LFlangeAngle <> 0 Or RFlangeAngle <> 0) Then
            'you can't have angles on both planes for now we will just throw an error
            Throw New Exception(part.PartNumber & " has angles on multiple planes ")
        ElseIf LWebAngle <> 0 Or RWebAngle <> 0 Then
            If stockUsed.SubType = "ST" Then
                'the material used is Square tube so cut orientation doesn't matter
                part.End1Angle = LWebAngle
                part.End2Angle = RWebAngle
                cw("MATERIAL IS SQUARE TUBE -SO- CUT ORIENTATION = 0", 0, 1)
                cw("End 1 Angle = " & part.End1Angle, 0, 1)
                cw("End 2 Angle = " & part.End2Angle, 0, 1)

            Else
                'the width side of the material will be facing down
                part.CutOrientation = 1
                part.End1Angle = LWebAngle
                part.End2Angle = RWebAngle
                cw("CUT ORIENTATION = " & part.CutOrientation, 0, 1)
                cw("End 1 Angle = " & part.End1Angle, 0, 1)
                cw("End 2 Angle = " & part.End2Angle, 0, 1)
            End If

        ElseIf LFlangeAngle <> 0 Or RFlangeAngle <> 0 Then
            If stockUsed.SubType = "ST" Then
                'the material used is Square tube so cut orientation doesn't matter
                part.End1Angle = LFlangeAngle
                part.End2Angle = RFlangeAngle
                cw("MATERIAL IS SQUARE TUBE -SO- CUT ORIENTATION = 0", 0, 1)
                cw("End 1 Angle = " & part.End1Angle, 0, 1)
                cw("End 2 Angle = " & part.End2Angle, 0, 1)
            Else
                'the height side of the material will be facing down
                part.CutOrientation = 2
                part.End1Angle = LFlangeAngle
                part.End2Angle = RFlangeAngle
                cw("CUT ORIENTATION = " & part.CutOrientation, 0, 1)
                cw("End 1 Angle = " & part.End1Angle, 0, 1)
                cw("End 2 Angle = " & part.End2Angle, 0, 1)
            End If
        End If

#End Region

        Return part
    End Function


End Module
