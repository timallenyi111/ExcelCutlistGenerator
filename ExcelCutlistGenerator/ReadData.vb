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
    ''' <param name="uniqueMaterialList"></param>Item 2: a list of every different partStockName used |
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
                cw("Number Of Stock Types: " & stockList.Count, 1, 0)
                ''setup the part list               

                partList = ReadPartData(workbook, stockList)
            End Using
        End Using

        Console.WriteLine(vbCrLf & "Checking for Duplicate Parts...")
        partList = CheckForDuplicateParts(partList)

        Return (partList, uniqueMaterialList, stockList)
    End Function

    Private Function ReadPartData(ByRef workbook As XLWorkbook, ByRef stockList As List(Of stockObject)) As List(Of partObject)
        Dim partList As New List(Of partObject)
        Dim uniqueMaterials As New List(Of String)
        'read data from spreadsheet

        Dim partWS = workbook.Worksheet(1) ' Assuming the first partWS
        Dim partLastRow = partWS.LastRowUsed().RowNumber()

        'these values have been modified to work with the AS partlist directly
        Dim partNumCol = 1
        Dim lenCol = 2
        Dim materialCol = 4
        Dim qtyCol = 3
        Dim LWebAngleCol = 5
        Dim RWebAngleCol = 6
        Dim LFlangeAngleCol = 7
        Dim RFlangeAngleCol = 8

        For row As Integer = 2 To partLastRow ' Assuming the first row is headers
            Dim partNumber = partWS.Cell(row, partNumCol).GetString().Trim()
            Dim lengthFrac = partWS.Cell(row, lenCol).GetString().Trim()
            Dim lengthDecimal = InchFracToDecimal(lengthFrac)
            Dim qty = partWS.Cell(row, qtyCol).GetValue(Of Integer)() ' Assuming quantities are in the third column
            Dim partStockName = partWS.Cell(row, materialCol).GetString().Trim().ToUpper() ' Assuming partStockName types are in the fourth column            

            'find the stock that this part will be cut from 
            Dim stockUsed As stockObject = Nothing
            'cw("Stock Check:", 0, 1)
            For Each stock As stockObject In stockList
                'cw(stock.Name.ToUpper & "*", 0, 1)
                If stock.Name = partStockName Then
                    stockUsed = stock
                    Exit For
                End If
            Next

            If stockUsed Is Nothing Then 'throw an exception if the partStockName isn't in the stocklist
                Throw New Exception("Part Number: " & partNumber & "is looking for stock: " & partStockName & "but it could not be found")
            End If

            Dim partItem As New partObject(partNumber, qty, lengthDecimal, stockUsed)

            'cw(partNumber, 1, 0)
            'cw(partStockName & "*", 0, 1)
            'cw(stockUsed.Name, 0, 1)

            'read in angle data and assign it to the part instance
            Dim LWebAngle As Integer = CInt(Math.Round(partWS.Cell(row, LWebAngleCol).GetValue(Of Double)()))
            Dim RWebAngle As Integer = CInt(Math.Round(partWS.Cell(row, RWebAngleCol).GetValue(Of Double)()))
            Dim LFlangeAngle As Integer = CInt(Math.Round(partWS.Cell(row, LFlangeAngleCol).GetValue(Of Double)()))
            Dim RFlangeAngle As Integer = CInt(Math.Round(partWS.Cell(row, RFlangeAngleCol).GetValue(Of Double)()))
            partItem.SetCutAngleAndOrientation(LWebAngle, RWebAngle, LFlangeAngle, RFlangeAngle)

            partItem.PrintSummary()

            'add the part to the part list and move on to the next row
            partList.Add(partItem)
        Next

        Return partList
    End Function

    Private Function ReadStockList(ByRef workbook As XLWorkbook) As List(Of stockObject)
        Dim stockList As New List(Of stockObject)
        Dim stockworksheet = workbook.Worksheet(2) ' Assuming the second worksheet is for partStockName
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

            newStock.PrintSummary()

            stockList.Add(newStock)
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
                If part.PartNumber = uniquePart.Item2.PartNumber And part.Length = uniquePart.Item2.Length And part.Stock.Name = uniquePart.Item2.Stock.Name Then
                    'identical parts found with same length and partStockName (probably a part used in different assemblies)
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
                ElseIf part.PartNumber = uniquePart.Item2.PartNumber And part.Stock.Name <> uniquePart.Item2.Stock.Name Then
                    'an identical part number with a different partStockName was found (this is probably user error)
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

    Private Function InchFracToDecimal(frac As String) As Double

        If frac.Contains("/") = False Then
            ' No fraction, just return the integer value           
            Return CDbl(frac)
        Else
            Dim inchInt As Integer = CInt(frac.Split(" ")(0))
            Dim inchFrac As String = frac.Split(" ")(1)
            Dim numerator As Integer = CInt(inchFrac.Split("/")(0))
            Dim denominator As Integer = CInt(inchFrac.Split("/")(1))
            Dim inchDecimal As Double = inchInt + (numerator / denominator)
            Return inchDecimal
        End If
    End Function

End Module
