Imports System
Imports System.IO
Imports System.Runtime.InteropServices
Imports ClosedXML.Excel
Imports Excel = Microsoft.Office.Interop.Excel

Module Program

    Sub Main(args As String())
        Console.WriteLine("Reading Data From File...")

        'used for debugging so I don't have to run the program through excel every time
        If debugMode Then
            Globals.inPath = "C:\Users\TimAllen\source\repos\timallenyi111\ExcelCutlistGenerator\CutlistTestSheet_angleNest.xlsm"
        Else
            Globals.inPath = args(0) ' get the input path from the command line arguments
        End If
        GetOutputPath() ' set the output path based on the input path
        Globals.bladeWidth = 0.1875 'inches

        Dim partList As New List(Of partObject)
        Dim uniqueMaterialsList As New List(Of String) 'list of the unique material types from the part list
        Dim stockList As New List(Of stockObject)

        'read data from spreadsheet
        Using fs = New FileStream(inPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using workbook = New XLWorkbook(fs)
                ''setup the part list
                Dim partData = ReadPartData(workbook)
                partList = partData.Item1
                uniqueMaterialsList = partData.Item2
                '''Setup Stock List'''
                stockList = ReadStockList(workbook)
            End Using
        End Using

        Console.WriteLine(vbCrLf & "Checking for Duplicate Parts...")
        partList = CheckForDuplicateParts(partList)

        Console.WriteLine("Generating Cutlist Nests...")
        'generate the nest list
        Dim nestList As List(Of List(Of StickObject)) = LegacyNesting(partList, stockList, uniqueMaterialsList) 'list of nests, each nest is a list of sticks        


        'output the nest results to the console for testing
        'WriteNestToConsole(nestList)

        'WriteOutputFS(nestList, inPath)
        Console.WriteLine("Writing Output File: " & Globals.outPath)
        WriteOutputNewFile(nestList)

        Console.WriteLine("Nesting Complete.")
        Console.WriteLine("Press any key to exit.")
        Console.ReadKey()
    End Sub

    Function ReadPartData(ByRef workbook As XLWorkbook) As (List(Of partObject), List(Of String))
        Dim partList As New List(Of partObject)
        Dim uniqueMaterials As New List(Of String)
        'read data from spreadsheet

        Dim partworksheet = workbook.Worksheet(1) ' Assuming the first partworksheet
        Dim partLastRow = partworksheet.LastRowUsed().RowNumber()
        For row As Integer = 2 To partLastRow ' Assuming the first row is headers
            Dim partNumber = partworksheet.Cell(row, 1).GetString().Trim()
            Dim lengthFrac = partworksheet.Cell(row, 2).GetString().Trim()
            Dim lengthDecimal = InchFracToDecimal(lengthFrac)
            Dim qty = partworksheet.Cell(row, 3).GetValue(Of Integer)() ' Assuming quantities are in the third column
            Dim material = partworksheet.Cell(row, 4).GetString().Trim() ' Assuming material types are in the fourth column
            If uniqueMaterials.Contains(material) = False Then 'add to unique materials list
                uniqueMaterials.Add(material)
            End If
            Dim partItem As New partObject(partNumber, qty, lengthDecimal, material)
            partList.Add(partItem)
        Next

        Return (partList, uniqueMaterials)
    End Function

    Function ReadStockList(ByRef workbook As XLWorkbook) As List(Of stockObject)
        Dim stockList As New List(Of stockObject)
        Dim stockworksheet = workbook.Worksheet(2) ' Assuming the second worksheet is for material
        Dim stockLastRow = stockworksheet.LastRowUsed().RowNumber()

        For row As Integer = 2 To stockLastRow
            Dim stockName = stockworksheet.Cell(row, 1).GetString().Trim()
            Dim stockLength = stockworksheet.Cell(row, 2).GetValue(Of Integer)()
            Dim newStock As New stockObject(stockName, stockLength)
            stockList.Add(newStock)
        Next

        Return stockList
    End Function

    Function InchFracToDecimal(frac As String) As Double
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

    ''' <summary>
    ''' Write the output to an open Excel workbook using the filestream method to avoid file locking issues.
    ''' </summary>
    ''' <param name="nestList"></param>
    ''' <param name="inPath"></param>
    Sub WriteOutputFS(ByRef nestList As List(Of List(Of StickObject)), ByRef inPath As String)
        Dim xlApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing
        Try
            xlApp = GetExcel(inPath)
            wb = xlApp.ActiveWorkbook
            If wb Is Nothing Then
                Throw New Exception("No active workbook found.")
                Return
            End If


            ' Get or create the "Results" sheet
            's = TryCast(GetWorksheetByName(wb, "CutSheet"), Excel.Worksheet)
            ws = TryCast(wb.Worksheets("CutSheet"), Excel.Worksheet)
            If ws Is Nothing Then
                ws = CType(wb.Worksheets.Add(), Excel.Worksheet)
                ws.Name = "CutSheet"
            End If

            ' Clear existing content
            ws.Cells.Clear()

            'write nest data to worksheet
            For Each nest As List(Of StickObject) In nestList


                For Each stick As StickObject In nest
                    Dim currentMaterial As String = stick.StockName
                    Dim currentRow As Integer = Nothing
                    If ws.LastRowUsed() Is Nothing Then
                        'first entry
                        currentRow = 1
                    Else
                        currentRow = ws.LastRowUsed().RowNumber() + 2 'leave a blank row between sticks
                    End If

                    'create the data set for the header
                    Dim header As Object(,) = {
                        {currentMaterial, "Stick " & stick.ID}
                    }
                    'write the header for the stick
                    Dim headerRange As Excel.Range = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 2))
                    headerRange.Value = header

                    'move to the next row for data
                    currentRow = +1
                    'generate the data for the cutlist
                    Dim nestData(,) = New String(0, 0) {}
                    ReDim nestData(1, stick.PartList.Count - 1)

                    Dim index As Integer = 0
                    For Each part As partObject In stick.PartList
                        nestData(0, index) = part.PartNumber
                        nestData(1, index) = part.Length
                        index += 1
                    Next

                    Dim dataRange As Excel.Range = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow + stick.PartList.Count - 1, 2))
                    dataRange.Value = nestData

                    wb.Save()
                    ws.Range("A1").Select()
                Next
            Next


        Catch ex As Exception
            Console.Error.WriteLine("Error writing to open workbook: " & ex.Message)
        Finally

            ' Release COM references you created (in reverse order)

            Marshal.ReleaseComObject(ws)
            Marshal.ReleaseComObject(wb)
            Marshal.ReleaseComObject(xlApp)
        End Try
    End Sub

    Sub WriteOutputNewFile(ByRef nestList As List(Of List(Of StickObject)))
        'delete the existing output file if it exists
        If File.Exists(Globals.outPath) Then
            File.Delete(Globals.outPath)
        End If

        Using wb As New XLWorkbook()
            Dim ws = wb.Worksheets.Add("CutSheet")
            ws.PageSetup.VerticalDpi = 300 'set to 300 dpi
            ws.Rows().Height = 20 ' set the height of all rows to 20 points
            Dim lastPageBreak As Integer = 0
            For Each nest As List(Of StickObject) In nestList
                Dim currentRow As Integer = Nothing

                For Each stick As StickObject In nest
                    Dim currentMaterial As String = stick.StockName

                    If ws.LastRowUsed() Is Nothing Then
                        'first entry
                        currentRow = 1
                    Else
                        currentRow = ws.LastRowUsed().RowNumber() + 2 'leave a blank row between sticks
                        lastPageBreak = CheckForPrintBreak(ws, currentRow, lastPageBreak, stick.PartList.Count) ' check if the next list will fit on the current page
                    End If
                    'write the header for the stick
                    ws.Cell(currentRow, 1).Value = currentMaterial
                    FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 1))
                    ws.Cell(currentRow, 2).Value = "Stick " & stick.ID
                    FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 2))



                    'move to the next row for data
                    currentRow += 1
                    'write the part data
                    For Each part As partObject In stick.PartList
                        InsertPartRow(ws, currentRow, part)
                        currentRow += 1
                    Next
                Next
            Next
            Console.WriteLine("Number of Page Breaks: " & ws.PageSetup.RowBreaks.Count)

            ws.Columns("A:Z").AdjustToContents()
            wb.SaveAs(Globals.outPath)


        End Using

        'open the new output file
        Process.Start(New ProcessStartInfo With {
            .FileName = Globals.outPath,
            .UseShellExecute = True
        })

    End Sub

    ''' <summary>
    ''' Checks if the next set of parts will fit on the current page. If not, adds a page break.
    ''' </summary>
    ''' <param name="ws"></param>
    ''' <param name="currentRow"></param>
    ''' <param name="lastPageBreak"></param>
    ''' <param name="numParts"></param>
    ''' <returns></returns>
    Function CheckForPrintBreak(ByRef ws As IXLWorksheet, ByRef currentRow As Integer, ByRef lastPageBreak As Integer, ByRef numParts As Integer) As Integer
        Dim verticalDPI = ws.PageSetup.VerticalDpi
        Dim verticalPageMargin As Double = ws.PageSetup.Margins.Top + ws.PageSetup.Margins.Bottom
        Dim verticalPrintArea As Double = 11 - verticalPageMargin 'assuming standard 11.5 inch height for now       
        Dim cellHeight As Double = ws.Row(1).Height
        Dim rowBreaks As List(Of Integer) = ws.PageSetup.RowBreaks.ToList()
        Console.WriteLine("")
        'Console.WriteLine("Vertical Print Area (inches): " & verticalPrintArea)
        'Console.WriteLine("Current Row: " & currentRow)
        'Console.WriteLine("Height Since Last Page Break + next set of parts (and 1 header) (inches): " & ((currentRow + numParts + 1 - lastPageBreak) * cellHeight) / 72)

        'we need to check if there are any existing page breaks and adjust the lastPageBreak accordingly
        Console.WriteLine("Last Page Break Inserted: " & lastPageBreak)
        Console.WriteLine("Current Row: " & currentRow)
        Console.WriteLine("Current Row Breaks:")
        For Each break As Integer In rowBreaks
            Console.WriteLine(vbTab & break)

            If break > lastPageBreak And break < currentRow Then
                ws.PageSetup.RowBreaks.Remove(break)
            End If
        Next

        'if the height of the next set of parts (and 1 header) exceeds the printable area, add a page break
        If ((currentRow + numParts + 1 - lastPageBreak) * cellHeight) / 72 > verticalPrintArea Then
            Console.WriteLine("Adding Page Break at Row: " & (currentRow - 1))
            ws.PageSetup.AddHorizontalPageBreak(currentRow - 1)
            lastPageBreak = currentRow - 1
        End If
        Return lastPageBreak
    End Function

    Sub FormatNestXLHeaderCell(ByVal ws As IXLWorksheet, ByRef cell As IXLCell)
        cell.Style.Font.Bold = True
        cell.Style.Fill.BackgroundColor = XLColor.LightGray
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
    End Sub

    Sub InsertPartRow(ByRef ws As IXLWorksheet, ByRef currentRow As Integer, ByRef part As partObject)

        ws.Cell(currentRow, 1).Value = part.PartNumber
        ws.Cell(currentRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        ws.Cell(currentRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
        ws.Cell(currentRow, 2).Value = part.Length
        ws.Cell(currentRow, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        ws.Cell(currentRow, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
        'check for duplicate part flags and add warnings if necessary
        If part.WarningList.Count > 0 Then
            Dim colIndex As Integer = 3 'start adding warnings in column C
            For Each warning As String In part.WarningList
                ws.Cell(currentRow, colIndex).Value = warning
                ws.Cell(currentRow, colIndex).Style.Font.FontColor = XLColor.Red
                ws.Cell(currentRow, colIndex).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
                colIndex += 1
            Next
            If part.DuplicateDifMaterial Or part.DuplicateDifLength = True Then
                'this is a part number with different properties so create a strong warning
                ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, colIndex - 1)).Style.Fill.BackgroundColor = XLColor.Yellow
                ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, colIndex - 1)).Style.Font.FontColor = XLColor.Red
                ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, colIndex - 1)).Style.Font.Bold = True
            Else
                'this is a duplicate part number with identical properties so only use a weak warning
                ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, colIndex - 1)).Style.Fill.BackgroundColor = XLColor.LightYellow
            End If

        End If

    End Sub

    Function GetExcel(ByRef outDoc As String) As Excel.Application
        Dim xlApp As Excel.Application = Nothing
        If File.Exists(outDoc) = False Then
            Try
                xlApp = CType(GetObject(, "Excel.Application"), Excel.Application)
            Catch ex As Exception
                Throw New Exception("Excel is not running.")
            End Try
        End If
        Return xlApp
    End Function

    Function GetTargetWorkbook(xlApp As Excel.Application, ByRef inDoc As String) As Excel.Workbook

        Return xlApp.ActiveWorkbook

    End Function

    Sub WriteNestToConsole(ByRef nestList As List(Of List(Of StickObject)))
        For Each nest As List(Of StickObject) In nestList
            For Each stick As StickObject In nest
                Console.WriteLine("Material: " & stick.StockName & " Stick ID: " & stick.ID & ":")
                For Each part As partObject In stick.PartList
                    Console.WriteLine(" - Part Number: " & part.PartNumber)
                Next
                Console.WriteLine("Remaining Length (inches): " & stick.RemainingStockLengthInches)
                Console.WriteLine(vbCrLf)
            Next
        Next

    End Sub

    Function LegacyNesting(ByRef partList As List(Of partObject), ByRef stockList As List(Of stockObject), uniqueMaterialsList As List(Of String)) As List(Of List(Of StickObject))

        Dim nestList As New List(Of List(Of StickObject)) 'list of nests, each nest is a list of sticks

        'nest each material type one at a time
        For Each uniqueStock As String In uniqueMaterialsList
            Dim currentStock As stockObject = Nothing

            'retrieve the current material from the available stock list
            For Each stock As stockObject In stockList
                If stock.Name = uniqueStock Then
                    currentStock = stock
                    Exit For
                End If
            Next

            Dim stickList As New List(Of StickObject)
            Dim sortedPartsList As New List(Of partObject) 'parts sorted by length (longest to shortest) and matching material type

            'sort the parts by length (longest to shortest) and matching material type
            For Each part As partObject In partList
                If part.Stock = uniqueStock Then
                    If sortedPartsList.Count = 0 Then
                        sortedPartsList.Add(part)
                    Else
                        For index As Integer = 0 To sortedPartsList.Count - 1
                            If part.Length > sortedPartsList(index).Length Then
                                sortedPartsList.Insert(index, part)
                                Exit For
                            ElseIf index = sortedPartsList.Count - 1 Then
                                sortedPartsList.Add(part)
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next

            'add the first stick to the sticklist
            stickList.Add(New StickObject(currentStock, 1))

            'nesting algorithm
            For Each part As partObject In sortedPartsList
                'if the part is longer than the stock then throw an error
                If part.Length > currentStock.LengthInches Then
                    Throw New Exception("Part length exceeds stock length. Cannot nest.")
                End If

                While part.RemainingQty > 0
                    Dim partNested As Boolean = False
                    For Each stick As StickObject In stickList
                        If part.Length <= stick.RemainingStockLengthInches Then
                            'i don't know if adding the entire part object is the most efficient way to do this but it works for now
                            stick.AddPart(part)
                            part.ReduceRemainingQty(1)
                            partNested = True
                            Exit For
                        End If
                    Next

                    'there wasn't a stick long enough to add the part.
                    If partNested = False Then
                        'add one here and add the part to it
                        'add another stick and try again
                        Dim newStickID = stickList.Count + 1
                        stickList.Add(New StickObject(currentStock, newStickID))
                        'get the stick that we just added to the list and add the current part to it
                        Dim addedStick As StickObject = stickList(newStickID - 1)
                        addedStick.AddPart(part)
                        part.ReduceRemainingQty(1)
                    End If
                End While
            Next

            'add the completed stick list for that material to nest list and move on to the next material
            nestList.Add(stickList)
        Next

        Return nestList

    End Function

    Sub GetOutputPath()
        Dim newFilePostfix As String = "_Cutlist_"
        'Dim outputPath As String = inPath.Insert(inPath.LastIndexOf("."), newFilePostfix)
        Dim outputPath As String = inPath.Substring(0, Len(inPath) - 5) & newFilePostfix & ".xlsx"
        Globals.outPath = outputPath
    End Sub

    Function CheckForDuplicateParts(ByRef partList As List(Of partObject)) As List(Of partObject)
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

End Module
