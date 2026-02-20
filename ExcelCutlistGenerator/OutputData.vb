Imports ClosedXML.Excel
Imports System.IO
'Imports Excel = Microsoft.Office.Interop.Excel
Module OutputData

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

                    ws.Cell(currentRow, 1).Value = "Drop(in):"
                    FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 1))
                    ws.Cell(currentRow, 2).Value = stick.RemainingStockLengthInches
                    FormatNestXLHeaderCell(ws, ws.Cell(currentRow, 2))
                    currentRow += 1

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
    Private Function CheckForPrintBreak(ByRef ws As IXLWorksheet, ByRef currentRow As Integer, ByRef lastPageBreak As Integer, ByRef numParts As Integer) As Integer
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

    Private Sub FormatNestXLHeaderCell(ByVal ws As IXLWorksheet, ByRef cell As IXLCell)
        cell.Style.Font.Bold = True
        cell.Style.Fill.BackgroundColor = XLColor.LightGray
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center
    End Sub

    Private Sub InsertPartRow(ByRef ws As IXLWorksheet, ByRef currentRow As Integer, ByRef part As partObject)

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

End Module
