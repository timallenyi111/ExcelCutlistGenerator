''' <summary>
''' This module isn't compiled and is for holding old functions for future reference
''' </summary>
Module OldFunctions

    ''' <summary>
    ''' Write the output to an open Excel workbook using the filestream method to avoid file locking issues.
    ''' </summary>
    ''' <param name="nestList"></param>
    ''' <param name="inPath"></param>
    'Sub WriteOutputFS(ByRef nestList As List(Of List(Of StickObject)), ByRef inPath As String)
    '    Dim xlApp As Excel.Application = Nothing
    '    Dim wb As Excel.Workbook = Nothing
    '    Dim ws As Excel.Worksheet = Nothing
    '    Try
    '        xlApp = GetExcel(inPath)
    '        wb = xlApp.ActiveWorkbook
    '        If wb Is Nothing Then
    '            Throw New Exception("No active workbook found.")
    '            Return
    '        End If


    '        ' Get or create the "Results" sheet
    '        's = TryCast(GetWorksheetByName(wb, "CutSheet"), Excel.Worksheet)
    '        ws = TryCast(wb.Worksheets("CutSheet"), Excel.Worksheet)
    '        If ws Is Nothing Then
    '            ws = CType(wb.Worksheets.Add(), Excel.Worksheet)
    '            ws.Name = "CutSheet"
    '        End If

    '        ' Clear existing content
    '        ws.Cells.Clear()

    '        'write nest data to worksheet
    '        For Each nest As List(Of StickObject) In nestList


    '            For Each stick As StickObject In nest
    '                Dim currentMaterial As String = stick.StockName
    '                Dim currentRow As Integer = Nothing
    '                If ws.LastRowUsed() Is Nothing Then
    '                    'first entry
    '                    currentRow = 1
    '                Else
    '                    currentRow = ws.LastRowUsed().RowNumber() + 2 'leave a blank row between sticks
    '                End If

    '                'create the data set for the header
    '                Dim header As Object(,) = {
    '                    {currentMaterial, "Stick " & stick.ID}
    '                }
    '                'write the header for the stick
    '                Dim headerRange As Excel.Range = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 2))
    '                headerRange.Value = header

    '                'move to the next row for data
    '                currentRow = +1
    '                'generate the data for the cutlist
    '                Dim nestData(,) = New String(0, 0) {}
    '                ReDim nestData(1, stick.PartList.Count - 1)

    '                Dim index As Integer = 0
    '                For Each part As partObject In stick.PartList
    '                    nestData(0, index) = part.PartNumber
    '                    nestData(1, index) = part.Length
    '                    index += 1
    '                Next

    '                Dim dataRange As Excel.Range = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow + stick.PartList.Count - 1, 2))
    '                dataRange.Value = nestData

    '                wb.Save()
    '                ws.Range("A1").Select()
    '            Next
    '        Next


    '    Catch ex As Exception
    '        Console.Error.WriteLine("Error writing to open workbook: " & ex.Message)
    '    Finally

    '        ' Release COM references you created (in reverse order)

    '        Marshal.ReleaseComObject(ws)
    '        Marshal.ReleaseComObject(wb)
    '        Marshal.ReleaseComObject(xlApp)
    '    End Try
    'End Sub

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
End Module
