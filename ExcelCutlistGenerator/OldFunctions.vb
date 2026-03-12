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

    ''' <summary>
    ''' Retrieves the cut angles from the excel sheet and assigns them to the part object. It also determines the cut orientation based on the angles and partStockName used.
    ''' </summary>
    ''' <param name="angleRange"></param>
    ''' <param name="stockUsed"></param>
    ''' <param name="partObj"></param>
    ''' <returns></returns>
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
        'we need to determine which way the partStockName needs to be place in the saw
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
            Return part
        End If

        If LWebAngle <> 0 Or RWebAngle <> 0 Then
            If stockUsed.SubType = "ST" Then
                'the partStockName used is Square tube so cut orientation doesn't matter
                part.LeftAngle = LWebAngle
                part.RightAngle = RWebAngle
                cw("MATERIAL IS SQUARE TUBE -SO- CUT ORIENTATION = 0", 0, 1)
                cw("End 1 Angle = " & part.LeftAngle, 0, 1)
                cw("End 2 Angle = " & part.RightAngle, 0, 1)

            Else
                'the width side of the partStockName will be facing down
                part.CutOrientation = 1
                part.LeftAngle = LWebAngle
                part.RightAngle = RWebAngle
                cw("CUT ORIENTATION = " & part.CutOrientation, 0, 1)
                cw("End 1 Angle = " & part.LeftAngle, 0, 1)
                cw("End 2 Angle = " & part.RightAngle, 0, 1)
            End If

        ElseIf LFlangeAngle <> 0 Or RFlangeAngle <> 0 Then
            If stockUsed.SubType = "ST" Then
                'the partStockName used is Square tube so cut orientation doesn't matter
                part.LeftAngle = LFlangeAngle
                part.RightAngle = RFlangeAngle
                cw("MATERIAL IS SQUARE TUBE -SO- CUT ORIENTATION = 0", 0, 1)
                cw("End 1 Angle = " & part.LeftAngle, 0, 1)
                cw("End 2 Angle = " & part.RightAngle, 0, 1)
            Else
                'the height side of the partStockName will be facing down
                part.CutOrientation = 2
                part.LeftAngle = LFlangeAngle
                part.RightAngle = RFlangeAngle
                cw("CUT ORIENTATION = " & part.CutOrientation, 0, 1)
                cw("End 1 Angle = " & part.LeftAngle, 0, 1)
                cw("End 2 Angle = " & part.RightAngle, 0, 1)
            End If
        End If

        Return part
    End Function
#End Region
End Module
