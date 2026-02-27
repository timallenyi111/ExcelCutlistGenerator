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
            'Globals.inPath = "C:\Users\TimAllen\OneDrive - A & R Erectors, Inc\A&R Working Directory\TIEM\_Plant 5\Assembly Conveyor Platform\Frame Assembly Conveyor Part List.xlsm"
        Else
            Globals.inPath = args(0) ' get the input path from the command line arguments
        End If
        GetOutputPath() ' set the output path based on the input path
        Globals.bladeWidth = 0.1875 'inches

        'Read data from excel sheet
        Dim inputData As (List(Of partObject), List(Of String), List(Of stockObject)) = ReadExcelData()

        Dim partList As List(Of partObject) = inputData.Item1
        Dim uniqueMaterialsList As List(Of String) = inputData.Item2 'list of the unique material types from the part list
        Dim stockList As List(Of stockObject) = inputData.Item3

        cw("Generating Cutlist Nests...", 1)

        'generate the nest list
        Dim nestList As List(Of List(Of StickObject)) = LegacyNesting(partList, stockList, uniqueMaterialsList) 'list of nests, each nest is a list of sticks        

        If enableOutput Then
            WriteOutputNewFile(nestList)
        Else
            'don't write data to a new file to prevent having 
            cw("OUTPUT FILE DISABLED FOR DEBUGGING", 1, 0)
            cw("Change enableOutput to True in global parameters")

            WriteNestToConsole(nestList)
        End If

        cw("Nesting Complete.")
        cw("Press any key to exit.")
        Console.ReadKey()
    End Sub



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

    ''' <summary>
    ''' output message to the console
    ''' </summary>
    ''' <param name="msg"></param> message to output
    ''' <param name="blankLines"></param> number of blank lines before message
    ''' <param name="numTabs"></param> number of tabs before message
    Sub cw(ByRef msg As String, Optional blankLines As Integer = 0, Optional numTabs As Integer = 0)

        Dim spaceIndex = blankLines
        While spaceIndex > 0
            Console.WriteLine(vbCrLf)
            spaceIndex -= 1
        End While

        Dim tabIndex = numTabs
        While tabIndex > 0
            msg = vbTab & msg
            tabIndex -= 1
        End While

        Console.WriteLine(msg)
    End Sub

End Module
