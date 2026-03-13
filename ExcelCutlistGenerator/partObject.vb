Public Class partObject
    Dim _partNumber As String
    Dim _qty As Integer
    Dim _qtyOnStick As Integer = 0 'used for output to determine how many of the parts are being cut on the same stick
    Dim _length As Double
    Dim _partStock As stockObject
    Dim _remainingQty As Integer
    Dim _duplicateIdentical As Boolean = False
    Dim _duplicateDifLength As Boolean = False
    Dim _duplicateDifMaterial As Boolean = False
    Dim _warningList As New List(Of String)
    Dim _lAngle As Integer = 0
    Dim _rAngle As Integer = 0
    Dim _cutOrientation As Integer = 0
    Dim _multiPlaneCut As Boolean = False


    Public Sub New(ByRef partNumber As String, ByRef qty As Integer, ByRef length As Double, ByRef stock As stockObject)
        _partNumber = partNumber
        _qty = qty
        _remainingQty = qty
        _length = length
        _partStock = stock
    End Sub
    ReadOnly Property PartNumber As String
        Get
            Return _partNumber
        End Get
    End Property

    ReadOnly Property Qty As Integer
        Get
            Return _qty
        End Get
    End Property

    ReadOnly Property Length As Double
        Get
            Return _length
        End Get
    End Property

    ''' <summary>
    ''' Gets the "bottom length" that is needed for saw code based if the material is being cut in orientation 1 or 2
    ''' </summary>
    ''' <param name="cutOrientation"></param>
    ''' <returns>Bottom cut length rounded to 2 decimal places</returns>
    Public Function GetBottomLength(ByRef cutOrientation As Integer)
        'The amount to be added to the cut length from the left angle x1 = height/tan(left angle)
        Dim x1 As Double = 0
        'The amount to be added to the cut length from the right angle x2 = -height/tan(right angle)
        Dim x2 As Double = 0
        If cutOrientation = 1 Then
            'the material is being cut with tall side up
            If LeftAngle <> 0 Then
                x1 = Stock.Height / Math.Tan(((90 - LeftAngle) * Math.PI) / 180)
            End If
            If RightAngle <> 0 Then
                x2 = -Stock.Height / Math.Tan(((90 - RightAngle) * Math.PI) / 180)
            End If

        ElseIf cutOrientation = 2 Then
            'the material is being cut with the short side up
            If LeftAngle <> 0 Then
                x1 = Stock.Width / Math.Tan(((90 - LeftAngle) * Math.PI) / 180)
            End If
            If RightAngle <> 0 Then
                x2 = -Stock.Width / Math.Tan(((90 - RightAngle) * Math.PI) / 180)
            End If
        Else
            Throw New Exception("Cut Orientation: " & cutOrientation & "sent, but only material oreintations 1 and 2 are allowed")
        End If
        'round the bottom length to 3 decimal places
        Dim _bottomLength As Double = Math.Round(Length + x1 + x2, 3)
        Return _bottomLength
    End Function

    ReadOnly Property Stock As stockObject
        Get
            Return _partStock
        End Get
    End Property

    ReadOnly Property RemainingQty As Integer
        Get
            Return _remainingQty
        End Get
    End Property

    ''' <summary>
    ''' This reduces the remaining quantity of the part by the specified amount.
    ''' </summary>
    ''' <param name="amount"></param>
    Public Sub ReduceRemainingQty(ByRef amount As Integer)
        If amount <= _remainingQty Then
            _remainingQty -= amount
        Else
            Throw New Exception("Cannot reduce quantity below zero.")
        End If
    End Sub

    ReadOnly Property RemainingQtyOnStick As Integer
        Get
            Return _qtyOnStick
        End Get
    End Property

    ''' <summary>
    ''' add to the quantity on stick by the value provided
    ''' </summary>
    ''' <param name="stick"></param>
    Public Sub SetQtyOnStick(ByRef value As Integer)
        _qtyOnStick += value
    End Sub
    ''' <summary>
    ''' 'reduce the quantity of the part on the current stick that still needs to be cut by the specified amount.
    ''' </summary>
    ''' <param name="amount"></param>
    Public Sub ReduceQtyOnStick(ByRef amount As Integer)
        _qtyOnStick -= amount
    End Sub

    ReadOnly Property LeftAngle As Integer
        Get
            Return _lAngle
        End Get
    End Property

    ReadOnly Property RightAngle As Integer
        Get
            Return _rAngle
        End Get
    End Property

    ''' <summary>
    ''' The orienation the stock needs to be cut in the saw |
    ''' 0 doesn't matter |
    ''' 1 is width/short side down |
    ''' 2 is height/long side down 
    ''' </summary>
    ''' <returns></returns>
    Public Property CutOrientation As Integer
        Get
            Return _cutOrientation
        End Get
        Set(value As Integer)
            _cutOrientation = value
        End Set
    End Property

    ReadOnly Property MultiPlaneCut As Boolean
        Get
            Return _multiPlaneCut
        End Get
    End Property

    ''used for displaying warnings on the output cutlist.
#Region "Warning Properties and Subs"
    Public Property DuplicateIdentical As Boolean
        Get
            Return _duplicateIdentical
        End Get
        Set(value As Boolean)
            _duplicateIdentical = value
            If value = True Then
                AddWarning("Duplicate Part Numbers with identical properties")
            End If
        End Set
    End Property

    Public Property DuplicateDifLength As Boolean
        Get
            Return _duplicateDifLength
        End Get
        Set(value As Boolean)
            _duplicateDifLength = value
            If value = True Then
                AddWarning("Duplicate Part Number with Different Length")
            End If
        End Set
    End Property

    Public Property DuplicateDifMaterial As Boolean
        Get
            Return _duplicateDifMaterial
        End Get
        Set(value As Boolean)
            _duplicateDifMaterial = value
            If value = True Then
                AddWarning("Duplicate Part Number with Different Material")
            End If
        End Set
    End Property

    Public ReadOnly Property WarningList As List(Of String)
        Get
            Return _warningList
        End Get
    End Property

    Private Sub AddWarning(warning As String)
        _warningList.Add(warning)
    End Sub

#End Region

    Sub SetCutAngleAndOrientation(LWebAngle As Integer, RWebAngle As Integer, LFlangeAngle As Integer, RFlangeAngle As Integer)
        'all angles are 0 so no need to run the rest
        If LWebAngle = 0 And RWebAngle = 0 And LFlangeAngle = 0 And RFlangeAngle = 0 Then
            cw("No Angles, Cut Orientation = 0", 0, 1)
            Return
        End If

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


        'check for multiplane cuts and assign cut angles
        If (LWebAngle <> 0 Or RWebAngle <> 0) And (LFlangeAngle <> 0 Or RFlangeAngle <> 0) Then
            'check to see if there are cuts on both the web and flange.
            'if there is, then the part can't be programmed so it will be nested without the angles and a warning will be added to the part.
            _multiPlaneCut = True
            AddWarning("Part has angles on multiple planes.")
            Return

        ElseIf LFlangeAngle <> 0 Or RFlangeAngle <> 0 Then
            'the part angles are on the flange/short side
            _lAngle = LFlangeAngle
            _rAngle = RFlangeAngle
            Select Case _partStock.Type
                Case "HSS"
                    If _partStock.SubType = "SQ" Then
                        'square tube can be cut in either orientation so no need to assign one.
                        _cutOrientation = 0
                    Else
                        'rectangle tube needs to be cut with the height down so assign that orientation.
                        _cutOrientation = 2
                    End If
                Case "L"
                    If _partStock.SubType = "EQ" Then
                        'equal leg angle can be cut in either orientation so no need to assign one.
                        _cutOrientation = 0
                    Else
                        'this would represent a cut on the short leg of an unequal leg angle
                        _cutOrientation = 2
                    End If
                Case "W", "C", "S"
                    _cutOrientation = 2
                Case Else
                    'stock types aren't supported so throw an error to alert the user to check the part stock type.
                    Throw New Exception("Stock type" & _partStock.Type & " not supported for angle cuts. Please check the part stock type.")
            End Select
        ElseIf LWebAngle <> 0 Or RWebAngle <> 0 Then
            'the part angles are on the web/tall side
            _lAngle = LWebAngle
            _rAngle = RWebAngle
            Select Case _partStock.Type
                Case "HSS"
                    If _partStock.SubType = "SQ" Then
                        'square tube can be cut in either orientation so no need to assign one.
                        _cutOrientation = 0
                    Else
                        'rectangle tube needs to be cut with the width down so assign that orientation.
                        _cutOrientation = 1
                    End If
                Case "L"
                    If _partStock.SubType = "EQ" Then
                        'equal leg angle can be cut in either orientation so no need to assign one.
                        _cutOrientation = 0
                    Else
                        'this would represent a cut on the long leg of an unequal leg angle
                        _cutOrientation = 1
                    End If
                Case "W", "C", "S"
                    _cutOrientation = 1
                Case Else
                    'stock types aren't supported so throw an error to alert the user to check the part stock type.
                    Throw New Exception("Stock type" & _partStock.Type & " not supported for angle cuts. Please check the part stock type.")
            End Select
        End If

    End Sub

    Sub PrintSummary()
        cw(vbCrLf & "Part Number: " & _partNumber & " Summary:", 1, 0)
        cw("Length: " & _length, 0, 1)
        cw("Quantity: " & _qty, 0, 1)
        cw("Qty Remaining: " & _remainingQty, 0, 1)
        cw("Stock: " & Stock.Name, 0, 1)
        cw("Stock Subtype: " & Stock.SubType, 0, 2)
        'cw("Duplicate Identical: " & part.DuplicateIdentical, 0, 1)
        'cw("Duplicate Different Length: " & part.DuplicateDifLength, 0, 1)
        'cw("Duplicate Different Material: " & part.DuplicateDifMaterial, 0, 1)
        cw("Cut Orientation: " & CutOrientation, 0, 1)
        cw("Left Angle: " & LeftAngle, 0, 1)
        cw("Right Angle: " & RightAngle, 0, 1)
        For Each warning As String In _warningList
            cw("WARNING: " & warning, 0, 1)
        Next
    End Sub

End Class
