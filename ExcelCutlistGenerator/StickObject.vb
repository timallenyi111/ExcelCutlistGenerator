''' <summary>
''' Represents a stick of material used in nesting
''' </summary>
Public Class StickObject
    Dim _stock As stockObject
    Dim _stockName As String
    Dim _stockLengthInches As Double
    Dim _remainingStockLengthInches As Double
    Dim _partList As New List(Of partObject)
    Dim _id As Integer
    Dim _bladeWidth As Double
    Dim _orientation As Integer = 1

    Public Sub New(ByRef stock As stockObject, ByRef orientation As Integer, Optional ByRef bladeWidth As Double = 0.1875)
        _stock = stock
        _stockName = stock.Name
        _stockLengthInches = stock.LengthInches
        _remainingStockLengthInches = stock.LengthInches
        _bladeWidth = bladeWidth
        _orientation = orientation
    End Sub

    ReadOnly Property StockName As String
        Get
            Return _stockName
        End Get
    End Property

    ReadOnly Property StockLengthInches As Double
        Get
            Return _stockLengthInches
        End Get
    End Property

    ReadOnly Property RemainingStockLengthInches As Double
        Get
            Return _remainingStockLengthInches
        End Get
    End Property

    ReadOnly Property PartList As List(Of partObject)
        Get
            Return _partList
        End Get
    End Property

    'ReadOnly Property ID As Integer
    '    Get
    '        Return _id
    '    End Get
    'End Property

    ReadOnly Property Orientation As Integer
        Get
            Return _orientation
        End Get
    End Property

    ''' <summary>
    ''' This adds the part to the stick and reduces the remaining stock length by the part lenghth plus blade width.
    ''' </summary>
    ''' <param name="part"></param>
    Public Sub AddPart(ByRef part As partObject)
        If part.Length <= _remainingStockLengthInches Then
            _partList.Add(part)
            _remainingStockLengthInches -= part.Length + _bladeWidth 'account for blade width
        Else
            Throw New Exception("Part length exceeds remaining stock length.")
        End If
    End Sub

End Class
