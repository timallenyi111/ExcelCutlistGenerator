Public Class partObject
    Dim _partNumber As String
    Dim _qty As Integer
    Dim _length As Double
    Dim _material As String
    Dim _remainingQty As Integer

    Public Sub New(ByRef partNumber As String, ByRef qty As Integer, ByRef length As Double, ByRef material As String)
        _partNumber = partNumber
        _qty = qty
        _length = length
        _material = material
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

    ReadOnly Property Stock As String
        Get
            Return _material
        End Get
    End Property

    ReadOnly Property RemainingQty As Integer
        Get
            Return _qty
        End Get
    End Property

    ''' <summary>
    ''' This reduces the remaining quantity of the part by the specified amount.
    ''' </summary>
    ''' <param name="amount"></param>
    Public Sub ReduceRemainingQty(ByRef amount As Integer)
        If amount <= _qty Then
            _qty -= amount
        Else
            Throw New Exception("Cannot reduce quantity below zero.")
        End If
    End Sub

End Class
