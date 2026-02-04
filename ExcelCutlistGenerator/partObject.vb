Public Class partObject
    Dim _partNumber As String
    Dim _qty As Integer
    Dim _length As Double
    Dim _material As String
    Dim _remainingQty As Integer
    Dim _duplicateIdentical As Boolean = False
    Dim _duplicateDifLength As Boolean = False
    Dim _duplicateDifMaterial As Boolean = False
    Dim _warningList As New List(Of String)
    'Dim _end1Angle As Double = 0.0
    'Dim _end2Angle As Double = 0.0

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



End Class
