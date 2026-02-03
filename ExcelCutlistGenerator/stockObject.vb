Public Class stockObject
    Dim _name As String
    Dim _lengthFeet As Integer
    Dim _lengthInches As Double

    Public Sub New(ByRef name As String, ByRef lengthFeet As Integer)
        _name = name
        _lengthFeet = lengthFeet
        _lengthInches = lengthFeet * 12
    End Sub

    ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    ReadOnly Property Length As Double
        Get
            Return _lengthFeet
        End Get
    End Property

    ReadOnly Property LengthInches As Double
        Get
            Return _lengthInches
        End Get
    End Property

End Class
