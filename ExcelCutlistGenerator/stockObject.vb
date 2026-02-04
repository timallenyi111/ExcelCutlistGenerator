Public Class stockObject
    Dim _name As String
    Dim _lengthFeet As Integer
    Dim _lengthInches As Double
    Dim _type As String
    Dim _height As Double
    Dim _width As Double
    Dim _subType As String = Nothing 'used for classifying square tube v rectangle tube


    Public Sub New(ByRef name As String, ByRef lengthFeet As Integer, ByRef type As String, ByRef height As Double, ByRef width As Double)
        _name = name
        _lengthFeet = lengthFeet
        _lengthInches = lengthFeet * 12
        _type = type
        _height = height
        _width = width
        If type = "HSS" & height = width Then
            _subType = "ST"
        End If
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

    ReadOnly Property Type As String
        Get
            Return _type
        End Get
    End Property

    ReadOnly Property SubType As String
        Get
            Return _subType
        End Get
    End Property

    ReadOnly Property Height As Double
        Get
            Return _height
        End Get
    End Property

    ReadOnly Property Width As Double
        Get
            Return _width
        End Get
    End Property

End Class
