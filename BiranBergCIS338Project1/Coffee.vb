Public Class Coffee

    Private m_name As String
    Private m_price As Double

    Public Sub New()

    End Sub

    Public Sub New(ByVal name As String, ByVal price As Double)
        Me.m_name = name
        Me.m_price = price
    End Sub

    'sets name for coffee
    Public Property name As String
        Get
            Return m_name
        End Get
        Set(value As String)
            m_name = value
        End Set
    End Property

    'sets up price value for coffee
    Public Property price As Double
        Get
            Return m_price
        End Get
        Set(value As Double)
            m_price = value
        End Set
    End Property

    'gets description for a coffee
    Public Function getCDescription()
        Dim result As String
        result = "Coffee Name: " & m_name & vbCrLf & "Price: " & m_price.ToString & vbCrLf


        Return result


    End Function

End Class
