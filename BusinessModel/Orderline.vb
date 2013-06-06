Public Class Orderline

    Private m_Name As String
    Private m_Qty As Integer
    Private m_Price, m_subTotal As Double

    Public Property Name As String
        Get
            Return m_Name
        End Get
        Set(ByVal value As String)
            m_Name = value
        End Set
    End Property

    Public Property Qty As Integer
        Get
            Return m_Qty
        End Get
        Set(ByVal value As Integer)
            m_Qty = value
        End Set
    End Property

    Public Property Price As Double
        Get
            Return m_Price
        End Get
        Set(ByVal value As Double)
            m_Price = value
        End Set
    End Property

    Public ReadOnly Property LineTotal As Double
        Get
            Return m_Qty * m_Price
        End Get
    End Property
End Class
