Public Class Detail
    Private m_OrderDetails As New Collection

    'adds coffee to coffeeDB
    Public Sub New()

    End Sub

    'creates master DB
    Public Property MasterDatabase As Collection
        Get
            Return m_OrderDetails
        End Get
        Set(value As Collection)
            m_OrderDetails = value
        End Set
    End Property

    'adds order
    Public Sub addOrder(ByVal order As Order)
        m_OrderDetails.Add(order, order.ID)
    End Sub

    'gets order
    Public Function getOrder(ByVal id As String)
        Return m_OrderDetails.Item(id)
    End Function


End Class
