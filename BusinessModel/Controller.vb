Public Class Controller

    Shared ctrlOrderDetails As New Detail
    Shared ctrlMenu As New Menu
    Public orderNumber As String = 1

    Public Shared Function getMenu()
        Return ctrlMenu.MasterDatabase
    End Function

    Public Shared Sub updateMenu(ByVal db As Collection)
        ctrlMenu.MasterDatabase = db
    End Sub

    Public Shared Function getOderDB()
        Return ctrlOrderDetails.MasterDatabase
    End Function

    Public Shared Sub updateOrderDB(ByVal db As Collection)
        ctrlOrderDetails.MasterDatabase = db
    End Sub

    Public Function AddOrder(ByVal order As Order, Optional ByVal replace As Boolean = False) As String
        Try
            Dim msg As String = ""
            If ctrlOrderDetails.MasterDatabase.Contains(order.ID) Then
                If replace = True Then
                    ctrlOrderDetails.MasterDatabase.Remove(order.ID)
                    ctrlOrderDetails.MasterDatabase.Add(order, order.ID)
                Else
                    msg = "Duplicate key. Replace?"
                End If
            Else
                ctrlOrderDetails.MasterDatabase.Add(order, order.ID)
            End If
            Return msg
        Catch 
        End Try
    End Function

    Public Function GetOrder(ByVal key As String) As Order
        Try
            Return ctrlOrderDetails.MasterDatabase(key)
        Catch ex As Exception

        End Try
        
    End Function

    Public Function deleteOrder(ByVal order As Order,Optional ByVal del As Boolean = False) As String
        Try
            Dim msg As String = ""
            If ctrlOrderDetails.MasterDatabase.Contains(order.ID) Then
                If del = True Then
                    ctrlOrderDetails.MasterDatabase.Remove(order.ID)
                Else
                    msg = "Delete this order?"
                End If
            Else

            End If
            Return msg
        Catch 
        End Try
    End Function

    Public Function orderNo()
        orderNumber += 1
        Return orderNumber
    End Function


End Class
