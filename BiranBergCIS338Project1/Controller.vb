Public Class Controller
    Shared ctrlCoffeeDB As New CoffeeDataBase

    Public Shared Function getCoffeeDB()
        Return ctrlCoffeeDB.MasterDatabase
    End Function

    Public Shared Sub updateCoffeeDB(ByVal db As Collection)
        ctrlCoffeeDB.MasterDatabase = db
    End Sub


End Class
