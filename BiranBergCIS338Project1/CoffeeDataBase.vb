Public Class CoffeeDataBase

    Private m_coffeeDataBase As New Collection

    'adds coffee to coffeeDB
    Public Sub New()
        Dim tempCoffee As New Coffee
        tempCoffee = New Coffee("Antigua", 5.95)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Apanas", 11.95)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Bantu", 13.6)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Colombia", 9.92)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Costa Rica", 11.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Ethiopa", 13.6)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("French Roast", 9.92)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Huehuetenago", 11.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Kenya", 13.6)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Mexico", 11.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Morning Ed.", 9.92)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Nepenthe", 11.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Sumatra", 9.92)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Yemen", 8.96)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Yemen Mocha", 18.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
        tempCoffee = New Coffee("Zimbabwe", 11.52)
        m_coffeeDataBase.Add(tempCoffee, tempCoffee.name)
    End Sub

    'creates master DB
    Public Property MasterDatabase As Collection
        Get
            Return m_coffeeDataBase
        End Get
        Set(value As Collection)
            m_coffeeDataBase = value
        End Set
    End Property
    
    'adds coffee
    Public Sub addCoffee(ByVal coffee As Coffee)
        m_coffeeDataBase.Add(coffee, coffee.name)
    End Sub

    'gets coffee
    Public Function getCoffee(ByVal id As String)
        Return m_coffeeDataBase.Item(id)
    End Function
End Class
