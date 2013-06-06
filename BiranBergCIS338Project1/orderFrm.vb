Imports BusinessModel
Public Class orderFrm

    'Delcare variables
    Private m_coffeePrice As Double
    Private m_tax As Double = 0.09

    'establishing connection to outside classes
    Private ctrl As Controller
    Private order As Order
    Dim coffees As Collection = Controller.getMenu
    
    Dim coffeeList As New List(Of Coffee)

    'Variables for runtime created fields. 
    Public cboCoffees As New ArrayList
    Public lblCoffees As New ArrayList
    Public lblCoffeePrices As New ArrayList
    Public txtPrices As New ArrayList
    Public nudQtys As New ArrayList
    Public txtTotals As New ArrayList
    Public totalArray As New ArrayList

    'This will clear the entire form to allow for a new order to be added.
    Private Sub btnNewOrder_Click(sender As System.Object, e As System.EventArgs) Handles btnNewOrder.Click
        'clear controls
        Dim cbo As ComboBox
        Dim txt As TextBox
        Dim txt2 As TextBox
        Dim qty As NumericUpDown

        Dim last = cboCoffees.Count - 1
        Dim index As Integer
        For index = last To 1 Step -1
            cbo = cboCoffees(index)
            cboCoffees.RemoveAt(index)
            cbo.Dispose()

            txt = txtPrices(index)
            txtPrices.RemoveAt(index)
            txt.Dispose()

            txt2 = txtTotals(index)
            txtTotals.RemoveAt(index)
            txt2.Dispose()

            qty = nudQtys(index)
            nudQtys.RemoveAt(index)
            qty.Dispose()
        Next

        cboCoffee.Select()
        cboCoffee.SelectedIndex = -1
        txtTotalB4Tax.Clear()
        txtSalesTax.Clear()
        txtGrandTotal.Clear()
        txtPrice.Clear()
        txtTotal.Clear()
        nudQty.Value = 1
        txtCust.Clear()

        'server does not clear because of the assumption that the same server would continuously use the same station to fill multiple orders.

        'update the order number automatically
        ctrl.orderNumber = ctrl.orderNo
        nudOrderNumber.Value = CInt(ctrl.orderNumber)
        Me.Text = "Order Number: " + nudOrderNumber.Value.ToString
    End Sub

    Private Sub orderFrm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        btnClose.PerformClick()
    End Sub

    'handles loading everything for the menu
    Private Sub orderFrm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ctrl = CType(MdiParent, frmMain).ctrl

        Me.InitializeComboBox(cboCoffee)
        Me.InitializeTextBox(txtPrice)
        Me.InitializeNumericUpDown(nudQty)
        Me.InitializeTextBox2(txtTotal)

        cboCoffees.Add(cboCoffee)
        txtPrices.Add(txtPrice)
        nudQtys.Add(nudQty)
        txtTotals.Add(txtTotal)

        For Each coffee In coffees
            coffeeList.Add(coffee)
        Next

        addCoffeeList()
        populateList()

        'updates the order number
        nudOrderNumber.Value = CInt(ctrl.orderNumber)
        Me.Text = "Order Number: " + ctrl.orderNumber.ToString
    End Sub

    'populates the coffee menu
    Public Sub populateList()

        For Each Coffee as Coffee In coffeeList
           lblCoffeeList.Text += Coffee.name + vbCrLf
            lblCoffeePrice.Text += FormatCurrency(Coffee.price) + vbCrLf
        Next

    End Sub

    'adds all of the events so that you can add runtime created field to the menu
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        ' Add a new row
        Dim cbo As New ComboBox
        cbo.Left = cboCoffee.Left
        cbo.Size = cboCoffee.Size
        cbo.Tag = cboCoffee.Tag
        cbo.DropDownStyle = cboCoffee.DropDownStyle
        Dim n1 As Integer = cboCoffees.Count - 1
        cbo.Top = cboCoffees(n1).top + 25
        Me.InitializeComboBox(cbo)
        cboCoffees.Add(cbo)
        cbo.Tag = cboCoffees.Count - 1
        Me.pnlOrder.Controls.Add(cbo)

        'Register as event listener
        AddHandler cbo.SelectedIndexChanged, AddressOf cboCoffee_SelectedIndexChanged1

        ' Add a new price text box
        Dim txtP As New TextBox
        txtP.Left = txtPrice.Left
        txtP.Size = txtPrice.Size
        txtP.ReadOnly = txtPrice.ReadOnly
        txtP.Enabled = txtPrice.Enabled
        Dim n2 As Integer = txtPrices.Count - 1
        txtP.Top = txtPrices(n2).top + 25
        txtP.TextAlign = HorizontalAlignment.Right
        txtPrices.Add(txtP)
        txtP.Tag = txtPrices.Count - 1
        Me.pnlOrder.Controls.Add(txtP)

        Dim nud As New NumericUpDown
        nud.Left = nudQty.Left
        nud.Size = nudQty.Size
        nud.Minimum = 1
        nud.ReadOnly = nudQty.ReadOnly
        Dim n3 As Integer = nudQtys.Count - 1
        nud.Top = nudQtys(n3).top + 25
        nud.TextAlign = HorizontalAlignment.Center
        nudQtys.Add(nud)
        nud.Tag = nudQtys.Count - 1
        Me.pnlOrder.Controls.Add(nud)

        AddHandler nud.ValueChanged, AddressOf nudQty_ValueChanged1

        'create new total textbox 
        Dim txtT As New TextBox
        txtT.Left = txtTotal.Left
        txtT.Size = txtTotal.Size
        txtT.ReadOnly = txtTotal.ReadOnly
        txtT.Enabled = txtTotal.Enabled
        Dim n4 As Integer = txtTotals.Count - 1
        txtT.Top = txtTotals(n4).top + 25
        txtT.TextAlign = HorizontalAlignment.Right
        txtTotals.Add(txtT)
        txtT.Tag = txtTotals.Count - 1
        Me.pnlOrder.Controls.Add(txtT)

    End Sub


    Public Sub addCoffeeList()
        For Each Coffee As Coffee In coffeeList
            cboCoffee.Items.Add(Coffee.name)
        Next
    End Sub

    'initializes comboboxs 
    Public Function InitializeComboBox(ByVal cbo As ComboBox) As ComboBox
        ' Intialize values
        For Each Coffee As Coffee In coffeeList
            cbo.Items.Add(Coffee.name)
        Next
        Return cbo
    End Function

    'initializes the price text boxs
    Public Function InitializeTextBox(ByVal txt As TextBox) As TextBox
        Return txt
    End Function

    'initializes the total text boxs
    Public Function InitializeTextBox2(ByVal txt As TextBox) As TextBox
        Return txt
    End Function

    'initializes a nud
    Public Function InitializeNumericUpDown(ByVal nud As NumericUpDown) As NumericUpDown
        Return nud
    End Function

    'handles what happens when user selects a different coffee.
    Private Sub cboCoffee_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCoffee.SelectedIndexChanged
        ' Add the item to the textbox on the right
        Try
            txtPrice.Text = FormatCurrency(coffeeList(cboCoffee.SelectedIndex).price.ToString)
            
            Dim ol As New Orderline 
            ol.Price = CDbl(coffeeList(cboCoffee.SelectedIndex).price)
            ol.Qty = CInt(nudQty.Value)
            txtTotal.Text = FormatCurrency(ol.LineTotal)
        Catch

        End Try

    End Sub

    'handles what happens when user selects a different coffee for runtime created fields.
    Private Sub cboCoffee_SelectedIndexChanged1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Sender.tag looks up item selected and sends corresponding price to textbox price
        Try
            txtPrices(sender.tag).text = FormatCurrency(coffeeList(sender.SelectedIndex).price.ToString)
            Dim ol As New Orderline 
            ol.Price = CDbl(coffeeList(cboCoffees(sender.tag).SelectedIndex).price)
            ol.Qty = CInt(nudQtys(sender.tag).Value)
            txtTotals(sender.tag).Text = FormatCurrency(ol.LineTotal)
        Catch
        End Try
    End Sub

    'removes one row of runtime created fields at a time 
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        ' Delete textboxes in reverse order of being added

        Dim cbo As ComboBox
        Dim nud As NumericUpDown
        Dim txt As TextBox
        Dim txt2 As TextBox
        Dim last = cboCoffees.Count - 1
        If last > 0 Then
            Dim index As Integer
            For index = last To last Step -1
                cbo = cboCoffees(index)
                cboCoffees.RemoveAt(index)
                cbo.Dispose()
                txt = txtPrices(index)
                txtPrices.RemoveAt(index)
                txt.Dispose()
                nud = nudQtys(index)
                nudQtys.RemoveAt(index)
                nud.Dispose()
                txt2 = txtTotals(index)
                txtTotals.RemoveAt(index)
                txt2.Dispose()
            Next
        End If
        cboCoffee.Select()
    End Sub

    'changes the nudQty button and then changes the values for that item
    Private Sub nudQty_ValueChanged(sender As Object, e As EventArgs) Handles nudQty.ValueChanged
        Try
            Dim ol As New Orderline 
            ol.Price = CDbl(coffeeList(cboCoffees(sender.tag).SelectedIndex).price)
            ol.Qty = CInt(nudQtys(sender.tag).Value)
            txtTotals(sender.tag).Text = FormatCurrency(ol.LineTotal)
        Catch
            'do nothing
        End Try

    End Sub

    'changes one of the runtime added nudQty buttons and then changes the values for that item
    Private Sub nudQty_ValueChanged1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ol As New Orderline 
            ol.Price = CDbl(coffeeList(cboCoffees(sender.tag).SelectedIndex).price)
            ol.Qty = CInt(nudQtys(sender.tag).Value)
            txtTotals(sender.tag).Text = FormatCurrency(ol.LineTotal)
        Catch
            'do nothing
        End Try
    End Sub

    'totals up the order
    Private Sub btnTotal_Click(sender As Object, e As EventArgs) Handles btnTotal.Click
        Try 
            If txtServer.TextLength = Nothing and txtCust.TextLength = Nothing Then
                MessageBox.Show("Please enter both a server name and a customer name before submitting an order.")
            ElseIf txtServer.TextLength = Nothing then 
                MessageBox.Show("Please enter a server name before submitting an order.")
            ElseIf txtCust.TextLength = Nothing then 
                MessageBox.Show("Please enter a customer name before submitting an order. ")
            Else
                Try
                    If cboCoffee.SelectedIndex >= 0 then 
                        order = New Order
                        order.ID = nudOrderNumber.Value
                        order.ServerName = txtServer.Text
                        order.CustomerName = txtCust.Text
                        order.OrderDate = dtpDate.Value

                        Dim count = cboCoffees.Count - 1
                        Dim index As Integer
                        For index = 0 to count Step 1
                            If cboCoffees(index).selectedIndex >= 0 then
                                Dim oL As New Orderline
                                oL.Name = cboCoffees(index).SelectedItem
                                oL.Price = CDbl(txtPrices(index).text) 
                                oL.Qty = CInt(nudQtys(index).value)
                                order.AddLine(oL)
                            End If
                        Next

                        txtTotalB4Tax.Text = FormatCurrency(order.SubTotal)
                        txtSalesTax.text = FormatCurrency(order.SalesTax)
                        txtGrandTotal.text = FormatCurrency(order.Total)
                    Else
                        Messagebox.Show("You must select at least one coffee.")
                    end if
                Catch ex As Exception
                End Try
            End If
        Catch
            'do nothing
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        btnTotal.PerformClick()
        If txtServer.TextLength = Nothing and txtCust.TextLength = Nothing then 

        Else 
            Dim msg As String = ctrl.AddOrder(order)
            If msg <> "" Then
                Dim result As Integer = MessageBox.Show(msg, "Coffee Order", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
                If result = vbYes Then
                    ctrl.AddOrder(order, True)
                End If
            End If
        End If
    End Sub

    Private Sub btnDelete_Click( sender As Object,  e As EventArgs) Handles btnDelete.Click
        order = ctrl.GetOrder(nudOrderNumber.Value)

        Dim msg As String = ctrl.deleteOrder(order)
        If msg <> "" Then
            Dim result As Integer = MessageBox.Show(msg, "Coffee Order", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
            If result = vbYes Then
                ctrl.deleteOrder(order, True)
                btnNewOrder.PerformClick()
            End If
        End If
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Dim msg As String = "Would you like to save the order before you close the form?"
        Dim result As Integer = MessageBox.Show(msg, "Before you go.", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation,MessageBoxDefaultButton.Button2)
        If result = vbYes then 
            btnSave.PerformClick()
            Me.Dispose()
        ElseIf result = vbNo then 
            Me.Dispose()
        Else 
            'Hi there :)
        End If
    End Sub

Private Sub nudOrderNumber_ValueChanged( sender As Object,  e As EventArgs) Handles nudOrderNumber.ValueChanged
        Me.Text = "Order Number: " + nudOrderNumber.Value.ToString
End Sub
End Class
