Imports BusinessModel

Public Class OrderSummary

    Private ctrl As Controller
    Dim orders As Collection = Controller.getOderDB
    Dim orderList As new List(Of Order)    
    Dim tempOrder As new Order 
    Dim orderLines As New ArrayList 

    Private Sub OrderSummary_Load( sender As Object,  e As EventArgs) Handles MyBase.Load
        For Each ord In orders
            lbxSummary.Items.Add(ord.ID)
            orderList.Add(ord)
        Next
    End Sub

    Private Sub lbxSummary_SelectedIndexChanged( sender As Object,  e As EventArgs) Handles lbxSummary.SelectedIndexChanged
        'this populates the order information on the summary page. 
        Try
            lblServerPH.Text = orderList(lbxSummary.SelectedIndex).ServerName
            lblCustPH.Text = orderList(lbxSummary.SelectedIndex).CustomerName
            lblDatePH.Text = orderList(lbxSummary.SelectedIndex).OrderDate.ToShortDateString
            lblOrderTotalPH.Text = FormatCurrency(orderList(lbxSummary.SelectedIndex).Total)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub lbxSummar_doubleClick(sender As Object, e As EventArgs) Handles lbxSummary.DoubleClick, btnRetrieve.Click
        'this is a try catch loop.
        'this handles both a double click on the lbx and a button click on the retrieve button. 
        Try
            'sets up the variables for a new orderform and the two counts.
            Dim ChildForm As New orderFrm        
            Dim count As Integer = 0
            Dim lc As Integer
            
            'creates a count of the orderlines in the order
            For Each ol As orderline In orderList(lbxSummary.SelectedIndex).LineCount
                lc += 1
            Next
                lc = lc - 1
            
            'creates a new orderform to display the order that is being called and links it to the MDI parent
            ChildForm.MdiParent = frmMain
            ChildForm.Text = "Order Number" + orderList(lbxSummary.SelectedIndex).ID.ToString
            ChildForm.Show()
            
            'closes the order summary box
            Me.Close()
            
            'for each loop to populate the orderform 
            For Each oL As Orderline In orderList(lbxSummary.SelectedIndex).LineCount

                ChildForm.cboCoffees(count).SelectedItem = oL.name
                ChildForm.txtPrices(count).text = FormatCurrency(oL.Price)
                ChildForm.nudQtys(count).Value = oL.Qty
                ChildForm.txtTotals(count).text = FormatCurrency(oL.LineTotal)
                ChildForm.nudOrderNumber.Value = orderList(lbxSummary.SelectedIndex).ID
                ChildForm.Text = "Order Number: " + orderList(lbxSummary.SelectedIndex).ID.ToString
                ChildForm.dtpDate.Value = orderList(lbxSummary.SelectedIndex).OrderDate
                ChildForm.txtTotalB4Tax.Text = FormatCurrency(orderList(lbxSummary.SelectedIndex).SubTotal)
                ChildForm.txtSalesTax.Text = FormatCurrency(orderList(lbxSummary.SelectedIndex).SalesTax)
                ChildForm.txtGrandTotal.Text = FormatCurrency(orderList(lbxSummary.SelectedIndex).Total)
                ChildForm.txtCust.Text = orderList(lbxSummary.SelectedIndex).CustomerName
                ChildForm.txtServer.Text = orderList(lbxSummary.SelectedIndex).ServerName
                
                'checks to see if the number of orderlines has been reached. 
                If count < lc then 
                    ChildForm.btnAdd.PerformClick()
                    count += 1
                End If
            Next
        Catch
        
        End Try
    End Sub
End Class