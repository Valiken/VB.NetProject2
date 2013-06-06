<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class orderFrm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(orderFrm))
        Me.gbMenu = New System.Windows.Forms.GroupBox()
        Me.lblCoffeePrice = New System.Windows.Forms.Label()
        Me.lblCoffeeList = New System.Windows.Forms.Label()
        Me.gbSummary = New System.Windows.Forms.GroupBox()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.pnlOrder = New System.Windows.Forms.Panel()
        Me.lblCoffeeSelect = New System.Windows.Forms.Label()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.lblPrice = New System.Windows.Forms.Label()
        Me.txtPrice = New System.Windows.Forms.TextBox()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.nudQty = New System.Windows.Forms.NumericUpDown()
        Me.cboCoffee = New System.Windows.Forms.ComboBox()
        Me.gbTotal = New System.Windows.Forms.GroupBox()
        Me.txtGrandTotal = New System.Windows.Forms.TextBox()
        Me.txtSalesTax = New System.Windows.Forms.TextBox()
        Me.txtTotalB4Tax = New System.Windows.Forms.TextBox()
        Me.lblOrderTotal = New System.Windows.Forms.Label()
        Me.lblSalesTax = New System.Windows.Forms.Label()
        Me.lblFdTotal = New System.Windows.Forms.Label()
        Me.lblOrder = New System.Windows.Forms.Label()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.lblDateAndTime = New System.Windows.Forms.Label()
        Me.btnTotal = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnNewOrder = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblAddress = New System.Windows.Forms.Label()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.txtCust = New System.Windows.Forms.TextBox()
        Me.lblCust = New System.Windows.Forms.Label()
        Me.nudOrderNumber = New System.Windows.Forms.NumericUpDown()
        Me.gbMenu.SuspendLayout
        Me.gbSummary.SuspendLayout
        Me.pnlOrder.SuspendLayout
        CType(Me.nudQty,System.ComponentModel.ISupportInitialize).BeginInit
        Me.gbTotal.SuspendLayout
        CType(Me.nudOrderNumber,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SuspendLayout
        '
        'gbMenu
        '
        Me.gbMenu.BackColor = System.Drawing.Color.Transparent
        Me.gbMenu.Controls.Add(Me.lblCoffeePrice)
        Me.gbMenu.Controls.Add(Me.lblCoffeeList)
        Me.gbMenu.ForeColor = System.Drawing.Color.Black
        Me.gbMenu.Location = New System.Drawing.Point(12, 82)
        Me.gbMenu.Name = "gbMenu"
        Me.gbMenu.Size = New System.Drawing.Size(200, 338)
        Me.gbMenu.TabIndex = 0
        Me.gbMenu.TabStop = false
        Me.gbMenu.Text = "Menu of Coffee"
        '
        'lblCoffeePrice
        '
        Me.lblCoffeePrice.AutoSize = true
        Me.lblCoffeePrice.Location = New System.Drawing.Point(155, 19)
        Me.lblCoffeePrice.Name = "lblCoffeePrice"
        Me.lblCoffeePrice.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.lblCoffeePrice.Size = New System.Drawing.Size(0, 13)
        Me.lblCoffeePrice.TabIndex = 1
        '
        'lblCoffeeList
        '
        Me.lblCoffeeList.AutoSize = true
        Me.lblCoffeeList.ForeColor = System.Drawing.Color.Black
        Me.lblCoffeeList.Location = New System.Drawing.Point(7, 20)
        Me.lblCoffeeList.Margin = New System.Windows.Forms.Padding(3, 0, 3, 10)
        Me.lblCoffeeList.Name = "lblCoffeeList"
        Me.lblCoffeeList.Size = New System.Drawing.Size(0, 13)
        Me.lblCoffeeList.TabIndex = 0
        '
        'gbSummary
        '
        Me.gbSummary.AutoSize = true
        Me.gbSummary.BackColor = System.Drawing.Color.Transparent
        Me.gbSummary.Controls.Add(Me.btnDelete)
        Me.gbSummary.Controls.Add(Me.btnSave)
        Me.gbSummary.Controls.Add(Me.btnRemove)
        Me.gbSummary.Controls.Add(Me.btnAdd)
        Me.gbSummary.Controls.Add(Me.pnlOrder)
        Me.gbSummary.Location = New System.Drawing.Point(228, 82)
        Me.gbSummary.Name = "gbSummary"
        Me.gbSummary.Size = New System.Drawing.Size(394, 222)
        Me.gbSummary.TabIndex = 1
        Me.gbSummary.TabStop = false
        Me.gbSummary.Text = "Order Summary"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(313, 180)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 23)
        Me.btnDelete.TabIndex = 12
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = true
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(232, 180)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 11
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = true
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(87, 180)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(75, 23)
        Me.btnRemove.TabIndex = 10
        Me.btnRemove.Text = "Remove"
        Me.btnRemove.UseVisualStyleBackColor = true
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(6, 180)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 23)
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Add "
        Me.btnAdd.UseVisualStyleBackColor = true
        '
        'pnlOrder
        '
        Me.pnlOrder.AutoScroll = true
        Me.pnlOrder.BackColor = System.Drawing.Color.Transparent
        Me.pnlOrder.Controls.Add(Me.lblCoffeeSelect)
        Me.pnlOrder.Controls.Add(Me.lblQty)
        Me.pnlOrder.Controls.Add(Me.txtTotal)
        Me.pnlOrder.Controls.Add(Me.lblPrice)
        Me.pnlOrder.Controls.Add(Me.txtPrice)
        Me.pnlOrder.Controls.Add(Me.lblTotal)
        Me.pnlOrder.Controls.Add(Me.nudQty)
        Me.pnlOrder.Controls.Add(Me.cboCoffee)
        Me.pnlOrder.Location = New System.Drawing.Point(6, 19)
        Me.pnlOrder.Name = "pnlOrder"
        Me.pnlOrder.Size = New System.Drawing.Size(382, 155)
        Me.pnlOrder.TabIndex = 11
        '
        'lblCoffeeSelect
        '
        Me.lblCoffeeSelect.AutoSize = true
        Me.lblCoffeeSelect.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCoffeeSelect.Location = New System.Drawing.Point(3, 1)
        Me.lblCoffeeSelect.Name = "lblCoffeeSelect"
        Me.lblCoffeeSelect.Size = New System.Drawing.Size(88, 13)
        Me.lblCoffeeSelect.TabIndex = 0
        Me.lblCoffeeSelect.Text = "Coffee Selection:"
        '
        'lblQty
        '
        Me.lblQty.AutoSize = true
        Me.lblQty.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblQty.Location = New System.Drawing.Point(235, 2)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(49, 13)
        Me.lblQty.TabIndex = 1
        Me.lblQty.Text = "Quantity:"
        '
        'txtTotal
        '
        Me.txtTotal.Enabled = false
        Me.txtTotal.Location = New System.Drawing.Point(285, 18)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = true
        Me.txtTotal.ShortcutsEnabled = false
        Me.txtTotal.Size = New System.Drawing.Size(64, 20)
        Me.txtTotal.TabIndex = 8
        Me.txtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPrice
        '
        Me.lblPrice.AutoSize = true
        Me.lblPrice.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPrice.Location = New System.Drawing.Point(175, 1)
        Me.lblPrice.Name = "lblPrice"
        Me.lblPrice.Size = New System.Drawing.Size(37, 13)
        Me.lblPrice.TabIndex = 2
        Me.lblPrice.Text = "Price: "
        '
        'txtPrice
        '
        Me.txtPrice.Enabled = false
        Me.txtPrice.Location = New System.Drawing.Point(175, 18)
        Me.txtPrice.Name = "txtPrice"
        Me.txtPrice.ReadOnly = true
        Me.txtPrice.ShortcutsEnabled = false
        Me.txtPrice.Size = New System.Drawing.Size(54, 20)
        Me.txtPrice.TabIndex = 6
        Me.txtPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = true
        Me.lblTotal.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTotal.Location = New System.Drawing.Point(285, 2)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(34, 13)
        Me.lblTotal.TabIndex = 3
        Me.lblTotal.Text = "Total:"
        '
        'nudQty
        '
        Me.nudQty.Location = New System.Drawing.Point(235, 18)
        Me.nudQty.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudQty.Name = "nudQty"
        Me.nudQty.ReadOnly = true
        Me.nudQty.Size = New System.Drawing.Size(44, 20)
        Me.nudQty.TabIndex = 7
        Me.nudQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.nudQty.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'cboCoffee
        '
        Me.cboCoffee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCoffee.FormattingEnabled = true
        Me.cboCoffee.Location = New System.Drawing.Point(6, 17)
        Me.cboCoffee.Name = "cboCoffee"
        Me.cboCoffee.Size = New System.Drawing.Size(163, 21)
        Me.cboCoffee.TabIndex = 5
        Me.cboCoffee.Tag = "coffeeName"
        '
        'gbTotal
        '
        Me.gbTotal.BackColor = System.Drawing.Color.Transparent
        Me.gbTotal.Controls.Add(Me.txtGrandTotal)
        Me.gbTotal.Controls.Add(Me.txtSalesTax)
        Me.gbTotal.Controls.Add(Me.txtTotalB4Tax)
        Me.gbTotal.Controls.Add(Me.lblOrderTotal)
        Me.gbTotal.Controls.Add(Me.lblSalesTax)
        Me.gbTotal.Controls.Add(Me.lblFdTotal)
        Me.gbTotal.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbTotal.Location = New System.Drawing.Point(228, 320)
        Me.gbTotal.Name = "gbTotal"
        Me.gbTotal.Size = New System.Drawing.Size(237, 100)
        Me.gbTotal.TabIndex = 2
        Me.gbTotal.TabStop = false
        Me.gbTotal.Text = "Order Total"
        '
        'txtGrandTotal
        '
        Me.txtGrandTotal.Enabled = false
        Me.txtGrandTotal.Location = New System.Drawing.Point(81, 71)
        Me.txtGrandTotal.Name = "txtGrandTotal"
        Me.txtGrandTotal.ReadOnly = true
        Me.txtGrandTotal.ShortcutsEnabled = false
        Me.txtGrandTotal.Size = New System.Drawing.Size(100, 20)
        Me.txtGrandTotal.TabIndex = 18
        Me.txtGrandTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtSalesTax
        '
        Me.txtSalesTax.Enabled = false
        Me.txtSalesTax.Location = New System.Drawing.Point(81, 45)
        Me.txtSalesTax.Name = "txtSalesTax"
        Me.txtSalesTax.ReadOnly = true
        Me.txtSalesTax.ShortcutsEnabled = false
        Me.txtSalesTax.Size = New System.Drawing.Size(100, 20)
        Me.txtSalesTax.TabIndex = 17
        Me.txtSalesTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTotalB4Tax
        '
        Me.txtTotalB4Tax.Enabled = false
        Me.txtTotalB4Tax.Location = New System.Drawing.Point(81, 19)
        Me.txtTotalB4Tax.Name = "txtTotalB4Tax"
        Me.txtTotalB4Tax.ReadOnly = true
        Me.txtTotalB4Tax.ShortcutsEnabled = false
        Me.txtTotalB4Tax.Size = New System.Drawing.Size(100, 20)
        Me.txtTotalB4Tax.TabIndex = 16
        Me.txtTotalB4Tax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblOrderTotal
        '
        Me.lblOrderTotal.AutoSize = true
        Me.lblOrderTotal.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblOrderTotal.Location = New System.Drawing.Point(17, 74)
        Me.lblOrderTotal.Name = "lblOrderTotal"
        Me.lblOrderTotal.Size = New System.Drawing.Size(66, 13)
        Me.lblOrderTotal.TabIndex = 2
        Me.lblOrderTotal.Text = "Order Total: "
        '
        'lblSalesTax
        '
        Me.lblSalesTax.AutoSize = true
        Me.lblSalesTax.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblSalesTax.Location = New System.Drawing.Point(18, 48)
        Me.lblSalesTax.Name = "lblSalesTax"
        Me.lblSalesTax.Size = New System.Drawing.Size(57, 13)
        Me.lblSalesTax.TabIndex = 1
        Me.lblSalesTax.Text = "Sales Tax:"
        '
        'lblFdTotal
        '
        Me.lblFdTotal.AutoSize = true
        Me.lblFdTotal.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblFdTotal.Location = New System.Drawing.Point(18, 23)
        Me.lblFdTotal.Name = "lblFdTotal"
        Me.lblFdTotal.Size = New System.Drawing.Size(61, 13)
        Me.lblFdTotal.TabIndex = 0
        Me.lblFdTotal.Text = "Food Total:"
        '
        'lblOrder
        '
        Me.lblOrder.AutoSize = true
        Me.lblOrder.BackColor = System.Drawing.Color.Transparent
        Me.lblOrder.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblOrder.Location = New System.Drawing.Point(9, 45)
        Me.lblOrder.Name = "lblOrder"
        Me.lblOrder.Size = New System.Drawing.Size(53, 13)
        Me.lblOrder.TabIndex = 6
        Me.lblOrder.Text = "Order No."
        '
        'lblServer
        '
        Me.lblServer.AutoSize = true
        Me.lblServer.BackColor = System.Drawing.Color.Transparent
        Me.lblServer.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblServer.Location = New System.Drawing.Point(109, 45)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(41, 13)
        Me.lblServer.TabIndex = 8
        Me.lblServer.Text = "Server:"
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(112, 61)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(100, 20)
        Me.txtServer.TabIndex = 2
        '
        'lblDateAndTime
        '
        Me.lblDateAndTime.AutoSize = true
        Me.lblDateAndTime.BackColor = System.Drawing.Color.Transparent
        Me.lblDateAndTime.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblDateAndTime.Location = New System.Drawing.Point(423, 42)
        Me.lblDateAndTime.Name = "lblDateAndTime"
        Me.lblDateAndTime.Size = New System.Drawing.Size(61, 13)
        Me.lblDateAndTime.TabIndex = 10
        Me.lblDateAndTime.Text = "Date/Time:"
        '
        'btnTotal
        '
        Me.btnTotal.Location = New System.Drawing.Point(482, 339)
        Me.btnTotal.Name = "btnTotal"
        Me.btnTotal.Size = New System.Drawing.Size(144, 23)
        Me.btnTotal.TabIndex = 13
        Me.btnTotal.Text = "&Total"
        Me.btnTotal.UseVisualStyleBackColor = true
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(482, 391)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(144, 23)
        Me.btnClose.TabIndex = 15
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = true
        '
        'btnNewOrder
        '
        Me.btnNewOrder.Location = New System.Drawing.Point(482, 365)
        Me.btnNewOrder.Name = "btnNewOrder"
        Me.btnNewOrder.Size = New System.Drawing.Size(144, 23)
        Me.btnNewOrder.TabIndex = 14
        Me.btnNewOrder.Text = "&New Order"
        Me.btnNewOrder.UseVisualStyleBackColor = true
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = true
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Mistral", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.lblTitle.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTitle.Location = New System.Drawing.Point(13, 13)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(174, 26)
        Me.lblTitle.TabIndex = 14
        Me.lblTitle.Text = "Cal Poly Coffee Roasters"
        '
        'lblAddress
        '
        Me.lblAddress.AutoSize = true
        Me.lblAddress.BackColor = System.Drawing.Color.Transparent
        Me.lblAddress.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblAddress.Location = New System.Drawing.Point(423, 21)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(206, 13)
        Me.lblAddress.TabIndex = 15
        Me.lblAddress.Text = "3801 W. Temple Ave. Pomona, CA 91768"
        '
        'dtpDate
        '
        Me.dtpDate.Location = New System.Drawing.Point(426, 58)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(196, 20)
        Me.dtpDate.TabIndex = 4
        '
        'txtCust
        '
        Me.txtCust.Location = New System.Drawing.Point(315, 58)
        Me.txtCust.Name = "txtCust"
        Me.txtCust.Size = New System.Drawing.Size(100, 20)
        Me.txtCust.TabIndex = 3
        '
        'lblCust
        '
        Me.lblCust.AutoSize = true
        Me.lblCust.BackColor = System.Drawing.Color.Transparent
        Me.lblCust.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCust.Location = New System.Drawing.Point(312, 41)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(54, 13)
        Me.lblCust.TabIndex = 8
        Me.lblCust.Text = "Customer:"
        '
        'nudOrderNumber
        '
        Me.nudOrderNumber.Location = New System.Drawing.Point(12, 61)
        Me.nudOrderNumber.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudOrderNumber.Name = "nudOrderNumber"
        Me.nudOrderNumber.ReadOnly = true
        Me.nudOrderNumber.Size = New System.Drawing.Size(50, 20)
        Me.nudOrderNumber.TabIndex = 1
        Me.nudOrderNumber.Tag = "ordFrm"
        Me.nudOrderNumber.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'orderFrm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLight
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"),System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(671, 511)
        Me.Controls.Add(Me.nudOrderNumber)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnNewOrder)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnTotal)
        Me.Controls.Add(Me.lblDateAndTime)
        Me.Controls.Add(Me.lblCust)
        Me.Controls.Add(Me.txtCust)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.lblOrder)
        Me.Controls.Add(Me.gbTotal)
        Me.Controls.Add(Me.gbSummary)
        Me.Controls.Add(Me.gbMenu)
        Me.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "orderFrm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Order Form"
        Me.gbMenu.ResumeLayout(false)
        Me.gbMenu.PerformLayout
        Me.gbSummary.ResumeLayout(false)
        Me.pnlOrder.ResumeLayout(false)
        Me.pnlOrder.PerformLayout
        CType(Me.nudQty,System.ComponentModel.ISupportInitialize).EndInit
        Me.gbTotal.ResumeLayout(false)
        Me.gbTotal.PerformLayout
        CType(Me.nudOrderNumber,System.ComponentModel.ISupportInitialize).EndInit
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents gbMenu As System.Windows.Forms.GroupBox
    Friend WithEvents gbSummary As System.Windows.Forms.GroupBox
    Friend WithEvents gbTotal As System.Windows.Forms.GroupBox
    Friend WithEvents lblOrder As System.Windows.Forms.Label
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents lblDateAndTime As System.Windows.Forms.Label
    Friend WithEvents txtGrandTotal As System.Windows.Forms.TextBox
    Friend WithEvents txtSalesTax As System.Windows.Forms.TextBox
    Friend WithEvents txtTotalB4Tax As System.Windows.Forms.TextBox
    Friend WithEvents lblOrderTotal As System.Windows.Forms.Label
    Friend WithEvents lblSalesTax As System.Windows.Forms.Label
    Friend WithEvents lblFdTotal As System.Windows.Forms.Label
    Friend WithEvents btnTotal As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnNewOrder As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents nudQty As System.Windows.Forms.NumericUpDown
    Friend WithEvents cboCoffee As System.Windows.Forms.ComboBox
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents lblPrice As System.Windows.Forms.Label
    Friend WithEvents lblQty As System.Windows.Forms.Label
    Friend WithEvents lblCoffeeSelect As System.Windows.Forms.Label
    Friend WithEvents txtPrice As System.Windows.Forms.TextBox
    Friend WithEvents txtTotal As System.Windows.Forms.TextBox
    Friend WithEvents lblCoffeeList As System.Windows.Forms.Label
    Friend WithEvents lblCoffeePrice As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents pnlOrder As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtCust As System.Windows.Forms.TextBox
    Friend WithEvents lblCust As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents nudOrderNumber As System.Windows.Forms.NumericUpDown

End Class
