<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OrderSummary
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OrderSummary))
        Me.lbxSummary = New System.Windows.Forms.ListBox()
        Me.btnRetrieve = New System.Windows.Forms.Button()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.lblCust = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.lblServerPH = New System.Windows.Forms.Label()
        Me.lblCustPH = New System.Windows.Forms.Label()
        Me.lblDatePH = New System.Windows.Forms.Label()
        Me.lblOrderTotalPH = New System.Windows.Forms.Label()
        Me.lblOrderTotal = New System.Windows.Forms.Label()
        Me.SuspendLayout
        '
        'lbxSummary
        '
        Me.lbxSummary.BackColor = System.Drawing.Color.White
        Me.lbxSummary.FormattingEnabled = true
        Me.lbxSummary.Location = New System.Drawing.Point(13, 13)
        Me.lbxSummary.Name = "lbxSummary"
        Me.lbxSummary.Size = New System.Drawing.Size(120, 290)
        Me.lbxSummary.TabIndex = 0
        '
        'btnRetrieve
        '
        Me.btnRetrieve.Location = New System.Drawing.Point(139, 280)
        Me.btnRetrieve.Name = "btnRetrieve"
        Me.btnRetrieve.Size = New System.Drawing.Size(204, 23)
        Me.btnRetrieve.TabIndex = 1
        Me.btnRetrieve.Text = "Retrieve Order"
        Me.btnRetrieve.UseVisualStyleBackColor = true
        '
        'lblServer
        '
        Me.lblServer.AutoSize = true
        Me.lblServer.BackColor = System.Drawing.Color.Transparent
        Me.lblServer.Location = New System.Drawing.Point(139, 13)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(41, 13)
        Me.lblServer.TabIndex = 2
        Me.lblServer.Text = "Server:"
        '
        'lblCust
        '
        Me.lblCust.AutoSize = true
        Me.lblCust.BackColor = System.Drawing.Color.Transparent
        Me.lblCust.Location = New System.Drawing.Point(139, 38)
        Me.lblCust.Name = "lblCust"
        Me.lblCust.Size = New System.Drawing.Size(51, 13)
        Me.lblCust.TabIndex = 2
        Me.lblCust.Text = "Customer"
        '
        'lblDate
        '
        Me.lblDate.AutoSize = true
        Me.lblDate.BackColor = System.Drawing.Color.Transparent
        Me.lblDate.Location = New System.Drawing.Point(139, 62)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(33, 13)
        Me.lblDate.TabIndex = 2
        Me.lblDate.Text = "Date:"
        '
        'lblServerPH
        '
        Me.lblServerPH.AutoSize = true
        Me.lblServerPH.BackColor = System.Drawing.Color.Transparent
        Me.lblServerPH.Location = New System.Drawing.Point(243, 13)
        Me.lblServerPH.Name = "lblServerPH"
        Me.lblServerPH.Size = New System.Drawing.Size(0, 13)
        Me.lblServerPH.TabIndex = 2
        '
        'lblCustPH
        '
        Me.lblCustPH.AutoSize = true
        Me.lblCustPH.BackColor = System.Drawing.Color.Transparent
        Me.lblCustPH.Location = New System.Drawing.Point(243, 38)
        Me.lblCustPH.Name = "lblCustPH"
        Me.lblCustPH.Size = New System.Drawing.Size(0, 13)
        Me.lblCustPH.TabIndex = 2
        '
        'lblDatePH
        '
        Me.lblDatePH.AutoSize = true
        Me.lblDatePH.BackColor = System.Drawing.Color.Transparent
        Me.lblDatePH.Location = New System.Drawing.Point(243, 62)
        Me.lblDatePH.Name = "lblDatePH"
        Me.lblDatePH.Size = New System.Drawing.Size(0, 13)
        Me.lblDatePH.TabIndex = 2
        '
        'lblOrderTotalPH
        '
        Me.lblOrderTotalPH.AutoSize = true
        Me.lblOrderTotalPH.BackColor = System.Drawing.Color.Transparent
        Me.lblOrderTotalPH.Location = New System.Drawing.Point(243, 87)
        Me.lblOrderTotalPH.Name = "lblOrderTotalPH"
        Me.lblOrderTotalPH.Size = New System.Drawing.Size(0, 13)
        Me.lblOrderTotalPH.TabIndex = 2
        '
        'lblOrderTotal
        '
        Me.lblOrderTotal.AutoSize = true
        Me.lblOrderTotal.BackColor = System.Drawing.Color.Transparent
        Me.lblOrderTotal.Location = New System.Drawing.Point(139, 87)
        Me.lblOrderTotal.Name = "lblOrderTotal"
        Me.lblOrderTotal.Size = New System.Drawing.Size(63, 13)
        Me.lblOrderTotal.TabIndex = 2
        Me.lblOrderTotal.Text = "Order Total:"
        '
        'OrderSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"),System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(351, 313)
        Me.Controls.Add(Me.lblOrderTotal)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.lblOrderTotalPH)
        Me.Controls.Add(Me.lblCust)
        Me.Controls.Add(Me.lblDatePH)
        Me.Controls.Add(Me.lblCustPH)
        Me.Controls.Add(Me.lblServerPH)
        Me.Controls.Add(Me.lblServer)
        Me.Controls.Add(Me.btnRetrieve)
        Me.Controls.Add(Me.lbxSummary)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "OrderSummary"
        Me.Tag = "ordSum"
        Me.Text = "Summary Form"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents lbxSummary As System.Windows.Forms.ListBox
    Friend WithEvents btnRetrieve As System.Windows.Forms.Button
    Friend WithEvents lblServer As System.Windows.Forms.Label
    Friend WithEvents lblCust As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblServerPH As System.Windows.Forms.Label
    Friend WithEvents lblCustPH As System.Windows.Forms.Label
    Friend WithEvents lblDatePH As System.Windows.Forms.Label
    Friend WithEvents lblOrderTotalPH As System.Windows.Forms.Label
    Friend WithEvents lblOrderTotal As System.Windows.Forms.Label
End Class
