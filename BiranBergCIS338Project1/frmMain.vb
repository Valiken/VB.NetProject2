Imports System.Windows.Forms
Imports BusinessModel

Public Class frmMain
    Public ctrl As New Controller

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles OrderFormToolStripMenuItem.Click
        ' Create a new instance of the child form.
        Dim ChildForm As New orderFrm
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me
        
        ChildForm.Text = "Order Number: " & ctrl.orderNumber

        ChildForm.Show()
    End Sub
    
    Private Sub SummaryFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryFormToolStripMenuItem.Click
                ' Create a new instance of the child form.
        Dim ChildForm As New OrderSummary
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me
        
        ChildForm.Text = "Order Summary"

        ChildForm.Show()
    End Sub

    Private Sub FileMenu_DropDownOpening( sender As Object,  e As EventArgs) Handles FileMenu.DropDownOpening
        Dim activeChild As form = Me.ActiveMdiChild
        Try
            if Me.ActiveMdiChild Is Nothing then
                CloseToolStripMenuItem.Enabled = false
                SaveToolStripMenuItem.Enabled = false 
            ElseIf Me.ActiveMdiChild.Tag = orderFrm.Tag then     
                CloseToolStripMenuItem.Enabled = true
                SaveToolStripMenuItem.Enabled = true
            ElseIf Me.ActiveMdiChild.Tag = orderSummary.Tag then
                CloseToolStripMenuItem.Enabled = true
                SaveToolStripMenuItem.Enabled = false
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
        CheckEditMenu()
        If CutToolStripMenuItem.Enabled = True Then
            If TypeOf Me.ActiveMdiChild.ActiveControl Is TextBox Then
                Clipboard.SetText(CType(ActiveMdiChild.ActiveControl, TextBox).SelectedText())
                CType(ActiveMdiChild.ActiveControl, TextBox).SelectedText = ""
            End If
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
        CheckEditMenu()
        If CopyToolStripMenuItem.Enabled = True Then
            If TypeOf Me.ActiveMdiChild.ActiveControl Is TextBox Then
                Clipboard.SetText(CType(ActiveMdiChild.ActiveControl, TextBox).SelectedText())
            End If
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
        CheckEditMenu()
        If PasteToolStripMenuItem.Enabled Then
            If TypeOf Me.ActiveMdiChild.ActiveControl Is TextBox Then
                CType(ActiveMdiChild.ActiveControl, TextBox).SelectedText = Clipboard.GetText()
            End If
        End If
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub CloseToolStripMenuItem_Click( sender As Object,  e As EventArgs) Handles CloseToolStripMenuItem.Click
        
        Try
            If Me.ActiveMdiChild.Tag = orderFrm.Tag then     
                Dim activeChild As orderfrm = Me.ActiveMdiChild
                activeChild.btnClose.PerformClick()
            ElseIf Me.ActiveMdiChild.Tag = orderSummary.Tag then
                Dim activeChild2 As ordersummary = Me.ActiveMdiChild
                activeChild2.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub

        Private Sub CheckEditMenu()
        ' Only allow menu choices as appropriate
        CutToolStripMenuItem.Enabled = False
        CopyToolStripMenuItem.Enabled = False
        PasteToolStripMenuItem.Enabled = False

        If Not Me.ActiveMdiChild Is Nothing Then
            If Not Me.ActiveMdiChild.ActiveControl Is Nothing Then
                If TypeOf Me.ActiveMdiChild.ActiveControl Is TextBox Then
                    If CType(ActiveMdiChild.ActiveControl, TextBox).SelectedText <> "" Then
                        CutToolStripMenuItem.Enabled = True
                        CopyToolStripMenuItem.Enabled = True
                    End If
                    If Clipboard.ContainsText Then
                        PasteToolStripMenuItem.Enabled = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click( sender As Object,  e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim activeChild As orderFrm = Me.ActiveMdiChild
        activeChild.btnSave.PerformClick()
    End Sub
End Class
