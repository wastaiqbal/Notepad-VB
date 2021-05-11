Imports System.IO
Public Class Form1

    Private kode1 As Integer = 0
    Dim Open As New OpenFileDialog
    Dim fileinfo As FileInfo
    Dim Save As New SaveFileDialog
    Dim SaveAs As New SaveFileDialog

    'MenuStrip
    Private Sub NewToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NewToolStripMenuItem.Click
        If MsgBox("Apakah anda ingin membuat text baru ?", MsgBoxStyle.YesNo + MessageBoxIcon.Question, "New") = DialogResult.Yes Then
            RichTextBoxNote.Clear()
            RichTextBoxNote.Focus()
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Open.Title = "Open"
        Open.Filter = "RichTextBox (*.rtf)|*.rtf"
        If Open.ShowDialog() = Windows.Forms.DialogResult.OK Then
            RichTextBoxNote.LoadFile(Open.FileName, RichTextBoxStreamType.PlainText)
            kode1 = 1
            ToolStripStatusLabel1.Text = "Terbuka"
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Save.Title = "Save"
        Save.Filter = "RichTextBox (*.rtf)|*.rtf"
        If kode1 = 0 Then
            If Save.ShowDialog() = DialogResult.OK Then
                RichTextBoxNote.SaveFile(Save.FileName, RichTextBoxStreamType.TextTextOleObjs)
                ToolStripStatusLabel1.Text = "Tersimpan"
                kode1 = 2
            End If
        ElseIf kode1 = 1 Then
            fileinfo = New FileInfo(Open.FileName)
            fileinfo.Delete()
            RichTextBoxNote.SaveFile(Open.FileName, RichTextBoxStreamType.TextTextOleObjs)
            ToolStripStatusLabel1.Text = "Tersimpan"
        ElseIf kode1 = 2 Then
            fileinfo = New FileInfo(Save.FileName)
            fileinfo.Delete()
            RichTextBoxNote.SaveFile(Save.FileName, RichTextBoxStreamType.TextTextOleObjs)
            ToolStripStatusLabel1.Text = "Tersimpan"
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveAs.Title = "Save As..."
        SaveAs.Filter = "Richtextbox (*.rtf)|*.rtf"
        If SaveAs.ShowDialog() = Windows.Forms.DialogResult.OK Then
            RichTextBoxNote.SaveFile(SaveAs.FileName, RichTextBoxStreamType.TextTextOleObjs)
            ToolStripStatusLabel1.Text = "Tersimpan"
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBoxNote.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBoxNote.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBoxNote.Paste()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBoxNote.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBoxNote.Redo()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBoxNote.SelectAll()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        If PrintDialogMenu.ShowDialog() = DialogResult.OK Then
            PrintDocumentMenu.Print()
        End If
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        PrintPreviewDialogMenu.ShowDialog()
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ToolBarToolStripMenuItem.Click
        If ToolBarToolStripMenuItem.Checked = True Then
            ToolBarToolStripMenuItem.Checked = False
            ToolStripMenu.Visible = False
        Else
            ToolStripMenu.Visible = True
            ToolBarToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click
        If StatusBarToolStripMenuItem.Checked = True Then
            StatusBarToolStripMenuItem.Checked = False
            StatusStripMenu.Visible = False
        Else
            StatusStripMenu.Visible = True
            StatusBarToolStripMenuItem.Checked = True
        End If
    End Sub

    'ToolStrip
    Private Sub NewToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles NewToolStripButton.Click
        If MsgBox("Apakah anda ingin membuat text baru ?", MsgBoxStyle.YesNo + MessageBoxIcon.Question, "New") = DialogResult.Yes Then
            RichTextBoxNote.Clear()
            RichTextBoxNote.Focus()
        End If
    End Sub

    Private Sub SaveToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles SaveToolStripButton.Click
        Save.Title = "Save"
        Save.Filter = "RichTextBox (*.rtf)|*.rtf"
        If kode1 = 0 Then
            If Save.ShowDialog() = DialogResult.OK Then
                RichTextBoxNote.SaveFile(Save.FileName, RichTextBoxStreamType.TextTextOleObjs)
                ToolStripStatusLabel1.Text = "Tersimpan"
                kode1 = 2
            End If
        ElseIf kode1 = 1 Then
            fileinfo = New FileInfo(Open.FileName)
            fileinfo.Delete()
            RichTextBoxNote.SaveFile(Open.FileName, RichTextBoxStreamType.TextTextOleObjs)
            ToolStripStatusLabel1.Text = "Tersimpan"
        ElseIf kode1 = 2 Then
            fileinfo = New FileInfo(Save.FileName)
            fileinfo.Delete()
            RichTextBoxNote.SaveFile(Save.FileName, RichTextBoxStreamType.TextTextOleObjs)
            ToolStripStatusLabel1.Text = "Tersimpan"
        End If
    End Sub

    Private Sub OpenToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles OpenToolStripButton.Click
        Open.Title = "Open"
        Open.Filter = "RichTextBox (*.rtf)|*.rtf"
        If Open.ShowDialog() = Windows.Forms.DialogResult.OK Then
            RichTextBoxNote.LoadFile(Open.FileName, RichTextBoxStreamType.PlainText)
            kode1 = 1
            ToolStripStatusLabel1.Text = "Terbuka"
        End If
    End Sub

    Private Sub CutToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles CutToolStripButton.Click
        RichTextBoxNote.Cut()
    End Sub

    Private Sub CopyToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles CopyToolStripButton.Click
        RichTextBoxNote.Copy()
    End Sub

    Private Sub PasteToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles PasteToolStripButton.Click
        RichTextBoxNote.Paste()
    End Sub

    Private Sub PrintToolStripButton_Click(sender As System.Object, e As System.EventArgs) Handles PrintToolStripButton.Click
        If PrintDialogMenu.ShowDialog() = DialogResult.OK Then
            PrintDocumentMenu.Print()
        End If
    End Sub

    Private Sub RichTextBoxNote_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBoxNote.TextChanged
        If RichTextBoxNote.TextLength + 1 Or RichTextBoxNote.TextLength - 1 Then
            ToolStripStatusLabel1.Text = "Belum Tersimpan"
        End If
    End Sub
End Class
