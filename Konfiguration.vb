Imports System.IO
Imports System.Configuration
Imports System.Collections.Specialized
Public Class Konfiguration
    Public Farbe As String
    Public Merkfarbe As String
    Public Anzeigefarbe As String
    Public Anzeigepfad As String

    Public Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtPfaddisplay.TextChanged
    End Sub
    Public Sub FarbeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FarbeToolStripMenuItem.Click
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub GrauToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GrauToolStripMenuItem1.Click
        Config.Farbe = Color.LightGray
        BackColor = Config.Farbe
        Form1.BackColor = Config.Farbe
        Anzeigefarbe = "Systemstandart"
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub RotToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RotToolStripMenuItem.Click
        Config.Farbe = Color.Red
        BackColor = Config.Farbe
        Form1.BackColor = Config.Farbe
        Anzeigefarbe = "Rot"
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub BlauToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlauToolStripMenuItem.Click
        Config.Farbe = Color.Blue
        BackColor = Config.Farbe
        Form1.BackColor = Config.Farbe
        Anzeigefarbe = "Blau"
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub GelbToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GelbToolStripMenuItem.Click
        Config.Farbe = Color.Yellow
        BackColor = Config.Farbe
        Form1.BackColor = Config.Farbe
        Anzeigefarbe = "Gelb"
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub GrünToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GrünToolStripMenuItem.Click
        Config.Farbe = Color.Green
        BackColor = Config.Farbe
        Form1.BackColor = Config.Farbe
        Anzeigefarbe = "Grün"
        FarbeToolStripMenuItem.Text = Anzeigefarbe
        My.Settings.Save()
    End Sub
    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles ButAbbruch.Click
        BackColor = Config.Merkfarbe
        Form1.BackColor = Config.Merkfarbe
        Config.Farbe = Config.Merkfarbe
        Config.Pfad = Config.Merkpfad
        Close()
    End Sub
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles butOK.Click
        Config.Merkfarbe = Config.Farbe
        Config.Merkpfad = Config.Pfad
        My.Settings.Save()
        Close()
    End Sub
    Public Sub Konfiguration_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BackColor = Config.Farbe
        Config.Merkfarbe = Config.Farbe
        Config.Merkpfad = Config.Pfad
        txtPfaddisplay.Text = Config.Pfad
        If Config.Farbe = Color.Green Then
            Anzeigefarbe = "Grün"
        ElseIf Config.Farbe = Color.Red Then
            Anzeigefarbe = "Rot"
        ElseIf Config.Farbe = Color.Yellow Then
            Anzeigefarbe = "Gelb"
        ElseIf Config.Farbe = Color.Blue Then
            Anzeigefarbe = "Gelb"
        Else
            Anzeigefarbe = "Systemstandart"
        End If
        FarbeToolStripMenuItem.Text = Anzeigefarbe
    End Sub
    Public Sub ButBrowse_Click_1(sender As Object, e As EventArgs) Handles ButBrowse.Click
        Dim sfd As New SaveFileDialog()
        sfd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        sfd.FilterIndex = 1
        sfd.RestoreDirectory = True
        If sfd.ShowDialog() = DialogResult.OK Then
            Dim FileName As String = sfd.FileName
            txtPfaddisplay.Text = FileName
            Config.Pfad = FileName
        End If
        My.Settings.Save()
    End Sub
End Class