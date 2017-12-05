Imports System.IO
Public Class Form1
    Public erstzahl As Double
    Public zweitzahl As Double
    Public ergebnis As Double
    Public memory As Double
    Public kommazahl As Double
    Public zweitkommazahl As Double
    Public isistgleich As Boolean
    Public iskomma As Boolean
    Public iskommazwei As Boolean
    Public iseinsgeteilt As Boolean
    Public isprozent As Boolean
    Public ishochzwei As Boolean
    Public iswurzel As Boolean
    Public rechenzeichen As String
    Public ispreviewhandeled As Boolean
    Public pfad As String, txtTmp As String
    'Dim Config As New Taschenrechner.My.MySettings
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Config.Merkpfad = Config.Pfad
        Config.Merkfarbe = Config.Farbe
        BackColor = Config.Farbe
        My.Settings.Save()

    End Sub
    Private Sub Nummer0_Click(sender As Object, e As MouseEventArgs) Handles butZahl0.MouseClick
        ZifferneingabeF(0)
    End Sub
    Private Sub Nummer1_Click(sender As Object, e As MouseEventArgs) Handles butZahl1.MouseClick
        ZifferneingabeF(1)
    End Sub
    Private Sub Nummer2_Click(sender As Object, e As MouseEventArgs) Handles butZahl2.MouseClick
        ZifferneingabeF(2)
    End Sub
    Private Sub Nummer3_Click(sender As Object, e As MouseEventArgs) Handles butZahl3.MouseClick
        ZifferneingabeF(3)
    End Sub
    Private Sub Nummer4_Click(sender As Object, e As MouseEventArgs) Handles butZahl4.MouseClick
        ZifferneingabeF(4)
    End Sub
    Private Sub Nummer5_Click(sender As Object, e As MouseEventArgs) Handles butZahl5.MouseClick
        ZifferneingabeF(5)
    End Sub
    Private Sub Nummer6_Click(sender As Object, e As MouseEventArgs) Handles butZahl6.MouseClick
        ZifferneingabeF(6)
    End Sub
    Private Sub Nummer7_Click(sender As Object, e As MouseEventArgs) Handles butZahl7.MouseClick
        ZifferneingabeF(7)
    End Sub
    Private Sub Nummer8_Click(sender As Object, e As MouseEventArgs) Handles butZahl8.MouseClick
        ZifferneingabeF(8)
    End Sub
    Private Sub Nummer9_Click(sender As Object, e As MouseEventArgs) Handles butZahl9.MouseClick
        ZifferneingabeF(9)
    End Sub
    Private Sub PlusMinus_Click(sender As Object, e As MouseEventArgs) Handles butPlusMinus.MouseClick
        If isistgleich = True Then
            ergebnis = -ergebnis
            txtDisplay.Text = ergebnis
        End If
        If rechenzeichen = "" Then
            erstzahl = -erstzahl
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
        Else
            zweitzahl = -zweitzahl
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub Clear_Click(sender As Object, e As MouseEventArgs) Handles butClear.MouseClick
        ReinigungF()
    End Sub
    Private Sub Cleareverything_Click(sender As Object, e As MouseEventArgs) Handles butCleareverything.MouseClick
        ReinigungF()
    End Sub
    Private Sub Komma_Click(sender As Object, e As MouseEventArgs) Handles butKomma.MouseClick
        KommaklickF()
    End Sub
    Private Sub Istgleich_Click(sender As Object, e As MouseEventArgs) Handles butIstgleich.MouseClick
        IsgleichF()
    End Sub
    Private Sub Plus_Click(sender As Object, e As MouseEventArgs) Handles butPlus.MouseClick
        PlusklickF()
    End Sub
    Private Sub Minus_Click(sender As Object, e As MouseEventArgs) Handles butMinus.MouseClick
        MinusclickF()
    End Sub
    Private Sub Mal_Click(sender As Object, e As MouseEventArgs) Handles butMal.MouseClick
        MalF()
    End Sub
    Private Sub Geteilt_Click(sender As Object, e As MouseEventArgs) Handles butGeteilt.MouseClick
        GeteiltF()
    End Sub
    Private Sub Zurück_Click(sender As Object, e As MouseEventArgs) Handles butZurück.MouseClick
        ZurückF()
    End Sub
    Private Sub Button1_Click(sender As Object, e As MouseEventArgs) Handles butMClear.MouseClick
        memory = Nothing
    End Sub
    Private Sub Button2_Click(sender As Object, e As MouseEventArgs) Handles butMRecall.MouseClick
        txtDisplay.Text = memory
        If rechenzeichen = "" Then
            erstzahl = Nothing
        Else
            zweitzahl = Nothing
        End If
    End Sub
    Private Sub MPlus_Click(sender As Object, e As MouseEventArgs) Handles butMPlus.MouseClick
        If isistgleich = True Then
            memory = memory + ergebnis
            isistgleich = False
        ElseIf rechenzeichen = "" Then
            memory = memory + erstzahl
            erstzahl = Nothing
        Else
            memory = memory + zweitzahl
            zweitzahl = Nothing
        End If
    End Sub
    Private Sub MMinus_Click(sender As Object, e As MouseEventArgs) Handles butMMinus.MouseClick
        If isistgleich = True Then
            memory = memory - ergebnis
            isistgleich = False
        ElseIf rechenzeichen = "" Then
            memory = memory - erstzahl
            erstzahl = Nothing
        Else
            memory = memory - zweitzahl
            zweitzahl = Nothing
        End If
    End Sub
    Private Sub M_Click(sender As Object, e As MouseEventArgs) Handles butMS.MouseClick
        If rechenzeichen = "" Then
            erstzahl = Nothing
            erstzahl = memory
            txtDisplay.Text = erstzahl
        Else
            zweitzahl = Nothing
            zweitzahl = memory
            txtDisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub Prozent_Click(sender As Object, e As EventArgs) Handles butProzent.MouseClick
        isprozent = True
        If rechenzeichen = "" Then
            erstzahl = Nothing
            txtErstzahldisplay.Text = erstzahl
            txtDisplay.Text = erstzahl
        ElseIf rechenzeichen IsNot "" Then
            zweitzahl = erstzahl / 100 * zweitzahl
        End If
        txtDisplay.Text = zweitzahl
        txtZweitzahldisplay.Text = zweitzahl
    End Sub
    Private Sub Wurzel_Click(sender As Object, e As EventArgs) Handles butWurzel.MouseClick
        iswurzel = True
        If isistgleich = True Then
            ergebnis = Math.Sqrt(ergebnis)
            txtDisplay.Text = ergebnis
            isistgleich = False
        ElseIf rechenzeichen = "" Then
            erstzahl = Math.Sqrt(erstzahl)
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
        Else
            zweitzahl = Math.Sqrt(zweitzahl)
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub Hochzwei_Click(sender As Object, e As EventArgs) Handles butHochzwei.MouseClick
        ishochzwei = True
        If isistgleich = True Then
            ergebnis = ergebnis * ergebnis
            txtDisplay.Text = ergebnis
            isistgleich = False
        ElseIf rechenzeichen = "" Then
            erstzahl = erstzahl * erstzahl
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
        Else
            zweitzahl = zweitzahl * zweitzahl
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub EinsGeteilt_Click(sender As Object, e As EventArgs) Handles butEinsgeteilt.MouseClick
        iseinsgeteilt = True
        If isistgleich = True Then
            ergebnis = 1 / ergebnis
            txtDisplay.Text = ergebnis
            isistgleich = False

        ElseIf rechenzeichen = "" Then
            erstzahl = 1 / erstzahl
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
        Else
            zweitzahl = 1 / zweitzahl
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub Form1_Keydown(sender As Object, ByVal e As KeyEventArgs) Handles MyBase.KeyDown ', Zurück.KeyDown, Zahl9.KeyDown, Zahl8.KeyDown, Zahl7.KeyDown, Zahl6.KeyDown, Zahl5.KeyDown, Zahl4.KeyDown, Zahl3.KeyDown, Zahl2.KeyDown, Zahl1.KeyDown, Zahl0.KeyDown, Wurzel.KeyDown, Prozent.KeyDown, PlusMinus.KeyDown, Plus.KeyDown, MS.KeyDown, MRecall.KeyDown, MPlus.KeyDown, MMinus.KeyDown, Minus.KeyDown, MClear.KeyDown, Mal.KeyDown, Komma.KeyDown, Istgleich.KeyDown, Hochzwei.KeyDown, Geteilt.KeyDown, EinsGeteilt.KeyDown, Cleareverything.KeyDown, Clear.KeyDown
        If (Not ispreviewhandeled) Then
            TastatureingabeF(e.KeyCode)
            ispreviewhandeled = False
        End If
    End Sub
    Private Sub Wurzel_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles butZurück.PreviewKeyDown, butZahl9.PreviewKeyDown, butZahl8.PreviewKeyDown, butZahl7.PreviewKeyDown, butZahl6.PreviewKeyDown, butZahl5.PreviewKeyDown, butZahl4.PreviewKeyDown, butZahl3.PreviewKeyDown, butZahl2.PreviewKeyDown, butZahl1.PreviewKeyDown, butZahl0.PreviewKeyDown, butWurzel.PreviewKeyDown, butProzent.PreviewKeyDown, butPlusMinus.PreviewKeyDown, butPlus.PreviewKeyDown, butMS.PreviewKeyDown, butMRecall.PreviewKeyDown, butMPlus.PreviewKeyDown, butMMinus.PreviewKeyDown, butMinus.PreviewKeyDown, butMClear.PreviewKeyDown, butMal.PreviewKeyDown, butKomma.PreviewKeyDown, butIstgleich.PreviewKeyDown, butHochzwei.PreviewKeyDown, butGeteilt.PreviewKeyDown, butEinsgeteilt.PreviewKeyDown, butCleareverything.PreviewKeyDown, butClear.PreviewKeyDown
        TastatureingabeF(e.KeyCode)
        ispreviewhandeled = True
    End Sub
    Private Sub LR_Click(sender As Object, e As MouseEventArgs) Handles butLR.MouseClick
        Dim iDateiInhaltZeil() As String = Nothing
        If Config.Pfad IsNot "" Then
            iDateiInhaltZeil = File.ReadAllLines(Config.Pfad)
        End If
        If Config.Pfad = Nothing Then
            MsgBox("Es konnte kein Pfad gefunden werden! Gehen Sie in die Konfiguration und wählen Sie einen Pfad aus.")
        ElseIf iDateiInhaltZeil.Length = Nothing Then
            MsgBox("Letzte Rechnung: Keine Rechungen vorhanden")
        Else
            ZeilensortierenF()
            MsgBox("Letzte Rechnung: " & iDateiInhaltZeil(iDateiInhaltZeil.Length - 1))
        End If
    End Sub
    Private Sub butLRA_Click(sender As Object, e As MouseEventArgs) Handles butLRA.MouseClick
        Dim mydocpath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim iDateiInhaltZeile() As String = Nothing
        Dim iMessageBoxText As String = ""
        If Config.Pfad IsNot "" Then
            iDateiInhaltZeile = File.ReadAllLines(Config.Pfad)
            If iDateiInhaltZeile.Length = Nothing Then
                MsgBox("Letzte Rechnung: Keine Rechungen vorhanden")
            ElseIf iDateiInhaltZeile.Length - 1 = 1 Then
                ZeilensortierenF()
                MsgBox("Letzte Rechnung: " & iDateiInhaltZeile(iDateiInhaltZeile.Length - 1))
            Else
                For Each zeile In iDateiInhaltZeile
                    iMessageBoxText = iMessageBoxText & zeile & Environment.NewLine
                Next
                ZeilensortierenF()
                MessageBox.Show(iMessageBoxText, "Alle " & (iDateiInhaltZeile.Length - 1) & " Rechnungen!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MsgBox("Es konnte kein Pfad gefunden werden! Gehen Sie in die Konfiguration und wählen Sie einen Pfad aus.")
        End If
    End Sub
    Private Sub LetztenX_Click(sender As Object, e As MouseEventArgs) Handles butLRX.MouseClick
        Dim iDateiInhaltZeile() As String = Nothing
        Dim iMessageBoxText As String = ""
        Dim schleifenzahl As Integer
        Dim message, title, defaultValue As String
        Dim myValue As Object
        Dim z As Int32
        If Config.Pfad IsNot "" Then
            Dim sr As New StreamReader(Config.Pfad)
            Do While sr.Peek >= Nothing
                If sr.ReadLine = Nothing Then
                End If
                z = z + 1
            Loop
            sr.Close()
        End If
        message = "Geben Sie eine Zahl ein"
        title = "Letzen X Rechnungen"
        defaultValue = ""
        myValue = InputBox(message, title, defaultValue)
        If myValue Is "" Or "Nothing" Then myValue = defaultValue
        If Config.Pfad IsNot Nothing Then
            iDateiInhaltZeile = File.ReadAllLines(Config.Pfad)
        End If
        If Config.Pfad = Nothing Then
            MsgBox("Es konnte kein Pfad gefunden werden! Gehen Sie in die Konfiguration und wählen Sie einen Pfad aus.")
        ElseIf myValue = "" Or "Nothing" Then
            MessageBox.Show("Abgebrochen!", "Abbruch", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf myValue <= Nothing Then
            MsgBox("Geben Sie eine Zahl ein, die größer als Nothing ist")
        ElseIf iDateiInhaltZeile.Length = "Nothing" Then
            MsgBox("Keine Rechnungen Vorhanden")
        ElseIf (iDateiInhaltZeile.Length - 1 <= myValue) Then
            For Each zeile In iDateiInhaltZeile
                iMessageBoxText = iMessageBoxText & zeile & Environment.NewLine
            Next
            ZeilensortierenF()
            MessageBox.Show(iMessageBoxText, "Letzten " & (iDateiInhaltZeile.Length - 1) & " Rechnungen!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            schleifenzahl = 1
            Do Until schleifenzahl = myValue + 1
                iMessageBoxText = iMessageBoxText & (iDateiInhaltZeile(iDateiInhaltZeile.Length - schleifenzahl)) & Environment.NewLine
                schleifenzahl = schleifenzahl + 1
            Loop
            ZeilensortierenF()
            MessageBox.Show(iMessageBoxText, "Letzten " & myValue & " Rechnungen!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Sub ZeilensortierenF()
        Dim fs As Object, f As Object
        Dim txt As String = ""
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.OpenTextFile(Config.Pfad)
        Do While f.AtEndOfStream <> True
            txtTmp = f.ReadLine
            If Trim$(txtTmp) <> "" Then _
            txt = txt & vbCrLf & txtTmp
        Loop
        f.Close
        f = fs.CreateTextFile(Config.Pfad, True)
        f.write(txt)
        f.Close
    End Sub
    Public Sub ZifferneingabeF(vZiffer As Integer)
        If rechenzeichen = "" And
            isprozent = True Then
            erstzahl = Nothing
            isprozent = False
        ElseIf isprozent = True Then
            zweitzahl = Nothing
            isprozent = False
        End If
        If rechenzeichen = "" And
            ishochzwei = True Then
            erstzahl = Nothing
            ishochzwei = False
        ElseIf ishochzwei = True Then
            zweitzahl = Nothing
            ishochzwei = False
        End If
        If rechenzeichen = "" And
            iseinsgeteilt = True Then
            erstzahl = Nothing
            iseinsgeteilt = False
        ElseIf iseinsgeteilt = True Then
            zweitzahl = Nothing
            iseinsgeteilt = False
        End If
        If rechenzeichen = "" And
           iswurzel = True Then
            erstzahl = Nothing
            iswurzel = False

        ElseIf iseinsgeteilt = True Then
            zweitzahl = Nothing
            iswurzel = False
        End If
        If iskomma = True Then
            kommazahl = kommazahl * 10 + vZiffer
            txtDisplay.Text = kommazahl
        ElseIf iskommazwei = True Then
            zweitkommazahl = zweitkommazahl * 10 + vZiffer
            txtDisplay.Text = zweitkommazahl
        ElseIf rechenzeichen = "" Then
            erstzahl = erstzahl * 10 + vZiffer
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
            txtZweitzahldisplay.Text = zweitzahl
            txtRechenzeichendisplay.Text = rechenzeichen
        Else
            zweitzahl = zweitzahl * 10 + vZiffer
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
            txtErstzahldisplay.Text = erstzahl
            txtRechenzeichendisplay.Text = rechenzeichen
        End If
    End Sub
    Private Sub TastatureingabeF(vKey As Keys)
        If vKey = Keys.D0 OrElse vKey = Keys.NumPad0 Then
            ZifferneingabeF(Nothing)
        ElseIf vKey = Keys.D1 OrElse vKey = Keys.NumPad1 Then
            ZifferneingabeF(1)
        ElseIf vKey = Keys.D2 OrElse vKey = Keys.NumPad2 Then
            ZifferneingabeF(2)
        ElseIf vKey = Keys.D3 OrElse vKey = Keys.NumPad3 Then
            ZifferneingabeF(3)
        ElseIf vKey = Keys.D4 OrElse vKey = Keys.NumPad4 Then
            ZifferneingabeF(4)
        ElseIf vKey = Keys.D5 OrElse vKey = Keys.NumPad5 Then
            ZifferneingabeF(5)
        ElseIf vKey = Keys.D6 OrElse vKey = Keys.NumPad6 Then
            ZifferneingabeF(6)
        ElseIf vKey = Keys.D7 OrElse vKey = Keys.NumPad7 Then
            ZifferneingabeF(7)
        ElseIf vKey = Keys.D8 OrElse vKey = Keys.NumPad8 Then
            ZifferneingabeF(8)
        ElseIf vKey = Keys.D9 OrElse vKey = Keys.NumPad9 Then
            ZifferneingabeF(9)
        ElseIf vKey = Keys.Oemplus OrElse vKey = Keys.Add Then
            PlusklickF()
        ElseIf vKey = Keys.OemMinus OrElse vKey = Keys.Subtract Then
            MinusclickF()
        ElseIf vKey = Keys.Multiply Then
            MalF()
        ElseIf vKey = Keys.Divide Then
            GeteiltF()
        ElseIf vKey = Keys.Oemcomma OrElse vKey = Keys.Decimal Then
            KommaklickF()
        ElseIf vKey = Keys.Back Then
            ZurückF()
        ElseIf vKey = Keys.Enter Then
            IsgleichF()
        ElseIf vKey = Keys.Escape Then
            ReinigungF()
        End If
    End Sub
    Private Sub ReinigungF()
        erstzahl = Nothing
        zweitzahl = Nothing
        ergebnis = Nothing
        txtDisplay.Text = "Nothing"
        txtErstzahldisplay.Text = erstzahl
        txtZweitzahldisplay.Text = zweitzahl
        rechenzeichen = ""
        txtRechenzeichendisplay.Text = rechenzeichen
        iskomma = False
        iskommazwei = False
    End Sub
    Private Sub ZurückF()
        If rechenzeichen = "" Then
            erstzahl = Nothing
            kommazahl = Nothing
            iskomma = False
            txtDisplay.Text = erstzahl
            txtErstzahldisplay.Text = erstzahl
        Else
            zweitzahl = Nothing
            zweitkommazahl = Nothing
            iskommazwei = False
            txtDisplay.Text = zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
    End Sub
    Private Sub PlusklickF()
        rechenzeichen = "+"
        txtRechenzeichendisplay.Text = rechenzeichen
        If erstzahl = Nothing Then
            erstzahl = ergebnis
            txtErstzahldisplay.Text = erstzahl
        End If
        zweitzahl = Nothing
        txtZweitzahldisplay.Text = zweitzahl
        iskomma = False
    End Sub
    Private Sub MinusclickF()
        rechenzeichen = "-"
        txtRechenzeichendisplay.Text = rechenzeichen
        If erstzahl = Nothing Then
            erstzahl = ergebnis
            txtErstzahldisplay.Text = erstzahl
        End If
        zweitzahl = Nothing
        txtZweitzahldisplay.Text = zweitzahl
        iskomma = False
    End Sub
    Private Sub MalF()
        rechenzeichen = "x"
        txtRechenzeichendisplay.Text = rechenzeichen
        If erstzahl = Nothing Then
            erstzahl = ergebnis
            txtErstzahldisplay.Text = erstzahl
        End If
        zweitzahl = Nothing
        txtZweitzahldisplay.Text = zweitzahl
        iskomma = False
    End Sub
    Private Sub GeteiltF()
        rechenzeichen = "/"
        txtRechenzeichendisplay.Text = rechenzeichen
        If erstzahl = Nothing Then
            erstzahl = ergebnis
            txtErstzahldisplay.Text = erstzahl
        End If
        zweitzahl = Nothing
        txtZweitzahldisplay.Text = zweitzahl
        iskomma = False
    End Sub
    Private Sub IsgleichF()
        txtErstzahldisplay.Text = erstzahl
        txtZweitzahldisplay.Text = zweitzahl
        txtRechenzeichendisplay.Text = rechenzeichen
        If kommazahl = Nothing Then
            iskomma = True
        End If
        If iskomma = True Then
            erstzahl = kommazahl / (10 ^ kommazahl.ToString.Length) + erstzahl
            txtErstzahldisplay.Text = erstzahl
        End If
        If iskommazwei = True Then
            zweitzahl = zweitkommazahl / (10 ^ zweitkommazahl.ToString.Length) + zweitzahl
            txtZweitzahldisplay.Text = zweitzahl
        End If
        isistgleich = True
        If rechenzeichen = "+" Then
            ergebnis = erstzahl + zweitzahl
            txtDisplay.Text = ergebnis
        ElseIf rechenzeichen = "-" Then
            ergebnis = erstzahl - zweitzahl
            txtDisplay.Text = ergebnis
        ElseIf rechenzeichen = "x" Then
            ergebnis = erstzahl * zweitzahl
            txtDisplay.Text = ergebnis
        ElseIf rechenzeichen = "/" Then
            ergebnis = erstzahl / zweitzahl
            txtDisplay.Text = ergebnis
        Else
            ergebnis = erstzahl
            txtDisplay.Text = ergebnis
        End If
        LoggeRechnung()
        rechenzeichen = ""
        iskomma = False
        iskommazwei = False
        erstzahl = Nothing
        zweitzahl = Nothing
        kommazahl = Nothing
        zweitkommazahl = Nothing
    End Sub
    Private Sub LoggeRechnung()
        Using outputFile As New StreamWriter(Config.Pfad, True)
            outputFile.WriteLine(Environment.NewLine & DateTime.Now & "  " & erstzahl & " " & rechenzeichen & " " & zweitzahl & " " & "=" & " " & ergebnis)
        End Using
    End Sub
    Private Sub KommaklickF()
        If rechenzeichen = "" Then
            iskomma = True
        Else
            iskommazwei = True
        End If
        txtZweitzahldisplay.Text = zweitzahl
        txtErstzahldisplay.Text = erstzahl
    End Sub
    Private Sub KonfigurationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KonfigurationToolStripMenuItem.Click
        Konfiguration.Show()
    End Sub
    Private Sub SchließenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SchließenToolStripMenuItem.Click
        Close()
    End Sub
End Class