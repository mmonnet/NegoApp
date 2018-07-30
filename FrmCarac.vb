Public Class FrmCarac
    Private Sub FrmQuantitatif_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CmbAgents.SelectedIndex = 0         'l'agent 1 est affiché par défaut à l'ouverture de la fenêtre, avec ses valeurs dans les champs correspondant
        TxtCoeffAff.Text = CoeffAff
        TxtMin.Text = Min(0) * CoeffAff
        TxtMax.Text = Max(0) * CoeffAff
        TxtOptimumMoyenne.Text = Opt(0) * CoeffAff
        TxtOptimumEcartType.Text = OptET(0) * CoeffAff
        TxtRisk.Text = Risk(0)
        TxtRegt.Text = Regt(0)
        TxtMand.Text = Mand(0)
        TxtConf.Text = Conf(0)
        If Objectif(0) Then
            CmbObjectif.SelectedIndex = 1
        Else
            CmbObjectif.SelectedIndex = 0
        End If
        TxtCert.Text = Cert(0)
        TxtAlea.Text = Alea(0)

    End Sub

    Private Sub CmbAgents_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbAgents.SelectedIndexChanged
        Dim i As Integer
        i = CmbAgents.SelectedIndex
        TxtCoeffAff.Text = CoeffAff
        TxtMin.Text = Min(i) * CoeffAff
        TxtMax.Text = Max(i) * CoeffAff
        TxtOptimumMoyenne.Text = Opt(i) * CoeffAff
        TxtOptimumEcartType.Text = OptET(i) * CoeffAff
        TxtRisk.Text = Risk(i)
        TxtRegt.Text = Regt(i)
        TxtMand.Text = Mand(i)
        TxtConf.Text = Conf(i)
        If Objectif(i) Then
            CmbObjectif.SelectedIndex = 1
        Else
            CmbObjectif.SelectedIndex = 0
        End If
        TxtCert.Text = Cert(i)
        TxtAlea.Text = Alea(i)

    End Sub
    Private Sub TxtMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtMin.TextChanged

    End Sub

    Private Sub BtnEffacer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnEffacer.Click
        TxtMin.Text = ""
        TxtMax.Text = ""
        TxtOptimumMoyenne.Text = ""
        TxtOptimumEcartType.Text = ""
        TxtRisk.Text = ""
        TxtRegt.Text = ""
        TxtMand.Text = ""
        TxtConf.Text = ""
        CmbObjectif.SelectedIndex = 0

    End Sub

    Private Sub TxtAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAnnuler.Click
        Close()
    End Sub

    Private Sub BtnValider_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValider.Click
        Dim i As Integer
        i = CmbAgents.SelectedIndex

        CoeffAff = Valeur(TxtCoeffAff.Text)


        'Sauvegarde des valeurs entrées par l'utilisateur dans les champs correspondant
        If Valeur(TxtMin.Text) < 0 Or Valeur(TxtMin.Text) > 100 * CoeffAff Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Minimum requis : " & TxtMin.Text)
        Else
            Min(i) = Valeur(TxtMin.Text) / CoeffAff
        End If
        If Valeur(TxtMax.Text) < 0 Or Valeur(TxtMax.Text) > 100 * CoeffAff Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Maximum espéré : " & TxtMax.Text)
        Else
            Max(i) = Valeur(TxtMax.Text) / CoeffAff
        End If
        If Valeur(TxtOptimumMoyenne.Text) < 0 Or Valeur(TxtOptimumMoyenne.Text) > 100 * CoeffAff Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Optimum recherché : " & TxtOptimumMoyenne.Text)
        Else
            Opt(i) = Valeur(TxtOptimumMoyenne.Text) / CoeffAff
        End If
        If Valeur(TxtOptimumEcartType.Text) < 0 Or Valeur(TxtOptimumEcartType.Text) > 100 * CoeffAff Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de l'écart-type de l'optimum recherché : " & TxtOptimumEcartType.Text)
        Else
            OptET(i) = Valeur(TxtOptimumEcartType.Text) / CoeffAff
        End If

        If Valeur(TxtRisk.Text) < 0 Or Valeur(TxtRisk.Text) > 100 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Aversion au risque : " & TxtRisk.Text)
        Else
            Risk(i) = Valeur(TxtRisk.Text)
        End If
        If Valeur(TxtRegt.Text) < 0 Or Valeur(TxtRegt.Text) > 100 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Aversion au regret : " & TxtRegt.Text)
        Else
            Regt(i) = Valeur(TxtRegt.Text)
        End If
        If Valeur(TxtMand.Text) < 0 Or Valeur(TxtMand.Text) > 100 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Mandat de négociation : " & TxtMand.Text)
        Else
            Mand(i) = Valeur(TxtMand.Text)
        End If
        If Valeur(TxtConf.Text) < 0 Or Valeur(TxtConf.Text) > 100 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Degré de confiance : " & TxtConf.Text)
        Else
            Conf(i) = Valeur(TxtConf.Text)
        End If

        If CmbObjectif.SelectedIndex = 0 Then
            Objectif(i) = False
        Else
            Objectif(i) = True
        End If

        If Valeur(TxtAlea.Text) < 0 Or Valeur(TxtAlea.Text) > 1 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 1. Valeur de Amplitude des variations : " & TxtAlea.Text)
        Else
            Alea(i) = Valeur(TxtAlea.Text)
        End If
        If Valeur(TxtCert.Text) < 0 Or Valeur(TxtCert.Text) > 100 Then
            MsgBox("La valeur entrée doit être comprise entre 0 et 100. Valeur de Degré de certitude : " & TxtCert.Text)
        Else
            Cert(i) = Valeur(TxtCert.Text)
        End If
    End Sub

    Private Sub CmbObjectif_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbObjectif.SelectedIndexChanged
        Dim i As Integer
        i = CmbAgents.SelectedIndex
        Select Case CmbObjectif.SelectedIndex
            Case Is = 0
                Objectif(i) = False
            Case Else
                Objectif(i) = True
        End Select
    End Sub
End Class