Public Class FrmNego
    Private Sub BtnNegocier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNegocier.Click
        '//////// Dans le cas où l'utilisateur n'a pas cliqué sur Révéler avant de cliquer sur négocier, la révélation a quand même lieu
        Dim i, j, k, l, StratDcd, IndProp, IndDcd As Integer 'Indices désignant le Proposant ou le Décidant 

        TableauRev.RowCount = 1         'définit le nombre de ligne du tableau révélation

        For i = 1 To 2                  '2 représente ici le nombre d'agents envisagés

            '/ACTION DE L'INCERTITUDE sur l'exactitude des données sur lesquelles se fondent les négociations/
            If Objectif(i - 1) Then
                MinX(i - 1) = Min(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter le minimum par défaut
                MinX(i - 1) = Borner(MinX(i - 1))
                OptX(i - 1) = Opt(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter l'optimum par défaut
                OptX(i - 1) = Borner(OptX(i - 1))
                MaxX(i - 1) = Max(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter le maximum par défaut
                MaxX(i - 1) = Borner(MaxX(i - 1))
            Else
                MinX(i - 1) = Min(i - 1) - Cert(i - 1)      'Si l'objectif est de partager, diminuer le minimum par défaut
                MinX(i - 1) = Borner(MinX(i - 1))
                OptX(i - 1) = Opt(i - 1)                    'Si l'objectif est de partager, conserver l'optimum par défaut
                OptX(i - 1) = Borner(OptX(i - 1))
                MaxX(i - 1) = Max(i - 1) + Cert(i - 1)      'Si l'objectif est de partager, augmenter le maximum par défaut
                MaxX(i - 1) = Borner(MaxX(i - 1))
            End If

            Select Case ActRevelation(i - 1)            '///ACTION DE REVELATION///
                Case "Mentir"
                    ActProposition(i - 1) = MaxX(i - 1) + Alea(i - 1)           'Révèle le MAXIMUM affecté de l'incertitude, augmenté d'une variation
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Mentir"
                Case ("Bluffer")
                    ActProposition(i - 1) = OptX(i - 1) + Alea(i - 1)           'Révèle l'OPTIMUM affecté de l'incertitude, augmenté d'une variation
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Bluffer"
                Case "Dire la vérité"
                    ActProposition(i - 1) = OptX(i - 1)                         'Révèle l'OPTIMUM affecté de l'incertitude, sans changement
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Dire la vérité"
            End Select
        Next i

        '////// Fin de la révélation (bis)


        TableauNeg.RowCount = 20         'Définit le nombre de lignes de l'objet tableau dédié à la négociation

        jmax = 10                        'Définit le nombre de tours de négociations aller-retour maximum

        For l = 0 To 19                  'Remise à zéro visuelle du tableau des négociations à chaque clic du bouton "Valider"
            TableauNeg(0, l).Value = ""
            TableauNeg(1, l).Value = ""
            TableauNeg(2, l).Value = ""
            TableauNeg(3, l).Value = ""
            Nego(l, 0) = 0
            Nego(l, 1) = 0
            Nego(l, 2) = 0
        Next

        For j = 1 To jmax               'Début de la négociation 'j désigne le nombre de tours
            For i = 1 To 2
                k = 2 - i               'k désigne l'agent B
                l = 2 * j + i - 3       'l désigne l'indice de ligne de la variable Nego et de l'objet tableau TableauNeg
                IndProp = i - 1         'Désigne i-1 comme le proposant
                IndDcd = k              'Désigne k comme le décidant

                If i = 1 Then
                    StratDcd = CmbStratAgt1.SelectedIndex 'Désigne la stratégie choisie par le décidant 
                Else
                    StratDcd = CmbStratAgt2.SelectedIndex 'Désigne la stratégie choisie par le décidant 
                End If


                TableauNeg(0, l).Value = l + 1                               'Affiche le numéro de ligne tout à gauche
                TableauNeg(IndProp + 1, l).Value = ActProposition(IndProp)   'Affichage de la valeur proposée dans la colonne proposant
                Nego(l, IndProp) = ActProposition(IndProp)
                TableauNeg(IndDcd + 1, l).Value = ""
                Nego(l, IndDcd) = 0
                TableauNeg(3, l).Value = "False"
                Nego(l, 2) = "False"

                Select Case StratDcd
                    Case 0                          'Stratégie Médiane indice 0
                        If ActProposition(IndProp) <= (OptX(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptX(IndDcd) - Alea(IndDcd)) Then
                            Accord = True
                            TableauNeg(3, l).Value = "True"
                            Nego(l, 2) = "True"
                            GoTo Sortie
                        Else
                            ActProposition(IndDcd) = (ActProposition(IndProp) + ActProposition(IndDcd)) / 2 'Proposition médiane au prochain tour par le Décidant
                            ActProposition(IndDcd) = ActVerifMax(ActProposition(IndDcd), IndDcd)    'Vérifie que la proposition médiane ne dépasse pas le maximum du Décidant
                        End If

                    Case 1                          'Stratégie Rigide indice 1 Si la proposition est dans la zone optimum plus ou moins l'aléa, la proposition est acceptée
                        If ActProposition(IndProp) <= (OptX(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptX(IndDcd) - Alea(IndDcd)) Then
                            Accord = True
                            TableauNeg(3, l).Value = "True"
                            Nego(l, 2) = "True"
                            GoTo Sortie
                        End If

                    Case 2                          'Stratégie Rigide avec Regret indice 2 : Si aucun accord n'est trouvé dans les derniers tours, la stratégie médiane est alors adoptée
                        If ActProposition(IndProp) <= (OptX(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptX(IndDcd) - Alea(IndDcd)) Then
                            Accord = True
                            TableauNeg(3, l).Value = "True"
                            Nego(l, 2) = "True"
                            GoTo Sortie
                        Else
                            If j >= jmax - 1 Then
                                ActProposition(IndDcd) = (ActProposition(IndProp) + ActProposition(IndDcd)) / 2 'Proposition médiane au prochain tour par le Décidant
                                ActProposition(IndDcd) = ActVerifMax(ActProposition(IndDcd), IndDcd)    'Vérifie que la proposition médiane ne dépasse pas le maximum du Décidant
                            End If
                        End If

                    Case 3                          'Stratégie Rigide avec Risque indice 3
                        If ActProposition(IndProp) <= (OptX(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptX(IndDcd) - Alea(IndDcd)) Then
                            Accord = True
                            TableauNeg(3, l).Value = "True"
                            Nego(l, 2) = "True"
                            GoTo Sortie
                        Else
                            If j >= jmax / 2 Then
                                If ActVerifRisk(ActProposition(IndDcd), IndProp, Risk) Then
                                    Nego(l, 2) = "False"
                                    GoTo Sortie
                                End If
                            End If
                        End If

                    Case 4                          'Stratégie Rigide avec Regret et Risque indice 4
                        If ActProposition(IndProp) <= (OptX(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptX(IndDcd) - Alea(IndDcd)) Then
                            Accord = True
                            TableauNeg(3, l).Value = "True"
                            Nego(l, 2) = "True"
                            GoTo Sortie
                        Else
                            If j >= jmax / 2 Then
                                If ActVerifRisk(ActProposition(IndDcd), IndProp, Risk) Then
                                    Nego(l, 2) = "False"
                                    GoTo Sortie
                                End If
                            End If

                            If j >= jmax - 1 Then
                                ActProposition(IndDcd) = (ActProposition(IndProp) + ActProposition(IndDcd)) / 2 'Proposition médiane au prochain tour par le Décidant
                                ActProposition(IndDcd) = ActVerifMax(ActProposition(IndDcd), IndDcd)    'Vérifie que la proposition médiane ne dépasse pas le maximum du Décidant
                            End If
                        End If

                End Select
            Next i
        Next j
Sortie:  'repère de sortie de la boucle, marque la fin des tours de négociation
        jmax = l

    End Sub

    Private Sub TxtAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtAnnuler.Click
        Me.Close()
    End Sub

    Private Sub FrmNego_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Integer
        Accord = False
        CmbObjectifAgt1.SelectedIndex = Objectif(0) 'affiche l'objectif de négociation par défaut pour l'Agent 1
        CmbObjectifAgt2.SelectedIndex = Objectif(1) 'affiche l'objectif de négociation par défaut pour l'Agent 2
        CmbStratAgt1.SelectedIndex = 0      'affiche la stratégie de négociation "Médiane" par défaut pour l'Agent 1
        CmbStratAgt2.SelectedIndex = 0      'affiche la stratégie de négociation "Médiane" par défaut pour l'Agent 2


        For i = 1 To 2                      '2 représente ici le nombre d'agents envisagées (il s'agit d'un indice)

            'Randomize()
            'Alea(i - 1) = Rnd() * Alea(i - 1)  'Calcul de la variation d'amplitude spécifiée pour chaque agent

            If Conf(i - 1) <= 0.3 Then
                ActRevelation(i - 1) = "Mentir"
            End If
            If Conf(i - 1) > 0.3 And Conf(i - 1) <= 0.6 Then
                ActRevelation(i - 1) = "Bluffer"
            End If
            If Conf(i - 1) > 0.6 Then
                ActRevelation(i - 1) = "Dire la vérité"
            End If
        Next i

    End Sub

    Private Sub BtnReveler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReveler.Click
        Dim i, j, k, l As Integer

        TableauRev.RowCount = 1         'définit le nombre de ligne du tableau révélation

        For i = 1 To 2                  '2 représente ici le nombre d'agents envisagés

            '/ACTION DE L'INCERTITUDE sur l'exactitude des données sur lesquelles se fondent les négociations/
            If Objectif(i - 1) Then
                MinX(i - 1) = Min(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter le minimum par défaut
                MinX(i - 1) = Borner(MinX(i - 1))
                OptX(i - 1) = Opt(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter l'optimum par défaut
                OptX(i - 1) = Borner(OptX(i - 1))
                MaxX(i - 1) = Max(i - 1) + Cert(i - 1)      'Si l'objectif est de maximiser, augmenter le maximum par défaut
                MaxX(i - 1) = Borner(MaxX(i - 1))
            Else
                MinX(i - 1) = Min(i - 1) - Cert(i - 1)      'Si l'objectif est de partager, diminuer le minimum par défaut
                MinX(i - 1) = Borner(MinX(i - 1))
                OptX(i - 1) = Opt(i - 1)                    'Si l'objectif est de partager, conserver l'optimum par défaut
                OptX(i - 1) = Borner(OptX(i - 1))
                MaxX(i - 1) = Max(i - 1) + Cert(i - 1)      'Si l'objectif est de partager, augmenter le maximum par défaut
                MaxX(i - 1) = Borner(MaxX(i - 1))
            End If

            Select Case ActRevelation(i - 1)            '///ACTION DE REVELATION///
                Case "Mentir"
                    ActProposition(i - 1) = MaxX(i - 1) + Alea(i - 1)           'Révèle le MAXIMUM affecté de l'incertitude, augmenté d'une variation
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Mentir"
                Case ("Bluffer")
                    ActProposition(i - 1) = OptX(i - 1) + Alea(i - 1)           'Révèle l'OPTIMUM affecté de l'incertitude, augmenté d'une variation
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Bluffer"
                Case "Dire la vérité"
                    ActProposition(i - 1) = OptX(i - 1)                         'Révèle l'OPTIMUM affecté de l'incertitude, sans changement
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    TableauRev.Rows(j).Cells(i - 1).Value = "Dire la vérité"
            End Select
        Next i
    End Sub
    Private Sub CmbObjectifAgt1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbObjectifAgt1.SelectedIndexChanged
        Select Case CmbObjectifAgt1.SelectedIndex
            Case Is = 0
                Objectif(0) = False
            Case Else
                Objectif(0) = True
        End Select
    End Sub

    Private Sub CmbObjectifAgt2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbObjectifAgt2.SelectedIndexChanged
        Select Case CmbObjectifAgt2.SelectedIndex
            Case Is = 0
                Objectif(1) = False
            Case Else
                Objectif(1) = True
        End Select
    End Sub
End Class