Public Class FrmNegoDyn
    Dim TableauResStats() As Single


    Private Sub FrmNegoDyn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Accord = False

        CmbObjectifAgt1.SelectedIndex = Objectif(0) 'affiche l'objectif de négociation pour l'Agent 1
        CmbObjectifAgt2.SelectedIndex = Objectif(1) 'affiche l'objectif de négociation pour l'Agent 2
        CmbStratAgt1.SelectedIndex = 0      'affiche la stratégie de négociation "Médiane" par défaut pour l'Agent 1
        CmbStratAgt2.SelectedIndex = 0      'affiche la stratégie de négociation "Médiane" par défaut pour l'Agent 2

        LblNegDynCoeffAff.Text = CoeffAff

        TxtLimTours.Text = 20               'nombre de tours max de négociation par défaut
        TxtNbTirages.Text = 20              'nombre de tirage (cycle de programme) par défaut

        '///// DECISION DE REVELATION fixée pour une fois pour tous les tirages

        For i = 1 To 2                      '2 représente ici le nombre d'agents envisagées (il s'agit d'un indice)
            'Randomize()
            'Alea(i - 1) = Rnd() * Alea(i - 1)  'Calcul de la variation d'amplitude spécifiée pour chaque Agent

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
                    'TableauRev.Rows(j).Cells(i - 1).Value = "Mentir"
                Case ("Bluffer")
                    ActProposition(i - 1) = OptX(i - 1) + Alea(i - 1)           'Révèle l'OPTIMUM affecté de l'incertitude, augmenté d'une variation
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    'TableauRev.Rows(j).Cells(i - 1).Value = "Bluffer"
                Case "Dire la vérité"
                    ActProposition(i - 1) = OptX(i - 1)                         'Révèle l'OPTIMUM affecté de l'incertitude, sans changement
                    ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    'TableauRev.Rows(j).Cells(i - 1).Value = "Dire la vérité"
            End Select
        Next i

    End Sub

    Private Sub BtnAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAnnuler.Click
        Me.Close()
    End Sub

    Private Sub TxtNbTirages_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtNbTirages.TextChanged
        NbTirages = Valeur(TxtNbTirages.Text)
    End Sub

    Private Sub BtnAfficher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAfficher.Click
        Dim Resultats_statistiques As Statistique
        Dim i, j, jmax As Integer
        ReDim TableauResStats(NbTirages)

        '///Calcul des moyennes et des écart-types pour chaque Agent et pour les accords
        j = 0                                                                       'j représente l'indice de la ligne de la variable temporaire
        For i = 1 To NbTirages                                                      'i représente le numéro du tirage
            TableauResStats(i) = 0
            If NegoDyn(i, 0) <> 0 Then
                j = Inc(j)
                jmax = j
                TableauResStats(j) = NegoDyn(i, 0)                                   'variable temporaire pour stocker les résultats des calculs précèdants
            End If
        Next i

        Resultats_statistiques = MoyVariance(1, jmax - 1, TableauResStats)          'appel à la fonction à 1 dimension pour obtenir la moyenne et l'écart-type
        Ntours = jmax                                                               'stockage temporaire pour écriture

        With Resultats_statistiques
            LblMoyenneNbTours.Text = .Moyenne                                       'affichage de la moyenne 
            MoyenneNbTours = .Moyenne                                               'stockage temporaire pour écriture
            LblEcartTypeNbTours.Text = .EcartType                                   'affichage de l'écart-type 
            EcartTypeNbTours = .EcartType                                           'stockage temporaire pour écriture
        End With

        j = 0                                                                       'j représente l'indice de la ligne de la variable temporaire
        For i = 1 To NbTirages                                                      'i représente le numéro du tirage
            TableauResStats(i) = 0
            If NegoDyn(i, 1) <> 0 Then
                j = Inc(j)
                jmax = j
                TableauResStats(j) = NegoDyn(i, 1)                                   'variable temporaire pour stocker les résultats des calculs précèdants
            End If
        Next i

        Resultats_statistiques = MoyVariance(1, jmax - 1, TableauResStats)          'appel à la fonction à 1 dimension pour obtenir la moyenne et l'écart-type
        NAgt1 = jmax                                                                'stockage temporaire pour écriture

        With Resultats_statistiques
            LblValMoyenneAgt1.Text = .Moyenne * CoeffAff                            'affichage de la moyenne affecté du coefficient d'affichage
            MoyenneValAgt1 = .Moyenne                                               'stockage temporaire pour écriture 
            LblValEcartTypeAgt1.Text = .EcartType * CoeffAff                         'affichage de l'écart-type affecté du coefficient d'affichage
            EcartTypeValAgt1 = .EcartType                                           'stockage temporaire pour écriture
        End With

        j = 0
        For i = 1 To NbTirages                                                      'i représente le numéro du tirage
            TableauResStats(i) = 0
            If NegoDyn(i, 2) <> 0 Then
                j = Inc(j)
                jmax = j
                TableauResStats(j) = NegoDyn(i, 2)                                  'variable temporaire pour stocker les résultats des calculs précèdants
            End If
        Next i

        Resultats_statistiques = MoyVariance(1, jmax - 1, TableauResStats)          'appel à la fonction à 1 dimension pour obtenir la moyenne et l'écart-type
        NAgt2 = jmax                                                                'stockage temporaire pour écriture
        With Resultats_statistiques
            LblValMoyenneAgt2.Text = .Moyenne * CoeffAff                            'affichage de la moyenne affecté du coefficient d'affichage
            MoyenneValAgt2 = .Moyenne                                               'stockage temporaire pour écriture 
            LblValEcartTypeAgt2.Text = .EcartType * CoeffAff                         'affichage de l'écart-type affecté du coefficient d'affichage
            EcartTypeValAgt2 = .EcartType                                           'stockage temporaire pour écriture
        End With

        j = 0
        For i = 1 To NbTirages                                                      'i représente le numéro du tirage
            TableauResStats(i) = 0
            If NegoDyn(i, 4) <> 0 Then
                j = Inc(j)
                jmax = j
                TableauResStats(j) = NegoDyn(i, 4)                                   'variable temporaire pour stocker les résultats des calculs précèdants
            End If
        Next i

        Resultats_statistiques = MoyVariance(1, jmax - 1, TableauResStats)           'appel à la fonction à 1 dimension pour obtenir la moyenne et l'écart-type
        NAccords = jmax                                                              'stockage temporaire pour écriture
        With Resultats_statistiques
            LblValMoyenneTot.Text = .Moyenne * CoeffAff                              'affichage de la moyenne affecté du coefficient d'affichage
            MoyenneValAccord = .Moyenne                                              'stockage temporaire pour écriture 
            LblValEcartTypeTot.Text = .EcartType * CoeffAff                          'affichage de l'écart-type affecté du coefficient d'affichage
            EcartTypeValAccord = .EcartType                                          'stockage temporaire pour écriture
        End With

    End Sub
    Private Sub BtnCalculer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCalculer.Click
        Dim i, j, k, l, tir, IndProp, IndDcd, StratDcd As Integer
        Dim OptGauss(1) As Double
        jmax = Valeur(TxtLimTours.Text)

        TableauNegDyn.ColumnCount = 6
        TableauNegDyn.RowCount = NbTirages + 1

        For tir = 1 To NbTirages
            For j = 1 To jmax
                For i = 1 To 2
                    k = 2 - i               'k désigne l'agent B
                    l = 2 * j + i - 3       'l désigne le numéro du tour de négociation 
                    IndProp = i - 1         'Désigne i-1 comme le proposant
                    IndDcd = k              'Désigne k comme le décidant

                    OptGauss(i - 1) = GaussNumDist(Opt(i - 1), OptET(i - 1), 10)   'Tirage au sort d'une valeur normale de l'optimum du Proposant selon les moyennes et écart-types fournis dans un échantillon généré de taille 10
                    OptGauss(k) = GaussNumDist(Opt(k), OptET(k), 10)               'Tirage au sort d'une valeur normale de l'optimum du Décidant selon les moyennes et écart-types fournis dans un échantillon généré de taille 10

                    Select Case ActRevelation(i - 1)            '///ACTION DE REVELATION///
                        Case "Mentir"
                            ActProposition(i - 1) = MaxX(i - 1) + Alea(i - 1)           'Révèle le MAXIMUM affecté de l'incertitude, augmenté d'une variation
                            ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                        Case ("Bluffer")
                            ActProposition(i - 1) = OptGauss(i - 1) + Alea(i - 1)       'Révèle l'OPTIMUM affecté de l'incertitude, augmenté d'une variation
                            ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                        Case "Dire la vérité"
                            ActProposition(i - 1) = OptGauss(i - 1)                     'Révèle l'OPTIMUM affecté de l'incertitude, sans changement
                            ActProposition(i - 1) = Borner(ActProposition(i - 1))       'Borne de valeur de proposition à 1
                    End Select

                    If j = 1 Then                                              'Enregistrement de la valeur de la proposition initiale 
                        ActPropositionInitiale(IndProp) = ActProposition(IndProp)
                    End If


                    If i = 1 Then
                        StratDcd = CmbStratAgt1.SelectedIndex 'Désigne la stratégie choisie par le décidant 
                    Else
                        StratDcd = CmbStratAgt2.SelectedIndex 'Désigne la stratégie choisie par le décidant 
                    End If

                    Select Case StratDcd
                        Case 0                          'Stratégie Médiane indice 0
                            If ActProposition(IndProp) <= (OptGauss(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptGauss(IndDcd) - Alea(IndDcd)) Then
                                Accord = True
                                GoTo Sortie
                            Else
                                ActProposition(IndDcd) = (ActProposition(IndProp) + ActProposition(IndDcd)) / 2 'Proposition médiane au prochain tour par le Décidant
                                ActProposition(IndDcd) = ActVerifMax(ActProposition(IndDcd), IndDcd)    'Vérifie que la proposition médiane ne dépasse pas le maximum du Décidant
                            End If

                        Case 1                          'Stratégie Rigide indice 1 Si la proposition est dans la zone optimum plus ou moins l'aléa, la proposition est acceptée
                            If ActProposition(IndProp) <= (OptGauss(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptGauss(IndDcd) - Alea(IndDcd)) Then
                                Accord = True
                                GoTo Sortie
                            End If

                        Case 2                          'Stratégie Rigide avec Regret indice 2 : Si aucun accord n'est trouvé dans les derniers tours, la stratégie médiane est alors adoptée
                            If ActProposition(IndProp) <= (OptGauss(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptGauss(IndDcd) - Alea(IndDcd)) Then
                                Accord = True
                                GoTo Sortie
                            Else
                                If j >= jmax - 1 Then
                                    ActProposition(IndDcd) = (ActProposition(IndProp) + ActProposition(IndDcd)) / 2 'Proposition médiane au prochain tour par le Décidant
                                    ActProposition(IndDcd) = ActVerifMax(ActProposition(IndDcd), IndDcd)    'Vérifie que la proposition médiane ne dépasse pas le maximum du Décidant
                                End If
                            End If

                        Case 3                          'Stratégie Rigide avec Risque indice 3
                            If ActProposition(IndProp) <= (OptGauss(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptGauss(IndDcd) - Alea(IndDcd)) Then
                                Accord = True
                                GoTo Sortie
                            Else
                                If j >= jmax / 2 Then
                                    If ActVerifRisk(ActProposition(IndDcd), IndProp, Risk) Then
                                        GoTo Sortie
                                    End If
                                End If
                            End If

                        Case 4                          'Stratégie Rigide avec Regret et Risque indice 4
                            If ActProposition(IndProp) <= (OptGauss(IndDcd) + Alea(IndDcd)) And ActProposition(IndProp) >= (OptGauss(IndDcd) - Alea(IndDcd)) Then
                                Accord = True
                                GoTo Sortie
                            Else
                                If j >= jmax / 2 Then
                                    If ActVerifRisk(ActProposition(IndDcd), IndProp, Risk) Then
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


Sortie:     'Marque la fin du processus de négociation

            TableauNegDyn(0, tir - 1).Value = tir                                 'Affiche le numéro de ligne tout à gauche
            TableauNegDyn(1, tir - 1).Value = l + 1                                   'Affiche le nombre de tours de négociation de ligne tout à gauche
            TableauNegDyn(IndProp + 2, tir - 1).Value = ActPropositionInitiale(IndProp)   'Affiche la valeur initiale dans la colonne de l'agant proposant
            TableauNegDyn(IndDcd + 2, tir - 1).Value = ""
            TableauNegDyn(4, tir - 1).Value = Accord                              'Affiche le résultat de la négociation dans la 5e colonne
            TableauNegDyn(5, tir - 1).Value = ActProposition(IndProp)              'Affiche la valeur du résultat de la négociation 

            NegoDyn(tir - 1, 0) = l + 1
            NegoDyn(tir - 1, IndProp + 1) = ActPropositionInitiale(IndProp)
            NegoDyn(tir - 1, IndDcd + 1) = 0
            NegoDyn(tir - 1, 3) = Accord
            NegoDyn(tir - 1, 4) = ActProposition(IndProp)
        Next tir


    End Sub

    Private Sub CmbObjectifAgt2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbObjectifAgt2.SelectedIndexChanged
        Select Case CmbObjectifAgt2.SelectedIndex
            Case Is = 0
                Objectif(1) = False
            Case Else
                Objectif(1) = True
        End Select
    End Sub

    Private Sub CmbObjectifAgt1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbObjectifAgt1.SelectedIndexChanged
     
        Select Case CmbObjectifAgt1.SelectedIndex
            Case Is = 0
                Objectif(0) = False
            Case Else
                Objectif(0) = True
        End Select
    End Sub

    Private Sub TxtLimTours_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtLimTours.TextChanged
        jmax = Valeur(TxtLimTours.Text)
    End Sub
End Class