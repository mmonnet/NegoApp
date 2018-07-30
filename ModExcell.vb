Imports System
Imports System.ComponentModel
Imports System.Object
Imports System.Drawing
Imports System.String
Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.IO

Module ModExcell
    '*******************************************************
    'Définition des fonctions Excell du logiciel
    '*******************************************************
    Public Function lireDansFichierExcel(ByVal nomFichier As String) As Boolean
        Dim i, k, Ideb As Integer
        Dim str As String

        On Error GoTo errorFind

        xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True  ' Rendre visible l'application
        wkbObj = xlApp.Workbooks.Open(nomFichier) ' Ouvrir le fichier Excel et affectation
        i = 1
        k = 0

        For i = 1 To 2
            With wkbObj.Worksheets(i)
                Ideb = 2 'Ideb = numéro de ligne dans la feuille de calcul Excel
                If i = 1 Then
                    If CStr(.Range("B" & Ideb).Value) <> "" Then    ' Lecture de la date
                        DateTest = .Range("B" & Ideb).Value
                    End If
                    Ideb = Ideb + 1
                    Ideb = Ideb + 1
                    If CStr(.Range("B" & Ideb).Value) <> "" Then    ' Lecture du nom de l'essai
                        nomTest = .Range("B" & Ideb).Value
                    End If
                End If

                Ideb = 6            ' Lecture du Miminum requis (Min)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Min(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture du Maximum envisagé (Max)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Max(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture de l'Optimum recherché ou de sa moyenne (Opt)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Opt(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture de l'écart-type de l'Optimum recherché (OptET)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    OptET(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture de l'Aversion au risque (Risk)
                Ideb = Ideb + 1
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Risk(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture de l'Aversion au regret (Rgt)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Regt(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture du mandat de négociation (Mand)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Mand(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture du Degré de confiance (Conf)
                Ideb = Ideb + 1
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Conf(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     ' Lecture de l'Objectif de négociation (Objectif)
                If Valeur(.Range("B" & Ideb).Value) <> Nothing Then
                    Objectif(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1      ' Lecture du Degré d'incertitude (Cert)
                Ideb = Ideb + 1
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Cert(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1      ' Lecture de l'Amplitude des variations (Alea)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    Alea(i - 1) = Valeur(.Range("B" & Ideb).Value)
                End If

                Ideb = Ideb + 1     'Lecture du Coefficient d'affichage (CoeffAff)
                If Valeur(.Range("B" & Ideb).Value) > 0 Then
                    CoeffAff = Valeur(.Range("B" & Ideb).Value)
                End If

            End With

        Next i      ' Prochaine feuille de calcul

        ' Fermeture du fichier Excel
        'wkbObj.Close(SaveChanges:=False, Filename:=nomFichier)
        wkbObj.Close(SaveChanges:=False)

        ''wkbObj(nomFichier).close(savechanges:=False)

        'wkbObj.Close(nomFichier)
        wkbObj = Nothing

        ' Fermeture du fichier Excel
        xlApp.Visible = False
        xlApp.Quit()

        'For Each item As Process In Process.GetProcessesByName("EXCEL")
        '    processId += item.Id
        'Next

        'Process.GetProcessById(processId).Kill()

        ' Suppression des objets
        lireDansFichierExcel = True

        Exit Function

errorFind:
        lireDansFichierExcel = False
        wkbObj.Close()
        wkbObj = Nothing
        xlApp.Visible = False
        xlApp = Nothing

    End Function

    Public Function EcrireDansFichierExcel(ByVal NomFichier As String) As Boolean

        Dim xlApp As Excel.Application
        Dim wkbObj As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        'Dim xlSheetGraph As Excel.Chart
        Dim MyCharts As Excel.ChartObjects
        Dim i, j, k, iDeb, Ideb1, iLon As Integer
        Dim Str, Str1 As String
        Dim Lim1, Lim2 As Single
        On Error GoTo errorFind

        ' ouvrir l'application excel ; Ok
        xlApp = CreateObject("Excel.Application")

        xlApp.Visible = True  ' Rendre visible l'application

        ' Ouvrir le fichier Excel et affectation
        wkbObj = xlApp.Workbooks.Open(NomFichier)

        For i = 1 To 2
            With wkbObj.Worksheets(i)

                ' xlSheet = CType(wkbObj.Worksheets("Graphiques"), Excel.Worksheet)
                'MyCharts = xlSheet.ChartObjects
                'xlSheetGraph = CType(wkbObj.Worksheets(4), Excel.Chart)
                'xlSheet = CType(wkbObj.Worksheets("Graphiques"), Excel.Worksheet)
                ' recherche de la première occurence à droite du caractère "\"

                iDeb = 2
                .Range("B" & iDeb).Value = DateTest 'Ecriture de la date
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = NomFichier 'Ecriture du nom de fichier
                iDeb = Inc(iDeb)
                iDeb = Inc(iDeb)
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Min(i - 1) 'Ecriture de la variable Min
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Max(i - 1) 'Ecriture de la variable Max
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Opt(i - 1) 'Ecriture de la variable Opt
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = OptET(i - 1) 'Ecriture de la variable OptET
                iDeb = Inc(iDeb)
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Risk(i - 1) 'Ecriture de la variable Risk
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Regt(i - 1) 'Ecriture de la variable Regt
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Mand(i - 1) 'Ecriture de la variable Mand
                iDeb = Inc(iDeb)
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Conf(i - 1) 'Ecriture de la variable Conf
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Objectif(i - 1) 'Ecriture de la variable Objectif
                iDeb = Inc(iDeb)
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Cert(i - 1) 'Ecriture de la variable Cert
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = Alea(i - 1) 'Ecriture de la variable Alea
                iDeb = Inc(iDeb)
                .Range("B" & iDeb).Value = CoeffAff    'Ecriture du Coefficient d'affichage (CoeffAff)

            End With
        Next

        With wkbObj.Worksheets(3)
            For j = 0 To jmax
                .Range("A" & (j + 2)).Value = j + 1

                If Nego(j, 0) <> Nothing Then
                    .Range("B" & (j + 2)).Value = Nego(j, 0)
                Else
                    .Range("B" & (j + 2)).Value = ""
                End If
                If Nego(j, 1) <> Nothing Then
                    .Range("C" & (j + 2)).Value = Nego(j, 1)
                Else
                    .Range("C" & (j + 2)).Value = ""
                End If
                If Nego(j, 2) <> Nothing Then
                    .Range("D" & (j + 2)).Value = Nego(j, 2)
                Else
                    .Range("D" & (j + 2)).Value = ""
                End If
            Next
        End With

        With wkbObj.Worksheets(4)

            'Ecriture de toutes les moyennes
            .Range("B" & 1).Value = MoyenneNbTours
            .Range("C" & 1).Value = MoyenneValAgt1
            .Range("D" & 1).Value = MoyenneValAgt2
            .Range("F" & 1).Value = MoyenneValAccord

            'Ecriture de tous les écarts-types
            .Range("B" & 2).Value = EcartTypeNbTours
            .Range("C" & 2).Value = EcartTypeValAgt1
            .Range("D" & 2).Value = EcartTypeValAgt2
            .Range("F" & 2).Value = EcartTypeValAccord

            'Ecriture de tous les N
            .Range("B" & 3).Value = Ntours
            .Range("C" & 3).Value = NAgt1
            .Range("D" & 3).Value = NAgt2
            .Range("F" & 3).Value = NAccords
            j = 6

            For i = 0 To NbTirages
                .Range("A" & j).Value = j - 5                   'Ecriture du numéro du tirage
                .Range("B" & j).Value = NegoDyn(i, 0)           'Ecriture du nombre de tours de négociation effectué
                .Range("C" & j).Value = NegoDyn(i, 1)           'Ecriture de la proposition initiale si faite par l'agent 1
                .Range("D" & j).Value = NegoDyn(i, 2)           'Ecriture de la proposition initiale si faite par l'agent 2
                .Range("E" & j).Value = NegoDyn(i, 3)           'Ecriture du résultat de la négociation, 0 = Absence, 1 = Accord
                .Range("F" & j).Value = NegoDyn(i, 4)           'Ecriture de la valeur sur laquelle les agents se sont accordés dans le cas d'un accord
                j = Inc(j)
            Next i

        End With

        wkbObj.Save()
        xlApp.Visible = False
        wkbObj.Close(NomFichier) 'Fermeture du fichier
        xlApp.Quit()
        wkbObj = Nothing
        xlApp = Nothing

        'Process.GetProcessById(processId).Kill()
        EcrireDansFichierExcel = True

        Exit Function

errorFind:
        MsgBox("Erreur d'enregistrement : " & Err.Description)
        EcrireDansFichierExcel = False
        wkbObj.Close(NomFichier)
        wkbObj = Nothing
        xlApp.Visible = False
        xlApp.Quit()
        xlApp = Nothing


    End Function

End Module
