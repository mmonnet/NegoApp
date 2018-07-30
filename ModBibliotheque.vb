Imports System.Math
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel


Module ModBibliotheque
    '***********************************
    'Déclaration des variables communes 
    '***********************************
    ' Variable d'aide
    Public helpfile, helpcontent, helpindex As String

    'Excel
    Public xlApp As Excel.Application
    Public wkbObj As Excel.Workbook
    Public xlSheet As Excel.Worksheet

    'Définitions des variables
    Public nomFichier, nomFichierFinal, nomTest As String
    Public DateTest As Date
    Public ChxStrat As Integer
    Public Structure Statistique
        Public Moyenne As Single
        Public EcartType As Single
        Public Erreur As Boolean
    End Structure

    'Définition des variables tableau
    Public n = 2                'Nombre d'agents participant à la négociation
    Public jmax As Integer      'Nombre de tours de négociations maximum
    Public NbTirages As Integer 'Nombre de tirages maximum
    Public Min(n - 1) As Single 'Minimum 
    Public Max(n - 1) As Single 'Maximum espéré 
    Public Opt(n - 1) As Single 'Optimum recherché ou Moyenne de l'optimum recherché
    Public OptET(n - 1) As Single 'Ecart-type de l'optimum recherché
    Public MinX(n - 1) As Single 'Minimum requis requis affecté par l'incertitude
    Public MaxX(n - 1) As Single 'Maximum espéré affecté par l'incertitude
    Public OptX(n - 1) As Single 'Optimum recherché affecté par l'incertitude
    Public Risk(n - 1) As Single 'Aversion au risque
    Public Regt(n - 1) As Single 'Aversion au regret
    Public Mand(n - 1) As Single 'Etendue du mandat des agents
    Public Conf(n - 1) As Single 'Degré de confiance
    Public Objectif(n - 1) As Boolean 'Objectif de la négociation
    Public Cert(n - 1) As Single 'Degré d'incertitude
    Public Alea(n - 1) As Single 'Amplitude des variations
    ' Public Alea(n - 1) As Single  'Variation aléatoire de 0 à Alea(n-1)
    Public CoeffAff As Single    'Coefficient d'affichage des données finales

    'Définition des dimensions de la variable qui affichée dans le tableau de négociation
    Public Nego(21, 2) As String '21+1 lignes, 2+1 colonnes
    Public NegoDyn(100000, 4) As String '100 000+1 lignes, 4+1 colonnes

    'Définition des actions 
    Public ActRevelation(n - 1) As String 'Action de révélation
    Public ActProposition(n - 1) As Single 'Action de proposition  
    Public ActPropositionInitiale(n - 1) As Single 'Action de proposition initiale 
    Public ActNegociation(n - 1) As Single 'Action de negociation
    Public Accord As Boolean 'Résultat de négociation True or False


    'Définition des variables temporaires pour l'écriture des données des négociations dynamiques
    Public EcartTypeValAgt1, EcartTypeValAgt2, EcartTypeValAccord, EcartTypeNbTours, MoyenneValAgt1, MoyenneValAgt2, MoyenneValAccord, MoyenneNbTours As Single
    Public Ntours, NAgt1, NAgt2, NAccords As Integer


    '*****************************************
    'Définition des fonctions de négociations
    '*****************************************
    'Fonction chargée de vérifier que la valeur proposée soit supérieure ou égale au minimum acceptable
    Public Function ActVerifMin(ByVal ValProp As Single, ByVal IndDcd As Integer) As Single
        If ValProp <= MinX(IndDcd) Then
            ActVerifMin = MinX(IndDcd)
        Else
            ActVerifMin = ValProp
        End If
    End Function

    'Fonction chargée de vérifier que la valeur proposée soit inférieure ou égale au maximum acceptable
    Public Function ActVerifMax(ByVal ValProp As Single, ByVal IndDcd As Integer) As Single
        If ValProp >= MaxX(IndDcd) Then
            ActVerifMax = MaxX(IndDcd)
        Else
            ActVerifMax = ValProp
        End If
    End Function

    'Fonction chargée de vérifier que la valeur proposée soit inférieure ou égale au mandat accepté
    Public Function ActVerifMandMax(ByVal ValProp As Single, ByVal IndDcd As Integer) As Boolean
        If ValProp <= Mand(IndDcd) Then
            ActVerifMandMax = True
        Else
            ActVerifMandMax = False
        End If
    End Function

    'Fonction chargée de vérifier que la valeur proposée entre dans les limites fixées par Min et Max du Décidant
    Public Function ActVerifRisk(ByVal ValProp As Single, ByVal IndDcd As Integer, ByRef Risk() As Single) As Boolean
        Dim nombre_Aleatoire, Sigma As Single

        Sigma = 0.5 / 2.7             '2.7 * écart-type = la tolérance, 50% des valeurs normales

        If ValProp < MinX(IndDcd) Or ValProp > MaxX(IndDcd) Then
            Randomize()
            nombre_Aleatoire = GaussNumDist(0.5, Sigma, 10) 'Nombre aléatoire entre 0 et 1 obéissant à une loi normale (Moyenne, Ecart-type, taille échantillon)

            If nombre_Aleatoire <= (1 - Risk(IndDcd)) Then
                ActVerifRisk = True
            Else
                ActVerifRisk = False
            End If
        Else
            ActVerifRisk = False
        End If
    End Function

    '*******************************************************
    'Définition des fonctions d'enregistrement des fichiers
    '*******************************************************
    ' Retourne le nom du fichier sans extension filename à partir de son arborescence 
    Public Function GetFilenameWithoutExtension(ByVal path As String) As String
        Dim pos As Integer
        Dim filename As String
        pos = InStrRev(path, "\")
        If pos > 0 Then
            filename = Mid$(path, pos + 1, Len(path))
            GetFilenameWithoutExtension = Left(filename, Len(filename) - Len(Mid$(filename, InStrRev(filename, "."), Len(filename))))
        Else
            GetFilenameWithoutExtension = ""
        End If
    End Function

    ' Retourne le nom du fichier avec extension filename à partir de son arborescence 
    Public Function GetFilenameWithExtension(ByVal path As String) As String
        Dim pos As Integer
        pos = InStrRev(path, "\")
        If pos > 0 Then
            GetFilenameWithExtension = Mid$(path, pos + 1, Len(path))
        Else
            GetFilenameWithExtension = ""
        End If
    End Function

    ' Retourne l'arborescence  du fichier 
    Public Function GetDirectoryFromPathFilename(ByVal path As String) As String
        Dim pos As Integer
        pos = InStrRev(path, "\")
        If pos > 0 Then
            GetDirectoryFromPathFilename = Left$(path, pos)
        Else
            GetDirectoryFromPathFilename = ""
        End If
    End Function
    'Détecte les chiffres à virgule dans un fichier excell pour les convertir en nombre décimal 
    Public Function Valeur(ByVal Chaine As String) As Single
        If Chaine = "," Then
            Valeur = 0
            Exit Function
        End If
        If InStr(Chaine, ",") = 0 Then
            Valeur = Val(Chaine)
        Else
            Valeur = CDbl(Chaine)
        End If
    End Function
    ' Incrémente la valeur de 1
    Public Function Inc(ByVal Valeur As Integer) As Integer
        Inc = Valeur + 1
    End Function
    'Borne la valeur entre 0 et 1
    Public Function Borner(ByVal Valeur As Single) As Single
        Select Case Valeur
            Case Is > 1
                Borner = 1
            Case Is < 0
                Borner = 0
            Case Else
                Borner = Valeur
        End Select
    End Function
    'Retourne une valeur comprise entre 0 et 1 inclusifs (!)
    'Public Function RandomSingle(ByVal MinValue As Single, ByVal MaxValue As Single) As Single
    '    SyncLock 

    '        Random.Next(MinValue, MaxValue)

    '    End SyncLock
    'End Function
End Module
