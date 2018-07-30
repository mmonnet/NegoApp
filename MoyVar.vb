Module ModMoyVar
    '*********************************************
    'fonction qui calcule la moyenne de l'ecart-
    'type d'une variable X qui prend K valeur
    '*********************************************
    Public Function MoyVariance(ByVal i1 As Integer, ByVal I2 As Integer, ByRef x() As Single) As Statistique
        Dim Moy, Sigma As Single
        Dim i, N As Integer

        N = I2 - i1 + 1
        Moy = 0
        If N >= 1 Then
            For i = i1 To I2
                Moy = Moy + x(i)
            Next i
            Moy = Moy / N

            Sigma = 0
            For i = i1 To I2
                Sigma = Sigma + (x(i) - Moy) * (x(i) - Moy)
            Next i

            Sigma = Math.Sqrt(Sigma / N)
            MoyVariance.Moyenne = Moy
            MoyVariance.EcartType = Sigma
            MoyVariance.Erreur = False
        Else
            MoyVariance.Erreur = True
        End If
    End Function

End Module
