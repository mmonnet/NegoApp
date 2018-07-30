Imports System.Math
Module ModGauss
    Friend GaussNumArray() As Double
    Friend intICell As Long

    Friend Function GaussNumDist(ByVal Mean As Double, ByVal StdDev As Double, ByVal SampleSize As Integer) As Double
        intICell = 1                'Loop variable

        ReDim GaussNumArray(SampleSize)

        Do While (intICell < (SampleSize))
            Call NumDist(Mean, StdDev)
            Application.DoEvents()
        Loop
        GaussNumDist = GaussNumArray(5)             'Prélève la 5e valeur aléatoire de la distribution normale générée de taille et d'écart-type spécifié par l'utilisateur
    End Function

    Sub NumDist(ByVal meanin As Double, ByVal sdin As Double)
        '---------------------------------------------------------------------------------
        'Converts uniform random numbers over the region 0 to 1 into Gaussian distributed
        'random numbers using Box-Muller algorithm.
        'Adapted from Numerical Recipes in C
        '---------------------------------------------------------------------------------

        'Defining variables
        Dim dblR1 As Double
        Dim dblR2 As Double
        Dim mean As Double
        Dim var As Double
        Dim circ As Double
        Dim trans As Double
        Dim dblY1 As Double
        Dim dblY2 As Double
        Dim Pi As Double
        Pi = 4 * Atan(1)

        'Get two random numbers
        dblR1 = (2 * UniformRandomNumber()) - 1
        dblR2 = (2 * UniformRandomNumber()) - 1

        circ = (dblR1 ^ 2) + (dblR2 ^ 2)        'Radius of circle

        If circ >= 1 Then       'If outside unit circle, then reject number
            Call NumDist(meanin, sdin)
            Exit Sub
        End If

        'Transform to Gaussian
        trans = Sqrt(-2 * Log(circ) / circ)

        dblY1 = (trans * dblR1 * sdin) + meanin
        dblY2 = (trans * dblR2 * sdin) + meanin

        GaussNumArray(intICell) = dblY1   'First number

        'Increase intICell for next random number
        intICell = (intICell + 1)

        GaussNumArray(intICell) = dblY2   'Second number

        'Increase intICell again ready for next call of ConvertNumberDistribution
        intICell = (intICell + 1)

    End Sub

    Friend Function UniformRandomNumber() As Double
        '-----------------------------------------------------------------------------------
        'Outputs random numbers with a period of > 2x10^18 in the range 0 to 1 (exclusive)
        'Implements a L'Ecuyer generator with Bays-Durham shuffle
        'Adapted from Numerical Recipes in C
        '-----------------------------------------------------------------------------------

        'Defining constants
        Const IM1 As Double = 2147483563
        Const IM2 As Double = 2147483399
        Const AM As Double = (1.0# / IM1)
        Const IMM1 As Double = (IM1 - 1.0#)
        Const IA1 As Double = 40014
        Const IA2 As Double = 40692
        Const IQ1 As Double = 53668
        Const IQ2 As Double = 52774
        Const IR1 As Double = 12211
        Const IR2 As Double = 3791
        Const NTAB As Double = 32
        Const NDIV As Double = (1.0# + IM1 / NTAB)
        Const ESP As Double = 0.00000012
        Const RNMX As Double = (1.0# - ESP)

        Dim iCell As Integer
        Dim idum As Double
        Dim j As Integer
        Dim k As Long
        Dim temp As Double

        Static idum2 As Long
        Static iy As Long
        Static iv(NTAB) As Long

        idum2 = 123456789
        iy = 0

        'Seed value required is a negative integer (idum)
        Randomize()
        idum = (-Rnd() * 1000)

        'For loop to generate a sequence of random numbers based on idum
        For iCell = 1 To 10
            'Initialize generator
            If (idum <= 0) Then
                'Prevent idum = 0
                If (-(idum) < 1) Then
                    idum = 1
                Else
                    idum = -(idum)
                End If
                idum2 = idum
                For j = (NTAB + 7) To 0
                    k = ((idum) / IQ1)
                    idum = ((IA1 * (idum - (k * IQ1))) - (k * IR1))
                    If (idum < 0) Then
                        idum = (idum + IM1)
                    End If
                    If (j < NTAB) Then
                        iv(j) = idum
                    End If
                Next j
                iy = iv(0)
            End If

            'Start here when not initializing
            k = (idum / IQ1)
            idum = ((IA1 * (idum - (k * IQ1))) - (k * IR1))
            If (idum < 0) Then
                idum = (idum + IM1)
            End If
            k = (idum2 / IQ2)
            idum2 = ((IA2 * (idum2 - (k * IQ2))) - (k * IR2))
            If (idum2 < 0) Then
                idum2 = idum2 + IM2
            End If
            j = (iy / NDIV)
            iy = (iv(j) - idum2)
            iv(j) = idum
            If (iy < 1) Then
                iy = (iy + IMM1)
            End If
            temp = AM * iy
            If (temp <= RNMX) Then
                'Return the value of the random number
                UniformRandomNumber = temp
            End If
        Next iCell
    End Function
End Module
