Module DiscreteCashFlow

    Function CompoundInterest(i As Double, n As Integer) As Double
        CompoundInterest = (1 + i) ^ n
    End Function

    Function CompoundMin1(i As Double, n As Integer) As Double
        CompoundMin1 = CompoundInterest(i, n) - 1
    End Function

    Function FGivenP(i As Double, n As Integer) As Double
        FGivenP = CompoundInterest(i, n)
    End Function

    Function PGivenF(i As Double, n As Integer) As Double
        PGivenF = 1 / FGivenP(i, n)
    End Function

    Function AGivenF(i As Double, n As Integer) As Double
        'AGivenF = i / (((1 + i) ^ n) - 1)
        AGivenF = i / CompoundMin1(i, n)
    End Function

    Function FGivenA(i As Double, n As Integer) As Double
        'FGivenA = (((1 + i) ^ n) - 1) / i
        'FGivenA = 1 / AGivenF ?
        FGivenA = CompoundMin1(i, n) / i
    End Function

    Function AGivenP(i As Double, n As Integer) As Double
        'AGivenP = (i * ((1 + i) ^ n)) / (((1 + i) ^ n) - 1)
        AGivenP = (i * CompoundInterest(i, n)) / CompoundMin1(i, n)
    End Function

    Function PGivenA(i As Double, n As Integer) As Double
        'PGivenA = (((1 + i) ^ n) - 1) / (i * ((1 + i) ^ n))
        'PGivenA = 1 / AGivenP
        PGivenA = CompoundMin1(i, n) / (i * CompoundInterest(i, n))
    End Function

    Function AGivenG(i As Double, n As Integer) As Double
        ' AGivenG = (1 / i) - (n / (((1 + i) ^ n) - 1))
        AGivenG = (1 / i) - (n / CompoundMin1(i, n))
    End Function

    Function GeoGradient(i As Double, n As Integer, g As Double)
        Dim io As Double
        io = (1 + i) / (1 + g) - 1
        GeoGradient = PGivenA(io, n)
    End Function

End Module
