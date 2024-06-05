Attribute VB_Name = "CashFlowFuncs"
Function CompoundInterest(i As Double, N As Integer) As Double
   If N > 0 Then
        CompoundInterest = (1 + i) ^ N
    Else
        CompoundInterest = 1
    End If
End Function

Function CompoundMin1(i As Double, N As Integer) As Double
    CompoundMin1 = CompoundInterest(i, N) - 1
End Function

Function FGivenP(i As Double, N As Integer) As Double
    FGivenP = CompoundInterest(i, N)
End Function

Function PGivenF(i As Double, N As Integer) As Double
    PGivenF = 1 / FGivenP(i, N)
End Function

Function AGivenF(i As Double, N As Integer) As Double
    'AGivenF = i / (((1 + i) ^ n) - 1)
    AGivenF = i / CompoundMin1(i, N)
End Function

Function FGivenA(i As Double, N As Integer) As Double
    'FGivenA = (((1 + i) ^ n) - 1) / i
    'FGivenA = 1 / AGivenF ?
    FGivenA = CompoundMin1(i, N) / i
End Function

Function AGivenP(i As Double, N As Integer) As Double
    'AGivenP = (i * ((1 + i) ^ n)) / (((1 + i) ^ n) - 1)
    If CompoundMin1(i, N) > 0 Then
        AGivenP = (i * CompoundInterest(i, N)) / CompoundMin1(i, N)
    Else
        AGivenP = 1
    End If
End Function

Function PGivenA(i As Double, N As Integer) As Double
    'PGivenA = (((1 + i) ^ n) - 1) / (i * ((1 + i) ^ n))
    'PGivenA = 1 / AGivenP
    PGivenA = CompoundMin1(i, N) / (i * CompoundInterest(i, N))
End Function

Function AGivenG(i As Double, N As Integer) As Double
   ' AGivenG = (1 / i) - (n / (((1 + i) ^ n) - 1))
   AGivenG = (1 / i) - (N / CompoundMin1(i, N))
End Function

Function GeoGradient(i As Double, N As Integer, g As Double)
    Dim io As Double
    io = (1 + i) / (1 + g) - 1
    GeoGradient = PGivenA(io, N) / (1 + g)
End Function

Function BookValDB(P As Double, d As Double, N As Integer)
    BookValDB = P * (1 - d) ^ N
End Function

Function AcDepDB(P As Double, d As Double, N As Integer)
    AcDepDB = P - BookValDB(P, d, N)
End Function
