Public Function GEARING(DepositVal, RepayOptionStr, InterestRate, Term, colRent, flatRent, colValue, flatValue) As Double
  Dim colMort As Double, colCapAndInt As Double
  Dim colInt As Double, colCap As Double
  Dim colIntOnlyProfit As Double, colCapAndIntProfit As Double
  Dim flatMort As Double, flatCapAndInt As Double
  Dim flatInt As Double, flatCap As Double
  Dim flatIntOnlyProfit As Double, flatCapAndIntProfit As Double
  Dim depVal As Double, repayOpt As String
    
  On Error GoTo Err_CAI

  ' Don't use function arguments in "if" tests
  depVal = DepositVal
  repayOpt = RepayOptionStr
  
  colMort = colValue - DepositVal
  colCapAndInt = CAPITALANDINTEREST(Term, InterestRate, colMort)
  colInt = (InterestRate / 100) * colMort / 12
  colCap = colCapAndInt - colInt
  colIntOnlyProfit = colRent - colInt
  colCapAndIntProfit = colRent - colCapAndInt
  
  flatMort = flatValue - DepositVal
  flatCapAndInt = CAPITALANDINTEREST(Term, InterestRate, flatMort)
  flatInt = (InterestRate / 100) * flatMort / 12
  flatCap = flatCapAndInt - flatInt
  flatIntOnlyProfit = flatRent - flatInt
  flatCapAndIntProfit = flatRent - flatCapAndInt
  If depVal = 30000 Then
    flatIntOnlyProfit = flatIntOnlyProfit * 4
    flatCapAndIntProfit = flatCapAndIntProfit * 4
  ElseIf depVal = 37500 Then
    flatIntOnlyProfit = flatIntOnlyProfit * 3
    flatCapAndIntProfit = flatCapAndIntProfit * 3
  ElseIf depVal = 50000 Then
    flatIntOnlyProfit = flatIntOnlyProfit * 2
    flatCapAndIntProfit = flatCapAndIntProfit * 2
  ElseIf depVal = 75000 Then
    flatIntOnlyProfit = flatIntOnlyProfit * 1
    flatCapAndIntProfit = flatCapAndIntProfit * 1
  Else
    flatIntOnlyProfit = 0
    flatCapAndIntProfit = 0
  End If
  
  If repayOpt = "capAndInt" Then
    GEARING = colCapAndIntProfit + flatCapAndIntProfit
  ElseIf repayOpt = "intOnly" Then
    GEARING = colIntOnlyProfit + flatIntOnlyProfit
  Else
    GEARING = 0
  End If
  

  Exit Function

Err_CAI:
  GEARING = 0#

End Function


