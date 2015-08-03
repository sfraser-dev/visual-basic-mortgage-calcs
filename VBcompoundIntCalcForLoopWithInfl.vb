Public Function COMPOUNDCALCINFFOR(Base, InterestRate, Years, RegMonthly, InfRate) As Double
  ' Adding interest every month
  ' Increasing regular deposit with inflation each year
  ' http://math.stackexchange.com/questions/1265528/compounded-interest-with-exponentially-increasing-periodic-payments
  Dim Index As Integer
  Dim Deposit1 As Double, Deposit2 As Double, Deposit3 As Double, Deposit4 As Double
  Dim Deposit5 As Double, Deposit6 As Double, Deposit7 As Double, Deposit8 As Double
  Dim Deposit9 As Double, Deposit10 As Double, Deposit11 As Double, Deposit12 As Double
  
  On Error GoTo Err_Jump

  Deposit12 = Base
  For Index = 1 To Years
    'RegMonthly = RegMonthly * (1 + InterestRate)               ' Yearly interest
    'RegMonthly = RegMonthly * (1 + InterestRate / N) ^ N       ' Monthly interest
    Deposit1 = (Deposit12 + RegMonthly) * (1 + InterestRate / 12)
    Deposit2 = (Deposit1 + RegMonthly) * (1 + InterestRate / 12)
    Deposit3 = (Deposit2 + RegMonthly) * (1 + InterestRate / 12)
    Deposit4 = (Deposit3 + RegMonthly) * (1 + InterestRate / 12)
    Deposit5 = (Deposit4 + RegMonthly) * (1 + InterestRate / 12)
    Deposit6 = (Deposit5 + RegMonthly) * (1 + InterestRate / 12)
    Deposit7 = (Deposit6 + RegMonthly) * (1 + InterestRate / 12)
    Deposit8 = (Deposit7 + RegMonthly) * (1 + InterestRate / 12)
    Deposit9 = (Deposit8 + RegMonthly) * (1 + InterestRate / 12)
    Deposit10 = (Deposit9 + RegMonthly) * (1 + InterestRate / 12)
    Deposit11 = (Deposit10 + RegMonthly) * (1 + InterestRate / 12)
    Deposit12 = (Deposit11 + RegMonthly) * (1 + InterestRate / 12)
    
    RegMonthly = RegMonthly * (1 + InfRate)
  Next
  
  COMPOUNDCALCINFFOR = Deposit12
  
  Exit Function

Err_Jump:
  COMPOUNDCALCINFFOR = 0#

End Function

