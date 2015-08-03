Public Function CAPITALANDINTEREST(Term, InterestRate, LoanAmount) As Currency
  ' Calculate the monthly payments for a repayment mortgage (capital and interest)
  ' http://www.haresoftware.com/xlent_udf_repaymentmortgage.htm
  Dim First As Single, Second As Single, intTerm As Integer, LoanPercent As Double

  On Error GoTo Err_CAI

  LoanPercent = InterestRate / 100
  intTerm = Term * 12

  First = LoanAmount * (LoanPercent / 12) * (1 + (LoanPercent / 12)) ^ intTerm
  Second = ((1 + (LoanPercent / 12)) ^ intTerm) - 1

  CAPITALANDINTEREST = First / Second

  Exit Function

Err_CAI:
  CAPITALANDINTEREST = 0#

End Function

