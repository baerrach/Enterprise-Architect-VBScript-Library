'[group=ArchiMateModelValidation]
option explicit

'
' VB6 is such a lack lustre language.
' No inheritance. Probably could do something roll-your own ala JavaScript.
' But since everything is a variant and what I am looking at is code-generation
' as long as the generated code has the correct function/method signatures it will just work.
'

Dim BonusRate, PayRate
BonusRate = 1.45
PayRate = 14.75

Class Payroll
	Function PayEmployee(HoursWorked, PayRate)
		PayEmployee = HoursWorked * Payrate
	End Function
End Class

Class BonusPayroll
	dim baseClass

	Private Sub Class_Initialize
		set baseClass = new Payroll
	End Sub

	Function PayEmployee(HoursWorked, PayRate)
		PayEmployee = baseClass.PayEmployee(HoursWorked, PayRate) * BonusRate
	End Function
End Class

Sub RunPayroll()
  Dim PayrollItem, BonusPayrollItem, HoursWorked
  set PayrollItem = new Payroll
  set BonusPayrollItem = new BonusPayroll
  HoursWorked = 40
  
  Dim list()
  redim list(1)
  set list(0) = PayrollItem
  set list(1) = BonusPayrollItem
  
  Dim item
  for each item in list
		MsgBox("item pay is: " & item.PayEmployee(HoursWorked, PayRate))
  next
End Sub

RunPayroll()
