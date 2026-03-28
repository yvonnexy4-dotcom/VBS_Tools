Dim num1, num2, operation, result

' Input of the first number
num1 = InputBox("Enter the first number:", "VBS Calculator")

' Input of the operator
operation = InputBox("Choose an operation (+, -, *, /):", "VBS Calculator")

' Input of the second number
num2 = InputBox("Enter the second number:", "VBS Calculator")

' Calculation based on selection
Select Case operation
    Case "+"
        result = CDbl(num1) + CDbl(num2)
    Case "-"
        result = CDbl(num1) - CDbl(num2)
    Case "*"
        result = CDbl(num1) * CDbl(num2)
    Case "/"
        If num2 = "0" Then
            result = "Error: Division by zero!"
        Else
            result = CDbl(num1) / CDbl(num2)
        End If
    Case Else
        result = "Invalid operation"
End Select

' Displaying the result
MsgBox "The result is: " & result, vbInformation, "Result"