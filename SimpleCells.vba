Sub CellExample()

Dim fint As Integer
Dim sum As Integer
Dim number As Integer
Dim num2 As Integer

num2 = 1
fint = InputBox(" Please enter a positive integer less than 100")

For number = 0 To fint
If number Mod 2 <> 0 Then
sum = sum + number
Cells(num2, "B").Value = number
num2 = num2 + 1
End If
Next

Cells(num2 + 1, "A").Value = "Sum"
Cells(num2, "B").Value = "------------"
Cells(num2 + 1, "B").Value = sum

End Sub
