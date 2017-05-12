Option Explicit

Sub Problem6()

Dim increase As Integer
Dim minrent As Double
Dim profit As Double
Dim profitc As Double
Dim number As Integer
Dim bestnum As Integer
Dim count As Integer

count = 0
number = 50
profitc = 1

minrent = InputBox("Enter price for max apartment occupation:")
increase = InputBox("Enter price increase for vacancy:")

For count = count To number Step 1
profit = (minrent + (increase * count)) * number
If profit > profitc Then
    profitc = profit
    bestnum = number
    End If
number = number - 1
Next
MsgBox ("The number of units: for best profit " & bestnum & ". Monthly Earnings: " & profitc & "$")

End Sub
