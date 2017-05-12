Option Explicit

Sub Problem5()
Dim salep As String
Dim salen As Double
Dim earnings As Double

Do While salep <> "X" Or salen <> -1
salep = InputBox("Enter salesperson's name:")
salen = InputBox("Enter salesperson's earning:")
    If salep <> "X" Or salen <> -1 Then
        earnings = salen * 0.095 + 550
        MsgBox (salep & " sold " & salen & "$. Week's Earnings: " & earnings & "$")
    Else
        MsgBox ("Bye")
    End If
Loop
End Sub
