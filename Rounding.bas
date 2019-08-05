Attribute VB_Name = "Rounding"
Function Round_Up(ByVal d As Double) As Integer
    Dim result As Integer
    result = Math.Round(d)
    If result >= d Then
        Round_Up = result
    Else
        Round_Up = result + 1
    End If
End Function
