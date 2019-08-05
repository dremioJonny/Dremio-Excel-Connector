Attribute VB_Name = "PrintData"

Function PrintToWS(ByVal data As Variant, _
    Optional rowi As Double = 1, _
    Optional roworig As Integer = 1, _
    Optional keyno As Integer = 0, _
    Optional coli As Integer = 1, _
    Optional wasdict As Integer = 0) As Double

    Dim key
    Dim j As Integer

    If (TypeName(data) = "Dictionary") Then
        For Each key In data.Keys:
            Cells(roworig + wasdict, coli + keyno).Value = key
            If roworig >= rowi Then
                rowi = PrintToWS(data.Item(key), rowi, roworig, keyno, coli + wasdict + keyno, 1)
            Else
                rowi = PrintToWS(data.Item(key), rowi, roworig, keyno, coli + wasdict + keyno, 1)
            End If
            keyno = keyno + 1
        wasdict = 0
        Next
    ElseIf (TypeName(data) = "Variant()") Then
        For j = LBound(data) To UBound(data)
            rowi = PrintToWS(data(j), rowi, roworig, keyno, coli)
        Next j
    Else
        If Not (IsEmpty(Cells(rowi, coli))) Then
            rowi = rowi + 1
        End If
        Cells(rowi, coli).Value = data
    End If

    PrintToWS = rowi
End Function

