
Module Base36
    Public Function DecToBase36(ByVal iDecimalValue As Long) As String

        Dim iInput As Long
        Dim iPower As Int16
        Dim iTemp As Long
        Dim cB36Digit(36) As Char
        Dim cB36Result(10) As Char
        Dim sB36Char As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        For iTemp = 0 To 35
            cB36Digit(iTemp) = Mid(sB36Char, iTemp + 1, 1)
        Next

        Try
            If iDecimalValue > 3656158440062975 Then
                MsgBox("Only numbers 0 - 3,656,158,440,062,975 allowed")
                Exit Function
            End If
            For iPower = 6 To 0 Step -1
                '
                iTemp = Int(iDecimalValue / (36 ^ iPower))
                DecToBase36 = DecToBase36 & cB36Digit(iTemp)
                cB36Result(iPower) = cB36Digit(iTemp)

                iDecimalValue = iDecimalValue - (iTemp * (36 ^ iPower))
            Next
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        End Try

    End Function

    Public Function Base36ToDecimal(ByVal sBase36Value As String) As Long
        Dim sB36Char As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        sBase36Value = sBase36Value.PadLeft(7, "0")
        Dim i As Int16
        Dim cPlace As Char
        Dim iPos As Int16
        Dim iDecimal As Long
        Dim iPower As Int16
        sBase36Value = UCase(sBase36Value)
        For i = 6 To 0 Step -1
            iPos = InStr(sB36Char, Mid(sBase36Value, i + 1, 1)) - 1
            iDecimal = iDecimal + (36 ^ (6 - i) * iPos)
            'Debug.Write()
            'Debug.Write(Mid(sBase36Value, i + 1, 1) & vbTab & iPos & vbTab _
            '& (36 ^ (6 - i) * iPos _
            ' & vbCrLf))
        Next
        Base36ToDecimal = iDecimal
    End Function

    Public Function AddOneToBase36(ByVal sBase36Value As String) As String

        Dim iTemp As Long

        If IsNothing(sBase36Value) Then sBase36Value = "0"

        AddOneToBase36 = DecToBase36(Base36ToDecimal(sBase36Value) + 1)


    End Function

    Public Function AddToBase36(ByVal sBase36Value As String, ByVal NumberToAdd As Long) As String

        Dim iTemp As Long

        If IsNothing(sBase36Value) Then sBase36Value = "0"

        AddToBase36 = DecToBase36(Base36ToDecimal(sBase36Value) + NumberToAdd)


    End Function

End Module
