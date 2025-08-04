Public Module PwCrypt

    Public Function EnCryptDsnPassword(ByVal sDsn As String) As String
        Return CryptDsnPassword(sDsn, True)
    End Function

    Public Function DeCryptDsnPassword(ByVal sDsn As String) As String
        Return CryptDsnPassword(sDsn, False)
    End Function

    Public Function DeCryptUserPassword(ByVal sPass As String) As String
        Return criptaPW(sPass, False)
    End Function

    Public Function CryptUserPassword(ByVal sPass As String) As String
        Return criptaPW(sPass, True)
    End Function

    Private Function CryptDsnPassword(ByVal Dsn As String, ByVal En As Boolean) As String
        Dim sDsn As String = Dsn
        If Not sDsn.EndsWith(";") Then sDsn &= ";"
        Dim posI As Integer = sDsn.ToLower.IndexOf("password=")
        If posI = -1 Then Return Dsn
        posI += "password=".Length
        Dim posF As Integer = sDsn.ToLower.IndexOf(";", posI)
        If posF = -1 Or posI = posF Then Return Dsn
        Dim pw As String = sDsn.Substring(posI, posF - posI)
        pw = Crypt(pw, En)
        Dim s As String = ""
        s &= sDsn.Substring(0, posI)
        s &= pw
        s &= sDsn.Substring(posF)
        If s.EndsWith(";") Then s = s.Substring(0, s.Length - 1)
        Return s
    End Function

    Private Function criptaPW(ByVal txtClear As String, ByVal En As Boolean) As String
        Dim valide As String
        Dim crypta As String
        'Dim myoff As Integer
        Dim TxtCrypt As String
        Dim i As Integer
        'Dim j As Integer
        Dim lunga As Integer
        Dim lungaClear As Integer
        Dim flag As Integer = 0

        valide = "ABCDEFGHIJKLMNOPQRTSUVWXYZ0123456789"
        crypta = "FGHIJKLMNUVWXYZ0123OPQRTS456789ABCDE"
        lunga = Len(valide)
        lungaClear = Len(txtClear)

        TxtCrypt = ""
        txtClear = txtClear.ToUpper
        If En = True Then
            If (lungaClear > 0) Then
                i = 1
                While (i <= lungaClear)
                    If InStr(valide, Mid(txtClear, i, 1)) = 0 Then Return ""
                    TxtCrypt &= Mid(crypta, InStr(valide, Mid(txtClear, i, 1)), 1)
                    i += 1
                End While
            End If
        Else
            If (lungaClear > 0) Then
                i = 1
                While (i <= lungaClear)
                    If InStr(crypta, Mid(txtClear, i, 1)) = 0 Then Return ""
                    TxtCrypt &= Mid(valide, InStr(crypta, Mid(txtClear, i, 1)), 1)
                    i += 1
                End While
            End If

        End If

        Return IIf(TxtCrypt.Length > 0, TxtCrypt, " ")
    End Function

    Private Function Crypt(ByVal s As String, ByVal En As Boolean) As String
        Dim i As Integer
        Dim s2 As String = ""
        Dim rk As Integer
        Dim hx As String
        Dim ox As Byte

        Try
            If En Then
                If s.Length = 0 Then Throw New Exception()
                rk = Now.Ticks Mod 65536
                Dim Rnd As New Random(rk)
                For i = 0 To Rnd.Next(100)
                    Rnd.Next()
                Next
                s2 = Chr(Fix(rk / 256)) & Chr(rk Mod 256)
                For i = 0 To s.Length - 1
                    ox = ox Xor Rnd.Next(255)
                    If i > 0 Then ox = ox Xor Asc(s.Chars(Rnd.Next(i - 1)))
                    s2 &= Chr(Asc(s.Chars(i)) Xor Rnd.Next(255) Xor ox)
                Next
                s = ""
                For i = 0 To s2.Length - 1
                    hx = Hex(Asc(s2.Chars(i)))
                    If hx.Length = 1 Then hx = "0" & hx
                    s &= hx
                Next
                Return s

            Else
                If s.Length < 4 Then Throw New Exception()
                If s.Length Mod 2 <> 0 Then Throw New Exception()
                For i = 0 To s.Length - 1 Step 2
                    hx = s.Substring(i, 2)
                    s2 &= Chr(Integer.Parse(hx, System.Globalization.NumberStyles.HexNumber))
                Next
                s = s2.Substring(2)
                s2 = s2.Substring(0, 2)
                rk = Asc(s2.Chars(0)) * 256 + Asc(s2.Chars(1))
                Dim Rnd As New Random(rk)
                For i = 0 To Rnd.Next(100)
                    Rnd.Next()
                Next
                s2 = ""
                For i = 0 To s.Length - 1
                    ox = ox Xor Rnd.Next(255)
                    If i > 0 Then ox = ox Xor Asc(s2.Chars(Rnd.Next(i - 1)))
                    s2 &= Chr(Asc(s.Chars(i)) Xor Rnd.Next(255) Xor ox)
                Next
                Return s2
            End If
        Catch
            Throw New Exception("Error in PwCrypt")
        End Try

    End Function


End Module

