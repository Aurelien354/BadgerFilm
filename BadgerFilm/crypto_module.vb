'Imports System.IO
'Imports System.Text

'Module crypto_module
'    Public Sub crypt(ByVal path As String, ByVal filename As String)
'        Dim abc(255) As Byte
'        For i As Integer = 0 To 255
'            abc(i) = Convert.ToByte(i)
'        Next

'        Dim table(255, 255) As Byte
'        For i As Integer = 0 To 255
'            For j As Integer = 0 To 255
'                table(i, j) = abc((i + j) Mod 256)
'            Next
'        Next

'        Dim fileContent As Byte() = File.ReadAllBytes(path & "\" & filename)
'        Dim passwordTmp As Byte() = Encoding.ASCII.GetBytes("PENELOPE")
'        Dim keys(fileContent.Length - 1) As Byte

'        For i As Integer = 0 To fileContent.Length - 1
'            keys(i) = passwordTmp(i Mod passwordTmp.Length)
'        Next

'        Dim result(fileContent.Length - 1) As Byte '= New Byte(fileContent.Length)

'        For i As Integer = 0 To fileContent.Length - 1
'            Dim value As Byte = fileContent(i)
'            Dim key As Byte = keys(i)
'            Dim valueIndex As Integer = -1
'            Dim keyIndex As Integer = -1
'            For j As Integer = 0 To 255
'                If abc(j) = value Then
'                    valueIndex = j
'                    Exit For
'                End If
'            Next
'            For j As Integer = 0 To 255
'                If abc(j) = key Then
'                    keyIndex = j
'                    Exit For
'                End If
'            Next
'            result(i) = table(keyIndex, valueIndex)
'        Next

'        File.WriteAllBytes(path & "\" & filename & ".crypt", result)

'    End Sub

'    Public Function decrypt(ByVal path As String, ByVal filename As String) As String
'        'Debug.WriteLine("decrypt")
'        Dim abc(255) As Byte
'        For i As Integer = 0 To 255
'            abc(i) = Convert.ToByte(i)
'        Next

'        Dim table(255, 255) As Byte
'        For i As Integer = 0 To 255
'            For j As Integer = 0 To 255
'                table(i, j) = abc((i + j) Mod 256)
'            Next
'        Next

'        Dim fileContent As Byte() = File.ReadAllBytes(path & "\" & filename)
'        Dim passwordTmp As Byte() = Encoding.ASCII.GetBytes("PENELOPE")
'        Dim keys(fileContent.Length - 1) As Byte

'        For i As Integer = 0 To fileContent.Length - 1
'            keys(i) = passwordTmp(i Mod passwordTmp.Length)
'        Next

'        Dim result(fileContent.Length - 1) As Byte

'        For i As Integer = 0 To fileContent.Length - 1
'            Dim value As Byte = fileContent(i)
'            Dim key As Byte = keys(i)
'            Dim valueIndex As Integer = -1
'            Dim keyIndex As Integer = -1
'            For j As Integer = 0 To 255
'                If abc(j) = key Then
'                    keyIndex = j
'                    Exit For
'                End If
'            Next
'            For j As Integer = 0 To 255
'                If table(keyIndex, j) = value Then
'                    valueIndex = j
'                    Exit For
'                End If
'            Next
'            result(i) = abc(valueIndex)
'        Next

'        decrypt = Encoding.ASCII.GetString(result)

'    End Function
'End Module
