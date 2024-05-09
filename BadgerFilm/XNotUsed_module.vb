Imports System.IO

Module XNotUsed_module
    Public Sub find_mac_EPDL(ByRef energy() As Double, ByRef xs() As Double, ByVal num As Integer, ByVal name As String)
        'num =
        '23501   Total cross sections                                   
        '23502   Coherent scattering cross sections                     
        '23504   Incoherent scattering cross sections                   
        '23515   Pair production cross sections, Electron field        
        '23516   Pair production cross sections, Total                  
        '23517   Pair production cross sections, Nuclear field         
        '23522   Total photoionization cross section                 
        '23534   K (1S1/2)    Photoionization subshell cross section   
        '23535   L1 (2S1/2)   Photoionization subshell cross section    
        '23536   L2 (2P1/2)   Photoionization subshell cross section    
        '23537   L3 (2P3/2)   Photoionization subshell cross section
        '23538   M1 (3S1/2)   Photoionization subshell cross section   
        '23539   M2 (3P1/2)   Photoionization subshell cross section 
        '23540   M3 (3P3/2)   Photoionization subshell cross section

        Dim flag As Integer = 0

        Dim data_file As String = name '"epdl92.txt"

        Dim mystream As New StreamReader(data_file)

        Try
            If (mystream IsNot Nothing) Then

                Dim tmp As String

                Dim cnt As Integer = 0
                Dim nbr_val As Integer

                While mystream.Peek <> -1
                    tmp = mystream.ReadLine()
                    'MsgBox(Mid(tmp, 71, 5))
                    Dim read_num As String = Mid(tmp, 71, 5)
                    If IsNumeric(read_num) = False Then Continue While

                    If (read_num <> num) Then
                        Continue While
                    End If
                    tmp = Mid(tmp, 1, 66)
                    'MsgBox(tmp)

                    flag = flag + 1

                    If flag = 3 Then
                        nbr_val = Trim(Mid(tmp, 1, 11))
                        'MsgBox(nbr_val)
                        ReDim energy(nbr_val - 1)
                        ReDim xs(nbr_val - 1)
                    End If

                    If flag > 3 Then
                        'tmp = Replace(tmp, ".", ",")
                        Dim tmp_str() As String
                        tmp_str = Trim(tmp).Split(" ")
                        'MsgBox(UBound(tmp_str))
                        'For i As Integer = 0 To UBound(tmp_str)
                        'If InStr(tmp_str(i), "+") <> 0 Then
                        '    tmp_str(i) = Split(tmp_str(i), "+")(0) * 10 ^ Split(tmp_str(i), "+")(1)
                        'End If
                        'If InStr(tmp_str(i), "-") <> 0 Then
                        '    tmp_str(i) = Split(tmp_str(i), "-")(0) * 10 ^ (-1 * Split(tmp_str(i), "-")(1))
                        'End If
                        'tmp_str(i) = Mid(tmp_str(i), 1, 8) * 10 ^ (Mid(tmp_str(i), 9))
                        'MsgBox(tmp_str(i))
                        'Next

                        For i As Integer = 0 To 2
                            energy((flag - 4) * 3 + i) = tmp_str(2 * i)
                            xs((flag - 4) * 3 + i) = tmp_str(2 * i + 1)
                            cnt = cnt + 1
                            'MsgBox(cnt)
                            If cnt = nbr_val Then

                                Exit For
                                Exit While
                            End If
                        Next

                    End If

                End While

            End If

        Catch Ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in find_mac_EPDL " & Ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)

        Finally
            ' Check this again, since we need to make sure we didn't throw an exception on open. 
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try
    End Sub

    Public Sub find_mac_EPDL(ByRef data() As String, ByVal num As Integer, ByVal name As String)
        'num =
        '23501   Total cross sections                                   
        '23502   Coherent scattering cross sections                     
        '23504   Incoherent scattering cross sections                   
        '23515   Pair production cross sections, Electron field        
        '23516   Pair production cross sections, Total                  
        '23517   Pair production cross sections, Nuclear field         
        '23522   Total photoionization cross section                 
        '23534   K (1S1/2)    Photoionization subshell cross section   
        '23535   L1 (2S1/2)   Photoionization subshell cross section    
        '23536   L2 (2P1/2)   Photoionization subshell cross section    
        '23537   L3 (2P3/2)   Photoionization subshell cross section
        '23538   M1 (3S1/2)   Photoionization subshell cross section   
        '23539   M2 (3P1/2)   Photoionization subshell cross section 
        '23540   M3 (3P3/2)   Photoionization subshell cross section

        Dim flag As Integer = 0

        Dim data_file As String = name '"epdl92.txt"

        Dim mystream As New StreamReader(data_file)

        Try
            If (mystream IsNot Nothing) Then

                Dim tmp As String

                Dim cnt As Integer = 0
                Dim nbr_val As Integer

                While mystream.Peek <> -1
                    tmp = mystream.ReadLine()
                    'MsgBox(Mid(tmp, 71, 5))
                    Dim read_num As String = Mid(tmp, 71, 5)
                    If IsNumeric(read_num) = False Then Continue While

                    If (read_num <> num) Then
                        Continue While
                    End If
                    tmp = Mid(tmp, 1, 66)
                    'MsgBox(tmp)

                    flag = flag + 1

                    If flag = 3 Then
                        nbr_val = Trim(Mid(tmp, 1, 11))
                        'MsgBox(nbr_val)
                        ReDim data(nbr_val - 1)
                    End If

                    If flag > 3 Then
                        'tmp = Replace(tmp, ".", ",")
                        Dim tmp_str() As String
                        tmp_str = Trim(tmp).Split(" ")
                        'MsgBox(UBound(tmp_str))
                        'For i As Integer = 0 To UBound(tmp_str)
                        'If InStr(tmp_str(i), "+") <> 0 Then
                        '    tmp_str(i) = Split(tmp_str(i), "+")(0) * 10 ^ Split(tmp_str(i), "+")(1)
                        'End If
                        'If InStr(tmp_str(i), "-") <> 0 Then
                        '    tmp_str(i) = Split(tmp_str(i), "-")(0) * 10 ^ (-1 * Split(tmp_str(i), "-")(1))
                        'End If
                        'tmp_str(i) = Mid(tmp_str(i), 1, 8) * 10 ^ (Mid(tmp_str(i), 9))
                        'MsgBox(tmp_str(i))
                        'Next

                        For i As Integer = 0 To 2
                            data((flag - 4) * 3 + i) = tmp_str(2 * i) & " " & tmp_str(2 * i + 1)
                            cnt = cnt + 1
                            'MsgBox(cnt)
                            If cnt = nbr_val Then

                                Exit For
                                Exit While
                            End If
                        Next

                    End If

                End While

            End If

        Catch Ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in find_mac_EPDL " & Ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)

        Finally
            ' Check this again, since we need to make sure we didn't throw an exception on open. 
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try
    End Sub
End Module
