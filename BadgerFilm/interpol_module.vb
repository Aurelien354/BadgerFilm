Imports System.IO

Module interpol_module
    '***********************************************
    'Interpole the input data data_xs at the energy E_inter using log log interpolation methode
    '***********************************************
    Public Function interpol_log_log(ByVal data_xs As data_xs, ByVal E_inter As Double) As Double()
        Try
            Dim indice As Integer = 0

            Dim res(data_xs.cross_section.GetUpperBound(1)) As Double
            If E_inter < data_xs.energy(0) Then
                For i As Integer = 0 To data_xs.cross_section.GetUpperBound(1)
                    res(i) = data_xs.cross_section(0, i)
                Next
                Return res
            End If

            While data_xs.energy(indice) < E_inter
                indice = indice + 1
            End While

            If data_xs.energy(indice + 1) = E_inter Then
                For i As Integer = 0 To data_xs.cross_section.GetUpperBound(1)
                    res(i) = data_xs.cross_section(indice + 1, i)
                Next
                Return res
            End If

            If data_xs.energy(indice) = E_inter Then
                For i As Integer = 0 To data_xs.cross_section.GetUpperBound(1)
                    res(i) = data_xs.cross_section(indice, i)
                Next
                Return res
            End If

            For i As Integer = 0 To data_xs.cross_section.GetUpperBound(1)
                res(i) = interpol_log_log_method({data_xs.energy(indice - 1), data_xs.energy(indice)}, {data_xs.cross_section(indice - 1, i), data_xs.cross_section(indice, i)}, E_inter)
            Next
            Return res

        Catch ex As Exception
            'MsgBox("Ionization cross section not found!")
            Console.WriteLine("Ionization cross section not found!")
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in interpol_log_log " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)

            Return {-1}

        End Try

    End Function

    Public Function interpol_log_log_method(energy() As Double, xs() As Double, energy_inter As Double) As Double
        Try
            interpol_log_log_method = 0

            If energy_inter < energy(0) Then
                Exit Function
            End If

            For j As Integer = 0 To UBound(energy) - 1
                If energy_inter < energy(j + 1) And energy_inter > energy(j) Then
                    Dim y2 As Double = xs(j + 1)
                    Dim y1 As Double = xs(j)
                    Dim x2 As Double = energy(j + 1)
                    Dim x1 As Double = energy(j)

                    If y2 = 0 Or y1 = 0 Then
                        interpol_log_log_method = 0
                    Else
                        interpol_log_log_method = 10 ^ (Math.Log10(y1) + Math.Log10(y2 / y1) * Math.Log10(energy_inter / x1) / Math.Log10(x2 / x1))
                    End If
                    Exit Function
                End If
                If energy_inter = energy(j) Then
                    interpol_log_log_method = energy(j)
                    Exit Function
                End If
            Next

            If energy_inter = energy(UBound(energy)) Then
                interpol_log_log_method = xs(UBound(energy))
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in interpol_log_log_method " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Function
End Module
