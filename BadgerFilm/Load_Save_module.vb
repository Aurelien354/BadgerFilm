Imports System.IO
Imports System.Text.RegularExpressions

Module Load_Save_module
    Public Sub load_data(ByVal data_file As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        '******************
        ' Load an entire system (samples, materials and X-ray lines)
        ' data_file: path and name of the saved file
        ' analysis_cond_handler: structure to store the retrived data
        '******************

        Dim temp As String = Nothing
        Dim mystream As New StreamReader(data_file)
        Try
            If (mystream IsNot Nothing) Then
                temp = mystream.ReadToEnd()
            End If
        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            ' Check this again, since we need to make sure we didn't throw an exception on open. 
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try

        'If data_file = Nothing Then Exit Sub

        Try
            Dim version As String = ""
            Dim indice As Integer = 0
            Dim num_analysis_cond_handler As Integer
            Dim line() As String = Split(temp, vbCrLf)

            If line(0) Like "[#]v*" Then
                version = Split(line(0), "#")(1)
                indice = indice + 1

            End If


            Dim toa_line() As String = Split(line(indice), vbTab)

            'If toa_line(0) = "toa" Then
            '    num_analysis_cond_handler = Split(line(1), vbTab).Last
            '    indice = 4
            '    toa = toa_line(1)
            '    'TextBox1.Text = toa(1)
            'Else
            '    num_analysis_cond_handler = Split(line(0), vbTab).Last
            '    indice = 3
            '    toa = 40
            '    'TextBox1.Text = 40
            'End If

            toa = Split(line(indice), vbTab).Last
            indice = indice + 1
            num_analysis_cond_handler = Split(line(indice), vbTab).Last
            ReDim layer_handler(num_analysis_cond_handler - 1)

            indice = indice + 1
            'Nope



            For i As Integer = 0 To num_analysis_cond_handler - 1
                indice = indice + 1
                'Nope
                indice = indice + 1
                layer_handler(i).density = Split(line(indice), vbTab).Last
                indice = indice + 1
                layer_handler(i).isfix = Split(line(indice), vbTab).Last
                indice = indice + 1
                layer_handler(i).thickness = Split(line(indice), vbTab).Last
                indice = indice + 1
                layer_handler(i).wt_fraction = Split(line(indice), vbTab).Last

                layer_handler(i).id = i

                indice = indice + 1
                Dim num_elt As Integer = Split(line(indice), vbTab).Last
                ReDim layer_handler(i).element(num_elt - 1)

                indice = indice + 1
                'Nope

                'Dim tot As Double = 0
                For j As Integer = 0 To num_elt - 1
                    indice = indice + 1
                    'Nope
                    indice = indice + 1
                    layer_handler(i).element(j).elt_name = Split(line(indice), vbTab).Last
                    indice = indice + 1
                    layer_handler(i).element(j).isConcFixed = Split(line(indice), vbTab).Last
                    indice = indice + 1
                    layer_handler(i).element(j).conc_wt = Split(line(indice), vbTab).Last
                    'tot = tot + layer_handler(i).element(j).conc_wt

                    layer_handler(i).element(j).mother_layer_id = i

                Next

                'For j As Integer = 0 To num_elt - 1
                '    layer_handler(i).element(j).conc_wt = layer_handler(i).element(j).conc_wt / tot
                'Next


                indice = indice + 1
                'Nope
                convert_wt_to_at(layer_handler, i)
            Next

            indice = indice + 1
            Dim num_exp_elt As String = Split(line(indice), vbTab).Last

            If num_exp_elt = "" Then

            Else

                If num_exp_elt > 0 Then
                    ReDim elt_exp_handler(num_exp_elt - 1)

                    indice = indice + 1
                    'Nope

                    For j As Integer = 0 To num_exp_elt - 1
                        indice = indice + 1
                        elt_exp_handler(j).a = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        elt_exp_handler(j).elt_name = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        elt_exp_handler(j).z = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        Dim num_lines As Integer = Split(line(indice), vbTab).Last
                        ReDim elt_exp_handler(j).line(num_lines - 1)

                        For k As Integer = 0 To num_lines - 1
                            indice = indice + 1
                            elt_exp_handler(j).line(k).Ec = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            elt_exp_handler(j).line(k).xray_energy = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            elt_exp_handler(j).line(k).xray_name = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            elt_exp_handler(j).line(k).std = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            elt_exp_handler(j).line(k).std_filename = Split(line(indice), vbTab).Last

                            indice = indice + 1
                            Dim num_kratios As Integer = Split(line(indice), vbTab).Last
                            ReDim elt_exp_handler(j).line(k).k_ratio(num_kratios - 1)

                            For l As Integer = 0 To num_kratios - 1
                                indice = indice + 1
                                elt_exp_handler(j).line(k).k_ratio(l).elt_intensity = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                elt_exp_handler(j).line(k).k_ratio(l).kv = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                elt_exp_handler(j).line(k).k_ratio(l).experimental_value = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                If version <> "" Then
                                    elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = Split(line(indice), vbTab).Last
                                    indice = indice + 1
                                Else
                                    elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = 0
                                End If
                                elt_exp_handler(j).line(k).k_ratio(l).std_intensity = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                elt_exp_handler(j).line(k).k_ratio(l).theo_value = Split(line(indice), vbTab).Last
                            Next
                        Next
                    Next
                End If
            End If
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    init_element_layer(layer_handler(i).element(j).elt_name, vbNull, layer_handler(i).element(j))
                Next
            Next

            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            Next

        Catch Ex As Exception
            MessageBox.Show("Not a valid BadgerFilm input file. Original error: " & Ex.Message)
        End Try

    End Sub

    Public Sub export(ByVal file_name As String, ByVal layer_handler() As layer, ByVal elt_exp_handler() As Elt_exp, ByVal toa As Double, ByVal version As String)
        '*******************************************
        Dim tmp As String = ""
        tmp = "#" & version & vbCrLf
        tmp = tmp & "toa" & vbTab & toa & vbCrLf
        tmp = tmp & "#layer_handler" & vbTab & layer_handler.Count & vbCrLf
        tmp = tmp & "******************" & vbCrLf

        For i As Integer = 0 To UBound(layer_handler)
            tmp = tmp & "layer_handler" & vbTab & i & vbCrLf
            tmp = tmp & "density" & vbTab & layer_handler(i).density & vbCrLf
            tmp = tmp & "isfix" & vbTab & layer_handler(i).isfix & vbCrLf
            tmp = tmp & "thickness" & vbTab & layer_handler(i).thickness & vbCrLf
            tmp = tmp & "wt_fraction" & vbTab & layer_handler(i).wt_fraction & vbCrLf
            tmp = tmp & "#elt" & vbTab & layer_handler(i).element.Count & vbCrLf
            tmp = tmp & "***************" & vbCrLf

            For j As Integer = 0 To UBound(layer_handler(i).element)
                tmp = tmp & "elt" & vbTab & j & vbCrLf
                tmp = tmp & "name" & vbTab & layer_handler(i).element(j).elt_name & vbCrLf
                tmp = tmp & "isConcFixed" & vbTab & layer_handler(i).element(j).isConcFixed & vbCrLf
                tmp = tmp & "conc" & vbTab & layer_handler(i).element(j).conc_wt & vbCrLf
                'tmp = tmp & "line" & vbTab & analysis_cond_handler(i).elts(j).line & vbCrLf
            Next
            tmp = tmp & "***************" & vbCrLf
        Next

        If elt_exp_handler IsNot Nothing Then
            tmp = tmp & "#num_elt_exp" & vbTab & elt_exp_handler.Count & vbCrLf
            tmp = tmp & "*********" & vbCrLf
            For i As Integer = 0 To UBound(elt_exp_handler)
                tmp = tmp & "#a" & vbTab & elt_exp_handler(i).a & vbCrLf
                tmp = tmp & "#elt_name" & vbTab & elt_exp_handler(i).elt_name & vbCrLf
                tmp = tmp & "#z" & vbTab & elt_exp_handler(i).z & vbCrLf
                tmp = tmp & "#num_line" & vbTab & elt_exp_handler(i).line.Count & vbCrLf

                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    tmp = tmp & "#Ec" & vbTab & elt_exp_handler(i).line(j).Ec & vbCrLf
                    tmp = tmp & "#xray_energy" & vbTab & elt_exp_handler(i).line(j).xray_energy & vbCrLf
                    tmp = tmp & "#xray_name" & vbTab & elt_exp_handler(i).line(j).xray_name & vbCrLf
                    tmp = tmp & "#std" & vbTab & elt_exp_handler(i).line(j).std & vbCrLf
                    tmp = tmp & "#std_filename" & vbTab & elt_exp_handler(i).line(j).std_filename & vbCrLf
                    tmp = tmp & "#num_kratios" & vbTab & elt_exp_handler(i).line(j).k_ratio.Count & vbCrLf


                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        tmp = tmp & "#elt_intensity" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).elt_intensity & vbCrLf
                        tmp = tmp & "#kv" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).kv & vbCrLf
                        tmp = tmp & "#experimental_value" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).experimental_value & vbCrLf
                        tmp = tmp & "#err_experimental_value" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value & vbCrLf
                        tmp = tmp & "#std_intensity" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).std_intensity & vbCrLf
                        tmp = tmp & "#theo_value" & vbTab & elt_exp_handler(i).line(j).k_ratio(k).theo_value & vbCrLf

                    Next
                Next
            Next
        End If

        'tmp = tmp & "#energy_analysis" & vbTab & analysis_cond_handler(i).elts(j).energy_analysis.Count & vbCrLf
        '    tmp = tmp & "*********" & vbCrLf
        '    For k As Integer = 0 To UBound(analysis_cond_handler(i).elts(j).energy_analysis)
        '        tmp = tmp & "energy_analysis" & vbTab & analysis_cond_handler(i).elts(j).energy_analysis(k) & vbCrLf
        '    Next

        '    tmp = tmp & "#kratio" & vbTab & analysis_cond_handler(i).elts(j).kratio.Count & vbCrLf
        '    tmp = tmp & "*********" & vbCrLf
        '    For k As Integer = 0 To UBound(analysis_cond_handler(i).elts(j).kratio)
        '        tmp = tmp & "kratio" & vbTab & analysis_cond_handler(i).elts(j).kratio(k) & vbCrLf
        '    Next

        '    tmp = tmp & "#kratio_measured" & vbTab & analysis_cond_handler(i).elts(j).kratio_measured.Count & vbCrLf
        '    tmp = tmp & "*********" & vbCrLf
        '    For k As Integer = 0 To UBound(analysis_cond_handler(i).elts(j).kratio)
        '        tmp = tmp & "kratio_measured" & vbTab & analysis_cond_handler(i).elts(j).kratio_measured(k) & vbCrLf
        '    Next

        '    tmp = tmp & "#std_file" & vbTab & analysis_cond_handler(i).elts(j).std_filename.Count & vbCrLf
        '    tmp = tmp & "*********" & vbCrLf
        '    For k As Integer = 0 To UBound(analysis_cond_handler(i).elts(j).kratio)
        '        tmp = tmp & "std_file" & vbTab & analysis_cond_handler(i).elts(j).std_filename(k) & vbCrLf
        '    Next
        'Next
        'tmp = tmp & "***************" & vbCrLf
        'Next
        '*******************************************

        Dim sw As New StreamWriter(file_name, False)
        sw.Write(tmp)
        sw.Close()

        'TextBox12.Text = tmp
    End Sub

    Public Sub import_Stratagem(ByVal data_file As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        Dim temp As String = Nothing
        Dim mystream As New StreamReader(data_file)
        Try
            If (mystream IsNot Nothing) Then
                temp = mystream.ReadToEnd()
            End If
        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            ' Check this again, since we need to make sure we didn't throw an exception on open. 
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try

        'temp = Replace(temp, vbCr, vbCrLf)
        Dim lines() As String = Split(temp, vbCrLf)
        import_Stratagem_method(lines, layer_handler, elt_exp_handler, toa)
    End Sub

    Public Sub import_Stratagem_method(ByVal lines() As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        Try
            layer_handler = Nothing
            toa = 40 'by default

            For i As Integer = 0 To UBound(lines)
                lines(i) = Regex.Replace(lines(i), "#.*$", "")
            Next

            For i As Integer = 0 To UBound(lines)
                Dim line As String = Trim(lines(i))
                If line = "" Then Continue For   'OrElse line(0) = "#"
                If line Like "$Layer*" Then
                    If layer_handler Is Nothing Then
                        ReDim layer_handler(0)
                    Else
                        ReDim Preserve layer_handler(UBound(layer_handler) + 1)
                    End If
                    Dim layer_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                    If layer_def.Count = 4 Or layer_def.Count = 3 Then
                        layer_handler(UBound(layer_handler)).density = layer_def(1)
                        layer_handler(UBound(layer_handler)).thickness = layer_def(2)
                        If layer_def.Count = 4 Then
                            If layer_def(3) = "k" Then
                                layer_handler(UBound(layer_handler)).isfix = True
                            Else
                                layer_handler(UBound(layer_handler)).isfix = False
                            End If
                        Else
                            layer_handler(UBound(layer_handler)).isfix = True
                        End If
                    Else
                        layer_handler(UBound(layer_handler)).density = 2.0
                        layer_handler(UBound(layer_handler)).thickness = 1000000000.0
                        layer_handler(UBound(layer_handler)).isfix = True
                    End If
                    layer_handler(UBound(layer_handler)).wt_fraction = True
                    layer_handler(UBound(layer_handler)).id = UBound(layer_handler)
                End If

                If line Like "$Elt*" Then
                    If layer_handler(UBound(layer_handler)).element Is Nothing Then
                        ReDim layer_handler(UBound(layer_handler)).element(0)
                    Else
                        ReDim Preserve layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element) + 1)
                    End If

                    Dim elt_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                    If elt_def.Count = 4 Or elt_def.Count = 3 Then
                        layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).z = elt_def(1)
                        layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).conc_wt = elt_def(2)

                        If elt_def.Count = 4 Then
                            If elt_def(3) = "k" Then
                                layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).isConcFixed = True
                            Else
                                layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).isConcFixed = False
                            End If
                        Else
                            layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).isConcFixed = True
                        End If
                        layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).mother_layer_id = UBound(layer_handler)
                    End If
                End If

                If line Like "$Geom*" Then
                    toa = Split(Regex.Replace(line, " {2,}", " "), " ")(1)
                End If

                If line Like "$K*" Then
                    Dim k_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                    Dim indice_elt As Integer = -1
                    Dim indice_line As Integer = -1

                    If elt_exp_handler Is Nothing Then
                        ReDim elt_exp_handler(0)
                        indice_elt = 0
                    Else
                        For j As Integer = 0 To UBound(elt_exp_handler)
                            If elt_exp_handler(j).z = k_def(1) Then
                                indice_elt = j
                                Exit For
                            End If
                        Next
                        If indice_elt <> -1 Then
                            If elt_exp_handler(indice_elt).line IsNot Nothing Then
                                For j As Integer = 0 To UBound(elt_exp_handler(indice_elt).line)
                                    If elt_exp_handler(indice_elt).line(j).xray_name = k_def(2) Then
                                        indice_line = j
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If


                    If indice_elt = -1 Then
                        ReDim Preserve elt_exp_handler(UBound(elt_exp_handler) + 1)
                        indice_elt = UBound(elt_exp_handler)
                    End If

                    'Dim k_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                    If k_def.Count = 5 Or k_def.Count = 6 Then
                        elt_exp_handler(indice_elt).z = k_def(1)
                        elt_exp_handler(indice_elt).elt_name = Z_to_symbol(elt_exp_handler(indice_elt).z)

                        If indice_line = -1 Then
                            If elt_exp_handler(indice_elt).line Is Nothing Then
                                ReDim elt_exp_handler(indice_elt).line(0)
                                indice_line = 0
                            Else
                                ReDim Preserve elt_exp_handler(indice_elt).line(UBound(elt_exp_handler(indice_elt).line) + 1)
                                indice_line = UBound(elt_exp_handler(indice_elt).line)
                            End If
                            elt_exp_handler(indice_elt).line(indice_line).xray_name = k_def(2)
                        Else

                        End If



                        If elt_exp_handler(indice_elt).line(indice_line).k_ratio Is Nothing Then
                            ReDim elt_exp_handler(indice_elt).line(indice_line).k_ratio(0)
                        Else
                            ReDim Preserve elt_exp_handler(indice_elt).line(indice_line).k_ratio(UBound(elt_exp_handler(indice_elt).line(indice_line).k_ratio) + 1)
                        End If
                        elt_exp_handler(indice_elt).line(indice_line).k_ratio(UBound(elt_exp_handler(indice_elt).line(indice_line).k_ratio)).kv = k_def(3)

                        If k_def.Count = 6 Then
                            elt_exp_handler(indice_elt).line(indice_line).std_filename = Application.StartupPath & "\" & k_def(5) & ".txt"
                        Else
                            elt_exp_handler(indice_elt).line(indice_line).std_filename = ""
                        End If

                        i = i + 1
                        line = Trim(lines(i))
                        Dim k_value() As Double = Nothing
                        While line Like "$*" = False
                            If line = "" Then
                                i = i + 1
                                line = Trim(lines(i))
                                Continue While
                            End If
                            Dim tmp() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                            For j As Integer = 0 To UBound(tmp)
                                If k_value Is Nothing Then
                                    ReDim k_value(0)
                                    k_value(0) = tmp(j)
                                Else
                                    ReDim Preserve k_value(UBound(k_value) + 1)
                                    k_value(UBound(k_value)) = tmp(j)
                                End If
                            Next
                            i = i + 1
                            line = Trim(lines(i))
                        End While

                        Dim res As Double = 0
                        For j As Integer = 0 To UBound(k_value)
                            res = res + k_value(j)
                        Next
                        res = res / k_value.Count

                        elt_exp_handler(indice_elt).line(indice_line).k_ratio(UBound(elt_exp_handler(indice_elt).line(indice_line).k_ratio)).experimental_value = res
                        elt_exp_handler(indice_elt).line(indice_line).k_ratio(UBound(elt_exp_handler(indice_elt).line(indice_line).k_ratio)).err_experimental_value = res * 0.05
                        i = i - 1
                        line = Trim(lines(i))
                    End If
                End If


                If line Like "$Std*" Then
                    Dim std_file_name As String = Split(Regex.Replace(line, " {2,}", " "), " ")(1)
                    i = i + 1
                    i = i + 1

                    line = Trim(lines(i))
                    Dim std_lines() As String = Nothing
                    While line Like "$$*" = True
                        If std_lines Is Nothing Then
                            ReDim std_lines(0)
                        Else
                            ReDim Preserve std_lines(UBound(std_lines) + 1)
                        End If
                        std_lines(UBound(std_lines)) = Replace(line, "$$", "$")

                        i = i + 1
                        line = Trim(lines(i))
                    End While

                    Dim layer_handler_std() As layer = Nothing
                    Dim elt_exp_std_handler() As Elt_exp = Nothing
                    Dim toa_std As Double
                    import_Stratagem_method(std_lines, layer_handler_std, elt_exp_std_handler, toa_std)
                    export(std_file_name & ".txt", layer_handler_std, elt_exp_std_handler, toa_std, "vstratagem")

                    i = i - 1
                End If
            Next

            For i As Integer = 0 To UBound(layer_handler)
                Dim tot As Double = 0
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    layer_handler(i).element(j).elt_name = Z_to_symbol(layer_handler(i).element(j).z)
                    init_element_layer(layer_handler(i).element(j).elt_name, vbNull, layer_handler(i).element(j))
                    tot = tot + layer_handler(i).element(j).conc_wt
                Next
                If tot > 2.0 Then
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        layer_handler(i).element(j).conc_wt = layer_handler(i).element(j).conc_wt / tot
                    Next
                End If
            Next

            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            Next

            For i As Integer = 0 To UBound(layer_handler)
                convert_wt_to_at(layer_handler, i)
            Next
        Catch Ex As Exception
            MessageBox.Show("Not a valid STRATAgem file. Original error: " & Ex.Message)
        End Try
    End Sub
End Module
