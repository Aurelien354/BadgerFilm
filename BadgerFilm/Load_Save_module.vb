Imports System.IO
Imports System.Text.RegularExpressions

Module Load_Save_module
    Public Sub load_data(ByVal data_file As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        '******************
        ' Load an entire system (samples, materials and X-ray lines)
        ' data_file: path and name of the saved file
        ' layer_handler: structure to store the retrived data corresponding to the sample's geometry
        ' elt_exp_handler: structure to store the retrived data corresponding to the expremiental data (analyzed elements, X-ray lines k-ratios, kV, ...)
        ' toa is the takeoff angle in degree
        '******************

        'Try to open the file and store its content into the temp varaible
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
            Dim num_layers As Integer

            Dim line() As String = Split(temp, vbCrLf)

            'Retrieve the BadgerFilm version number used to save the data.
            If line(0) Like "[#]v*" Then
                version = Split(line(0), "#")(1)
                indice = indice + 1
            End If

            'Retrieve the takeoff angle in deg
            Dim toa_line() As String = Split(line(indice), vbTab)
            toa = Split(line(indice), vbTab).Last
            indice = indice + 1

            'Retrieve the number of layers and redim layer_handler accordingly
            num_layers = Split(line(indice), vbTab).Last
            ReDim layer_handler(num_layers - 1)
            indice = indice + 1

            'For each layer:
            For i As Integer = 0 To num_layers - 1
                'Skip unused line in the file
                indice = indice + 1
                indice = indice + 1
                'Retrieve the density for the current layer (in g/cm3)
                layer_handler(i).density = Split(line(indice), vbTab).Last
                indice = indice + 1
                'Retrieve the fal indicating if the thickness of the layer is fixed
                layer_handler(i).isfix = Split(line(indice), vbTab).Last
                indice = indice + 1
                'Retrieve the thickness of the current layer (in Angstrom)
                layer_handler(i).thickness = Split(line(indice), vbTab).Last
                indice = indice + 1
                'Retrieve the flag indicating if the layer composition is defined by weight fraction or by atomic fraction
                layer_handler(i).wt_fraction = Split(line(indice), vbTab).Last

                'Set the id (number) of the current layer (the top layer always has id=0)
                layer_handler(i).id = i
                indice = indice + 1

                'Retrieve the number of element in the current layer and redim element accordingly
                Dim num_elt As Integer = Split(line(indice), vbTab).Last
                ReDim layer_handler(i).element(num_elt - 1)
                indice = indice + 1

                For j As Integer = 0 To num_elt - 1
                    'Skip unused lines
                    indice = indice + 1
                    indice = indice + 1
                    'Retrieve the element's name
                    layer_handler(i).element(j).elt_name = Split(line(indice), vbTab).Last
                    indice = indice + 1
                    'Retrieve  the flag indicating if the element's concentration is fixed
                    layer_handler(i).element(j).isConcFixed = Split(line(indice), vbTab).Last
                    indice = indice + 1
                    'Retrieve the element's weight fraction
                    layer_handler(i).element(j).conc_wt = Split(line(indice), vbTab).Last

                    layer_handler(i).element(j).mother_layer_id = i

                Next
                indice = indice + 1

                'Convert the weight fractions of the elements into atomic fraction
                convert_wt_to_at(layer_handler, i)
            Next
            indice = indice + 1

            'Retrieve the nuumber of experimental data
            Dim num_exp_elt As String = Split(line(indice), vbTab).Last

            If num_exp_elt = "" Then
                MessageBox.Show("Error: no experimental data ond in " & data_file)
            Else
                If num_exp_elt > 0 Then
                    ReDim elt_exp_handler(num_exp_elt - 1)
                    indice = indice + 1

                    'For each experimental data block found:
                    For j As Integer = 0 To num_exp_elt - 1
                        indice = indice + 1
                        'Retrieve the element's atomic weight
                        elt_exp_handler(j).a = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        'Retrieve the element's name
                        elt_exp_handler(j).elt_name = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        'Retrieve the element's atomic number
                        elt_exp_handler(j).z = Split(line(indice), vbTab).Last
                        indice = indice + 1
                        'Retrieve the number of different X-ray lines used to analyze the current element (usually only 1) and redim the line array accordingly
                        Dim num_lines As Integer = Split(line(indice), vbTab).Last
                        ReDim elt_exp_handler(j).line(num_lines - 1)

                        'For each X-ray lines:
                        For k As Integer = 0 To num_lines - 1
                            indice = indice + 1
                            'Retrieve the critical ionization energy (not used)
                            elt_exp_handler(j).line(k).Ec = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            'Retrieve the X-ray energy
                            elt_exp_handler(j).line(k).xray_energy = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            'Retrieve the X-ray name
                            elt_exp_handler(j).line(k).xray_name = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            'Retrieve the standard's name used to analyzed this element and X-ray line (blank if pure element used as a standard)
                            elt_exp_handler(j).line(k).std = Split(line(indice), vbTab).Last
                            indice = indice + 1
                            'Retrieve the path of the standard file (blank if pure element used as a standard)
                            elt_exp_handler(j).line(k).std_filename = Split(line(indice), vbTab).Last
                            indice = indice + 1

                            'Retrieve the number of k-ratios for the current element and X-ray lines (one k-ratio per kV)
                            Dim num_kratios As Integer = Split(line(indice), vbTab).Last
                            ReDim elt_exp_handler(j).line(k).k_ratio(num_kratios - 1)

                            'For each X-ray lines:
                            For l As Integer = 0 To num_kratios - 1
                                indice = indice + 1
                                'Retrieve the theoretical X-ray intensity
                                elt_exp_handler(j).line(k).k_ratio(l).elt_intensity = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                'Retrieve the accelerating voltage value in kV
                                elt_exp_handler(j).line(k).k_ratio(l).kv = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                'Retrieve the experimental k-ratio value
                                elt_exp_handler(j).line(k).k_ratio(l).experimental_value = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                'Previous versions did not have experimental errors
                                If version <> "" Then
                                    'Retrieve the experimental error
                                    elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = Split(line(indice), vbTab).Last
                                    indice = indice + 1
                                Else
                                    'If old file, set the experimental error to 0
                                    elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = 0
                                End If
                                'Retrieve the theoreticla standard intensity
                                elt_exp_handler(j).line(k).k_ratio(l).std_intensity = Split(line(indice), vbTab).Last
                                indice = indice + 1
                                'Retrieve the theoretical k-ratio
                                elt_exp_handler(j).line(k).k_ratio(l).theo_value = Split(line(indice), vbTab).Last
                            Next
                        Next
                    Next
                End If
            End If

            'Initialize the elements (retrive the electron shell energies, ionization cross sections, atomic parameters, MAC, ...)
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    init_element_layer(layer_handler(i).element(j).elt_name, vbNull, layer_handler(i).element(j))
                Next
            Next

            'Calculate the mass thickness for each layer in g/cm² (based on the density and layer thickness)
            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Not a valid BadgerFilm input file. Original error: " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try

    End Sub

    Public Sub export(ByVal file_name As String, ByVal layer_handler() As layer, ByVal elt_exp_handler() As Elt_exp, ByVal toa As Double, ByVal version As String)
        Try
            '*******************************************************
            ' Function used to save the data
            'file_name is the name of the saved file
            'layer_handler contains the geomoetry of the sample
            'elt_exp_handler contains the experimental data (elements, Xray lines, experimental kratios, kV, ...)
            'toa is the takeoff angle in degree
            'version is used for compatibility with older BadgerFilm import functions
            '*******************************************************
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

        'Write the data stored in the tmp variable to the file.
        Dim sw As New StreamWriter(file_name, False)
        sw.Write(tmp)
        sw.Close()

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in export " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub import_Stratagem(ByVal data_file As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        Try
            '******************
            ' Import a Stratagem file
            ' data_file: path and name of the saved file
            ' layer_handler: structure to store the retrived data corresponding to the sample's geometry
            ' elt_exp_handler: structure to store the retrived data corresponding to the expremiental data (analyzed elements, X-ray lines k-ratios, kV, ...)
            ' toa is the takeoff angle in degree
            '******************

            'Tries to open the file and copy its content into the temp variable
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


        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in import_Stratagem " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub import_Stratagem_method(ByVal lines() As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef toa As Double)
        Try
            Dim import_method As Integer = 0
            layer_handler = Nothing
            toa = 40 'by default

            'Format the lines.
            For i As Integer = 0 To UBound(lines)
                lines(i) = Regex.Replace(lines(i), "#.*$", "")
            Next

            For i As Integer = 0 To UBound(lines)
                Dim line As String = Trim(lines(i))
                If line = "" Then Continue For
                If line Like "$StrataImport_1" Then
                    import_method = 1
                    Exit For
                ElseIf line Like "$StrataImport_3" Then
                    import_method = 0
                    Exit For
                End If
            Next

            Select Case import_method
                Case 1
                    For i As Integer = 0 To UBound(lines)
                        Dim line As String = Trim(lines(i))
                        'if blank line, skip to the next one
                        If line = "" Then Continue For   'OrElse line(0) = "#"

                        If line Like "$Geom*" Then
                            toa = Split(Regex.Replace(line, " {2,}", " "), " ")(1)
                        End If

                        If line Like "$K*" Then
                            Dim k_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ") 'Replace 2 or more spaces by 1 space.
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
                                While line Like "$*" = False And i < lines.Count - 1
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


                    Next
                    If layer_handler Is Nothing Then
                        ReDim layer_handler(0)
                        'Else
                        '    ReDim Preserve layer_handler(UBound(layer_handler) + 1)
                    End If
                    layer_handler(UBound(layer_handler)).density = 2.0
                    layer_handler(UBound(layer_handler)).thickness = 1000000000.0
                    layer_handler(UBound(layer_handler)).isfix = True
                    'Set the layer concentration definition by weight fraction
                    layer_handler(UBound(layer_handler)).wt_fraction = True
                    'Set the layer id
                    layer_handler(UBound(layer_handler)).id = UBound(layer_handler)

                    ReDim layer_handler(UBound(layer_handler)).element(UBound(elt_exp_handler))
                    For i As Integer = 0 To UBound(elt_exp_handler)
                        layer_handler(UBound(layer_handler)).element(i).z = elt_exp_handler(i).z
                        layer_handler(UBound(layer_handler)).element(i).conc_wt = 1 / elt_exp_handler.Count


                        layer_handler(UBound(layer_handler)).element(i).isConcFixed = False
                        layer_handler(UBound(layer_handler)).element(i).mother_layer_id = UBound(layer_handler)
                    Next

                Case 0

                    'For each lines:
                    For i As Integer = 0 To UBound(lines)
                        Dim line As String = Trim(lines(i))
                        'if blank line, skip to the next one
                        If line = "" Then Continue For   'OrElse line(0) = "#"
                        'if the line starts by Layer, then add a new layer (a new entry) to layer_handler
                        If line Like "$Layer*" Then
                            If layer_handler Is Nothing Then
                                ReDim layer_handler(0)
                            Else
                                ReDim Preserve layer_handler(UBound(layer_handler) + 1)
                            End If
                            'Retrieve the layer definition
                            Dim layer_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                            If layer_def.Count = 4 Or layer_def.Count = 3 Then
                                'Retrieve the layer density
                                layer_handler(UBound(layer_handler)).density = layer_def(1)
                                'Retrieve the layer thickness
                                layer_handler(UBound(layer_handler)).thickness = layer_def(2)
                                If layer_def.Count = 4 Then
                                    'Retrieve the flag indicating if the layer thickness is fixed
                                    If layer_def(3) = "k" Then
                                        layer_handler(UBound(layer_handler)).isfix = True
                                    Else
                                        layer_handler(UBound(layer_handler)).isfix = False
                                    End If
                                Else
                                    'By deflaut the thickness is fixed
                                    layer_handler(UBound(layer_handler)).isfix = True
                                End If
                            Else
                                'Load default values for the density, layer thickness and if the layer thickness is fixed
                                layer_handler(UBound(layer_handler)).density = 2.0
                                layer_handler(UBound(layer_handler)).thickness = 1000000000.0
                                layer_handler(UBound(layer_handler)).isfix = True
                            End If
                            'Set the layer concentration definition by weight fraction
                            layer_handler(UBound(layer_handler)).wt_fraction = True
                            'Set the layer id
                            layer_handler(UBound(layer_handler)).id = UBound(layer_handler)
                        End If

                        'If the line starts by Elt then retrieve the current element's data
                        If line Like "$Elt*" Then
                            If layer_handler(UBound(layer_handler)).element Is Nothing Then
                                ReDim layer_handler(UBound(layer_handler)).element(0)
                            Else
                                ReDim Preserve layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element) + 1)
                            End If
                            'Format the line.
                            Dim elt_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ")
                            If elt_def.Count = 4 Or elt_def.Count = 3 Then
                                'Retrieve the atomic number
                                layer_handler(UBound(layer_handler)).element(UBound(layer_handler(UBound(layer_handler)).element)).z = elt_def(1)
                                'Retrieve the element's weight concentration
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

                        'if the line starts by Geom the load the takeoff angle
                        If line Like "$Geom*" Then
                            toa = Split(Regex.Replace(line, " {2,}", " "), " ")(1)
                        End If

                        'if the line starts by K, load the k-ratios
                        If line Like "$K*" Then
                            Dim k_def() As String = Split(Regex.Replace(line, " {2,}", " "), " ") 'Replace 2 or more spaces by 1 space.
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

                        'if the line starts by Std, load the standards
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

            End Select


            'initialize the remaining variables which are not in the Stratagem file but are necessary for BadgerFilm
            For i As Integer = 0 To UBound(layer_handler)
                Dim tot As Double = 0
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    'Convert the atomic number of the element into its symbol
                    layer_handler(i).element(j).elt_name = Z_to_symbol(layer_handler(i).element(j).z)
                    'initialize the current element (load the atomic parameters, MAC, ...)
                    init_element_layer(layer_handler(i).element(j).elt_name, vbNull, layer_handler(i).element(j))
                    tot = tot + layer_handler(i).element(j).conc_wt
                Next
                ' if the sum of all the elemental concentration in the current layer are >2.0 then normalize the concentrations
                If tot > 2.0 Then
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        layer_handler(i).element(j).conc_wt = layer_handler(i).element(j).conc_wt / tot
                    Next
                End If
            Next

            'Calculate the mass thickness
            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            Next

            'Calculate the atomic fractions based on the weight fractions
            For i As Integer = 0 To UBound(layer_handler)
                convert_wt_to_at(layer_handler, i)
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Not a valid STRATAgem file. Original error: " & ex.Message

        Using err As StreamWriter = New StreamWriter("log.txt", True)
            err.WriteLine(tmp)
        End Using
        MessageBox.Show(tmp)
        End Try
    End Sub
End Module
