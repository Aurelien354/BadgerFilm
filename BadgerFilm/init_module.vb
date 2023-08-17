Imports System.IO

Module init_module
    Public Function read_data(ByVal path As String) As String
        Dim sr As New StreamReader(path)
        Try
            read_data = sr.ReadToEnd
        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            If (sr IsNot Nothing) Then
                sr.Close()
            End If
        End Try
    End Function
    Public Sub init_atomic_parameters(ByVal pen_path As String, ByVal eadl_path As String, ByVal ffast_path As String, ByRef at_data() As String, ByRef el_ion_xs()() As String, ByRef ph_ion_xs()() As String,
                                      ByRef MAC_data_PEN14()() As String, ByRef MAC_data_PEN18()() As String, ByRef MAC_data_FFAST()() As String, ByVal options As options)

        Dim data_file As String = eadl_path & "\pdrelax.p11"
        Dim sr As New StreamReader(data_file)
        Try
            Dim temp As String = sr.ReadToEnd
            at_data = Split(temp, vbCrLf)
        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            If (sr IsNot Nothing) Then
                sr.Close()
            End If
        End Try

        Dim path As String = pen_path & "\PENELOPE2018"
        Dim filename As String
        ReDim el_ion_xs(98)
        For z As Integer = 1 To 99
            If z < 10 Then
                filename = "pdesi0" & z & ".p14" '".p14.crypt"
            Else
                filename = "pdesi" & z & ".p14"
            End If
            Dim temp As String = read_data(path & "\" & filename) ' decrypt(path, filename)

            'Separate each lines
            el_ion_xs(z - 1) = Split(temp, vbCrLf)
        Next

        'path = pen_path '& "\PenelopeData"
        'Dim filename As String
        ReDim ph_ion_xs(98)
        For z As Integer = 1 To 99
            If z < 10 Then
                filename = "phpixs0" & z '& ".crypt"
            Else
                filename = "phpixs" & z '& ".crypt"
            End If
            Dim temp As String = read_data(path & "\" & filename) 'decrypt(path, filename)

            ph_ion_xs(z - 1) = Split(temp, vbCrLf)
        Next

        'Dim MAC_model As String = options.MAC_mode
        'Dim path As String = pen_path '& "\PenelopeData"
        'If MAC_model = "PENELOPE2014" Then
        '    path = path & "\ori"
        'End If

        'Dim filename As String
        ReDim MAC_data_PEN18(98)
        For z As Integer = 1 To 99
            If z < 10 Then
                filename = "phmaxs0" & z '& ".crypt"
            Else
                filename = "phmaxs" & z '& ".crypt"
            End If
            Dim temp As String = read_data(path & "\" & filename) 'decrypt(path, filename)

            'Separate each lines
            MAC_data_PEN18(z - 1) = Split(temp, vbCrLf)
        Next

        'Dim MAC_model As String = options.MAC_mode
        'Dim path As String = pen_path '& "\PenelopeData"
        'If MAC_model = "PENELOPE2014" Then
        path = pen_path & "\PENELOPE2014"
        'End If

        'Dim filename As String
        ReDim MAC_data_PEN14(98)
        For z As Integer = 1 To 99
            If z < 10 Then
                filename = "phmaxs0" & z '& ".crypt"
            Else
                filename = "phmaxs" & z '& ".crypt"
            End If
            Dim temp As String = read_data(path & "\" & filename) 'decrypt(path, filename)

            'Separate each lines
            MAC_data_PEN14(z - 1) = Split(temp, vbCrLf)
        Next

        path = ffast_path
        ReDim MAC_data_FFAST(98)
        For z As Integer = 1 To 92 'FFAST goes only up to 92 (U)!!!!!!!!!!!
            If z < 10 Then
                filename = "FFAST_0" & z & ".txt"
            Else
                filename = "FFAST_" & z & ".txt"
            End If
            Dim temp As String = read_data(path & "\" & filename) 'decrypt(path, filename)

            'Separate each lines
            MAC_data_FFAST(z - 1) = Split(temp, vbCrLf)
        Next

        For z As Integer = 93 To 99
            MAC_data_FFAST(z - 1) = MAC_data_PEN14(z - 1)
        Next


    End Sub

    '***********************************************
    'Read the caracteristic energy of all the elements and all the electronic shells available 
    'in the file pdatconf.pen (from PENELOPE database)
    'Fill the input string array with the unformatted data (one shell by line)
    '***********************************************
    Public Sub init_Ec(ByRef Ec_data() As String, ByVal pen_path As String)
        Dim data_file As String = pen_path & "\PENELOPE2018\pdatconf.p14"
        Dim tmp As String = ""

        'Dim path As String = pen_path '& "\PenelopeData"
        'Dim filename As String = "pdatconf.p14.crypt"
        'tmp = decrypt(path, filename)

        Dim mystream As New StreamReader(data_file)

        Try
            If (mystream IsNot Nothing) Then
                tmp = mystream.ReadToEnd
            End If

        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try

        If tmp = "" Then
            MsgBox("Error in init_Ec")
            Stop
            Exit Sub
        End If

        Dim lines() As String = Split(tmp, vbCrLf)

        Dim indice_end As Integer = 0
        For i As Integer = UBound(lines) To 0 Step -1
            If Trim(lines(i)) <> "" Then Exit For
            indice_end = indice_end + 1
        Next

        Dim indice As Integer = 0
        While (lines(indice))(0) = "#"
            indice = indice + 1
        End While

        ReDim Ec_data(UBound(lines) - indice - indice_end)
        For i As Integer = indice To UBound(lines) - indice_end
            Ec_data(i - indice) = lines(i)
        Next

    End Sub

    Public Sub init_element_layer(ByVal name As String, ByVal atomic_mass As Double, ByRef element As Elt_layer)

        element.elt_name = correct_symbol(name)
        element.z = symbol_to_Z(element.elt_name)

        If atomic_mass = vbNull Then
            element.a = zaro(element.z)(0)
        Else
            element.a = atomic_mass
        End If


    End Sub
    Public Sub init_element(ByVal name As String, ByVal xray_name As String, ByVal atomic_mass As Double, ByVal Ec_data() As String, ByRef element As Elt_exp,
                            ByVal at_data() As String, ByVal el_ion_xs()() As String, ByVal ph_ion_xs()() As String, ByVal MAC_data()() As String, ByVal options As options) ', Optional mother_layer_id As Integer = 0)


        element.elt_name = correct_symbol(name)
        'element.concentration = concentration
        element.z = symbol_to_Z(element.elt_name)
        If element.z = 0 Then
            MsgBox("Atomic number not found in symbol_to_Z")
            Stop
            'Return -1
        End If

        If atomic_mass = vbNull Or atomic_mass < 0 Then
            element.a = zaro(element.z)(0)
        Else
            element.a = atomic_mass
        End If

        element.at_data = init_atomic_parameters_pdrelax(element.z, at_data)
        element.el_ion_xs = init_el_ionization_xs(element.z, el_ion_xs)
        element.ph_ion_xs = init_photoionization_xs(element.z, ph_ion_xs)
        element.mac_data = init_mac(element.z, MAC_data, options)

        If xray_name <> "" Then
            If element.line Is Nothing Then
                ReDim element.line(0)
            Else
                ReDim Preserve element.line(UBound(element.line) + 1)
            End If

            element.line(UBound(element.line)).xray_name = xray_name
            Dim shell1, shell2 As Integer
            Siegbahn_to_transition_num(element.line(UBound(element.line)).xray_name, shell1, shell2, element.elt_name)
            element.line(UBound(element.line)).Ec = find_Ec(element.z, shell1, Ec_data) '0.288 'in keV  'function find_Ec.
            Dim Ec_shell2 As Double = find_Ec(element.z, shell2, Ec_data)
            element.line(UBound(element.line)).xray_energy = element.line(UBound(element.line)).Ec - Ec_shell2  'in keV 
        End If

        Dim first_line_of_interest As Integer = -1
        Dim count As Integer = 0
        For i As Integer = 0 To UBound(Ec_data)
            If (Trim(Mid(Ec_data(i), 2, 2)) <> element.z) Then
                If first_line_of_interest = -1 Then
                    Continue For
                Else
                    Exit For
                End If
            End If

            If first_line_of_interest = -1 Then first_line_of_interest = i
            count = count + 1
        Next

        ReDim element.Ec_data(count - 1)
        For i As Integer = 0 To count - 1
            element.Ec_data(i) = Ec_data(i + first_line_of_interest)
        Next

        'element.mother_layer_id = mother_layer_id

    End Sub

    Public Sub init_element_Xray_line_only(ByVal xray_name As String, ByVal Ec_data() As String, ByRef element As Elt_exp, ByVal line_indice As Integer)
        If xray_name <> "" Then
            element.line(line_indice).xray_name = xray_name
            Dim shell1, shell2 As Integer
            Siegbahn_to_transition_num(xray_name, shell1, shell2, element.elt_name)
            element.line(line_indice).Ec = find_Ec(element.z, shell1, Ec_data) '0.288 'in keV  'function find_Ec.
            Dim Ec_shell2 As Double = find_Ec(element.z, shell2, Ec_data)
            element.line(line_indice).xray_energy = element.line(line_indice).Ec - Ec_shell2  'in keV  'Gives the X-ray line energy based on the energy of the electron shells (fair approximation).
            'Debug.Print(element.elt_name & " " & element.line(line_indice).xray_name & " " & element.line(line_indice).xray_energy)
        End If

    End Sub
    '***********************************************
    'Extract the atomic data from pdrelax.p11 (PENELOPE databse)
    '***********************************************
    Public Function init_atomic_parameters_pdrelax(ByVal Z As Integer, ByVal at_data() As String) As data_atomic_parameters
        init_atomic_parameters_pdrelax.Z = Z
        Dim shell1() As Integer = Nothing
        Dim shell2() As Integer = Nothing
        Dim shell3() As Integer = Nothing
        Dim energy() As Double = Nothing
        Dim trans_prob() As Double = Nothing

        'Dim data_file As String = eadl_path & "\pdrelax.p11"

        'Dim mystream As New StreamReader(data_file)
        Try
            'If (mystream IsNot Nothing) Then
            Dim flag_finished As Boolean = False
            Dim start_line As Integer = 0
            Dim end_line As Integer = 0
            'mystream.ReadLine()
            'While mystream.EndOfStream = False
            For i As Integer = 1 To UBound(at_data)
                Dim tmp_line As String = at_data(i) 'mystream.ReadLine()
                If Trim(Mid(tmp_line, 3, 2)) = Z Then
                    flag_finished = True

                    If shell1 Is Nothing Then
                        ReDim shell1(0)
                        ReDim shell2(0)
                        ReDim shell3(0)
                        ReDim trans_prob(0)
                        ReDim energy(0)
                    Else
                        ReDim Preserve shell1(UBound(shell1) + 1)
                        ReDim Preserve shell2(UBound(shell2) + 1)
                        ReDim Preserve shell3(UBound(shell3) + 1)
                        ReDim Preserve trans_prob(UBound(trans_prob) + 1)
                        ReDim Preserve energy(UBound(energy) + 1)
                    End If
                    shell1(UBound(shell1)) = Trim(Mid(tmp_line, 5, 3))
                    shell2(UBound(shell2)) = Trim(Mid(tmp_line, 8, 3))
                    shell3(UBound(shell3)) = Trim(Mid(tmp_line, 11, 3))
                    trans_prob(UBound(trans_prob)) = Trim(Mid(tmp_line, 14, 13))
                    energy(UBound(energy)) = Trim(Mid(tmp_line, 27, 13)) / 1000

                End If
                If flag_finished = True And Trim(Mid(tmp_line, 3, 2)) <> Z Then
                    'Exit While
                    Exit For
                End If
            Next
            'End While
            ' End If

        Catch Ex As Exception
            MessageBox.Show(Ex.Message)
            'Finally
            'If (mystream IsNot Nothing) Then
            '    mystream.Close()
            'End If
        End Try

        init_atomic_parameters_pdrelax.shell1 = shell1
        init_atomic_parameters_pdrelax.shell2 = shell2
        init_atomic_parameters_pdrelax.shell3 = shell3
        init_atomic_parameters_pdrelax.energy = energy
        init_atomic_parameters_pdrelax.transition_probability = trans_prob

        Return init_atomic_parameters_pdrelax
    End Function

    '***********************************************
    'Extract the atomic data from EADL file for a given element definied by the input parameter Z (atomic number)
    '***********************************************
    Public Function init_atomic_parameters(ByVal Z As Integer, ByVal eadl_path As String) As data_atomic_parameters
        init_atomic_parameters.Z = Z
        Dim shell1() As Integer = Nothing
        Dim shell2() As Integer = Nothing
        Dim shell3() As Integer = Nothing
        Dim energy() As Double = Nothing
        Dim trans_prob() As Double = Nothing


        Dim num As Integer = 28533 'No other choice

        Dim data_file As String
        If Z = 3 Then
            ReDim shell1(1)
            ReDim shell2(1)
            ReDim shell3(1)
            ReDim trans_prob(1)
            ReDim energy(1)
            shell1(0) = 1
            shell2(0) = 2
            shell3(0) = 0
            trans_prob(0) = 0.00029
            energy(0) = 54.3

            shell1(1) = 1
            shell2(1) = 2
            shell3(1) = 2
            trans_prob(1) = 0.99971
            energy(1) = 48.0

            init_atomic_parameters.shell1 = shell1
            init_atomic_parameters.shell2 = shell2
            init_atomic_parameters.shell3 = shell3
            init_atomic_parameters.energy = energy
            init_atomic_parameters.transition_probability = trans_prob

            Return init_atomic_parameters
        ElseIf Z = 4 Then
            ReDim shell1(1)
            ReDim shell2(1)
            ReDim shell3(1)
            ReDim trans_prob(1)
            ReDim energy(1)
            shell1(0) = 1
            shell2(0) = 2
            shell3(0) = 0
            trans_prob(0) = 0.00036
            energy(0) = 108.5

            shell1(1) = 1
            shell2(1) = 2
            shell3(1) = 2
            trans_prob(1) = 0.99964
            energy(1) = 96.4

            init_atomic_parameters.shell1 = shell1
            init_atomic_parameters.shell2 = shell2
            init_atomic_parameters.shell3 = shell3
            init_atomic_parameters.energy = energy
            init_atomic_parameters.transition_probability = trans_prob

            Return init_atomic_parameters
        ElseIf Z = 5 Then
            ReDim shell1(1)
            ReDim shell2(1)
            ReDim shell3(1)
            ReDim trans_prob(1)
            ReDim energy(1)
            shell1(0) = 1
            shell2(0) = 3
            shell3(0) = 0
            trans_prob(0) = 0.00057
            energy(0) = 183.3

            shell1(1) = 1
            shell2(1) = 3
            shell3(1) = 3
            trans_prob(1) = 0.99943
            energy(1) = 169.0

            init_atomic_parameters.shell1 = shell1
            init_atomic_parameters.shell2 = shell2
            init_atomic_parameters.shell3 = shell3
            init_atomic_parameters.energy = energy
            init_atomic_parameters.transition_probability = trans_prob

            Return init_atomic_parameters
        End If

        If Z < 10 Then
            data_file = eadl_path & "\EADL0" & Z & ".txt" '& "\EADL\EADL0" & Z & ".txt"
        Else
            data_file = eadl_path & "\EADL" & Z & ".txt" '& "\EADL\EADL" & Z & ".txt"
        End If

        Dim mystream As New StreamReader(data_file)
        Try
            If (mystream IsNot Nothing) Then
                Dim tmp As String = Nothing
                Dim cnt As Integer = 0

                While mystream.Peek <> -1
                    tmp = mystream.ReadLine()
                    If IsNumeric(Mid(tmp, 71, 5)) = False Then Continue While
                    If (Mid(tmp, 71, 5) <> num) Then
                        Continue While
                    Else
                        Exit While
                    End If
                End While

                Dim nbr_shell As Integer = Trim(Mid(tmp, 45, 11))

                Dim nbr_data = 0
                For i As Integer = 1 To nbr_shell
                    tmp = mystream.ReadLine()
                    Dim tmp2 = Replace(tmp, "+", "E+")
                    tmp2 = Replace(tmp2, "-", "E-")

                    Dim current_shell As Integer = Mid(tmp2, 1, 12)
                    Dim nbr_transitions As Integer = Trim(Mid(tmp, 56, 11))

                    mystream.ReadLine()

                    If shell1 Is Nothing Then
                        ReDim shell1(nbr_transitions - 1)
                        ReDim shell2(nbr_transitions - 1)
                        ReDim shell3(nbr_transitions - 1)
                        ReDim trans_prob(nbr_transitions - 1)
                        ReDim energy(nbr_transitions - 1)
                    Else
                        ReDim Preserve shell1(UBound(shell1) + nbr_transitions)
                        ReDim Preserve shell2(UBound(shell2) + nbr_transitions)
                        ReDim Preserve shell3(UBound(shell3) + nbr_transitions)
                        ReDim Preserve trans_prob(UBound(trans_prob) + nbr_transitions)
                        ReDim Preserve energy(UBound(energy) + nbr_transitions)
                    End If

                    For j As Integer = 1 To nbr_transitions
                        tmp = Trim(mystream.ReadLine())
                        Dim tmp3() As String = Split(tmp, " ")
                        shell1(nbr_data) = current_shell
                        shell2(nbr_data) = Replace(Replace(tmp3(0), "+", "E+"), "-", "E-")
                        shell3(nbr_data) = Replace(Replace(tmp3(1), "+", "E+"), "-", "E-")
                        energy(nbr_data) = Replace(Replace(tmp3(2), "+", "E+"), "-", "E-") / 1000 'convert eV to keV
                        trans_prob(nbr_data) = Replace(Replace(tmp3(3), "+", "E+"), "-", "E-")

                        nbr_data = nbr_data + 1
                    Next
                Next
            End If

        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)

        Finally
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try

        init_atomic_parameters.shell1 = shell1
        init_atomic_parameters.shell2 = shell2
        init_atomic_parameters.shell3 = shell3
        init_atomic_parameters.energy = energy
        init_atomic_parameters.transition_probability = trans_prob

        Return init_atomic_parameters
    End Function

    '***********************************************
    'Extract the electron ionization cross section from PENELOPE databse for a given element
    'definied by the input parameter Z (atomic number)
    '***********************************************
    Public Function init_el_ionization_xs(ByVal Z As Integer, ByVal el_ion_xs()() As String) As data_xs
        init_el_ionization_xs.Z = Z

        'Dim file_name As String
        'If Z < 10 Then
        '    file_name = pen_path & "\PenelopeData\pdesi0" & Z & ".p14"
        'Else
        '    file_name = pen_path & "\PenelopeData\pdesi" & Z & ".p14"
        'End If

        ''read all the data
        'Dim sr As StreamReader = New StreamReader(file_name)
        'Dim temp = sr.ReadToEnd()
        'sr.Close()

        'Dim path As String = pen_path '& "\PenelopeData"
        'Dim filename As String
        'If Z < 10 Then
        '    filename = "pdesi0" & Z & ".p14.crypt"
        'Else
        '    filename = "pdesi" & Z & ".p14.crypt"
        'End If
        'Dim temp As String = decrypt(path, filename)

        ''Separate each lines
        Dim lines() As String = el_ion_xs(Z - 1) 'Split(temp, vbCrLf)

        Dim FIRST_LINE As Integer = 0 'skip the first lines
        While lines(FIRST_LINE)(0) = "#"
            FIRST_LINE = FIRST_LINE + 1
        End While

        ReDim init_el_ionization_xs.energy(UBound(lines) - FIRST_LINE)
        Dim tmp_init() As String = Split(Trim(lines(FIRST_LINE)), " ")
        ReDim init_el_ionization_xs.cross_section(UBound(lines) - FIRST_LINE, UBound(tmp_init) - 1)

        For i As Integer = FIRST_LINE To UBound(lines) 'skip the first lines
            If Trim(lines(i)) = "" Then Continue For
            Dim tmp() As String = Split(Trim(lines(i)), " ")
            init_el_ionization_xs.energy(i - FIRST_LINE) = tmp(0) / 1000 'convert eV to keV
            For j As Integer = 1 To UBound(tmp)
                init_el_ionization_xs.cross_section(i - FIRST_LINE, j - 1) = tmp(j)
            Next
        Next

        Return init_el_ionization_xs

    End Function

    '***********************************************
    'Extract the photon ionization cross section from PENELOPE databse for a given element
    'definied by the input parameter Z (atomic number)
    '***********************************************
    Public Function init_photoionization_xs(ByVal Z As Integer, ByVal ph_ion_xs()() As String) As data_xs
        init_photoionization_xs.Z = Z

        'Dim file_name As String
        'If Z < 10 Then
        '    file_name = pen_path & "\PenelopeData\phpixs0" & Z
        'Else
        '    file_name = pen_path & "\PenelopeData\phpixs" & Z
        'End If

        ''Debug.Print(Z)

        'Dim sr As StreamReader = New StreamReader(file_name)
        'Dim temp = sr.ReadToEnd()
        'sr.Close()

        'Dim path As String = pen_path '& "\PenelopeData"
        'Dim filename As String
        'If Z < 10 Then
        '    filename = "phpixs0" & Z & ".crypt"
        'Else
        '    filename = "phpixs" & Z & ".crypt"
        'End If
        'Dim temp As String = decrypt(path, filename)

        Dim lines() As String = ph_ion_xs(Z - 1) 'Split(temp, vbCrLf)
        Dim read_energy As Double = 0
        Dim FIRST_LINE As Integer = 2


        ReDim init_photoionization_xs.energy(UBound(lines) - FIRST_LINE)
        Dim tmp_init() As String = Split(Trim(lines(FIRST_LINE)), " ")
        ReDim init_photoionization_xs.cross_section(UBound(lines) - FIRST_LINE, UBound(tmp_init) - 1)

        For i As Integer = FIRST_LINE To UBound(lines) 'skip the first 2 lines
            If Trim(lines(i)) = "" Then Continue For
            Dim tmp() As String = Split(Trim(lines(i)), " ")
            init_photoionization_xs.energy(i - FIRST_LINE) = tmp(0) / 1000 'convert eV to keV
            For j As Integer = 1 To UBound(tmp)
                init_photoionization_xs.cross_section(i - FIRST_LINE, j - 1) = tmp(j)
            Next
        Next

        Return init_photoionization_xs

    End Function

    Public Function init_mac(ByVal Z As Integer, ByVal MAC_data()() As String, ByVal options As options) As data_xs
        If options.MAC_mode = "MAC30" Then Exit Function
        init_mac.Z = Z
        'Dim file_name As String
        'If Z < 10 Then
        '    file_name = pen_path & "\PenelopeData\phmaxs0" & Z
        'Else
        '    file_name = pen_path & "\PenelopeData\phmaxs" & Z
        'End If

        'Dim temp As String = Nothing

        'Dim sr As StreamReader = New StreamReader(file_name)
        'Try
        '    If (sr IsNot Nothing) Then
        '        temp = sr.ReadToEnd()
        '        sr.Close()
        '    End If
        'Catch Ex As Exception
        '    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        'Finally
        '    If (sr IsNot Nothing) Then
        '        sr.Close()
        '    End If
        'End Try
        'Dim MAC_model As String = options.MAC_mode
        'Dim path As String = pen_path '& "\PenelopeData"
        'If MAC_model = "PENELOPE2014" Then
        '    path = path & "\ori crypt"
        'End If

        'Dim filename As String
        'If Z < 10 Then
        '    filename = "phmaxs0" & Z & ".crypt"
        'Else
        '    filename = "phmaxs" & Z & ".crypt"
        'End If
        'Dim temp As String = decrypt(path, filename)


        'Separate each lines
        Dim lines() As String = MAC_data(Z - 1) 'Split(temp, vbCrLf)

        Dim FIRST_LINE As Integer = 0 'skip the first lines
        Dim energy_conversion_factor As Integer = 1

        If options.MAC_mode = "PENELOPE2014" Or options.MAC_mode = "PENELOPE2018" Or (options.MAC_mode = "FFAST" And Z > 92) Then
            While Trim(lines(FIRST_LINE))(0) = "#"
                FIRST_LINE = FIRST_LINE + 1
            End While
            energy_conversion_factor = 1000
        ElseIf options.MAC_mode = "FFAST" Then
            FIRST_LINE = 3
        End If

        For i As Integer = FIRST_LINE To UBound(lines)
            lines(i) = lines(i).Replace("  ", " ")
        Next

        ReDim init_mac.energy(UBound(lines) - FIRST_LINE)
        Dim tmp_init() As String = Split(Trim(lines(FIRST_LINE)), " ")
        ReDim init_mac.cross_section(UBound(lines) - FIRST_LINE, UBound(tmp_init) - 1)

        For i As Integer = FIRST_LINE To UBound(lines) 'skip the first lines
            If Trim(lines(i)) = "" Then Continue For
            Dim tmp() As String = Split(Trim(lines(i)), " ")
            init_mac.energy(i - FIRST_LINE) = tmp(0) / energy_conversion_factor 'convert eV to keV for PENELOPE data
            For j As Integer = 1 To UBound(tmp)
                init_mac.cross_section(i - FIRST_LINE, j - 1) = tmp(j)
            Next
        Next


    End Function

    '***********************************************
    'Find a critical ionization energy for the given element z and the given subshell shell
    '***********************************************
    Public Function find_Ec(ByVal z As Integer, ByVal shell As Integer, ByVal Ec_data() As String) As Double
        If Ec_data Is Nothing Then
            MsgBox("Error: Ec_data is nothing in find_Ec")
            Stop
        End If

        For i As Integer = 0 To UBound(Ec_data)
            'Dim lol As String = Mid(Ec_data(i), 2, 2)
            If (Trim(Mid(Ec_data(i), 2, 2)) <> z) Then Continue For

            'Dim lol2 As String = Mid(Ec_data(i), 4, 4)
            If (Trim(Mid(Ec_data(i), 4, 4)) <> shell) Then Continue For

            'Dim lol As Double = Trim(Mid(Ec_data(i), 22, 11))
            'Dim oldDecimalSeparator As String = Application.CurrentCulture.NumberFormat.NumberDecimalSeparator
            find_Ec = Trim(Mid(Ec_data(i), 22, 11)) / 1000 'convert eV to keV!!!!!!!!!!!!!!!!!!!!!!!!!!!
            Return find_Ec
        Next

        'MsgBox("Error: Ec not found in find_Ec")
        'Stop
        Return 0
    End Function

    '***********************************************
    'Find and return all the critical ionization energies for the entire system layer_handler()
    '***********************************************
    'Public Function find_all_Ec(ByVal layer_handler() As layer) As Double()
    '    If layer_handler Is Nothing Then
    '        MsgBox("Error: layer_handler is nothing in find_all_Ec")
    '        Stop
    '    End If
    '    Dim results() As Double = Nothing

    '    For i As Integer = 0 To UBound(layer_handler)
    '        For j As Integer = 0 To UBound(layer_handler(i).element)
    '            Dim curr_indice As Integer = 0
    '            If results Is Nothing Then
    '                ReDim results(UBound(layer_handler(i).element(j).Ec_data))
    '            Else
    '                curr_indice = UBound(results) + 1
    '                ReDim Preserve results(UBound(results) + layer_handler(i).element(j).Ec_data.Count)
    '            End If
    '            For k As Integer = 0 To UBound(layer_handler(i).element(j).Ec_data)
    '                results(curr_indice + k) = Trim(Mid(layer_handler(i).element(j).Ec_data(k), 22, 11)) / 1000 'convert eV to keV!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '            Next
    '        Next
    '    Next

    '    Return results
    'End Function

    '***********************************************
    'Find and return all the critical ionization energies for the entire system layer_handler() that are between Ec and E0
    '***********************************************
    Public Function find_all_Ec(ByVal elt_exp_all() As Elt_exp, ByVal Ec As Double, ByVal E0 As Double) As Double()
        If elt_exp_all Is Nothing Then
            MsgBox("Error: elt_exp_all is nothing in find_all_Ec")
            Stop
        End If
        Dim results() As Double = {Ec, E0}

        For i As Integer = 0 To UBound(elt_exp_all)
            Dim ind_num_edge_between_Ec_E0 As Integer = 0
            Dim val_between_Ec_E0() As Double = Nothing
            For j As Integer = 0 To UBound(elt_exp_all(i).Ec_data)
                Dim val As Double = Trim(Mid(elt_exp_all(i).Ec_data(j), 22, 11)) / 1000 'convert eV to keV!
                If val > Ec And val < E0 Then
                    ReDim Preserve results(UBound(results) + 1)
                    results(UBound(results)) = val
                End If
            Next
        Next

        Array.Sort(results)
        Array.Reverse(results)

        Return results
    End Function

    '***********************************************
    'Extract the fluorescence rate
    '***********************************************
    Public Function extract_fluorescence_yield(ByVal atomic_parameters As data_atomic_parameters, ByVal shell1_to_find As Integer) As Double

        extract_fluorescence_yield = 0
        With atomic_parameters
            extract_fluorescence_yield = fluorescence_rate(.shell1, .shell2, .shell3, .energy, .transition_probability, shell1_to_find)
        End With

    End Function

    '***********************************************
    'Extract the emission rate
    '***********************************************
    Public Function extract_emission_rate(ByVal atomic_parameters As data_atomic_parameters, ByVal shell1_to_find As Integer, ByVal shell2_to_find As Integer) As Double

        extract_emission_rate = 0
        With atomic_parameters
            extract_emission_rate = emission_rate(.shell1, .shell2, .shell3, .energy, .transition_probability, shell1_to_find, shell2_to_find)
        End With

    End Function

    '***********************************************
    'Extract the transition rate
    '***********************************************
    Public Function extract_transition_rate(ByVal atomic_parameters As data_atomic_parameters, ByVal shell1_to_find As Integer, ByVal shell2_to_find As Integer, Optional ByVal starting_index As Integer = 0) As Double

        extract_transition_rate = 0
        With atomic_parameters
            extract_transition_rate = transition_rate(.shell1, .shell2, .shell3, .energy, .transition_probability, shell1_to_find, shell2_to_find, starting_index)
        End With

    End Function

    '***********************************************
    'Calculate the fluorescence rate for a given transition
    '***********************************************
    Public Function fluorescence_rate(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                           shell1_to_find As Integer) As Double

        fluorescence_rate = trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, -1, 0) '/ trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, -1, -1)

    End Function

    '***********************************************
    'Calculate the emission rate for a given transition
    '***********************************************
    Public Function emission_rate(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                           shell1_to_find As Integer, shell2_to_find As Integer) As Double

        emission_rate = trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, shell2_to_find, 0) / trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, -1, 0)

    End Function

    '***********************************************
    'Calculate the transition rate for a given transition
    '***********************************************
    Public Function transition_rate(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                           shell1_to_find As Integer, shell2_to_find As Integer, Optional ByVal starting_index As Integer = 0) As Double

        transition_rate = trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, shell2_to_find, -1, starting_index) +
            trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, -1, shell2_to_find, starting_index)
        'transition_rate = transition_rate / trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, -1, -1, starting_index)

    End Function

    '***********************************************
    'Calculate the Coster-Kronig coefficient for the given transition.
    'This also includes Super-Coster-Kronig transitions and Auger transitions.
    'Basically all the decay processes where an electron is emitted.
    '***********************************************
    Public Function Coster_Kronig(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                           shell1_to_find As Integer, shell2_to_find As Integer) As Double

        Coster_Kronig = trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, shell2_to_find, -1) - trouve(shell1, shell2, shell3, energy, trans_prob, shell1_to_find, shell2_to_find, 0)

    End Function

    Public Function trouve(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                           shell1_to_find As Integer, shell2_to_find As Integer, special As Integer, Optional ByVal starting_index As Integer = 0) As Double
        '********************************************************************
        '* The function TROUVE computes the data
        '* This function finds And sums the desired transition probabilities
        '* SHELL1 Is the shell number where the vacancy Is located at the beginning of the process
        '* SHELL2 Is the shell number where the vacancy Is located at the end of the process
        '*        if SHELL2 = -1, all the shells are taken into account
        '* SPECIAL Is the type of transition involved during the relaxation process
        '*         if SPECIAL = 0, a photon Is emitted
        '*         if SPECIAL = X with X>0, an Auger electron Is emitted from the shell number X
        '*         if SPECIAL = -1, all the processes are taken into account
        '******************************************************************** 

        trouve = 0
        Dim flag_trouve As Boolean = False
        For i As Integer = starting_index To UBound(shell1)
            If shell1(i) = shell1_to_find Then
                flag_trouve = True
                If shell2(i) = shell2_to_find Or shell2_to_find < 0 Then
                    If shell3(i) = special Or special < 0 Then
                        trouve = trouve + trans_prob(i)
                    End If
                End If
            Else
                If flag_trouve = True Then
                    Return trouve
                End If
            End If
        Next

    End Function

    '***********************************************
    'Find the energy of the emitted particle (electron or photon) after a given transition
    '***********************************************
    Public Function trouve_energy(ByVal shell1() As Integer, ByVal shell2() As Integer, ByVal shell3() As Integer, ByVal energy() As Double, ByVal trans_prob() As Double,
                          shell1_to_find As Integer, shell2_to_find As Integer, special As Integer) As Double
        'In keV !!!!!!!!!!!!!!!!!!
        trouve_energy = 0
        For i As Integer = 0 To UBound(shell1)
            If shell1(i) = shell1_to_find Then
                If shell2(i) = shell2_to_find Or shell2_to_find < 0 Then
                    If shell3(i) = special Or special < 0 Then
                        trouve_energy = trouve_energy + energy(i) '/ 1000
                    End If
                End If
            End If
        Next

    End Function

    '***********************************************
    'Calculate w * em_rate / A   of the studied element
    '***********************************************
    Public Function constante_simple(ByVal studied_element As Elt_exp, ByVal shell1 As Integer, ByVal shell2 As Integer) As Double
        Dim w As Double = extract_fluorescence_yield(studied_element.at_data, shell1)
        Dim em_rate As Double = extract_emission_rate(studied_element.at_data, shell1, shell2)

        constante_simple = w * em_rate / studied_element.a

    End Function

    '***********************************************
    'Extract the main x-ray lines of a given element (studied_element)
    'Populate the energy of the main x-rays, the name of these x-rays and their critical energies
    'The main x-ray lines are definied by the 2 dimentional array list_of_main_xrays
    '***********************************************
    Public Sub extract_main_xray_lines(ByVal studied_element As Elt_exp, ByRef Energy_Xrays() As Double, ByRef Name_Xrays() As String, ByRef Ec_of_xray() As Double)

        If studied_element.at_data.shell1 Is Nothing Then
            MsgBox("Error in extract_main_xray_lines")
            Stop
            Exit Sub
        End If

        Dim list_of_main_xrays(,) As Integer
        If studied_element.elt_name = "Li" Or studied_element.elt_name = "Be" Then
            list_of_main_xrays = {{1, 2}}
        Else
            '                      Ka1     Ka2     Kb      La1     La2     Lb      Lc1      Ln      Ll      L3N5     Ma1      Ma2      Mb       Mc1      M4O6     M3O5     M2N4     M2N1     M3O1     M3O4    M1N2      M1N3     M2O4     M1O2
            list_of_main_xrays = {{1, 4}, {1, 3}, {1, 7}, {4, 9}, {4, 8}, {3, 8}, {3, 13}, {3, 5}, {4, 5}, {4, 14}, {9, 16}, {9, 15}, {8, 15}, {7, 14}, {8, 22}, {7, 21}, {6, 13}, {6, 10}, {7, 17}, {7, 20}, {5, 11}, {5, 12}, {6, 20}, {5, 18}}
        End If
        Dim temp As Double

        For i As Integer = 0 To UBound(list_of_main_xrays)
            With studied_element.at_data
                If IsNothing(studied_element.at_data.shell1) Or studied_element.at_data.shell1.Length = 0 Then
                    Stop
                End If
                temp = trouve_energy(.shell1, .shell2, .shell3, .energy, .transition_probability, list_of_main_xrays(i, 0), list_of_main_xrays(i, 1), 0)
            End With

            If temp <> 0 Then
                If Energy_Xrays Is Nothing Then
                    ReDim Energy_Xrays(0)
                    ReDim Name_Xrays(0)
                    ReDim Ec_of_xray(0)
                Else
                    ReDim Preserve Energy_Xrays(UBound(Energy_Xrays) + 1)
                    ReDim Preserve Name_Xrays(UBound(Name_Xrays) + 1)
                    ReDim Preserve Ec_of_xray(UBound(Ec_of_xray) + 1)
                End If
                Energy_Xrays(UBound(Energy_Xrays)) = temp
                Name_Xrays(UBound(Name_Xrays)) = convert_transition_num_to_Siegbahn(list_of_main_xrays(i, 0), list_of_main_xrays(i, 1))
                Ec_of_xray(UBound(Ec_of_xray)) = find_Ec(studied_element.z, list_of_main_xrays(i, 0), studied_element.Ec_data)
            End If
        Next

    End Sub

    '***********************************************
    'Calculate the X-ray production cross section of shell1 by electron impact of energy En
    '***********************************************
    Public Function Xray_production_xs_el_impact(ByVal studied_elt As Elt_exp, ByVal shell_destination As Integer, ByVal En As Double) As Double
        Dim sigma_el_xs() As Double = interpol_log_log(studied_elt.el_ion_xs, En)

        Dim res As Double = 0

        For i As Integer = 1 To shell_destination
            If sigma_el_xs(i - 1) <> 0 Then
                Dim trans_rate As Double = recursive_transition_rate(studied_elt.at_data, i, shell_destination)
                res = res + trans_rate * sigma_el_xs(i - 1)
            End If
        Next

        Return res

    End Function

    '***********************************************
    'Calculate the X-ray production cross section of shell1 by photon impact of energy En
    '***********************************************
    Public Function Xray_production_xs_ph_impact(ByVal studied_elt As Elt_exp, ByVal shell_destination As Integer, ByVal En As Double) As Double
        Dim sigma_ph_xs() As Double = interpol_log_log(studied_elt.ph_ion_xs, En)
        If sigma_ph_xs.Length = 1 Then Return 0

        Dim res As Double = 0

        For i As Integer = 1 To shell_destination
            If sigma_ph_xs(i) <> 0 Then
                Dim trans_rate As Double = recursive_transition_rate(studied_elt.at_data, i, shell_destination)
                res = res + trans_rate * sigma_ph_xs(i)
            End If
        Next

        Return res

    End Function

    Public Function recursive_transition_rate(ByVal atomic_parameters As data_atomic_parameters, ByVal shell1 As Integer, ByVal shell2 As Integer, Optional ByVal starting_index As Integer = 0) As Double
        Dim pile As Double
        Dim poupi As Double
        Dim pop As Double

        pile = 0
        'pop = 0
        'poupi = 0

        If (shell1 = shell2) Then
            Return 1
        End If

        Dim indice As Integer = 0
        For i As Integer = starting_index To UBound(atomic_parameters.shell1)
            indice = i
            If atomic_parameters.shell1(i) = shell1 Then
                Exit For
            End If
        Next

        For i As Integer = shell1 + 1 To shell2
            pop = recursive_transition_rate(atomic_parameters, i, shell2, indice)
            poupi = extract_transition_rate(atomic_parameters, shell1, i, indice)
            pile = pile + poupi * pop
        Next

        Return pile

    End Function



    Public Sub init_stoichio(ByRef stoichio_elt As Integer(,), ByVal stoichio_file_path As String)
        Dim data_file As String = Path.Combine(Application.StartupPath, stoichio_file_path)
        Dim tmp As String = ""

        'Dim path As String = pen_path '& "\PenelopeData"
        'Dim filename As String = "pdatconf.p14.crypt"
        'tmp = decrypt(path, filename)

        Dim mystream As New StreamReader(data_file)

        Try
            If (mystream IsNot Nothing) Then
                tmp = mystream.ReadToEnd
            End If

        Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
        Finally
            If (mystream IsNot Nothing) Then
                mystream.Close()
            End If
        End Try

        If tmp = "" Then
            MsgBox("Error in init_stoichio")
            Stop
            Exit Sub
        End If

        Dim lines() As String = Split(tmp, vbCrLf)

        Dim indice_end As Integer = 0
        For i As Integer = UBound(lines) To 0 Step -1
            If Trim(lines(i)) <> "" Then Exit For
            indice_end = indice_end + 1
        Next

        Dim indice As Integer = 0
        While (lines(indice))(0) = "#"
            indice = indice + 1
        End While

        ReDim stoichio_elt(UBound(lines) - indice - indice_end, 1)
        For i As Integer = indice To UBound(lines) - indice_end
            Dim tmpo() As String = Split(lines(i), vbTab)
            stoichio_elt(i - indice, 0) = CInt(tmpo(1))
            stoichio_elt(i - indice, 1) = CInt(tmpo(2))
        Next

    End Sub


End Module
