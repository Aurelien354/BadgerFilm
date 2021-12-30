Imports System.IO

Module mac_module

    Public Function MAC_calculation(ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal layer_id As Integer, ByVal layer_handler() As layer, ByVal elt_exp_all() As Elt_exp,
                                    ByVal fit_MAC As fit_MAC, ByVal options As options) As Double
        Try
            'Sould the concentration be normalized to 1 ????????????????????????????????

            '***************************************************************************
            'MAC of compounds calculated with atomic concentration
            '***************************************************************************
            'Dim sum As Double = 0
            'For i As Integer = 0 To UBound(layer_handler(layer_id).element)
            '    sum = sum + layer_handler(layer_id).element(i).concentration / layer_handler(layer_id).element(i).a
            'Next

            'MAC_calculation = 0
            'For i As Integer = 0 To UBound(layer_handler(layer_id).element)
            '    Dim tmp As Double = layer_handler(layer_id).element(i).concentration / (layer_handler(layer_id).element(i).a * sum)
            '    MAC_calculation = MAC_calculation + tmp * find_mac(layer_handler(layer_id).element(i), Exray_absorbed, mac_handler, fit_param)
            'Next
            '***************************************************************************

            '***************************************************************************
            'MAC of compounds calculated with weight concentration
            '***************************************************************************
            MAC_calculation = 0
            For i As Integer = 0 To UBound(layer_handler(layer_id).element)
                For j As Integer = 0 To UBound(elt_exp_all)
                    If elt_exp_all(j).elt_name = layer_handler(layer_id).element(i).elt_name Then
                        Dim MAC As Double = find_mac(elt_exp_all(j), studied_element.elt_name, studied_element.line(line_indice).xray_name, studied_element.line(line_indice).xray_energy, fit_MAC, options)
                        'Debug.Print(studied_element.elt_name & vbTab & studied_element.line(line_indice).xray_name & vbTab & elt_exp_all(j).elt_name & vbTab & MAC)
                        MAC_calculation = MAC_calculation + layer_handler(layer_id).element(i).fictitious_concentration * MAC 'Oct 21, 2021 AMXX Changed concentration by fictitious_concentration

                        Exit For
                    End If
                Next
            Next
            '***************************************************************************
            'MAC_calculation = 0

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in MAC_calculation " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function find_mac(ByVal absorber_element As Elt_exp, ByVal studied_element_name As String, ByVal studied_element_xray_name As String, ByVal E_photon As Double,
                             ByVal fit_MAC As fit_MAC, ByVal options As options) As Double
        Try
            'studied_element is the absorber
            'E_photon is the radiation being absorbed

            'Convert E_photon which is in keV to eV
            E_photon = E_photon * 1000

            'If absorber_element.elt_name = "O" Then
            '    Dim shell1, shell2 As Integer
            '    Siegbahn_to_transition_num("Ka", shell1, shell2, absorber_element.elt_name)
            '    Dim z As Integer = symbol_to_Z(absorber_element.elt_name)
            '    Dim Ec_shell2 As Double = find_Ec(z, shell2, absorber_element.Ec_data)
            '    Dim energy As Double = find_Ec(z, shell1, absorber_element.Ec_data) - Ec_shell2 'in keV 

            '    If Math.Floor(E_photon) = Math.Floor(energy * 1000) Then
            '        find_mac = 1200
            '        Return find_mac
            '    End If
            'End If
            'If absorber_element.elt_name = "Ti" Then
            '    Dim shell1, shell2 As Integer
            '    Siegbahn_to_transition_num("Ka", shell1, shell2, "Al")
            '    Dim z As Integer = symbol_to_Z("Al")
            '    Dim Ec_shell2 As Double = find_Ec(z, shell2, absorber_element.Ec_data)
            '    Dim energy As Double = find_Ec(z, shell1, absorber_element.Ec_data) - Ec_shell2 'in keV 

            '    If Math.Floor(E_photon) > 1480 And Math.Floor(E_photon) < 1490 Then
            '        find_mac = 2201
            '        Return find_mac
            '    End If
            'End If
            If E_photon = 0 Then Return 0


            If options.experimental_MAC.experimental_MAC_enabled = True Then
                For i As Integer = 0 To UBound(options.experimental_MAC.emitter)
                    If options.experimental_MAC.emitter(i) = studied_element_name Then
                        If options.experimental_MAC.xray_line(i) = studied_element_xray_name Then
                            If options.experimental_MAC.absorber(i) = absorber_element.elt_name Then
                                Return options.experimental_MAC.exp_MAC(i)
                            End If
                        End If
                    End If
                Next
            End If


            If fit_MAC.activated = True Then
                If absorber_element.elt_name = fit_MAC.absorber_elt Then
                    If Math.Abs((E_photon - (fit_MAC.X_ray_energy * 1000)) / E_photon) < 0.001 Then
                        find_mac = fit_MAC.MAC
                        'find_mac = MacFeLaVsConc(studied_element.concentration, E_photon / 1000)
                        Return find_mac
                    End If

                    'If Math.Floor(E_photon) = Math.Floor(fit_MAC.X_ray_energy * 1000) Then
                    '        find_mac = fit_MAC.MAC
                    '        'find_mac = MacFeLaVsConc(studied_element.concentration, E_photon / 1000)
                    '        Return find_mac
                    'End If
                End If
            End If
            'If fit_param <> 0 Then
            'If (studied_element.xray_name = "La" Or studied_element.xray_name = "La1" Or studied_element.xray_name = "La2") And studied_element.z = 26 Then
            '    find_mac = fit_param
            '    'find_mac = MacFeLaVsConc(studied_element.concentration, E_photon / 1000)
            '    Return find_mac
            'End If

            'If studied_element.xray_name = "Lb" And Math.Round(E_photon) = 717 And studied_element.z = 26 Then
            '    find_mac = fit_param
            '    Return find_mac
            'End If
            'End If







            'If E_photon >= 275 And E_photon <= 285 And studied_element.z = 14 Then

            '    Return 35000
            'End If


            Dim MAC_model As String = options.MAC_mode
            If MAC_model = "PENELOPE2018" Or MAC_model = "PENELOPE2014" Then
                '*************************************************
                'PENELOPE MAC extraction
                '*************************************************
                Dim mac_xs() As Double = interpol_log_log(absorber_element.mac_data, E_photon / 1000)

                If mac_xs.Length < 4 Then
                    find_mac = 0
                Else
                    find_mac = mac_xs(UBound(mac_xs))
                End If
                '*************************************************

            ElseIf MAC_model = "MAC30" Then
                '*************************************************
                'Heinrich MAC calculation
                '*************************************************
                find_mac = Heinrich_MAC30(absorber_element.z, E_photon, absorber_element.a, absorber_element.Ec_data)
                '*************************************************

            ElseIf MAC_model = "FFAST" Then
                '*************************************************
                'Chantler MAC extraction
                '*************************************************
                Dim mac_xs() As Double = interpol_log_log(absorber_element.mac_data, E_photon / 1000)

                If mac_xs.Length < 1 Then
                    find_mac = 0
                Else
                    find_mac = mac_xs(UBound(mac_xs))
                End If
                '*************************************************

            Else
                Debug.WriteLine("Error: MAC model unknown!!!")
            End If

            Return find_mac

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in find_mac " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Sub load_experimental_MAC(ByVal path As String, ByRef experimental_MAC As experimental_MAC)
        Try
            Dim sr As New StreamReader(path)
            Dim num_line As Integer = 0
            Try
                Dim temp As String = sr.ReadToEnd
                Dim lines() As String = Split(temp, vbCrLf)
                For i As Integer = 0 To UBound(lines)
                    If Trim(lines(i)) = "" Then Continue For
                    num_line = num_line + 1
                Next

                ReDim experimental_MAC.absorber(num_line - 2)
                ReDim experimental_MAC.emitter(num_line - 2)
                ReDim experimental_MAC.exp_MAC(num_line - 2)
                ReDim experimental_MAC.xray_line(num_line - 2)

                Dim index As Integer = 0
                For i As Integer = 1 To UBound(lines)
                    If Trim(lines(i)) = "" Then Continue For
                    Dim data() As String = Split(lines(i), vbTab)
                    If data.Count <> 4 Then
                        experimental_MAC.emitter = Nothing
                        experimental_MAC.xray_line = Nothing
                        experimental_MAC.absorber = Nothing
                        experimental_MAC.exp_MAC = Nothing
                        experimental_MAC.experimental_MAC_enabled = False
                        MsgBox("Error in Experimental_MACs.txt, line " & i)
                        Exit For
                    End If
                    experimental_MAC.emitter(index) = data(0)
                    experimental_MAC.xray_line(index) = data(1)
                    experimental_MAC.absorber(index) = data(2)
                    experimental_MAC.exp_MAC(index) = data(3)
                    index = index + 1
                Next


            Catch Ex As Exception
                MessageBox.Show("Cannot read Experimental_MACs.txt from disk. Original error: " & Ex.Message)
            Finally
                If (sr IsNot Nothing) Then
                    sr.Close()
                End If
            End Try
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in load_experimental_MAC " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Sub

    'Heinrich KFJ. in Proc. 11th Int. Congr. X-ray Optics & Microanalysis, Brown JD, Packwood RH (eds). Univ. Western Ontario: London, 1986; 67
    Public Function Heinrich_MAC30(ByVal z As Integer, ByVal energy As Double, ByVal At_weight As Double, ByVal Ec_data() As String) As Double
        Try
            Dim EnDat(9) As Double 'shell ionization energies for elt z

            For i As Integer = 0 To 9
                EnDat(i) = find_Ec(z, i + 1, Ec_data) * 1000
            Next


            Dim cc As Double
            Dim az As Double
            Dim nm As Double
            Dim bias As Double = 0.0

            If (energy > EnDat(0)) Then
                If (z < 6) Then
                    cc = 0.001D * (1.808599D * z - 0.287536D)
                    az = (-14.15422D * z + 155.6055D) * z + 24.4545D
                    bias = 18.2D * z - 103D
                    nm = (-0.01273815D * z + 0.02652873D) * z + 3.34745D
                Else
                    cc = 0.00001D * (525.3D + 133.257D * z - 7.5937D * z * z + 0.169357D * z * z * z - 0.0013975D * z * z * z * z)
                    az = ((-0.152624D * z + 6.52D) * z + 47D) * z
                    nm = 3.112D - 0.0121D * z
                    If ((energy > EnDat(0)) And (z >= 50D)) Then
                        az = ((-0.015D * z + 3.52D) * z + 47D) * z
                    End If
                    If ((energy > EnDat(0)) And (z >= 57D)) Then
                        cc = 0.000001D * (200D + 100D * z - z * z)
                    End If
                End If

            ElseIf (energy > EnDat(3)) Then
                Dim c As Double = 0.001D * (-0.0924D + 0.141478D * z - 0.00524999D * z * z + 0.0000985296D * z * z * z)
                c = c - 0.000000000907306D * z * z * z * z + 0.00000000000319245D * z * z * z * z * z
                cc = c
                az = (((-0.000116286D * z + 0.01253775D) * z + 0.067429D) * z + 17.8096D) * z
                nm = (-0.00004982D * z + 0.001889D) * z + 2.7575D
                If ((energy < EnDat(1)) And (energy > EnDat(2))) Then
                    cc = c * 0.858D
                End If
                If (energy < EnDat(2)) Then
                    cc = c * (0.8933D - 0.00829D * z + 0.0000638D * z * z)
                End If

            ElseIf ((energy < EnDat(3)) And (energy > EnDat(4))) Then

                nm = ((0.0000044509D * z - 0.00108246D) * z + 0.084597D) * z + 0.5385D
                Dim c As Double
                If (z < 30D) Then
                    c = (((0.072773258D * z - 11.641145D) * z + 696.02789D) * z - 18517.159D) * z + 188975.7D
                Else
                    c = (((0.001497763D * z - 0.40585911D) * z + 40.424792D) * z - 1736.63566D) * z + 30039D
                End If
                cc = 0.0000001D * c
                az = (((-0.00018641019D * z + 0.0263199611D) * z - 0.822863477D) * z + 10.2575657D) * z
                If (z < 61D) Then
                    bias = (((-0.0001683474D * z + 0.018972278D) * z - 0.536839169D) * z + 5.654D) * z
                Else
                    bias = (((0.0031779619D * z - 0.699473097D) * z + 51.114164D) * z - 1232.4022D) * z
                End If

            ElseIf (energy >= EnDat(8)) Then

                az = (4.62D - 0.04D * z) * z
                Dim c As Double = 0.00000001D * (((-0.129086D * z + 22.09365D) * z - 783.544D) * z + 7770.8D)
                c *= (((0.000004865D * z - 0.0006561D) * z + 0.0162D) * z + 1.406D)
                cc = c * ((-0.0001285D * z + 0.01955D) * z + 0.584D)
                bias = ((0.000378D * z - 0.052D) * z + 2.51D) * EnDat(7)
                nm = 3D - 0.004D * z
                If ((energy < EnDat(5)) And (energy >= EnDat(6))) Then
                    cc = c * (0.001366D * z + 1.082D)
                End If
                If ((energy < EnDat(6)) And (energy >= EnDat(7))) Then
                    cc = 0.95D * c
                End If
                If ((energy < EnDat(7)) And (energy >= EnDat(8))) Then
                    cc = 0.8D * c * ((0.0005083D * z - 0.06D) * z + 2.0553D)
                End If

            ElseIf (energy < EnDat(8)) Then

                cc = 0.000000108D * (((-0.0669827D * z + 17.07073D) * z - 1465.3D) * z + 43156D)
                az = ((0.00539309D * z - 0.61239D) * z + 19.64D) * z
                bias = 4.5D * z - 113D
                nm = 0.3736D + 0.02401D * z
            End If
            Dim mu As Double
            If (energy > EnDat(9)) Then

                mu = cc * Math.Exp(nm * Math.Log(12397D / energy)) * z * z * z * z / At_weight
                mu *= (1D - Math.Exp((bias - energy) / az))
            Else
                mu = cc * Math.Exp(nm * Math.Log(12397D / energy)) * z * z * z * z / At_weight * (
              1D - Math.Exp((bias - EnDat(9)) / az))
                mu = 1.02D * mu * (energy - 10D) / (EnDat(9) - 10D)
            End If

            Return mu
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Heinrich_MAC30 " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

End Module
