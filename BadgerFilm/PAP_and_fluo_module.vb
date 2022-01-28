Imports System.IO

Module PAP_and_fluo_module
    'Function to hadle the case of the Ka, La, and Ma X-ray lines which are considered as the sum of two lines: Ka= Ka1+Ka2, La=La1+La2 and Ma=Ma1+Ma2
    'The function returns the total emitted X-ray intensity of the considered X-ray line (in photon/sr/electron).
    'In the case of an element present in more than one layer (or substrate), the returned intensity is the sum of all the contributions from the different layers (or substrate).
    Public Function pre_auto(ByVal layer_handler() As layer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal elt_exp_all() As Elt_exp,
                             ByVal E0 As Double, ByVal toa As Double, ByVal Ec_data() As String, ByVal options As options,
                              ByVal print_res As Boolean, ByRef save_results As String, ByVal fit_MAC As fit_MAC,
                             Optional ByVal mode As String = "PAP", Optional ByVal brem_fit() As Double = Nothing) As Double
        Try
            'Calculate the takeoff angle in radian
            Dim sin_toa_in_rad As Double = Math.Sin(toa * Math.PI / 180)
            pre_auto = 0

            'Loop over all the layers (and substrate)
            For i As Integer = 0 To UBound(layer_handler)
                'Loop over all the elements present in the current layer (or substrate)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    'If the current element corresponds to the studied element, calculate the X-ray intensity.
                    If layer_handler(i).element(j).elt_name = studied_element.elt_name Then
                        'Handle the case of the Ka X-ray line
                        If studied_element.line(line_indice).xray_name Like "[Kk][Aa]" Then
                            init_element_Xray_line_only("Ka1", Ec_data, studied_element, line_indice)
                            Dim result As Double = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "Ka"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            init_element_Xray_line_only("Ka2", Ec_data, studied_element, line_indice)
                            result = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                        options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "Ka"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            studied_element.line(line_indice).xray_name = "Ka"

                            'Handle the case of the La X-ray line
                        ElseIf studied_element.line(line_indice).xray_name Like "[Ll][Aa]" Then
                            init_element_Xray_line_only("La1", Ec_data, studied_element, line_indice)
                            Dim result As Double = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "La"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            init_element_Xray_line_only("La2", Ec_data, studied_element, line_indice)
                            result = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "La"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            studied_element.line(line_indice).xray_name = "La"

                            'Handle the case of the La X-ray line
                        ElseIf studied_element.line(line_indice).xray_name Like "[Mm][Aa]" Then
                            init_element_Xray_line_only("Ma1", Ec_data, studied_element, line_indice)
                            Dim result As Double = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "Ma"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            'layer_handler(i).element(j).xray_name = "Ka2"
                            init_element_Xray_line_only("Ma2", Ec_data, studied_element, line_indice)
                            result = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                studied_element.line(line_indice).xray_name = "Ma"
                                Return -1
                            End If
                            pre_auto = pre_auto + result

                            studied_element.line(line_indice).xray_name = "Ma"

                            'Handle the case of a single X-ray line (not the sum of two X-ray lines)
                        Else
                            init_element_Xray_line_only(studied_element.line(line_indice).xray_name, Ec_data, studied_element, line_indice)
                            Dim result As Double = auto(layer_handler, layer_handler(i).element(j).mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad,
                                                       options, print_res, save_results, fit_MAC, mode, brem_fit)
                            pre_auto = pre_auto + result

                            'In case of an error returned by the function auto.
                            If result = -1 Then
                                Return -1
                            End If
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in pre_auto " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    'Function calculating the total emitted X-ray intensity (in photon/sr/electron) 
    Public Function auto(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                         ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByVal sin_toa_in_rad As Double, ByVal options As options,
                         ByVal print_res As Boolean, ByRef save_results As String, ByVal fit_MAC As fit_MAC,
                         Optional ByVal mode As String = "PAP", Optional ByVal brem_fit() As Double = Nothing) As Double
        Try
            'Verify that the eam energy is correct
            If E0 <= 0 Then
                Return -1
            End If

            'If the ionisation threshold of the studied X-ray line is greater than the beam energy E0, return 0.
            If studied_element.line(line_indice).Ec > E0 Then
                'Debug.Print("E0 < El !!!")
                If print_res = True Then
                    save_results = save_results & E0 & vbTab & studied_element.elt_name & vbTab & studied_element.line(line_indice).xray_name & vbTab & 0 & vbTab & 0 & vbTab &
                                0 & vbTab & Format(0, "0.000") & vbTab & Format(0, "0.000") & vbCrLf
                End If
                Return 0
            End If

            'Dim watch_tot As Stopwatch = Stopwatch.StartNew()

            'Verify that the studied element has a real concentration (>0). Otherwise, the function returns 0.
            Dim flag_conc_zero As Boolean = True
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    If layer_handler(i).element(j).elt_name = studied_element.elt_name Then
                        If layer_handler(i).element(j).conc_wt <> 0 Then
                            flag_conc_zero = False
                            Exit For
                        End If
                    End If
                Next
                If flag_conc_zero = False Then Exit For
            Next
            If flag_conc_zero = True Then Return 0


            'Store the concentration of the studied element (in weight fraction).
            Dim concentration As Double = 0
            For j As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
                If layer_handler(mother_layer_id).element(j).elt_name = studied_element.elt_name Then
                    concentration = layer_handler(mother_layer_id).element(j).conc_wt
                    Exit For
                End If
            Next

            'Calculate the total emitted X-ray intensity (sum of the primary X-ray intensity, secondary fluorescence produced by the other characteristic X-rays
            'and secondary fluorescence produced by the Bremsstrahlung). The total emitted X-ray intensity is stored in res_final.
            Dim res_final As Double = 0
            Dim phi_rz As Double

            'Handles the different phi-rho-z models
            If options.phi_rz_mode = "PAP" Or options.phi_rz_mode = Nothing Then
                'Calculate the phi-rho-z distribution using the PAP model.
                Dim F As Double
                Dim phi0 As Double
                Dim R_bar As Double
                Dim P As Double
                Dim a_ As Double
                Dim b_ As Double
                Dim A_f As Double
                Dim Z_bar As Double

                'Parameters returned by the my_PAP function that describe the phi-rho-z function.
                Dim PAP_Rx As Double
                Dim PAP_Rm As Double
                Dim PAP_Rc As Double
                Dim PAP_A1 As Double
                Dim PAP_A2 As Double
                Dim PAP_B1 As Double

                Dim ierror As Integer = my_pap(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad, phi_rz, F, phi0, R_bar, P, a_, b_, A_f, Z_bar,
                       fit_MAC, options, PAP_Rx, PAP_Rm, PAP_Rc, PAP_A1, PAP_A2, PAP_B1)

                If ierror < 0 Then Return ierror

                'If studied_element.z = 12 And layer_handler(0).element.Count > 1 And E0 = 15 Then ' mother_layer_id <> 0 And
                '    Dim tmp As String = "Rx" & vbTab & PAP_Rx & vbCrLf
                '    tmp = tmp & "Rm" & vbTab & PAP_Rm & vbCrLf
                '    tmp = tmp & "Rc" & vbTab & PAP_Rc & vbCrLf
                '    tmp = tmp & "A1" & vbTab & PAP_A1 & vbCrLf
                '    tmp = tmp & "A2" & vbTab & PAP_A2 & vbCrLf
                '    tmp = tmp & "B1" & vbTab & PAP_B1 & vbCrLf
                '    tmp = tmp & "E0" & vbTab & E0 & vbCrLf
                '    Debug.Print(tmp)
                '    'My.Computer.Clipboard.SetText(tmp)
                'End If
                'Dim results As String = PAP_Rx & vbCrLf & PAP_Rm & vbCrLf & PAP_Rc & vbCrLf & PAP_A1 & vbCrLf & PAP_A2 & vbCrLf & PAP_B1 & vbCrLf
                'My.Computer.Clipboard.SetText(results)

            ElseIf options.phi_rz_mode = "PROZA96" Then
                'Tentative to implement the PROZA96 model. Do not seem to work for thin films!!!
                Dim F As Double
                Dim phi0 As Double
                Dim rzm As Double
                Dim alpha As Double
                Dim beta As Double
                Dim Rx As Double

                Dim ierror As Integer = PROZA96(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad, phi_rz, F, phi0, rzm, alpha, beta, Rx, options, fit_MAC)
                If ierror < 0 Then Return ierror

            ElseIf options.phi_rz_mode = "XPHI" Then
                'Tentative to implement the XPHI model. Need to be tested!!!
                Dim rzm As Double
                Dim rzx1 As Double
                Dim phi_rzm As Double
                Dim phi0 As Double
                Dim alpha_val As Double
                Dim beta_val As Double


                Dim ierror As Integer = calc_multi_layer(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad, phi_rz, rzm, rzx1, phi_rzm, phi0, alpha_val, beta_val, options, fit_MAC)
                If ierror < 0 Then Return ierror
                'Dim results As String = phi_rzm & vbCrLf & rzm & vbCrLf & rzx1 & vbCrLf & alpha_val & vbCrLf & beta_val
                'My.Computer.Clipboard.SetText(results)

            ElseIf options.phi_rz_mode = "XPP" Then

                'Tentative to implement the XPP model.
                Dim F As Single
                Dim phi0 As Single
                Dim R_bar As Single
                Dim A_XPP As Single
                Dim B_XPP As Single
                Dim P As Single

                'If layer_handler(mother_layer_id).element.Count > 2 And E0 = 5 Then Stop

                Dim ierror As Integer = my_xpp(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad, phi_rz, F, phi0, R_bar, P, A_XPP, B_XPP, fit_MAC, options)
                If ierror < 0 Then Return ierror

                Dim m_ As Single
                If (studied_element.line(line_indice).xray_name(0) = "K") Then
                    m_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
                ElseIf (studied_element.line(line_indice).xray_name(0) = "L") Then
                    m_ = 0.82
                ElseIf (studied_element.line(line_indice).xray_name(0) = "M") Then
                    m_ = 0.78
                End If

                Dim QA_l As Single
                Dim El As Single = studied_element.line(line_indice).Ec
                Dim U0 As Single = E0 / El
                QA_l = Math.Log(U0) / El ^ 2 / U0 ^ m_

                phi_rz = phi_rz * QA_l

            End If
            '******************************************************

            'Primary intensity constants.
            Dim pap_res As Double
            If options.phi_rz_mode <> "XPP" Then

                Dim shell1 As Integer
                Dim shell2 As Integer
                'From an X-ray transition name, recover the electron shells involved in the transition.
                Siegbahn_to_transition_num(studied_element.line(line_indice).xray_name, shell1, shell2)

                'For test:
                'Dim fc7 As Double = 0
                'Dim tmp() As Double = interpol_log_log(studied_element.el_ion_xs, E0) 'XX AM 02-10-17
                'fc7 = tmp(shell1 - 1)
                'fc7 = fc7 * 6.022 * 10 ^ 23 / (4 * Math.PI)

                Dim fc6 As Double
                Dim norm_xs As Double = 1
                'Calculate the X-ray production cross section from the ionisation cross section of Bote and Salvat.
                If options.ionizationXS_mode = "Bote" Then
                    fc6 = Xray_production_xs_el_impact(studied_element, shell1, E0)
                    fc6 = fc6 * 6.022 * 10 ^ 23 / (4 * Math.PI)

                    'Calculate the X-ray production cross section from the ionisation cross section used by Pouchou and Pichoir when developping their model (green book).
                Else
                    'the electron impact ionization cross sections are normalized to the value of Bote and Salvat at 15 kV to return an absolute X-ray intensity in photon/sr/electron.
                    norm_xs = Xray_production_xs_el_impact(studied_element, shell1, 15) * 6.022 * 10 ^ 23 / (4 * Math.PI) /
                                qe0(studied_element.line(line_indice).Ec, 15, studied_element.line(line_indice).xray_name, studied_element.z)
                    fc6 = qe0(studied_element.line(line_indice).Ec, E0, studied_element.line(line_indice).xray_name, studied_element.z)
                End If

                'Calculate all the constants usued to calculate the total emitted X-ray intensity (except the concentration of the studied element).
                Dim const2 As Double = constante_simple(studied_element, shell1, shell2) * fc6 * norm_xs

                'Calculate the total emitted PRIMARY X-ray intensity
                pap_res = const2 * concentration * phi_rz

            Else
                pap_res = phi_rz
            End If


            '******************************************************
            'Fluorescence calculation.
            Dim res_fluo_caract_PAP As Double = 0
            Dim res_fluo_Brem_PAP As Double = 0

            'Calculate the secondary fluorescence produced by characterisitic X-rays.
            If options.char_fluo_flag = True Then
                res_fluo_caract_PAP = test_fluo_charact(layer_handler, mother_layer_id, studied_element, line_indice, concentration, elt_exp_all, E0, sin_toa_in_rad, fit_MAC, options)
                'Debug.Print("Fluo char: " & watch.Elapsed.TotalSeconds)
            End If

            'Calculate the secondary fluorescence produced by the Bremsstrahlung.
            If options.brem_fluo_flag = True Then
                res_fluo_Brem_PAP = test_fluo_brem_v3(layer_handler, mother_layer_id, concentration, studied_element, line_indice, elt_exp_all, E0, sin_toa_in_rad, fit_MAC, options, brem_fit)
                'Debug.Print("Fluo brem: " & watch.Elapsed.TotalSeconds)
            End If

            'If print_res = True Then
            'Debug.Print(E0 & vbTab & studied_element.elt_name & vbTab & studied_element.line(line_indice).xray_name & vbTab & pap_res & vbTab & res_fluo_caract_PAP & vbTab &
            'res_fluo_Brem_PAP & vbTab & Format(res_fluo_caract_PAP / pap_res * 100, "0.000") & vbTab & Format(res_fluo_Brem_PAP / pap_res * 100, "0.000"))
            'End If

            'Check if the SF values are realistic. Otherwise, set them to 0.
            If res_fluo_caract_PAP < 0 Then res_fluo_caract_PAP = 0
            If res_fluo_Brem_PAP < 0 Then res_fluo_Brem_PAP = 0

            'Save the values. Needed to export the data.
            If print_res = True Then
                save_results = save_results & E0 & vbTab & studied_element.elt_name & vbTab & studied_element.line(line_indice).xray_name & vbTab & pap_res & vbTab & res_fluo_caract_PAP & vbTab &
                                res_fluo_Brem_PAP & vbTab & Format(res_fluo_caract_PAP / pap_res * 100, "0.000") & vbTab & Format(res_fluo_Brem_PAP / pap_res * 100, "0.000") & vbCrLf
            End If

            'This should never happen.
            If pap_res < 0 Or res_fluo_caract_PAP < 0 Or res_fluo_Brem_PAP < 0 Then
                Return -1
            End If

            'Calculate the totam emitted X-ray intensity, sum of the emitted primary characteristic X-ray intensity and SF contributions.
            res_final = pap_res + res_fluo_caract_PAP + res_fluo_Brem_PAP


            'ElseIf options.phi_rz_mode = "PENEPMA" Then
            'Dim indice As Integer = 0
            'For i As Integer = 0 To UBound(kvs)
            '    If E0 = kvs(i) Then
            '        indice = i
            '        Exit For
            '    End If
            'Next
            'Dim mac As Double = MAC_calculation(studied_element.xray_energy, studied_element.mother_layer_id, layer_handler, fit_param)
            'res_final = integrate_phirz(phirz(indice), rz, mac, toa) ' / (4 * Math.PI) * 0.838994


            'Transfert the content of res_final to the variable returned by the function.
            auto = res_final

            'Check that the value returned is not NaN or Infinity.
            If Double.IsNaN(auto) Or Double.IsInfinity(auto) Then
                auto = 0
            End If

            Return auto

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in auto " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function my_pap(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                      ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByVal sin_toa_in_rad As Double, ByRef phi_rz As Double, ByRef F As Double, ByRef phi0 As Double,
                      ByRef R_bar As Double, ByRef P As Double, ByRef a_ As Double, ByRef b_ As Double, ByRef A_f As Double, ByRef Z_bar As Double, fit_MAC As fit_MAC,
                      ByVal options As options,
                      Optional ByRef PAP_Rx As Double = 0, Optional ByRef PAP_Rm As Double = 0, Optional ByRef PAP_Rc As Double = 0,
                      Optional ByRef PAP_A1 As Double = 0, Optional ByRef PAP_A2 As Double = 0, Optional ByRef PAP_B1 As Double = 0) As Integer
        Try
            Dim El As Double = studied_element.line(line_indice).Ec
            'El = 0.7074
            Dim U0 As Double = E0 / El

            '************************
            'Normalize the concentrations as if the whole sample was a homogeneous bulk sample.
            'Fictitious composition stored in the layer_handler.
            '************************
            Dim total_concentration As Double = 0
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    'layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).conc_wt ' line removed AMXX 22/11/2019
                    total_concentration = total_concentration + layer_handler(i).element(k).conc_wt
                Next

            Next

            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).conc_wt / total_concentration
                Next
            Next
            '************************

            Dim Rx As Double
            Dim Rx_old As Double
            Const RX_CONVERGENCE As Double = 0.01
            Dim flag_first_iter = True

            Dim M As Double = 0
            Dim Zn_bar As Double = 0
            Dim J As Double = 0
            Dim Pi(2) As Double
            Dim Di(2) As Double

            Z_bar = 0

            While (Math.Abs(Rx - Rx_old) / Rx > RX_CONVERGENCE Or flag_first_iter = True)
                'If flag_first_iter = True And layer_handler.Count > 1 Then
                '    flag_first_iter = False
                '    Rx = 7 * (E0 ^ 1.7 * El ^ 1.7) * (1 + 3 / El ^ 0.5 / (E0 / El + 0.3) ^ 2) * 10 ^ -6
                '    'AM XX 06-09-2018 (Septembre)
                '    Call pap_weight(layer_handler, 0.8 * Rx, -0.4 * 0.8 * Rx)
                '    Continue While
                'End If

                flag_first_iter = False
                Rx_old = Rx
                '************************
                'Calculate M
                'Calculate Zn bar
                'Calculate Z bar
                'Calculate J
                '************************
                M = 0
                Zn_bar = 0
                Z_bar = 0
                J = 0
                For i As Integer = 0 To UBound(layer_handler)
                    For k As Integer = 0 To UBound(layer_handler(i).element)
                        With layer_handler(i).element(k)
                            M = M + .z / .a * .fictitious_concentration
                            Zn_bar = Zn_bar + Math.Log(.z) * .fictitious_concentration
                            Z_bar = Z_bar + .z * .fictitious_concentration
                            J = J + .fictitious_concentration * .z / .a * Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
                        End With
                    Next
                Next
                Zn_bar = Math.Exp(Zn_bar)
                J = Math.Exp(J / M)
                '************************


                '************************
                'Calculate Pi
                'Calculate Di
                '************************
                Pi(0) = 0.78
                Pi(1) = 0.1
                Pi(2) = -(0.5 - 0.25 * J)
                Di(0) = 6.6 * 10 ^ -6
                Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
                Di(2) = 2.2 * 10 ^ -6 / J
                '************************

                '************************
                'Calculate R0 from p.61 (green book)
                '************************
                Dim R0 As Double = 0
                For k As Integer = 0 To 2
                    R0 = R0 + J ^ (1 - Pi(k)) * Di(k) * (E0 ^ (1 + Pi(k)) - El ^ (1 + Pi(k))) * 1 / (1 + Pi(k))
                Next
                R0 = R0 / M
                '************************

                '************************
                'Calculate Q0
                'Calculate b
                'Calculate Q
                'Calculate h
                'Calculate D
                'Calculate Rx
                '************************
                Dim Q0 As Double
                Dim b As Double
                Dim Q As Double
                Dim h As Double
                Dim D As Double
                Q0 = 1 - 0.535 * Math.Exp(-(21 / Zn_bar) ^ 1.2) - 0.00025 * (Zn_bar / 20) ^ 3.5
                b = 40 / Z_bar
                Q = Q0 + (1 - Q0) * Math.Exp(-(U0 - 1) / b)
                h = Z_bar ^ 0.45
                D = 1 + 1 / (U0 ^ h)
                Rx = Q * D * R0
                '************************

                '************************
                'Iterate to calculate fictitious concentrations
                'in order to determine Rx (only in case of multilayer specimen?).
                '************************
                'If (Math.Abs(Rx * 1000000.0 - rt * 1000000.0) > 0.02) Then 'not the same condition as described by Pouchou and Pichoir p.54
                'Stop
                If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
                    'Call pap_weight(layer_handler, 0.8 * Rx, -0.4 * 0.8 * Rx)
                    Call pap_weight(layer_handler, Rx, -0.4 * Rx)
                Else
                    Exit While
                End If

                'rt = Rx
                ' flag_first_iter = False
                'End If
                '************************
            End While



            '************************
            'Calculate other fictitious concentrations
            'in order to determine Zb bar (only in case of multilayer specimen).
            '************************
            If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
                'Call pap_weight(layer_handler, 0.5 * Rx, -0.1 * (0.5 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                Call pap_weight(layer_handler, 0.5 * Rx, -0.4 * (0.5 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            End If

            Dim Zb_bar As Double = 0
200:
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    Zb_bar = Zb_bar + layer_handler(i).element(k).z ^ 0.5 * layer_handler(i).element(k).fictitious_concentration
                Next
            Next
            Zb_bar = Zb_bar ^ 2
            '************************

            '************************
            'Calculate parameters
            'Eta bar
            'W bar
            'q
            'J(U0)
            'G(U0)
            'R
            'r
            'phi(0)
            '************************
            Dim eta_bar As Double
            Dim W_bar As Double
            Dim q_ As Double
            Dim J_U0 As Double
            Dim G_U0 As Double
            Dim R As Double
            Dim r_ As Double
            'Dim phi0 As Double

            eta_bar = 0.00175 * Zb_bar + 0.37 * (1 - Math.Exp(-0.015 * (Zb_bar ^ 1.3)))
            W_bar = 0.595 + eta_bar / 3.7 + eta_bar ^ 4.55
            q_ = (2 * W_bar - 1) / (1 - W_bar)
            J_U0 = 1 + U0 * (Math.Log(U0) - 1)
            G_U0 = (U0 - 1 - ((1 - (1 / U0 ^ (q_ + 1))) / (1 + q_))) / ((2 + q_) * J_U0)
            R = 1 - eta_bar * W_bar * (1 - G_U0)
            'Debug.Print(Zb_bar & " " & R & " " & eta_bar)
            r_ = 2 - 2.3 * eta_bar

            phi0 = 1 + 3.3 * (1 - (1 / U0 ^ r_)) * eta_bar ^ 1.2 '* ((1 - 0.6) / (1 - 100) * Zb_bar + (1 * 100 - 0.6) / (100 - 1)) 'AMXXX
            '************************

            '************************
            'Calculate fictitious concentrations
            'in order to determine fictitious Z bar (only in case of multilayer specimen?).
            '************************
            If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
                'Call pap_weight(layer_handler, 0.65 * Rx, -0.6 * (0.65 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!0.7
                Call pap_weight(layer_handler, 0.7 * Rx, -0.6 * (0.7 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!0.7
            End If
            '************************

            '************************
            'With the fictitious concentrations
            'Calculate M
            'Calculate Z bar
            'Calculate J
            '************************
            Z_bar = 0
            J = 0
            M = 0
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    With layer_handler(i).element(k)
                        Z_bar = Z_bar + .z * .fictitious_concentration
                        M = M + .z / .a * .fictitious_concentration
                        J = J + .fictitious_concentration * .z / .a * Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
                    End With
                Next
            Next
            J = Math.Exp(J / M)
            '************************

            '************************
            'Calculate m_
            '************************
            Dim m_ As Double
            If (studied_element.line(line_indice).xray_name(0) = "L") Then
                m_ = 0.82 'corr AMXX
            ElseIf (studied_element.line(line_indice).xray_name(0) = "M") Then
                m_ = 0.78 'corr AMXX
            ElseIf (studied_element.line(line_indice).xray_name(0) = "K") Then
                m_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
            Else
                'Do not handle other X-ray lines.
                phi_rz = 0
                PAP_Rx = 0
                PAP_Rm = 0
                PAP_Rc = 0
                PAP_A1 = 0
                PAP_A2 = 0
                PAP_B1 = 0
                Return -1
            End If
            '************************

            '************************
            'Because now J can be different
            'Recalculate P3
            'Recalculate D2
            'Recalculate D3
            '************************
            Pi(2) = -(0.5 - 0.25 * J)
            Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
            Di(2) = 2.2 * 10 ^ -6 / J
            '************************

            '************************
            'Calculate V0 
            'Calculate QA_l
            'Calculate 1/S
            'Calculate F = R/S * 1/QA_l
            '************************
            Dim V0 As Double
            Dim QA_l As Double
            Dim inv_S As Double = 0
            'Dim F As Double
            V0 = E0 / J

            '**************************************************
            'Test: tried to replace the ionization cross section used by Pouchou and Pichoir by the model developped by Bote and Salvat.
            'Dim indice As Integer = 0
            'For i As Integer = 0 To UBound(elt_exp_all)
            '    If elt_exp_all(i).z = studied_element.z Then
            '        studied_element.el_ion_xs = elt_exp_all(i).el_ion_xs
            '        Exit For
            '    End If
            'Next
            'Dim shell1 As Integer
            'Dim shell2 As Integer
            'Siegbahn_to_transition_num(studied_element.line(line_indice).xray_name, shell1, shell2)
            'Dim fc6 As Double = 0
            'Dim norm_xs As Double = 1


            'If My.Forms.Form1.CheckBox3.Checked = True Then
            'fc6 = Xray_production_xs_el_impact(studied_element, shell1, E0)
            'fc6 = fc6 * 6.022 * 10 ^ 23 / (4 * Math.PI)
            'Else
            'norm_xs = Xray_production_xs_el_impact(studied_element, shell1, 15) * 6.022 * 10 ^ 23 / (4 * Math.PI) /
            '            qe0(studied_element.line(line_indice).Ec, 15, studied_element.line(line_indice).xray_name, studied_element.z)
            '    fc6 = qe0(studied_element.line(line_indice).Ec, E0, studied_element.line(line_indice).xray_name, studied_element.z)
            'End If

            'QA_l = fc6 * norm_xs
            '**************************************************

            QA_l = Math.Log(U0) / El ^ 2 / U0 ^ m_

            For k As Integer = 0 To 2
                Dim Tk As Double = 1 + Pi(k) - m_
                inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
            Next
            inv_S = inv_S * U0 / (V0 * M)
            'Debug.Print("Elt " & studied_element.elt_name)
            'Debug.Print("E0 " & E0)
            'Debug.Print("J " & J)
            'Debug.Print("inv_S " & inv_S)
            F = R * inv_S / QA_l 'Is it * QA_l or / QA_l? From equation 13 p.37 it is "*" but from equations 2 and 3 it seems that it is "/".
            '************************

            '************************
            'Calculate fictitious mean atomic number (z)
            'in order to determine Rm (dRm) -> in text p.55 (only in case of multilayer specimen?).
            '************************
            If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
                Z_bar = (Z_bar + Zb_bar) / 2
            End If
            '************************

            '************************
            'Calculate G1
            'Calculate G2
            'Calculate G3
            'Calculate Rm
            '************************
            Dim G1 As Double
            Dim G2 As Double
            Dim G3 As Double
            Dim Rm As Double
            G1 = 0.11 + 0.41 * Math.Exp(-(Z_bar / 12.75) ^ 0.75)
            G2 = 1 - Math.Exp(-(U0 - 1.0) ^ 0.35 / 1.19) 'AM XX 29-01-2018
            G3 = 1 - Math.Exp(-(U0 - 0.5) * Z_bar ^ 0.4 / 4)

            Rm = G1 * G2 * G3 * Rx '* ((1 - 1.5) / (1 - 100) * Zb_bar + (1 * 100 - 1.5) / (100 - 1)) 'AMXXX   'AMXXX
            '************************

            '************************
            'Calculate d
            'Check the consistency of d
            '************************
            Dim d_ As Double
            d_ = (Rx - Rm) * (F - (phi0 * Rx / 3)) * ((Rx - Rm) * F - phi0 * Rx * (Rm + Rx / 3))
            If (d_ < 0.0) Then
                Dim Rm_temp As Double = Rm
                Rm = Rx * (F - phi0 * Rx / 3) / (F + phi0 * Rx)
                d_ = 0
            End If
            '************************

            '************************
            'Calculate the 'wt.fraction averaged mass absorption coefficient'
            'for the current layer
            '************************
            Dim mac As Double
            Dim chi As Double
            'mac = MAC_calculation(studied_element.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
            mac = MAC_calculation(studied_element, line_indice, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
            'Debug.Print(mac)
            chi = mac / sin_toa_in_rad
            '************************

            '************************
            'Calculation of phi_rz
            'Handle patologic cases
            '************************
            Dim Rc As Double
            Dim A1 As Double
            Dim A2 As Double
            Dim B1 As Double

            If Rm < 0 Or Rm > Rx Then
                Dim s_min As Double = 4.5
                If F < phi0 / s_min Then
                    phi_rz = phi0 * F / (phi0 + chi * F)
                Else
                    Dim phi0_p As Double
                    phi0_p = 3 * (F * s_min - phi0) / (Rx * s_min - 3)
                    phi_rz = ((2 / chi ^ 3) * (1 - Math.Exp(-chi * Rx)) - 2 * Rx / chi ^ 2 + Rx ^ 2 / chi) * (phi0_p / Rx ^ 2) + (phi0 - phi0_p) / (s_min + chi)
                End If
            Else
                '************************
                'Calculate Rc
                'Calculate A1
                'Calculate A2
                'Calculate B1
                '************************
                Rc = 1.5 * ((F - phi0 * Rx / 3) / phi0 - Math.Sqrt(d_) / (phi0 * (Rx - Rm)))
                If Rc < 0 Then
                    Rc = 3 * Rm * (F + phi0 * Rx) / (2 * phi0 * Rx) '????
                End If
                A1 = phi0 / (Rm * (Rc - Rx * (Rc / Rm - 1)))
                A2 = A1 * (Rc - Rm) / (Rc - Rx)
                B1 = phi0 - A1 * Rm ^ 2
                '************************

                'If Rc < 0 Then
                '    Rc = 0  '???
                'End If

                '************************
                'Integrate the phi(rz) function
                'Calculate the cumulative mass depth for each layer
                '************************
                Dim cumulative_mass_depth() As Double = Nothing
                Dim temp As Double = 0
                For i As Integer = 0 To UBound(layer_handler)
                    If cumulative_mass_depth Is Nothing Then
                        ReDim cumulative_mass_depth(0)
                    Else
                        ReDim Preserve cumulative_mass_depth(UBound(cumulative_mass_depth) + 1)
                    End If
                    temp = temp + layer_handler(i).mass_thickness
                    cumulative_mass_depth(i) = temp
                Next

                Dim lim_min As Double
                Dim lim_max As Double
                If mother_layer_id = 0 Then
                    lim_min = 0
                Else
                    lim_min = cumulative_mass_depth(mother_layer_id - 1)
                End If
                lim_max = cumulative_mass_depth(mother_layer_id)
                If lim_max > Rx Then lim_max = Rx

                Dim H1 As Double
                Dim H2 As Double


                If lim_min >= Rc Then
                    H1 = 0
                Else
                    If chi = 0 Then
                        H1 = A1 * (ppint2(A1, B1, Rm, chi, Math.Min(lim_max, Rc)) - ppint2(A1, B1, Rm, chi, lim_min))
                    Else
                        H1 = -A1 / chi * (ppint2(A1, B1, Rm, chi, Math.Min(lim_max, Rc)) - ppint2(A1, B1, Rm, chi, lim_min))
                    End If
                End If

                If lim_max <= Rc Then
                    H2 = 0
                Else
                    If chi = 0 Then
                        H2 = A2 * (ppint2(A2, 0, Rx, chi, Math.Min(lim_max, Rx)) - ppint2(A2, 0, Rx, chi, Math.Max(Rc, lim_min)))
                    Else
                        H2 = -A2 / chi * (ppint2(A2, 0, Rx, chi, Math.Min(lim_max, Rx)) - ppint2(A2, 0, Rx, chi, Math.Max(Rc, lim_min)))
                    End If
                End If

                Dim H1_plus_H2_test As Double = H1 + H2
                If H1_plus_H2_test <= 0 Then
                    H1_plus_H2_test = 0.00000000001
                End If
                phi_rz = H1_plus_H2_test

            End If

            Dim abs As Double = abs_outer_layers(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, mac, sin_toa_in_rad, fit_MAC, options)
            phi_rz = phi_rz * abs
            '************************

            PAP_Rx = Rx
            PAP_Rm = Rm
            PAP_Rc = Rc
            PAP_A1 = A1
            PAP_A2 = A2
            PAP_B1 = B1

            Return 0


            ''************************
            ''Calculation for the xpp model
            ''Calculate R_bar
            ''Calculate g
            ''Calculate h_xpp
            ''Calculate P
            ''Calculate a_
            ''Calculate b_
            ''Calculate epsilon
            ''Calculate A_
            ''Calculate B_
            ''************************
            ''Dim R_bar As Double
            'Dim g As Double
            'Dim h_xpp As Double
            ''Dim P As Double
            ''Dim a_ As Double
            ''Dim b_ As Double
            'Dim epsilon As Double
            ''Dim A_f As Double
            'Dim B_f As Double

            'Dim X As Double = 1 + 1.3 * Math.Log(Zb_bar)
            'Dim Y As Double = 0.2 + Zb_bar / 200

            'R_bar = F / (1 + (X * Math.Log(1 + Y * (1 - 1 / U0 ^ 0.42))) / Math.Log(1 + Y))

            'If F / R_bar < phi0 Then R_bar = F / phi0

            'g = 0.22 * Math.Log(4 * Zb_bar) * (1 - 2 * Math.Exp(-Zb_bar * (U0 - 1) / 15))

            'h_xpp = 1 - 10 * (1 - 1 / (1 + U0 / 10)) / Zb_bar ^ 2

            'b_ = 2 ^ 0.5 * ((1 + (1 - R_bar * phi0 / F)) ^ 0.5) / R_bar 'AM XX 11-04-18



            'Dim lim1 As Double = g * h_xpp ^ 4
            'Dim lim2 As Double = 0.9 * b_ * R_bar ^ 2 * (b_ - 2 * phi0 / F)
            'If lim1 > lim2 Then
            '    P = lim2 * F / R_bar ^ 2
            'Else
            '    P = lim1 * F / R_bar ^ 2
            'End If


            'a_ = (P + b_ * (2 * phi0 - b_ * F)) / (b_ * F * (2 - b_ * R_bar) - phi0)

            'epsilon = Math.Max((a_ - b_) / b_, 10 ^ -6)

            'a_ = b_ * (1 + epsilon)

            'B_f = (b_ ^ 2 * F * (1 + epsilon) - P - phi0 * b_ * (2 + epsilon)) / epsilon

            'A_f = (B_f / b_ + phi0 - b_ * F) * (1 + epsilon) / epsilon

            ''Dim test As Double
            ''test = (phi0 + B_f / (b_ + chi) - A_f * b_ * epsilon / (b_ * (1 + epsilon) + chi)) / (b_ + chi)


            'Dim a1_test_ As Double
            ''Dim a2_test_ As Double
            'Dim A1_test As Double
            ''Dim A2_test As Double

            ''Second try of calculation of caracteristic fluorescence (p.71). 'AM XX 12-04-18
            'A1_test = (10 * phi0 - 13 * b_ * F) ^ 2 / (100 * phi0 + 13 * b_ * F * (13 * b_ * R_bar - 20))
            'a1_test_ = (10 * phi0 - 13 * b_ * F) / (10 * F - 13 * b_ * F * R_bar)
            ''A2_test = phi0 - A1_test
            ''a2_test_ = 1.3 * b_

            'a_ = a1_test_
            'A_f = A1_test
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in my_pap " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Sub pap_weight(ByRef layer_handler() As layer, ByVal R As Single, ByVal L As Single)
        Try
            '************************
            'Calculate N in eq. (39) p. 54
            '************************
            Dim N As Double
            N = 1.0 / (primitive_pap_weight(R, R, L) - primitive_pap_weight(0.0, R, L))
            '************************

            '************************
            'Calculate the cumulative mass depth for each layer
            '************************
            Dim cumulative_mass_depth() As Double = Nothing
            Dim temp As Double = 0
            For i As Integer = 0 To UBound(layer_handler)
                If cumulative_mass_depth Is Nothing Then
                    ReDim cumulative_mass_depth(0)
                Else
                    ReDim Preserve cumulative_mass_depth(UBound(cumulative_mass_depth) + 1)
                End If
                temp = temp + layer_handler(i).mass_thickness
                cumulative_mass_depth(i) = temp
            Next
            '************************

            '************************
            'Calculate the weighting factors for each element (each layer and the substrate)
            'and then weight the concentrations
            '************************
            Dim min As Double
            Dim max As Double
            For i As Integer = 0 To UBound(layer_handler)
                If i = 0 Then 'First layer
                    min = 0
                    max = Math.Min(layer_handler(i).mass_thickness, R)
                ElseIf i = UBound(layer_handler) Then 'Last 'layer' (substrate)
                    min = Math.Min(cumulative_mass_depth(i - 1), R)
                    max = R
                Else
                    min = Math.Min(cumulative_mass_depth(i - 1), R)
                    max = Math.Min(cumulative_mass_depth(i), R)
                End If
                Dim integral As Double = (primitive_pap_weight(max, R, L) - primitive_pap_weight(min, R, L))
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).conc_wt * N * integral
                Next
            Next

            '************************
            'Normalize the fictitious concentrations
            '************************
            Dim tempsum As Double = 0
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    tempsum = tempsum + layer_handler(i).element(k).fictitious_concentration
                Next
            Next

            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).fictitious_concentration / tempsum
                Next
            Next
            '************************
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in pap_weight " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Sub

    Public Function primitive_pap_weight(ByVal x As Double, ByVal R As Double, ByVal L As Double)
        Try
            'primitive_pap_weight = x * (x ^ 4 / 5.0 - x ^ 3 / 2.0 * (R + L) + x ^ 2 / 3.0 * (R ^ 2 + 4.0 * R * L + L ^ 2) - x * (R ^ 2 * L + L ^ 2 * R) + (R ^ 2 * L ^ 2))
            primitive_pap_weight = 1 / 3 * x ^ 3 * (L ^ 2 + 4 * L * R + R ^ 2) + L ^ 2 * R ^ 2 * x - 0.5 * x ^ 4 * (L + R) - L * R * x ^ 2 * (L + R) + x ^ 5 / 5
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in primitive_pap_weight " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function abs_outer_layers(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                                     ByVal elt_exp_all() As Elt_exp, ByVal mac As Double, ByVal sin_toa_in_rad As Double, ByVal fit_MAC As fit_MAC,
                                     ByVal options As options) As Double
        Try
            ' bottom of p. 49.
            abs_outer_layers = 1
            If mother_layer_id = 0 Then Exit Function

            Dim fact As Double = 0

            For i As Integer = 0 To mother_layer_id - 1
                Dim mac_tmp As Double = 0
                mac_tmp = mac_tmp + MAC_calculation(studied_element, line_indice, i, layer_handler, elt_exp_all, fit_MAC, options)
                fact = fact + (mac_tmp - mac) * layer_handler(i).mass_thickness
                'fact = fact + (mac_tmp) * layer_handler(i).mass_thickness 'AMXXX
            Next

            abs_outer_layers = Math.Exp(-fact / sin_toa_in_rad)

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in abs_outer_layers " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function ppint2(ByVal a As Double, ByVal b As Double, ByVal r As Double, ByVal c As Single, ByVal x As Single) As Double
        Try
            If (c = 0.0) Then
                ppint2 = x * (x * x / 3.0 - r * x + r * r + b / a)
                Return ppint2
            End If

            If (c * x > 88.0) Then
                ppint2 = 0.0
                Return ppint2
            End If
            ppint2 = Math.Exp(-c * x) * ((x - r) ^ 2 + 2.0 * (x - r) / c + 2.0 / c / c + b / a)
            Return ppint2

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in ppint2 " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function qe0(ByVal El As Single, ByVal E0 As Single, ByVal line As String, ByVal z As Single) As Single
        Try
            ' Ionization cross section for electron impact as described by Pouchou and Pichoir in Electron Probe Quantitation p.73
            Dim q As Single 'AMXX
            Dim m As Single 'AMXX
            Dim U0 As Single 'AMXX

            U0 = E0 / El

            If (line(0) = "L") Then m = 0.82
            If (line(0) = "M") Then m = 0.78
            If (line(0) = "K") Then
                If (z > 30) Then
                    m = 0.86
                Else
                    m = 0.86 + 0.12 * Math.Exp(-1.0 * (z / 5.0) ^ 2.0)
                End If
            End If

            q = Math.Log(U0) / (El ^ 2 * U0 ^ m)

            If (line(0) = "K") Then qe0 = 1.0E-20 * 3.8 * m * q
            If (line(0) = "L") Then qe0 = 1.0E-20 * 5.7 * m * q
            If (line(0) = "M") Then qe0 = 39229.0 * m * q

            Return qe0

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in qe0 " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

End Module
