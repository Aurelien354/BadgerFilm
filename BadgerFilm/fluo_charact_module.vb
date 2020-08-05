Module fluo_charact_module

    Public Function test_fluo_charact(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_elt As Elt_exp, ByVal line_indice As Integer,
                                      ByVal studied_elt_conc As Double, ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByVal sin_toa_in_rad As Double,
                                      ByVal fit_MAC As fit_MAC, ByVal options As options) As Double

        test_fluo_charact = 0
        'Dim toa_in_rad As Double = toa * Math.PI / 180

        'Calculate TT and x ***********************
        Dim TT As Double = 0
        If mother_layer_id <> 0 Then
            For i As Integer = 0 To mother_layer_id - 1
                TT = TT + MAC_calculation(studied_elt.line(line_indice).xray_energy, i, layer_handler, elt_exp_all, fit_MAC, options) / sin_toa_in_rad * layer_handler(i).mass_thickness
            Next
            TT = Math.Exp(-TT)
        Else
            TT = 1
        End If
        '******************************************

        'Calculation of Fa ************************
        'Dim A1 As Double
        'Dim a1_ As Double
        'Dim phi_rz As Double
        'Dim Fa As Double
        'Dim phi0 As Double
        'Dim R_bar As Double
        'Dim P As Double
        'Dim b_ As Double
        'Dim Z_bar As Double
        'my_pap(layer_handler, studied_elt, E0, phi_rz, Fa, phi0, R_bar, P, a1_, b_, A1, Z_bar, fit_param)
        '******************************************

        Dim chi_a As Double = MAC_calculation(studied_elt.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options) / sin_toa_in_rad 'Math.Sin(toa_in_rad)

        Dim fluo_char As Double = 0

        For i As Integer = 0 To UBound(layer_handler)
            For j As Integer = 0 To UBound(layer_handler(i).element)
                '*************************************
                ' Pourquoi je fais ca ???
                ' Pour trouver ou l'element en question a ete initialise
                Dim indice_elt_in_all As Integer
                For k As Integer = 0 To UBound(elt_exp_all)
                    If elt_exp_all(k).elt_name = layer_handler(i).element(j).elt_name Then
                        indice_elt_in_all = k
                        Exit For
                    End If
                Next
                '*************************************

                Dim Energy_Xray_lines_of_B() As Double = Nothing
                Dim Name_Xray_lines_of_B() As String = Nothing
                Dim Ec_of_each_Xray_line_of_B() As Double = Nothing
                extract_main_xray_lines(elt_exp_all(indice_elt_in_all), Energy_Xray_lines_of_B, Name_Xray_lines_of_B, Ec_of_each_Xray_line_of_B)

                For k As Integer = 0 To UBound(Energy_Xray_lines_of_B)
                    If Energy_Xray_lines_of_B(k) < studied_elt.line(line_indice).Ec Then 'Not enought energy
                        Continue For
                    End If

                    If Ec_of_each_Xray_line_of_B(k) > E0 Then 'Not enought energy to create the x-ray line of B
                        Continue For
                    End If


                    Dim Fb As Double
                    Dim B_Z_bar As Double
                    Dim B_phi_rz As Double
                    Dim B_a1_ As Double
                    Dim B_phi0 As Double
                    Dim B_R_bar As Double
                    Dim B_P As Double
                    Dim XPP_B_b_ As Double
                    Dim XPP_B_A1 As Double
                    Dim B_Rx As Double
                    Dim B_Rm As Double
                    Dim B_Rc As Double
                    Dim B_A1 As Double
                    Dim B_A2 As Double
                    Dim B_B1 As Double


                    Dim tmp_element As Elt_exp = elt_exp_all(indice_elt_in_all)
                    ReDim tmp_element.line(0)
                    'tmp_element = layer_handler(i).element(j)
                    tmp_element.line(0).xray_energy = Energy_Xray_lines_of_B(k)
                    tmp_element.line(0).xray_name = Name_Xray_lines_of_B(k)
                    tmp_element.line(0).Ec = Ec_of_each_Xray_line_of_B(k)
                    'tmp_element.symbol = ""

                    Dim ffact As Double = 0

                    If options.phi_rz_mode = "PAP" Or options.phi_rz_mode = Nothing Then
                        my_pap(layer_handler, i, tmp_element, 0, elt_exp_all, E0, sin_toa_in_rad, B_phi_rz, Fb, B_phi0, B_R_bar, B_P, B_a1_, XPP_B_b_, XPP_B_A1, B_Z_bar, fit_MAC, options,
                           B_Rx, B_Rm, B_Rc, B_A1, B_A2, B_B1)
                        papfluor(layer_handler, elt_exp_all, tmp_element.line(0).xray_energy, chi_a, mother_layer_id, layer_handler(i).element(j).mother_layer_id,
                                                 ffact, B_A1, B_A2, B_B1, B_Rc, B_Rm, B_Rx, fit_MAC, options)

                    ElseIf options.phi_rz_mode = "PROZA96" Then
                        Dim F As Double
                        Dim phi0 As Double
                        Dim rzm As Double
                        Dim alpha As Double
                        Dim beta As Double
                        Dim phi_rz As Double
                        Dim Rx As Double

                        PROZA96(layer_handler, i, tmp_element, 0, elt_exp_all, E0, sin_toa_in_rad, phi_rz, F, phi0, rzm, alpha, beta, Rx, options, fit_MAC)

                        Bastin_fluor(layer_handler, elt_exp_all, tmp_element.line(0).xray_energy, chi_a, mother_layer_id, layer_handler(i).element(j).mother_layer_id,
                                     ffact, F, phi0, rzm, alpha, beta, Rx, options, fit_MAC)

                    End If


                    Dim shell1_A As Integer
                    Dim shell2_A As Integer
                    Siegbahn_to_transition_num(studied_elt.line(0).xray_name, shell1_A, shell2_A)
                    Dim omega_gamma_sur_A_A As Double = constante_simple(studied_elt, shell1_A, shell2_A)
                    'Dim sigma_ph_xs() As Double = interpol_log_log(studied_elt.ph_ion_xs, Energy_Xray_lines_of_B(k))
                    Dim sigma_ph_A_E As Double '= sigma_ph_xs(shell1_A)

                    sigma_ph_A_E = Xray_production_xs_ph_impact(studied_elt, shell1_A, Energy_Xray_lines_of_B(k))

                    Dim shell1_B As Integer
                    Dim shell2_B As Integer
                    Siegbahn_to_transition_num(Name_Xray_lines_of_B(k), shell1_B, shell2_B)
                    Dim omega_gamma_sur_A_B As Double = constante_simple(tmp_element, shell1_B, shell2_B)


                    Dim sigma_el_B_E As Double '= sigma_el_xs(shell1_B - 1)
                    'sigma_el_B_E = Xray_production_xs_el_impact(layer_handler(i).element(j), shell1_B, E0)

                    Dim norm_xs As Double = 1

                    If options.ionizationXS_mode = "Bote" Then
                        sigma_el_B_E = Xray_production_xs_el_impact(tmp_element, shell1_B, E0)
                        sigma_el_B_E = sigma_el_B_E * 6.022 * 10 ^ 23 / (4 * Math.PI)
                    Else
                        norm_xs = Xray_production_xs_el_impact(tmp_element, shell1_B, 15) * 6.022 * 10 ^ 23 / (4 * Math.PI) /
                                qe0(tmp_element.line(0).Ec, 15, tmp_element.line(0).xray_name, layer_handler(i).element(j).z)
                        sigma_el_B_E = qe0(tmp_element.line(0).Ec, E0, tmp_element.line(0).xray_name, layer_handler(i).element(j).z)
                    End If

                    '*******************************
                    'Dim norm_xs As Double = Xray_production_xs_el_impact(layer_handler(i).element(j), shell1_B, 15) * 6.022 * 10 ^ 23 / (4 * Math.PI) /
                    '        qe0(layer_handler(i).element(j).Ec, 15, layer_handler(i).element(j).xray_name, layer_handler(i).element(j).z, "E")

                    'sigma_el_B_E = qe0(layer_handler(i).element(j).Ec, E0, layer_handler(i).element(j).xray_name, layer_handler(i).element(j).z, "E")
                    '*******************************


                    Dim const_A As Double = omega_gamma_sur_A_A * studied_elt_conc * sigma_ph_A_E * 6.022E+23
                    'constante_ph_xs(studied_elt, shell1_A, shell2_A, Energy_Xray_lines_of_B(j)) '/ (4 * Math.PI)

                    Dim const_B As Double = omega_gamma_sur_A_B * layer_handler(i).element(j).conc_wt * sigma_el_B_E * norm_xs
                    'Dim const_B As Double = omega_gamma_sur_A_B * layer_handler(i).element(j).concentration * sigma_el_B_E * 6.022E+23 / (4 * Math.PI)
                    'constante_el_xs(layer_handler(studied_elt.mother_layer_id).element(i), shell1, shell2, E0) / (4 * Math.PI)


                    'Dim Pij As Double = 0
                    'If layer_handler(i).element(j).xray_name(0) = studied_elt.xray_name(0) Then
                    '    Pij = 1
                    'ElseIf layer_handler(i).element(j).xray_name(0) = "K" And studied_elt.xray_name(0) = "L" Then
                    '    Pij = 0.24
                    'ElseIf layer_handler(i).element(j).xray_name(0) = "L" And studied_elt.xray_name(0) = "K" Then
                    '    Pij = 4.2
                    'Else
                    '    'MsgBox("Error in Kab for Pij")
                    '    Debug.WriteLine("Error in Kab for Pij")
                    'End If

                    'Dim Ua As Double = E0 / studied_elt.Ec
                    'Dim Ub As Double = E0 / layer_handler(i).element(j).Ec
                    'fluo_char = fluo_char + 0.5 * Pij * sigma_ph_A_E * 6.022E+23 * omega_gamma_sur_A_B * ((Ub - 1) / (Ua - 1)) ^ 1.67 * Fa / Fb * ffact
                    '((Ub * Math.Log(Ub) - Ub + 1) / (Ua * Math.Log(Ua) - Ua + 1)) 
                    'fluo_char = fluo_char + 0.5 * ffact * TT * const_A * const_B * Pij * ((Ub * Math.Log(Ub) - Ub + 1) / (Ua * Math.Log(Ua) - Ua + 1)) * Fa / Fb

                    fluo_char = fluo_char + 0.5 * ffact * TT * const_A * const_B '/ B_phi_rz

                Next

            Next
        Next

        test_fluo_charact = fluo_char

    End Function
End Module
