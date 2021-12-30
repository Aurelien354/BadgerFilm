Imports System.IO

Module fluo_brem_v3_module

    Public Function test_fluo_brem_v3(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_elt_conc As Double, ByVal studied_elt As Elt_exp,
                                      ByVal line_indice As Integer, ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByVal sin_toa_in_rad As Double,
                                      ByVal fit_MAC As fit_MAC, ByVal options As options, ByVal brem_fit() As Double) As Double
        Try
            If E0 < studied_elt.line(line_indice).Ec Then Return 0
            'Dim lol As Double = Form1.TextBox6.Text
            'If Double.IsNaN(lol) Then lol = 1
            'E0 = E0 * lol

            'E0 = E0 * -0.000000763341 * studied_elt.z ^ 5 + 0.000106495 * studied_elt.z ^ 4 - 0.00577959 * studied_elt.z ^ 3 + 0.152984 * studied_elt.z ^ 2 - 1.99947 * studied_elt.z + 11.6347
            If brem_fit IsNot Nothing Then
                E0 = E0 * brem_fit(1)
            Else
                If studied_elt.line(line_indice).xray_name Like "[Kk]*" Then
                    If studied_elt.z <= 14 Then
                        E0 = E0 * (-0.003737141 * studied_elt.z ^ 4 + 0.1519312 * studied_elt.z ^ 3 - 2.267005 * studied_elt.z ^ 2 + 14.78875 * studied_elt.z - 33.29064)

                        'E0 = E0 * (-0.00001248424 * studied_elt.z ^ 4 - 0.001463123 * studied_elt.z ^ 3 + 0.04666448 * studied_elt.z ^ 2 - 0.2494477 * studied_elt.z + 1.773316)
                        'E0 = E0 * (0.00003808513 * studied_elt.z ^ 4 - 0.004643621 * studied_elt.z ^ 3 + 0.207022 * studied_elt.z ^ 2 - 3.99248 * studied_elt.z + 29.038)
                    End If
                    'E0 = E0 * (-0.000000763341 * studied_elt.z ^ 5 + 0.000106495 * studied_elt.z ^ 4 - 0.00577959 * studied_elt.z ^ 3 + 0.152984 * studied_elt.z ^ 2 - 1.99947 * studied_elt.z + 11.6347)
                ElseIf studied_elt.line(line_indice).xray_name Like "[Ll]*" Then
                    If studied_elt.z <= 52 Then
                        E0 = E0 * (-0.0005985668 * studied_elt.z ^ 2 + 0.04117917 * studied_elt.z + 0.4325065)
                    ElseIf studied_elt.z <= 72 Then
                        E0 = E0 * (0.0009081775 * studied_elt.z ^ 2 - 0.1215882 * studied_elt.z + 5.01147)
                    ElseIf studied_elt.z <= 86 Then
                        E0 = E0 * (0.0009001503 * studied_elt.z ^ 2 - 0.1412799 * studied_elt.z + 6.468995)
                    Else
                        E0 = E0 * (0.0005361637 * studied_elt.z ^ 2 - 0.09322002 * studied_elt.z + 4.962445)
                    End If
                    'E0 = E0 * (-0.0000004330739 * studied_elt.z ^ 4 + 0.0001030287 * studied_elt.z ^ 3 - 0.008865975 * studied_elt.z ^ 2 + 0.3213563 * studied_elt.z - 3.036796)
                    'E0 = E0 * (0.00000000836595 * studied_elt.z ^ 5 - 0.00000271563 * studied_elt.z ^ 4 + 0.000343512 * studied_elt.z ^ 3 - 0.021044 * studied_elt.z ^ 2 + 0.616591 * studied_elt.z - 5.7698)
                ElseIf studied_elt.line(line_indice).xray_name Like "[Mm]*" Then
                    E0 = E0 * (0.0000009958682 * studied_elt.z ^ 4 - 0.0003100635 * studied_elt.z ^ 3 + 0.03588155 * studied_elt.z ^ 2 - 1.83211 * studied_elt.z + 35.86223)

                End If
            End If
            'Dim toa_in_rad As Double = toa * Math.PI / 180

            'Calculate TT and x ***********************
            Dim TT As Double = 0
            If mother_layer_id <> 0 Then
                For i As Integer = 0 To mother_layer_id - 1
                    TT = TT + MAC_calculation(studied_elt, line_indice, i, layer_handler, elt_exp_all, fit_MAC, options) / sin_toa_in_rad * layer_handler(i).mass_thickness
                Next
                TT = Math.Exp(-TT)
            Else
                TT = 1
            End If
            '******************************************

            Dim chi_a As Double = MAC_calculation(studied_elt, line_indice, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options) / sin_toa_in_rad 'Math.Sin(toa_in_rad)


            Dim shell1 As Integer
            Dim shell2 As Integer
            Siegbahn_to_transition_num(studied_elt.line(line_indice).xray_name, shell1, shell2)

            Dim energy_division() As Double = find_all_Ec(elt_exp_all, studied_elt.line(line_indice).Ec, E0)  'contains all the energy subdivisions from Ec to E0. N divisions = (N-2) absorption edges + Ec + E0.

            Dim M As Double = 5.99 * 10 ^ -3 * E0 + 1.05
            Dim B As Double = -3.22 * 10 ^ -2 * E0 + 5.8
            Dim q As Double = 10 ^ -8
            Dim C As Double = 6 * 10 ^ -10

            Dim Ifluo_brem_tot As Double = 0

            For mm As Integer = 0 To UBound(energy_division) - 1
                Dim Ifluo_brem_Simpson As Double = 0

                For nn As Integer = 0 To 4 ' Perform integration of Y(E)dE (p.74) with Simpson's rule (in 5 points).
                    Dim E_actual As Double = energy_division(mm) - (energy_division(mm) - energy_division(mm + 1)) / 4 * nn

                    If nn = 4 Then
                        E_actual = energy_division(mm + 1) + 0.001
                    End If

                    If (E_actual > E0 - 0.001) Then Continue For


                    Dim Fb As Double = 0
                    Dim B_Z_bar As Double = 0
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

                    Dim Ifluo_brem As Double = 0

                    For l As Integer = 0 To UBound(layer_handler)

                        '**************************************************************************
                        Dim tmp_element As Elt_exp = Nothing
                        ReDim tmp_element.line(0)
                        'tmp_element = studied_elt  'layer_handler(studied_elt.mother_layer_id).element(i)

                        tmp_element.z = studied_elt.z
                        'tmp_element.z = 0
                        'For k As Integer = 0 To UBound(layer_handler(l).element)
                        '    tmp_element.z = tmp_element.z + layer_handler(l).element(k).z * layer_handler(l).element(k).concentration
                        'Next

                        'tmp_element.mother_layer_id = l 'UBound(layer_handler)
                        tmp_element.line(0).xray_energy = 0 'E_actual 'energy_division(mm) + 0.001   'Energy_Xray_lines_of_B(j)
                        tmp_element.line(0).xray_name = "Ka"
                        tmp_element.line(0).Ec = E_actual '(energy_division(mm + 1) + energy_division(mm)) / 2

                        my_pap(layer_handler, l, tmp_element, 0, elt_exp_all, E0, sin_toa_in_rad, B_phi_rz, Fb, B_phi0, B_R_bar, B_P, B_a1_, XPP_B_b_, XPP_B_A1, B_Z_bar,
                               fit_MAC, options, B_Rx, B_Rm, B_Rc, B_A1, B_A2, B_B1)
                        '***********************************************************************************
                        'If tmp_element.z = 13 Then ' And layer_handler(0).element.Count > 1 Then ' mother_layer_id <> 0 And
                        '    Dim tmp As String = "Rx" & vbTab & B_Rx & vbCrLf
                        '    tmp = tmp & "Rm" & vbTab & B_Rm & vbCrLf
                        '    tmp = tmp & "Rc" & vbTab & B_Rc & vbCrLf
                        '    tmp = tmp & "A1" & vbTab & B_A1 & vbCrLf
                        '    tmp = tmp & "A2" & vbTab & B_A2 & vbCrLf
                        '    tmp = tmp & "B1" & vbTab & B_B1 & vbCrLf
                        '    Debug.Print(tmp)
                        '    'My.Computer.Clipboard.SetText(tmp)
                        'End If

                        tmp_element.line(0).xray_energy = E_actual

                        Dim ffact As Double = 0
                        papfluor(layer_handler, elt_exp_all, tmp_element, chi_a, mother_layer_id, l, ffact, B_A1, B_A2, B_B1, B_Rc, B_Rm, B_Rx, fit_MAC, options)

                        If ffact < 0 Then
                            ffact = 0
                        End If

                        If Double.IsNaN(ffact) Then
                            Debug.Print("Error in brem fluorescence module (by other layers): ffact < 0 or Nan.")
                            ffact = 0
                        End If
                        '***********************************************************************

                        Ifluo_brem = Ifluo_brem + ffact / B_phi_rz 'Fb
                    Next


                    'Dim M As Double = 5.99 * 10 ^ -3 * E0 + 1.05
                    'Dim B As Double = -3.22 * 10 ^ -2 * E0 + 5.8
                    'Dim q As Double = 10 ^ -8
                    'Dim C As Double = 6 * 10 ^ -10

                    'Dim correction_curve As Double = 1
                    'If studied_elt.xray_name(0) = "K" Then
                    'correction_curve = -452.2436841 / (-2.513378579 * (studied_elt.z - 4.274881851)) + 7.597748504
                    'End If
                    Dim twiking_factor As Double
                    If brem_fit Is Nothing Then
                        If studied_elt.line(line_indice).xray_name Like "[Kk]*" Then

                            If studied_elt.z <= 14 Then
                                twiking_factor = 0.000004476461 * studied_elt.z ^ 4 - 0.000153563 * studied_elt.z ^ 3 + 0.001743002 * studied_elt.z ^ 2 - 0.007318512 * studied_elt.z + 0.0193749
                                ' twiking_factor = -0.00002398831 * studied_elt.z ^ 4 + 0.0009881932 * studied_elt.z ^ 3 - 0.01510198 * studied_elt.z ^ 2 + 0.1004173 * studied_elt.z - 0.2293788
                            Else
                                twiking_factor = -0.0000001153359 * studied_elt.z ^ 4 + 0.00001138019 * studied_elt.z ^ 3 - 0.0004570208 * studied_elt.z ^ 2 + 0.009165271 * studied_elt.z - 0.009048994
                                'twiking_factor = 0.0000001381234 * studied_elt.z ^ 4 - 0.00002119958 * studied_elt.z ^ 3 + 0.001064032 * studied_elt.z ^ 2 - 0.02112858 * studied_elt.z + 0.2060117
                            End If
                            '    twiking_factor = -0.0000262617 * studied_elt.z ^ 2 + 0.00353401 * studied_elt.z - 0.0168204
                        ElseIf studied_elt.line(line_indice).xray_name Like "[Ll]*" Then
                            If studied_elt.z <= 52 Then
                                twiking_factor = 0.00005604014 * studied_elt.z ^ 2 - 0.004304944 * studied_elt.z + 0.1181698
                            ElseIf studied_elt.z <= 72 Then
                                twiking_factor = -0.0000004489685 * studied_elt.z ^ 4 + 0.0001081349 * studied_elt.z ^ 3 - 0.009801686 * studied_elt.z ^ 2 + 0.397427 * studied_elt.z - 6.055677
                            ElseIf studied_elt.z <= 86 Then
                                twiking_factor = -0.00003568135 * studied_elt.z ^ 3 + 0.008211648 * studied_elt.z ^ 2 - 0.6290447 * studied_elt.z + 16.08798
                            Else
                                twiking_factor = -0.0006732208 * studied_elt.z ^ 2 + 0.1195678 * studied_elt.z - 5.261812
                            End If
                            'twiking_factor = 0.000000007144458 * studied_elt.z ^ 4 - 0.000002195895 * studied_elt.z ^ 3 + 0.0002342551 * studied_elt.z ^ 2 - 0.01018744 * studied_elt.z + 0.1928773
                            '    twiking_factor = -0.00000063491 * studied_elt.z ^ 3 + 0.000110601 * studied_elt.z ^ 2 - 0.00597444 * studied_elt.z + 0.139793

                        ElseIf studied_elt.line(line_indice).xray_name Like "[Mm]*" Then
                            twiking_factor = -0.00000002861734 * studied_elt.z ^ 4 + 0.000008821865 * studied_elt.z ^ 3 - 0.001005127 * studied_elt.z ^ 2 + 0.05026985 * studied_elt.z - 0.9006125

                        Else
                            twiking_factor = 0.05
                        End If
                        'twiking_factor = -0.0000002028089 * studied_elt.z ^ 4 + 0.00002166194 * studied_elt.z ^ 3 - 0.0008968171 * studied_elt.z ^ 2 + 0.01728938 * studied_elt.z - 0.06327038 '0.05 '-0.0000262617 * studied_elt.z ^ 2 + 0.00353401 * studied_elt.z - 0.0168204 
                    Else
                        twiking_factor = brem_fit(0)
                    End If

                    'If layer_handler.Length > 1 Then
                    '    If studied_elt.z = 26 Then
                    '        twiking_factor = 0.021857866
                    '    End If
                    '    If studied_elt.z = 14 Then
                    '        twiking_factor = 0.00599018
                    '    End If
                    'End If
                    'Dim tmplol1 As Double = q * Math.Exp(B) * (B_Z_bar * (E0 / E_actual - 1)) ^ M
                    'Dim tmplol2 As Double = C * B_Z_bar ^ 2
                    'Dim I_E As Double = twiking_factor * q / correction_curve * Math.Exp(B) * (B_Z_bar * (E0 / E_actual - 1)) ^ M + C * B_Z_bar ^ 2
                    Dim I_E As Double = twiking_factor * q * Math.Exp(B) * (B_Z_bar * (E0 / E_actual - 1)) ^ M + C * B_Z_bar ^ 2
                    'Console.WriteLine(B_Z_bar)

                    '************************************************
                    'Dim tmp1 As Double = Math.Sqrt(B_Z_bar * (E0 / E_actual - 1))
                    'Dim tmp2 As Double = -54.86 - 1.072 * E_actual + 0.2835 * E0 + 30.4 * Math.Log(B_Z_bar) + 875 / (B_Z_bar ^ 2 * E0 ^ 0.08)
                    'I_E = tmp1 * tmp2 * 10 ^ -8 '* 2.4

                    'tmp1 = Math.Sqrt(B_Z_bar) * (E0 / E_actual - 1)
                    'tmp2 = -73.9 - 1.2446 * E_actual + 36.502 * Math.Log(B_Z_bar) + (148.5 * E0 ^ 0.1293) / B_Z_bar
                    'Dim tmp3 As Double = 1 + (-0.006624 + 0.0002906 * E0) * B_Z_bar / E_actual
                    'Dim tmp4 As Double = 0.00193581 * E0 ^ 2 + 0.0310694 * E0 + 0.346158
                    'I_E = tmp1 * tmp2 * tmp3 * 10 ^ -8 / tmp4 '* 0.91 '3.78
                    '************************************************

                    Dim sigma_ph_A_E As Double
                    sigma_ph_A_E = Xray_production_xs_ph_impact(studied_elt, shell1, E_actual)


                    Dim simpson_coeff As Integer = 2
                    If nn = 0 Or nn = 4 Then
                        simpson_coeff = 1
                    ElseIf nn = 1 Or nn = 3 Then
                        simpson_coeff = 4
                    End If

                    Ifluo_brem_Simpson = Ifluo_brem_Simpson + sigma_ph_A_E * 6.022E+23 / studied_elt.a * Ifluo_brem * I_E * simpson_coeff
                Next

                Ifluo_brem_tot = Ifluo_brem_tot + Ifluo_brem_Simpson * ((energy_division(mm) - energy_division(mm + 1)) / 4) / 3
            Next

            Dim W_A As Double = extract_fluorescence_yield(studied_elt.at_data, shell1)
            Dim em_rate_A As Double = extract_emission_rate(studied_elt.at_data, shell1, shell2) '/ (4 * Math.PI)

            Dim Kf As Double = 0.5 * W_A * studied_elt_conc * em_rate_A * TT

            test_fluo_brem_v3 = Ifluo_brem_tot * Kf

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in test_fluo_brem_v3 " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    ' Integrate (ax2+bx+c)*Exp(r.x)*EI(epsilon.x+beta)dx
    Public Function my_papexei(ByVal limit_integ(,) As Double, ByVal ap() As Double, ByVal bp() As Double, ByVal cp() As Double,
                               ByVal r As Double, ByVal epsilon As Double, ByVal beta As Double) As Double
        Try
            Dim q2, t2, u2, eiu2, eiq2, eu2, et2, rv, w, i1(1), i2(1), i3(1, 2), i3a, i3b, i3c As Double

            Dim ev, p1, p2, p3, h, v As Double 'AMXX
            'Dim sign As Integer 'AMXX

            'dr = r
            'dq = q
            If ((r + epsilon) <> 0.0) Then w = Math.Log(Math.Abs(1 + r / epsilon))
            ev = Math.Exp(-r * beta / epsilon)
            'p = 0.0
            p1 = 0.0
            p2 = 0.0
            p3 = 0.0
            h = beta / epsilon
            v = r + epsilon
            rv = r / v

            For i As Integer = 0 To 1
                If (limit_integ(i, 0) = 1) Then Continue For

                For j As Integer = 0 To 1
                    'sign = (j + 1) * 2 - 3
                    q2 = epsilon * limit_integ(i, j) + beta
                    If (q2 = 0.0) Then
                        i1(j) = (-w) / r
                        i2(j) = (-rv + w) / r ^ 2 - h * i1(j)
                        i3(j, 0) = -2 * w
                        i3(j, 1) = rv ^ 2 + 2 * rv - h * (2 * i2(j) + h * i1(j)) * r ^ 3
                        i3(j, 2) = 0.0
                    Else
                        t2 = r / epsilon * q2
                        u2 = (r / epsilon + 1) * q2
                        eiq2 = expint(q2)
                        eiu2 = expint(u2)
                        et2 = Math.Exp(t2)
                        eu2 = Math.Exp(u2)
                        i1(j) = (et2 * eiq2 - eiu2) / r
                        i2(j) = (et2 * eiq2 * (t2 - 1) - rv * eu2 + eiu2) / r ^ 2 - h * i1(j)

                        ' due To roundoff errors, the third integral
                        ' must be evaluated in three parts

                        i3(j, 0) = et2 * eiq2 * (t2 ^ 2 - 2 * t2 + 2)
                        i3(j, 1) = -rv ^ 2 * eu2 * (u2 - 1) + 2 * rv * eu2 - h * (2 * i2(j) + h * i1(j)) * r ^ 3
                        i3(j, 2) = -2 * eiu2

                    End If
                Next
                i3a = i3(1, 0) - i3(0, 0)
                i3b = i3(1, 1) - i3(0, 1)
                i3c = i3(1, 2) - i3(0, 2)
                p1 = p1 + (i1(1) - i1(0)) * ev * cp(i)
                p2 = p2 + (i2(1) - i2(0)) * ev * bp(i)
                p3 = p3 + (i3a + i3b + i3c) * ev / r ^ 3 * ap(i)
            Next
            my_papexei = p1 + p2 + p3
            'my_papexei = p
            Return my_papexei

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in my_papexei " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    ' Integrate (ax^2+bx+c)*EI(b.x+c)dx
    Public Function my_papei(ByVal limit_integ(,) As Double, ByVal ap() As Double, ByVal bp() As Double, ByVal cp() As Double,
                           ByVal b As Double, ByVal c As Double) As Double
        Try
            Dim i1(1), i2(1), i3(1), u2, eiu2, eu2 As Double

            Dim p1, p2, p3, h As Double 'AMXX

            p1 = 0.0
            p2 = 0.0
            p3 = 0.0

            For i As Integer = 0 To 1
                If (limit_integ(i, 0) = 1) Then Continue For

                For j As Integer = 0 To 1
                    u2 = b * limit_integ(i, j) + c
                    If (u2 = 0.0) Then
                        eiu2 = 1.0
                    Else
                        eiu2 = expint(u2)
                    End If
                    eu2 = Math.Exp(u2)
                    h = c / b
                    i1(j) = (u2 ^ 1 * eiu2 - eu2) / b
                    i2(j) = (u2 ^ 2 * eiu2 - eu2 * (u2 - 1.0)) / (2 * b ^ 2) - h * i1(j)
                    i3(j) = (u2 ^ 3 * eiu2 - eu2 * (u2 ^ 2 - 2 * u2 + 2)) / (3 * b ^ 3) - h * (2 * i2(j) + h * i1(j))
                Next
                p1 = p1 + (i1(1) - i1(0)) * cp(i)
                p2 = p2 + (i2(1) - i2(0)) * bp(i)
                p3 = p3 + (i3(1) - i3(0)) * ap(i)
            Next
            my_papei = p1 + p2 + p3
            'my_papei = p
            Return my_papei

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in my_papei " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    ' Integrate (ax^2+bx+c)*exp(r*x)dx
    Public Function my_papex(ByVal limit_integ(,) As Double, ByVal ap() As Double, ByVal bp() As Double, ByVal cp() As Double, ByVal r As Double) As Double
        Try
            Dim i1(1), i2(1), i3(1), ru, ru2 As Double

            Dim p1, p2, p3 As Double 'AMXX

            p1 = 0.0
            p2 = 0.0
            p3 = 0.0

            For i As Integer = 0 To 1
                If (limit_integ(i, 0) = 1) Then Continue For

                For j As Integer = 0 To 1 'Calculate upper and lower integral limits
                    ru = r * limit_integ(i, j)
                    ru2 = Math.Exp(ru)
                    i1(j) = ru2
                    i2(j) = ru2 * (ru - 1)
                    i3(j) = ru2 * (ru ^ 2 - 2 * ru + 2)
                Next

                p1 = p1 + (i1(1) - i1(0)) * cp(i) / r
                p2 = p2 + (i2(1) - i2(0)) * bp(i) / r ^ 2
                p3 = p3 + (i3(1) - i3(0)) * ap(i) / r ^ 3
            Next

            my_papex = p1 + p2 + p3

            Return my_papex

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in my_papex " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function


    Public Sub papfluor(ByVal layer_handler() As layer, ByVal elt_exp_all() As Elt_exp, ByVal tmp_element As Elt_exp, ByVal chi As Double, ByVal ma As Integer, ByVal mb As Integer,
                        ByRef ffact As Double, ByVal pa1 As Double, ByVal pa2 As Double, ByVal pb1 As Double, ByVal rc As Double, ByVal rm As Double, ByVal rx As Double,
                        ByVal fit_MAC As fit_MAC, ByVal options As options)
        Try
            Dim sumtxmu, d, uv, a1, f1, f2, f3, f4, b1, b2, g1, g2, g3, g4, debb1 As Double 'AMXX
            Dim n As Integer 'AMXX

            Dim de(UBound(layer_handler) + 1) As Double
            de(0) = 0
            For i As Integer = 0 To UBound(layer_handler)
                de(i + 1) = de(i) + layer_handler(i).mass_thickness
            Next

            Dim mu(UBound(layer_handler)) As Double
            For i As Integer = 0 To UBound(layer_handler)
                mu(i) = MAC_calculation(tmp_element, 0, i, layer_handler, elt_exp_all, fit_MAC, options)
            Next



            If (mb - ma) >= 0 Then 'AMXX
                n = 1
            Else
                n = -1
            End If 'AMXX

            Dim t(UBound(layer_handler)) As Double
            For kk As Integer = 0 To UBound(de) - 1
                t(kk) = de(kk + 1) - de(kk)
            Next
            t(UBound(layer_handler)) = 0.0

            sumtxmu = 0.0
            If (Math.Abs(ma - mb) > 1) Then
                For kk As Integer = ma + n To mb - n Step n
                    'For kk As Integer = ma + n * 1 To mb - n * 1 Step n
                    sumtxmu = sumtxmu + n * mu(kk) * t(kk)
                Next
            End If

            Dim ap(1) As Double
            Dim bp(1) As Double
            Dim cp(1) As Double
            ap(0) = pa1
            ap(1) = pa2
            bp(0) = -2 * rm * pa1
            bp(1) = -2 * rx * pa2
            cp(0) = pa1 * rm ^ 2 + pb1
            cp(1) = pa2 * rx ^ 2

            d = 1.0 / chi

            Dim lim_min As Double = de(mb)
            Dim lim_max As Double = de(mb + 1)

            Dim limit_integ(1, 1) As Double
            limit_integ(0, 0) = 0
            limit_integ(0, 1) = rc
            limit_integ(1, 0) = rc
            limit_integ(1, 1) = rx

            If mb = 0 Then
                If (rc > lim_max) Then
                    limit_integ(0, 1) = lim_max
                    limit_integ(1, 0) = 1.0
                ElseIf (lim_max < rx) Then
                    limit_integ(1, 1) = lim_max
                End If
            Else
                If (rc < lim_min) Then
                    limit_integ(0, 0) = 1.0
                    limit_integ(1, 0) = lim_min
                    If (lim_max < rx) Then limit_integ(1, 1) = lim_max
                Else
                    limit_integ(0, 0) = lim_min
                    If (rc > lim_max) Then
                        limit_integ(0, 1) = lim_max
                        limit_integ(1, 0) = 1.0
                    Else
                        If (lim_max < rx) Then limit_integ(1, 1) = lim_max
                    End If
                End If
            End If
            If (limit_integ(1, 0) > rx) Then limit_integ(1, 0) = 1.0



            uv = Math.Log(Math.Abs((mu(ma) + chi) / (mu(ma) - chi)))
            'If (mode = "B") Then
            '    a1 = -chi
            '    f1 = -mu(mb)
            '    f2 = chi - mu(mb)
            '    ffact = -d * (my_papei(limit_integ, ap, bp, cp, f1, 0.0) - (my_papexei(limit_integ, ap, bp, cp, a1, f2, 0.0) +
            '        uv * my_papex(limit_integ, ap, bp, cp, a1)))
            '    Return
            'ElseIf (ma = mb) Then
            If (ma = mb) Then
                a1 = -chi
                b1 = chi * de(mb)
                f1 = -mu(mb)
                f2 = chi - mu(mb)
                g1 = mu(mb) * de(mb)
                g2 = (-chi + mu(mb)) * de(mb)
                ffact = -d * (my_papei(limit_integ, ap, bp, cp, f1, g1) - Math.Exp(b1) * (my_papexei(limit_integ, ap, bp, cp, a1, f2, g2) +
                    uv * my_papex(limit_integ, ap, bp, cp, a1)))
                If (mb = UBound(layer_handler)) Then Return
                b2 = -chi * t(mb)
                f3 = mu(mb)
                f4 = (chi + mu(mb))
                g3 = -mu(mb) * de(mb + 1)
                g4 = -(chi + mu(mb)) * de(mb + 1)
                ffact = ffact + d * (Math.Exp(b2) * my_papei(limit_integ, ap, bp, cp, f3, g3) - Math.Exp(b1) * my_papexei(limit_integ, ap, bp, cp, a1, f4, g4))
                Return
            Else
                debb1 = de(mb)
                If (ma > mb) Then debb1 = de(mb + 1)
                a1 = -chi * mu(mb) / mu(ma)
                b1 = chi / mu(ma) * (mu(mb) * debb1 - sumtxmu) - chi * t(ma)
                b2 = -chi * t(ma)
                f1 = -n * mu(mb)
                f2 = mu(mb) * (chi / mu(ma) - n * 1)
                g1 = n * (mu(mb) * debb1 - sumtxmu) - mu(ma) * t(ma)
                g2 = -(mu(mb) * debb1 - sumtxmu) * (chi / mu(ma) - n * 1)
                g3 = g2 + n * (chi - n * mu(ma)) * t(ma)
                g4 = n * (mu(mb) * debb1 - sumtxmu)
                If (ma > mb) Then
                    b1 = b1 + chi * t(ma)
                    g1 = g1 + mu(ma) * t(ma)
                    g4 = g4 - mu(ma) * t(ma)
                End If
                ffact = -d * (my_papei(limit_integ, ap, bp, cp, f1, g1) + n * Math.Exp(b1) * my_papexei(limit_integ, ap, bp, cp, a1, f2, g2))
                If (ma = UBound(layer_handler)) Then Return
                ffact = ffact + d * (n * Math.Exp(b1) * my_papexei(limit_integ, ap, bp, cp, a1, f2, g3) + Math.Exp(b2) * my_papei(limit_integ, ap, bp, cp, f1, g4))
                Return
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in papfluor " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

End Module
