Module Bastin_phrz_module

    Public Sub PROZA96(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                      ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByVal sin_toa_in_rad As Double, ByRef phi_rz_ As Double, ByRef FI_ As Double, ByRef phi0_ As Double,
                       ByRef rzm_ As Double, ByRef alpha_ As Double, ByRef beta_ As Double, ByRef Rx As Double, ByVal options As options, Optional fit_MAC As fit_MAC = Nothing)

        'If studied_element.z = 26 And E0 = 19 Then
        '    Stop
        'End If

        If studied_element.line(line_indice).Ec > E0 Then
            phi_rz_ = 0
            FI_ = 0
            phi0_ = 0
            rzm_ = 0
            alpha_ = 0
            beta_ = 0
            Exit Sub
        End If

        Dim phi0(UBound(layer_handler)) As Double '= phi0_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        Dim rzm(UBound(layer_handler)) As Double '= rzm_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        Dim FI(UBound(layer_handler)) As Double '= FI_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        Dim alpha(UBound(layer_handler)) As Double '= alpha_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        Dim beta(UBound(layer_handler)) As Double

        For i As Integer = 0 To UBound(layer_handler)
            PROZA96_coeff(layer_handler, i, studied_element, line_indice, elt_exp_all, E0, phi0(i), rzm(i), FI(i), alpha(i), beta(i), True)

        Next

        '************************
        'Calculate the cumulative mass depth (ad) for each layer
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

        Dim layer_limit_min As Double = 0
        Dim layer_limit_max As Double = cumulative_mass_depth(mother_layer_id)
        If mother_layer_id <> 0 Then
            layer_limit_min = cumulative_mass_depth(mother_layer_id - 1)
        End If


        Dim alpha_final As Double = 0
        Dim rzm_final As Double = 0
        Dim FI_final As Double = 0
        Dim phi0_final As Double = 0
        Dim beta_final As Double = 0
        Dim Rx_new As Double = 0

        If UBound(layer_handler) <> 0 And mother_layer_id <> UBound(layer_handler) Then
            Dim Rx_old As Double = 0
            For i As Integer = 0 To UBound(layer_handler)
                Rx_new = Rx_new + rzm(i) + 2.5 / alpha(i)
            Next
            Rx_new = Rx_new / (UBound(layer_handler) + 1)
            'Rx_new = ((rzm(mother_layer_id) + 2.5 / alpha(mother_layer_id)) + (rzm(mother_layer_id + 1) + 2.5 / alpha(mother_layer_id + 1))) / 2

            While Math.Abs(Rx_old - Rx_new) / Rx_old > 0.001
                Rx_old = Rx_new
                Rx_new = Bastin_weight(Rx_old, Rx_old, -0.4 * Rx_old, layer_limit_min, layer_limit_max) 'There is a problem here!
                If Rx_new < layer_limit_min Then
                    'For i As Integer = 0 To UBound(layer_handler)
                    '    Rx_new = Rx_new + rzm(i) + 2.5 / alpha(i)
                    'Next
                    'Rx_new = Rx_new / (UBound(layer_handler) + 1)
                    'Rx_old = Rx_new
                    Rx_new = (layer_limit_max + layer_limit_min) / 2
                End If
            End While

            For i As Integer = 0 To UBound(layer_handler)
                If i = 0 Then
                    alpha_final = alpha_final + Bastin_weight(alpha(i), Rx_new, -0.3 * Rx_new, 0, cumulative_mass_depth(i))
                    rzm_final = rzm_final + Bastin_weight(rzm(i), 0.7 * Rx_new, -0.9 * Rx_new, 0, cumulative_mass_depth(i))
                    FI_final = FI_final + Bastin_weight(FI(i), Rx_new, -0.8 * Rx_new, 0, cumulative_mass_depth(i))
                    phi0_final = phi0_final + Bastin_weight(phi0(i), 0.5 * Rx_new, -0.5 * Rx_new, 0, cumulative_mass_depth(i))
                Else
                    alpha_final = alpha_final + Bastin_weight(alpha(i), Rx_new, -0.3 * Rx_new, cumulative_mass_depth(i - 1), cumulative_mass_depth(i))
                    rzm_final = rzm_final + Bastin_weight(rzm(i), 0.7 * Rx_new, -0.9 * Rx_new, cumulative_mass_depth(i - 1), cumulative_mass_depth(i))
                    FI_final = FI_final + Bastin_weight(FI(i), Rx_new, -0.8 * Rx_new, cumulative_mass_depth(i - 1), cumulative_mass_depth(i))
                    phi0_final = phi0_final + Bastin_weight(phi0(i), 0.5 * Rx_new, -0.5 * Rx_new, cumulative_mass_depth(i - 1), cumulative_mass_depth(i))
                End If

            Next

            'alpha_final = Bastin_weight(alpha(mother_layer_id), Rx_new, -0.3 * Rx_new, layer_limit_min, layer_limit_max) 'AMXX Is it alpha(mother_layer_id) or (alpha(mother_layer_id)+alpha(mother_layer_id+1))/2 ?????????
            'rzm_final = Bastin_weight(rzm(mother_layer_id), 0.7 * Rx_new, -0.9 * Rx_new, layer_limit_min, layer_limit_max)
            'FI_final = Bastin_weight(FI(mother_layer_id), Rx_new, -0.8 * Rx_new, layer_limit_min, layer_limit_max)
            'phi0_final = Bastin_weight(phi0(mother_layer_id), 0.5 * Rx_new, -0.5 * Rx_new, layer_limit_min, layer_limit_max)


            PROZA96_coeff(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, E0, phi0_final, rzm_final, FI_final, alpha_final, beta_final, False)

        Else
            Rx_new = (rzm(mother_layer_id) + 2.5 / alpha(mother_layer_id))
            alpha_final = alpha(mother_layer_id)
            rzm_final = rzm(mother_layer_id)
            FI_final = FI(mother_layer_id)
            phi0_final = phi0(mother_layer_id)
            beta_final = beta(mother_layer_id)

        End If

        '************************
        'Calculate the 'wt.fraction averaged mass absorption coefficient'
        'for the current layer
        '************************
        Dim mac As Double
        Dim chi As Double
        mac = MAC_calculation(studied_element.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)

        chi = mac / sin_toa_in_rad 'Math.Sin(toa * Math.PI / 180)
        '************************

        Dim phim As Double = phi0_final * Math.Exp((beta_final * rzm_final) ^ 2)

        Dim I_E1 As Double
        Dim I_E2 As Double

        If mother_layer_id = 0 Then
            layer_limit_min = 0
        Else
            layer_limit_min = cumulative_mass_depth(mother_layer_id - 1)
        End If
        layer_limit_max = cumulative_mass_depth(mother_layer_id)
        'If layer_limit_max > Rx Then layer_limit_max = Rx

        If layer_limit_min > Rx_new Then
            I_E1 = 0
            I_E2 = 0
        ElseIf layer_limit_max <= rzm_final Then
            I_E1 = integrate_phirz_exp_murz(layer_limit_min, layer_limit_max, phim, beta_final, rzm_final, chi)
            I_E2 = 0

        ElseIf layer_limit_min >= rzm_final Then
            I_E1 = 0
            I_E2 = integrate_phirz_exp_murz(layer_limit_min, Math.Min(layer_limit_max, Rx_new), phim, alpha_final, rzm_final, chi)

        Else
            I_E1 = integrate_phirz_exp_murz(layer_limit_min, rzm_final, phim, beta_final, rzm_final, chi)
            I_E2 = integrate_phirz_exp_murz(rzm_final, Math.Min(layer_limit_max, Rx_new), phim, alpha_final, rzm_final, chi)

        End If

        Dim abs As Double = abs_outer_layers(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, mac, sin_toa_in_rad, fit_MAC, options)

        phi_rz_ = (I_E1 + I_E2) * abs
        FI_ = FI_final
        phi0_ = phi0_final
        rzm_ = rzm_final
        alpha_ = alpha_final
        beta_ = beta_final
        Rx = Rx_new

    End Sub

    Public Sub PROZA96_coeff(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                      ByVal elt_exp_all() As Elt_exp, ByVal E0 As Double, ByRef phi0 As Double, ByRef rzm As Double, ByRef FI As Double, ByRef alpha As Double, ByRef beta As Double,
                      ByVal calc_phi0_rzm_FI_alpha As Boolean)

        If calc_phi0_rzm_FI_alpha = True Then
            phi0 = phi0_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
            rzm = rzm_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
            FI = FI_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
            alpha = alpha_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        End If
        'Dim beta As Double
        'Dim FCN As Double

        Dim beta_0 As Double = alpha
        Dim beta_1 As Double = 1.5 * alpha
        Dim beta_2 As Double

        While True
            Dim FCN_0 As Double = FCN_function(phi0, rzm, alpha, beta_0, FI)
            Dim FCN_1 As Double = FCN_function(phi0, rzm, alpha, beta_1, FI)

            Dim tangent As Double
            Dim intercept As Double
            line_coeff(beta_0, beta_1, FCN_0, FCN_1, tangent, intercept)
            beta_2 = -intercept / tangent

            If Double.IsNaN(beta_2) Then Stop
            If beta_2 < 0 Then
                alpha = alpha * 1.01
                beta_0 = alpha
                beta_1 = 1.5 * alpha
                Continue While
            End If
            'Dim alpha_temp As Double = alpha 'AMXX check that! Is it really a temporary value or the new alpha value?
            'While beta_2 < 0
            '    'alpha_temp = alpha_temp * 1.01
            '    beta_0 = beta_0 * 0.9
            '    beta_1 = beta_1 * 0.9
            '    line_coeff(beta_0, beta_1, FCN_0, FCN_1, tangent, intercept)
            '    beta_2 = -intercept / tangent
            'End While

            'Dim FCN_2 As Double = FCN_function(phi0, rzm, alpha, beta_2, FI) 'AMXX This step seems to be useless

            Dim F As Double = FI_from_phirz_function(phi0, rzm, alpha, beta_2)

            If Math.Abs((F - FI) / FI) < 10 ^ -5 Then Exit While

            beta_1 = (beta_0 + beta_1) / 2
            beta_0 = beta_2
        End While
        beta = beta_2

    End Sub

    Public Function Bastin_weight(ByVal coeff_to_change As Double, ByVal R As Single, ByVal L As Single, ByVal rz_min As Double, ByVal rz_max As Double) As Double
        ' AMXX What happen when R is lower than rz_min??
        ' AMXX What happen when R is lower than rz_max but greater than rz_min??
        If R < rz_min Then Return 0
        '************************
        'Calculate N in eq. (39) p. 54 (xnorm)
        '************************
        Dim N As Double
        N = 1.0 / (primitive_pap_weight(R, R, L) - primitive_pap_weight(0.0, R, L))
        '************************
        Dim res As Double = coeff_to_change * N * (primitive_pap_weight(Math.Min(rz_max, R), R, L) - primitive_pap_weight(rz_min, R, L))
        Return res


    End Function

    Public Function FI_function(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double) As Double

        Dim Ec As Double = studied_element.line(line_indice).Ec
        Dim U0 As Double = E0 / Ec

        '************************
        'Calculate M 
        'Calculate Z bar
        'Calculate J
        '************************
        'Dim Z_bar As Double = 0
        Dim J As Double = 0
        Dim M As Double = 0
        'For i As Integer = 0 To UBound(layer_handler)
        For k As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            With layer_handler(mother_layer_id).element(k)
                'Z_bar = Z_bar + .z * .conc_wt
                M = M + .z / .a * .conc_wt
                J = J + .conc_wt * .z / .a * Math.Log(ionization_potential_Sternheimer(.z)) 'Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
            End With
        Next
        'Next
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
            'm_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
            Dim temp_Z As Integer = studied_element.z 'correction AMXX 11-March-2020
            If temp_Z = 6 Then
                m_ = 0.888
            ElseIf temp_Z = 7 Then
                m_ = 0.86
            ElseIf temp_Z = 8 Then
                m_ = 0.89
            Else
                m_ = 0.9
            End If
        End If

        '************************
        'Calculate Pi (p)
        'Calculate Di (ad)
        '************************
        Dim Pi(2) As Double
        Dim Di(2) As Double
        Pi(0) = 0.78
        Pi(1) = 0.1
        Pi(2) = -(0.5 - 0.25 * J)
        Di(0) = 6.6 * 10 ^ -6
        Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
        Di(2) = 2.2 * 10 ^ -6 / J

        '************************
        'Calculate V0 (v0)
        'Calculate QA_l (qe)
        'Calculate 1/S (inv_S)
        'Calculate F = R/S * 1/QA_l (df)
        '************************
        Dim V0 As Double
        Dim QA_l As Double
        Dim inv_S As Double = 0
        'Dim F As Double
        V0 = E0 / J
        QA_l = Math.Log(U0) / Ec ^ 2 / U0 ^ m_

        For k As Integer = 0 To 2
            Dim Tk As Double = 1 + Pi(k) - m_
            inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
        Next
        'inv_S = inv_S * (U0 / V0) / M
        inv_S = inv_S * U0 / (V0 * M)

        '************************
        'Calculate parameters
        'Eta bar
        'W bar
        'q
        'J(U0)
        'G(U0)
        'R
        '************************
        Dim eta_bar As Double
        Dim W_bar As Double
        Dim q_ As Double
        Dim J_U0 As Double
        Dim G_U0 As Double
        Dim R As Double
        Dim Zb_bar As Double = 0

        'For i As Integer = 0 To UBound(layer_handler)
        For k As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            Zb_bar = Zb_bar + layer_handler(mother_layer_id).element(k).z ^ 0.5 * layer_handler(mother_layer_id).element(k).conc_wt
        Next
        'Next
        Zb_bar = Zb_bar ^ 2

        eta_bar = 0.00175 * Zb_bar + 0.37 * (1 - Math.Exp(-0.015 * Zb_bar ^ 1.3))
        W_bar = 0.595 + eta_bar / 3.7 + eta_bar ^ 4.55
        q_ = (2 * W_bar - 1) / (1 - W_bar)
        J_U0 = 1 + U0 * (Math.Log(U0) - 1)
        G_U0 = (U0 - 1 - ((1 - (1 / U0 ^ (q_ + 1))) / (1 + q_))) / ((2 + q_) * J_U0)
        R = 1 - eta_bar * W_bar * (1 - G_U0)


        Dim FI As Double = R * inv_S / QA_l

        Return FI
    End Function

    Public Function phi0_function(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double) As Double

        Dim Ec As Double = studied_element.line(line_indice).Ec
        Dim U0 As Double = E0 / Ec
        Dim Z_total As Double = 0

        For i As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            Dim Z As Integer = layer_handler(mother_layer_id).element(i).z
            Dim C As Double = layer_handler(mother_layer_id).element(i).conc_wt
            Z_total = Z_total + C * Z
        Next

        Dim a As Double = 0.61747243 + 1.0991805 * 10 ^ -3 * Z_total + 1.224221 / Math.Sqrt(Z_total)
        Dim b As Double = -0.21964478 + 0.11332964 * Z_total - 2.0638629 * 10 ^ -2 * Z_total * Math.Log(Z_total)

        Return 1 + (1 - 1 / Math.Sqrt(U0)) ^ a * b

    End Function


    Public Function rzm_function(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double) As Double

        Dim Ec As Double = studied_element.line(line_indice).Ec
        Dim Jc As Double

        Dim temp_num As Double = 0
        Dim temp_den As Double = 0
        For i As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            Dim Z As Integer = layer_handler(mother_layer_id).element(i).z
            Dim A As Double = layer_handler(mother_layer_id).element(i).a
            Dim J As Double = ionization_potential_Sternheimer(Z)
            Dim C As Double = layer_handler(mother_layer_id).element(i).conc_wt

            temp_den = temp_den + C * Z / A
            temp_num = temp_num + temp_den * Math.Log(J)
        Next

        Jc = Math.Exp(temp_num / temp_den)

        Dim R0 As Double = 0

        Dim D(2) As Double
        Dim P(2) As Double
        D(0) = 6.6 * 10 ^ -6
        D(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * Jc ^ 2)
        D(2) = 2.2 * 10 ^ -6 / Jc
        P(0) = 0.78
        P(1) = 0.1
        P(2) = -(0.5 - 0.25 * Jc)

        For i As Integer = 0 To 2
            R0 = R0 + Jc ^ (1 - P(i)) * D(i) * (E0 ^ (1 + P(i)) - Ec ^ (1 + P(i))) / (1 + P(i))
        Next
        R0 = R0 * 1 / temp_den

        Dim Zc As Double = 0
        For i As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            Dim Z As Integer = layer_handler(mother_layer_id).element(i).z
            Dim C As Double = layer_handler(mother_layer_id).element(i).conc_wt
            Zc = Zc + C * Z
        Next

        'Dim a_ As Double = Math.Exp(-0.6801 + 2.254 * 10 ^ -3 * Zc * Math.Sqrt(Zc) - 0.09457 * Math.Sqrt(Zc) * Math.Log(Zc))
        'Dim b_ As Double = -0.02887 - 2.035 / (Zc * Math.Sqrt(Zc)) + 3.87 * Math.Exp(-Zc)
        'Dim c_ As Double = 1 / (27.78 - 6.123 * Math.Sqrt(Zc) - 52.96 * Math.Log(Zc) / Zc)

        'Dim U As Double = E0 / Ec
        'Dim P_ As Double = R0 * (a_ + b_ * Math.Log(U) / U + c_ / U)

        Dim a_ As Double = 2.1040483 + 0.044934014 * Zc * Math.Sqrt(Zc) - 5.518453 * 10 ^ -4 * Zc ^ 2 * Math.Sqrt(Zc) + 2.46257718 * 10 ^ -5 * Zc ^ 3
        Dim b_ As Double = Math.Exp(2.6219621 - 2.5091694 / Math.Sqrt(Zc) - 4.6725352 / Zc)
        Dim c_ As Double = ((3.2561755 - 0.060134019 * Zc + 1.2310844 * 10 ^ -3 * Zc ^ 2) / (1 - 1.7374036 * 10 ^ -2 * Zc + 2.524835 * 10 ^ -4 * Zc ^ 2)) ^ 2

        Dim U As Double = E0 / Ec
        Dim P_ As Double = R0 / (a_ + b_ / Math.Sqrt(U) + c_ / (U * Math.Sqrt(U)))


        Return P_

    End Function

    Public Function alpha_function(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double) As Double
        Dim alpha_total As Double
        Dim numerator As Double = 0
        Dim denominator As Double = 0

        Dim Ec As Double = studied_element.line(line_indice).Ec


        For i As Integer = 0 To UBound(layer_handler(mother_layer_id).element)
            Dim Z As Integer = layer_handler(mother_layer_id).element(i).z
            Dim A As Double = layer_handler(mother_layer_id).element(i).a
            Dim J As Double = ionization_potential_Sternheimer(Z)
            Dim C As Double = layer_handler(mother_layer_id).element(i).conc_wt

            Dim alpha_partial As Double = Math.Exp(12.93774 - 0.003426515 * Z) / (E0 ^ (1 / (1.139231 + 0.002775625 * Z))) * Math.Sqrt((Z / A) * Math.Log(1.166 * E0 / J) / (E0 ^ 2 - Ec ^ 2))

            Dim tmp As Double = C * Z / A
            numerator = numerator + tmp
            denominator = denominator + tmp * 1 / alpha_partial

        Next

        alpha_total = numerator / denominator

        Return alpha_total

    End Function

    Public Function ionization_potential_Sternheimer(ByRef Z As Integer) As Double 'in eV.

        Return (9.76 + 58.8 * Z ^ (-1.19)) * Z * 10 ^ -3

    End Function

    Public Sub line_coeff(ByVal x1 As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double, ByRef tangent As Double, ByRef intercept As Double)

        tangent = (y2 - y1) / (x2 - x1)
        'intercept = ((y1 + y2) - tangent * (x1 + x2)) / 2
        intercept = (y1 * x2 - y2 * x1) / (x2 - x1)

    End Sub

    Public Function FCN_function(ByVal phi0 As Double, ByVal rzm As Double, ByVal alpha As Double, ByVal beta As Double, ByVal FI As Double) As Double
        Dim a1 As Double = 0.254829592
        Dim a2 As Double = -0.284496736
        Dim a3 As Double = 1.421413741
        Dim a4 As Double = -1.453152027
        Dim a5 As Double = 1.061405429
        Dim t As Double = 1 / (1 + 0.3275911 * beta * rzm) 'AMXX check that!

        Dim P5 As Double = a1 * t + a2 * t ^ 2 + a3 * t ^ 3 + a4 * t ^ 4 + a5 * t ^ 5

        Return (Math.Exp((beta * rzm) ^ 2) - P5) / beta + Math.Exp((beta * rzm) ^ 2) / alpha - 2 * FI / (phi0 * Math.Sqrt(Math.PI))

    End Function

    Public Function FI_from_phirz_function(ByVal phi0 As Double, ByVal rzm As Double, ByVal alpha As Double, ByVal beta As Double) As Double

        Return phi0 * Math.Sqrt(Math.PI) / 2 * Math.Exp((beta * rzm) ^ 2) * (erf_Bastin(beta * rzm) / beta + 1 / alpha)

    End Function

    Public Function integrate_phirz_exp_murz(ByVal x1 As Double, ByVal x2 As Double, ByVal phim As Double, ByVal gamma As Double, ByVal rzm As Double, ByVal chi As Double) As Double

        Return Math.Sqrt(Math.PI) * phim * Math.Exp(1 / 4 * chi * (chi / gamma ^ 2 - 4 * rzm)) * (erf_Bastin(chi / (2 * gamma) + gamma * (x2 - rzm)) - erf_Bastin(chi / (2 * gamma) + gamma * (x1 - rzm))) / (2 * gamma)

    End Function

    Public Function erf_Bastin(ByVal x As Double) As Double
        Dim a1 As Double = 0.254829592
        Dim a2 As Double = -0.284496736
        Dim a3 As Double = 1.421413741
        Dim a4 As Double = -1.453152027
        Dim a5 As Double = 1.061405429
        Dim t As Double = 1 / (1 + 0.3275911 * x)

        Dim P5 As Double = a1 * t + a2 * t ^ 2 + a3 * t ^ 3 + a4 * t ^ 4 + a5 * t ^ 5

        erf_Bastin = 1 - P5 * Math.Exp(-x ^ 2)
        Return erf_Bastin
    End Function

    Public Sub Bastin_fluor(ByVal layer_handler() As layer, ByVal elt_exp_all() As Elt_exp, ByVal E1 As Double, ByVal chi_a As Double, ByVal ma As Integer, ByVal mb As Integer,
                    ByRef ffact As Double, ByVal F As Double, ByVal phi0 As Double, ByVal rzm As Double, ByVal alpha As Double, ByVal beta As Double, ByVal Rx As Double,
                            ByVal options As options, ByRef fit_MAC As fit_MAC)

        Dim de(UBound(layer_handler) + 1) As Double
        de(0) = 0
        For i As Integer = 0 To UBound(layer_handler)
            de(i + 1) = de(i) + layer_handler(i).mass_thickness
        Next

        Dim t(UBound(layer_handler)) As Double
        For kk As Integer = 0 To UBound(de) - 1
            t(kk) = de(kk + 1) - de(kk)
        Next
        t(UBound(layer_handler)) = 0.0

        Dim mu(UBound(layer_handler)) As Double
        For i As Integer = 0 To UBound(layer_handler)
            mu(i) = MAC_calculation(E1, i, layer_handler, elt_exp_all, fit_MAC, options)
        Next

        Dim a1 As Double
        Dim b1 As Double
        Dim b2 As Double
        Dim f1 As Double
        Dim f2 As Double
        Dim f3 As Double
        Dim f4 As Double
        Dim g1 As Double
        Dim g2 As Double
        Dim g3 As Double
        Dim g4 As Double
        Dim r1 As Double
        Dim r2 As Double

        Dim sum_mu_times_ti As Double = 0
        For i As Integer = ma To mb
            sum_mu_times_ti = sum_mu_times_ti + mu(i) * t(i)
        Next

        If ma > mb Then
            a1 = -chi_a / mu(ma) * mu(mb)
            b1 = chi_a / mu(ma) * (mu(mb) * de(mb) - sum_mu_times_ti) - chi_a * t(ma)
            b2 = -chi_a * t(ma)
            f1 = -mu(mb)
            f2 = mu(mb) * (chi_a / mu(ma) - 1)
            g1 = mu(mb) * de(mb) - sum_mu_times_ti - mu(ma) * t(ma)
            g2 = -(mu(mb) * de(mb) - sum_mu_times_ti) * (chi_a / mu(ma) - 1)
            g3 = g2 + (chi_a - mu(ma)) * t(ma)
            g4 = mu(mb) * de(mb) - sum_mu_times_ti

        ElseIf ma = mb Then
            'a1 = -chi_a
            'b1 = chi_a * de(ma)
            'b2 = -chi_a * t(ma)
            'f1 = -mu(ma)
            'f2 = chi_a - mu(ma)
            'f3 = mu(ma)
            'f4 = chi_a + mu(ma)
            'g1 = mu(ma) * de(ma)
            'g2 = (-chi_a + mu(ma)) * de(ma)
            'g3 = -mu(ma) * de(ma + 1)
            'g4 = -(chi_a + mu(ma)) * de(ma + 1)
            'r1 = Math.Log(1 + chi_a / mu(ma))
            'r2 = Math.Log(Math.Abs(1 - chi_a / mu(ma)))
            a1 = -chi_a
            b1 = chi_a * de(mb)
            b2 = -chi_a * t(mb)
            f1 = -mu(mb)
            f2 = chi_a - mu(mb)
            f3 = mu(mb)
            f4 = chi_a + mu(mb)
            g1 = mu(mb) * de(mb)
            g2 = (-chi_a + mu(mb)) * de(mb)
            g3 = -mu(mb) * de(mb + 1)
            g4 = -(chi_a + mu(mb)) * de(mb + 1)
            r1 = Math.Log(1 + chi_a / mu(ma))
            r2 = Math.Log(Math.Abs(1 - chi_a / mu(ma)))

        Else
            a1 = -chi_a / mu(ma) * mu(mb)
            b1 = chi_a / mu(ma) * (mu(mb) * de(mb + 1) + sum_mu_times_ti)
            b2 = -chi_a * t(ma)
            f1 = mu(mb)
            f2 = mu(mb) * (chi_a / mu(ma) + 1)
            g1 = -(mu(mb) * de(mb + 1) + sum_mu_times_ti)
            g2 = -(mu(mb) * de(mb + 1) + sum_mu_times_ti) * (chi_a / mu(ma) + 1)
            g3 = g2 - (chi_a + mu(ma)) * t(ma)
            g4 = -(mu(mb) * de(mb + 1) + sum_mu_times_ti) - mu(ma) * t(ma)

        End If

        Dim N As Integer = 999
        Dim x(N) As Double
        Dim y(N) As Double

        Dim de_min As Double = de(mb)
        Dim de_max As Double = Math.Min(de(mb + 1), Rx)
        For i As Integer = 0 To N
            x(i) = de_min + (de_max - de_min) / N * i
            y(i) = phi_Bastin(x(i), phi0, rzm, alpha, beta) * Y_function(x(i), layer_handler, ma, mb, a1, b1, b2, f1, f2, f3, f4, g1, g2, g3, g4, r1, r2, chi_a)

        Next

        ffact = 0
        For i As Integer = 0 To N - 1
            ffact = ffact + (y(i) + y(i + 1)) / 2 * (x(i + 1) - x(i))
        Next

    End Sub

    Public Function phi_Bastin(ByVal x As Double, ByVal phi0 As Double, ByVal rzm As Double, ByVal alpha As Double, ByVal beta As Double) As Double
        Dim phi_m As Double = phi0 * Math.Exp((beta * rzm) ^ 2)

        If x <= rzm Then
            phi_Bastin = phi_m * Math.Exp(-beta ^ 2 * (x - rzm) ^ 2)

        Else
            phi_Bastin = phi_m * Math.Exp(-alpha ^ 2 * (x - rzm) ^ 2)

        End If

        Return phi_Bastin

    End Function

    Public Function Y_function(ByVal x As Double, ByVal layer_handler() As layer, ByVal ma As Integer, ByVal mb As Integer, ByVal a1 As Double, ByVal b1 As Double, ByVal b2 As Double, ByVal f1 As Double, ByVal f2 As Double,
                               ByVal f3 As Double, ByVal f4 As Double, ByVal g1 As Double, ByVal g2 As Double, ByVal g3 As Double, ByVal g4 As Double, ByVal r1 As Double,
                               ByVal r2 As Double, ByVal chi_a As Double) As Double

        If ma > mb Then
            Return 1 / chi_a * (-expint(f1 * x + g1) - Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) - expint(f2 * x + g3)) + Math.Exp(b2) * expint(f1 * x + g4))

        ElseIf ma = mb Then
            Dim tmp_up As Double
            If ma = UBound(layer_handler) Then
                If x = 0 And (g1 = 0 Or g2 = 0) Then
                    tmp_up = Math.Log(f2 / f1) + r1 - r2
                    Return tmp_up
                Else
                    tmp_up = 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) + r1 - r2))
                    Return tmp_up
                End If
            Else
                If x = 0 And (g1 = 0 Or g2 = 0) Then
                    Stop
                    If b1 = 0 Then
                        Dim tmp As Double = Math.Log(f2 / f1)
                        tmp_up = Math.Log(f2 / f1) - r2
                    Else
                        Stop
                    End If
                Else
                    tmp_up = 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) - r2))
                End If
                ' tmp_up = 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) - r2))
                Dim tmp_down As Double = 1 / chi_a * (Math.Exp(b2) * expint(f3 * x + g3) - Math.Exp(a1 * x + b1) * (expint(f4 * x + g4) + r1))
                Return tmp_up + tmp_down

            End If

            'If x = 0 And (g1 = 0 Or g2 = 0) Then
            '    If b1 = 0 Then
            '        tmp_up = Math.Log(f2 / f1) - r2
            '    Else
            '        Stop
            '    End If
            'Else
            '    tmp_up = 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) - r2))
            'End If

            'Dim tmp_down As Double = 1 / chi_a * (Math.Exp(b2) * expint(f3 * x + g3) - Math.Exp(a1 * x + b1) * (expint(f4 * x + g4) + r1))
            'Return tmp_up + tmp_down

        Else
            Return 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) - expint(f2 * x + g3)) + Math.Exp(b2) * expint(f1 * x + g4))

        End If
        'Return 1 / chi_a * (-expint(f1 * x + g1) + Math.Exp(a1 * x + b1) * (expint(f2 * x + g2) + r1 - r2))

    End Function

End Module
