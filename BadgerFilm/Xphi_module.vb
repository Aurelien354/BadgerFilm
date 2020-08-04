Module Xphi_module
    Public Function phi0(ByVal radiation As String, ByVal Z As Double, ByVal E0 As Double, ByVal U0 As Double) As Double
        Dim m As Double = 1
        If radiation = "M" Then
            m = 0.8
        End If

        Dim a As Double = 1.87 * Z * (-0.00391 * Math.Log(Z + 1) + 0.00721 * (Math.Log(Z + 1) ^ 2) - 0.001067 * (Math.Log(Z + 1) ^ 3)) * (E0 / 30) ^ P(Z)

        Dim D1 As Double = 0.02 * Z '- m + 1 ' terms -m+1 present in Mikrochim. Acta (1992) 12 107-115 but not present in Mikrochim. Acta 114/115, 363-376 (1994) or in Lavrent'ev Journal of Analytical Chemistry, Vol. 59, No. 7, 2004, pp. 600–616
        Dim D2 As Double = 0.1 * Z '- m + 1
        Dim D3 As Double = 0.4 * Z '- m + 1

        Dim b1 As Double = (1 - (1 - 1 / U0 ^ D1) / (D1 * Math.Log(U0))) / D1
        Dim b2 As Double = (1 - (1 - 1 / U0 ^ D2) / (D2 * Math.Log(U0))) / D2
        Dim b3 As Double = (1 - (1 - 1 / U0 ^ D3) / (D3 * Math.Log(U0))) / D3

        phi0 = 1 + a * (0.27 * b1 + (1.1 + 5 / Z) * (b2 - 1.1 * b3))

    End Function

    Public Function P(ByVal z As Double)
        P = -0.25 + 0.0787 * z ^ 0.3
    End Function

    Public Function rzm(ByVal Z As Double, ByVal U0 As Double, ByVal rzx1 As Double) As Double
        rzm = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z)) / (1 + 1000 / (Z * U0 ^ 10))
    End Function

    Public Function rzx1(ByVal element() As Elt_layer, ByVal E0 As Double, ByVal Ec As Double) As Double

        Dim temp1 As Double = 0
        For i As Integer = 0 To UBound(element)

            temp1 = temp1 + element(i).conc_wt * element(i).z / element(i).a * rzx1_j(element(i).z, E0, Ec)
        Next

        Dim temp2 As Double = 0
        For i As Integer = 0 To UBound(element)
            temp2 = temp2 + element(i).conc_wt * element(i).z / element(i).a
        Next

        rzx1 = temp1 / temp2
    End Function

    Public Function rzx1_j(ByVal Z As Double, ByVal E0 As Double, ByVal Ec As Double) As Double
        Dim a As Double = 1.845 * 10 ^ -6 * (2.6 - 0.216 * E0 + 0.015 * E0 ^ 2 - 0.000387 * E0 ^ 3 + 0.00000501 * E0 ^ 4)
        Dim g As Double = 2.2 - 0.0166 * E0
        Dim f As Double = 1.2 - 0.04 * E0

        rzx1_j = a * (E0 ^ g - Ec ^ g) / (1.078 - 0.015 * Z ^ 0.7) ^ f '* (1 + 2 / E0 ^ 2) 'factor (1 + 2 / E0 ^ 2) from Mikrochim. Acta 114/115, 363-376 (1994)
        'Debug.Print(rzx1_j)
        'Debug.Print((1 + 2 / E0 ^ 2))
    End Function

    Public Function rzm_M(ByVal Z As Double, ByVal rzx1 As Double, ByVal U0 As Double) As Double
        rzm_M = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z))
    End Function

    Public Function Q(ByVal cst As Double, ByVal U As Double, ByVal El As Double, ByVal element As Elt_layer, ByVal radiation As String) As Double
        Dim m As Double
        If radiation = "K" Then
            m = 0.95
        Else
            m = 0.8 ' m = 0.7 in Mikrochim. Acta 114/115, 363-376 (1994), m = 0.8 in Microbeam Analysis 4, 239-253 (1995)
        End If
        'Q = cst * Math.Log(U) / U ^ m 'in Mikrochim. Acta 114/115, 363-376 (1994)
        Q = cst * Math.Log(U) / (El ^ 2 * U ^ m) 'in Microbeam Analysis 4, 239-253 (1995)
    End Function

    Public Function phi_rzm(ByVal element() As Elt_layer, ByVal radiation As String, ByVal E0 As Double, ByVal Ec As Double, ByVal rzm As Double, ByVal rzm_M As Double) As Double

        Dim temp1 As Double = 0
        For i As Integer = 0 To UBound(element)
            Dim phi_j_rzm_val As Double = phi_j_rzm(element(i), radiation, rzm, rzm_M, E0, Ec)
            temp1 = temp1 + element(i).conc_wt * element(i).z / element(i).a * phi_j_rzm_val
        Next

        Dim temp2 As Double = 0
        For i As Integer = 0 To UBound(element)
            temp2 = temp2 + element(i).conc_wt * element(i).z / element(i).a
        Next

        phi_rzm = temp1 / temp2

    End Function

    Public Function phi_j_rzm(ByVal element As Elt_layer, ByVal radiation As String, ByVal rzm As Double, ByVal rzm_M As Double, ByVal E0 As Double, ByVal Ec As Double) As Double
        Dim U0 As Double = E0 / Ec
        Dim Z As Double = element.z
        Dim A As Double = element.a

        Dim Ud_val As Double = Ud(rzm, rzm_M, Z, A, E0, U0)

        phi_j_rzm = phi_tau_j(rzm, Ud_val, E0, U0, Ec, element, radiation) + phi_eta_j(Ud_val, U0, Ec, element) +
            (phi_delta_j(Ud_val, U0, Ec, element, radiation) - phi_eta_j(Ud_val, U0, Ec, element)) * rzm / rzm_M

    End Function

    Public Function phi_tau_j(ByVal rzm As Double, ByVal Ud As Double, ByVal E0 As Double, ByVal U0 As Double, ByVal Ec As Double, ByVal element As Elt_layer,
                              ByVal radiation As String) As Double

        Dim Z As Double = element.z
        Dim A As Double = element.a

        Dim Rx_j As Double = rzx1_j(Z, E0, 0)

        Dim s1 As Double = 4.65 + 0.0356 * Z
        Dim s2 As Double = 1.112 + 0.00414 * Z ^ 2

        Dim tau1 As Double = (1 - rzm / Rx_j) ^ s1
        Dim tau2 As Double = (1 - rzm / Rx_j) ^ s2

        Dim tau_j As Double = tau1 + 4.65 * (tau2 - tau1) / (3.54 + 0.0356 * Z - 0.00414 * Z ^ 2)

        phi_tau_j = tau_j * Q(1, Ud, Ec, element, radiation) / Q(1, U0, Ec, element, radiation)
    End Function

    Public Function Ud(ByVal rzm As Double, ByVal rzm_M As Double, ByVal Z As Double, ByVal A As Double, ByVal E0 As Double, ByVal U0 As Double) As Double
        Dim Jj As Double = 0.00929 * (Z + 1.287 * Z ^ 0.33333)

        Ud = U0 * (1 - 7.85 * 10 ^ 4 * Z * (1 + 0.15 * rzm / rzm_M) * rzm / (E0 ^ 1.61 * A * Jj ^ 0.3)) ^ 0.61
    End Function

    Public Function phi_eta_j(ByVal Ud As Double, ByVal U0 As Double, ByVal Ec As Double, ByVal element As Elt_layer) As Double
        Dim Z As Double = element.z
        'Dim Ec As Double = element.Ec

        Dim D As Double = Z * (-0.00391 * Math.Log(Z + 1) + 0.00721 * Math.Log(Z + 1) ^ 2 - 0.001067 * Math.Log(Z + 1) ^ 3) * (Ud * Ec / 30) ^ P(Z)

        Dim d1 As Double = 0.02 * Z
        Dim d2 As Double = 0.1 * Z
        Dim d3 As Double = 0.4 * Z

        Dim b1 As Double = bi(d1, Ud, U0)
        Dim b2 As Double = bi(d2, Ud, U0)
        Dim b3 As Double = bi(d3, Ud, U0)

        phi_eta_j = D * (0.27 * b1 + (1.1 + 5 / Z) * (b2 - 1.1 * b3))

    End Function

    Public Function bi(ByVal di As Double, ByVal Ud As Double, ByVal U0 As Double) As Double
        bi = (Ud / U0) ^ di * (Math.Log(Ud) / Math.Log(U0)) * (1 / di) * (1 - (1 - 1 / Ud ^ di) / (di * Math.Log(Ud)))
    End Function

    Public Function phi_delta_j(ByVal Ud As Double, ByVal U0 As Double, ByVal Ec As Double, ByVal element As Elt_layer, ByVal radiation As String) As Double
        Dim Z As Double = element.z
        'Dim Ec As Double = element.Ec

        Dim m As Double
        If radiation = "K" Then
            m = 0.95
        Else
            m = 0.8
        End If

        Dim B As Double = 5 * (1 - 1 / (1 + Z) ^ 0.8) * (Ud * Ec / 30) ^ P(Z)

        Dim d1 As Double = 1 - m
        Dim d2 As Double = 7 - 4 * Math.Exp(-0.1 * Z) - m ' 7 - 4 * Math.Exp(0.1 * Z) - m in Lavrent'ev Journal of Analytical Chemistry, Vol. 59, No. 7, 2004, pp. 600–616

        phi_delta_j = B * (0.28 * (1 - 0.5 * Math.Exp(-0.1 * Z)) * bi(d1, Ud, U0) + 0.165 * Z ^ 0.6 * bi(d2, Ud, U0))

    End Function

    Public Function beta(ByVal rzm_val As Double, ByVal phi_rzm_val As Double, ByVal phi0_val As Double) As Double

        If phi_rzm_val < phi0_val Then
            phi_rzm_val = phi0_val * 1.00000001
            Debug.Print("Warning in XPHI caclulation of beta!")
        End If

        beta = rzm_val / Math.Sqrt(Math.Log(phi_rzm_val / phi0_val))
        'beta = rzm_val * Math.Log(phi_rzm_val) / Math.Sqrt(phi0_val) 'Llovet and Merlet, Microsc. Microanal. 16, 21–32, 2010

    End Function

    Public Function alpha(ByVal rzx1_val As Double, ByVal rzm_val As Double, ByVal phi_rzm_val As Double) As Double

        alpha = 0.46598 * (rzx1_val - rzm_val) 'Llovet and Merlet, Microsc. Microanal. 16, 21–32, 2010

        'alpha = (rzx1_val - rzm_val) / Math.Sqrt(Math.Log(phi_rzm_val / 0.01))

    End Function

    Public Sub xphi(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double, ByVal Ec As Double,
                    ByRef rzm_val As Double, ByRef rzx1_val As Double, ByRef phi_rzm_val As Double, ByRef phi0_val As Double, ByRef alpha_val As Double, ByRef beta_val As Double)

        Dim Z_bar As Double = 0
        Dim element As Elt_layer() = layer_handler(mother_layer_id).element
        For i As Integer = 0 To UBound(element)
            Z_bar = Z_bar + element(i).conc_wt * element(i).z
        Next

        rzx1_val = rzx1(element, E0, Ec) '* 0.975 '* 1.2
        'rzx1_val = rzx1(element, E0, Ec) * (1 + Z_bar * 2 / 100) '* 1.5

        rzm_val = rzm(Z_bar, E0 / Ec, rzx1_val)

        Dim rzm_M_val As Double = rzm_M(Z_bar, rzx1_val, E0 / Ec)

        phi_rzm_val = phi_rzm(element, studied_element.line(line_indice).xray_name(0), E0, Ec, rzm_val, rzm_M_val)

        phi0_val = phi0(studied_element.line(line_indice).xray_name(0), Z_bar, E0, E0 / Ec)

        beta_val = beta(rzm_val, phi_rzm_val, phi0_val)

        'rzx1_val = rzx1(element, E0, Ec) '* 0.975
        rzx1_val = rzx1(element, E0, Ec) * (1 + Z_bar * 2 / 100) '* 1.5

        Dim FI_val As Double = FI_XPHI_function(layer_handler, mother_layer_id, studied_element, line_indice, E0)
        Dim rzx_val As Double = rzm_val + 2.42147 * FI_val / phi_rzm_val - 2.145966 * beta_val * erf_Bastin(rzm_val / beta_val)
        rzx1_val = rzx_val
        'rzm_val = rzm(Z_bar, E0 / Ec, rzx1_val)

        'alpha_val = alpha(rzx1_val, rzm_val, phi_rzm_val)
        alpha_val = alpha(rzx_val, rzm_val, phi_rzm_val)


    End Sub

    Public Function FI_XPHI_function(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double) As Double

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
                J = J + .conc_wt * .z / .a * Math.Log(0.00929 * (.z + 1.287 * .z ^ 0.33333)) 'Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
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
            m_ = 0.8 'corr AMXX
        ElseIf (studied_element.line(line_indice).xray_name(0) = "M") Then
            m_ = 0.8 'corr AMXX
        ElseIf (studied_element.line(line_indice).xray_name(0) = "K") Then
            'm_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
            m_ = 0.95
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
    Public Sub multi_layer(ByVal layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal E0 As Double, ByVal Ec As Double,
                          ByRef rzm As Double, ByRef rzx1 As Double, ByRef phi_rzm As Double, ByRef phi0 As Double)

        Dim rzm_(UBound(layer_handler)) As Double
        Dim rzx1_(UBound(layer_handler)) As Double
        Dim phi_rzm_(UBound(layer_handler)) As Double
        Dim phi0_(UBound(layer_handler)) As Double
        Dim alpha As Double
        Dim beta As Double

        For i As Integer = 0 To UBound(layer_handler)
            xphi(layer_handler, i, studied_element, line_indice, E0, Ec, rzm_(i), rzx1_(i), phi_rzm_(i), phi0_(i), alpha, beta)
        Next

        Dim rzm_i_minus_1 As Double = rzm_(UBound(layer_handler))
        Dim rzx1_i_minus_1 As Double = rzx1_(UBound(layer_handler))
        Dim phi_rzm_i_minus_1 As Double = phi_rzm_(UBound(layer_handler))
        Dim phi0_i_minus_1 As Double = phi0_(UBound(layer_handler))

        'xphi(layer_handler(UBound(layer_handler)).element, radiation, E0, Ec, rzm_i_minus_1, rzx1_i_minus_1, phi_rzm_i_minus_1, phi0_i_minus_1, alpha, beta)


        'Dim test_rzx1 As Double = rzx1_i_minus_1

        If UBound(layer_handler) <> 0 Then ' Handle the case where the sample is a bulk sample
            'rzm_i_minus_1 = 0
            'rzx1_i_minus_1 = 0
            'phi_rzm_i_minus_1 = 0
            'phi0_i_minus_1 = 0
            For i As Integer = UBound(layer_handler) - 1 To 0 Step -1
                'Dim rzm_i As Double
                'Dim rzx1_i As Double
                'Dim phi_rzm_i As Double
                'Dim phi0_i As Double
                'Dim alpha_i As Double
                'Dim beta_i As Double

                Dim Z_Ln As Double = 0
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    Z_Ln = Z_Ln + layer_handler(i).element(j).conc_wt * layer_handler(i).element(j).z
                Next

                'xphi(layer_handler(i).element, radiation, E0, Ec, rzm_i, rzx1_i, phi_rzm_i, phi0_i, alpha_i, beta_i)

                Dim x As Double = 4 * layer_handler(i).mass_thickness / rzx1_(i) ' ???
                'x = layer_handler(i).mass_thickness * 10 ^ 6 / ((0.00682 * Z_Ln + 0.821) * E0 ^ (1.801 - 0.00318 * Z_Ln)) 'From H.-J. Hunger, S. Rogaschewski, SCANNING Vol. 8, 6 (1986) 
                'x = 1.3

                Dim lol As Double = rzm_(i) - rzm_i_minus_1
                Dim lol2 As Double = Math.Tanh(A(Z_Ln) * x + B(Z_Ln) * x ^ 2)
                rzm_i_minus_1 = (rzm_(i) - rzm_i_minus_1) * Math.Tanh(A(Z_Ln) * x + B(Z_Ln) * x ^ 2) + rzm_i_minus_1

                'x = 4 * layer_handler(i).mass_thickness / rzx1_i '
                rzx1_i_minus_1 = (rzx1_(i) - rzx1_i_minus_1) * Math.Tanh(A(Z_Ln) * x + B(Z_Ln) * x ^ 2) + rzx1_i_minus_1

                'x = 4 * layer_handler(i).mass_thickness / phi_rzm_i '
                phi_rzm_i_minus_1 = (phi_rzm_(i) - phi_rzm_i_minus_1) * Math.Tanh(A(Z_Ln) * x + B(Z_Ln) * x ^ 2) + phi_rzm_i_minus_1

                'x = 4 * layer_handler(i).mass_thickness / phi0_i '
                'x = layer_handler(i).mass_thickness * 10 ^ 6 / ((0.00682 * Z_Ln + 0.821) * E0 ^ (1.801 - 0.00318 * Z_Ln))
                phi0_i_minus_1 = (phi0_(i) - phi0_i_minus_1) * Math.Tanh(A(Z_Ln) * x + B(Z_Ln) * x ^ 2) + phi0_i_minus_1
            Next
            rzm = rzm_i_minus_1
            rzx1 = rzx1_i_minus_1
            phi_rzm = phi_rzm_i_minus_1
            phi0 = phi0_i_minus_1

        Else
            rzm = rzm_(0)
            rzx1 = rzx1_(0)
            phi_rzm = phi_rzm_(0)
            phi0 = phi0_(0)
        End If



    End Sub

    Public Function A(ByVal Zf As Double) As Double
        A = (2.153 * Zf - 14.789) / (3.706 * Zf + 17.822)
    End Function

    Public Function B(ByVal Zf As Double) As Double
        B = (0.3618 * Zf - 27.803) / (1.235 * Zf - 82.055)
    End Function

    Public Sub calc_multi_layer(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer, ByVal elt_exp_all() As Elt_exp,
                                ByVal E0 As Double, ByVal sin_toa_in_rad As Double, ByRef phi_rz As Double, ByRef rzm As Double, ByRef rzx1 As Double, ByRef phi_rzm As Double,
                                ByRef phi0 As Double, ByRef alpha_val As Double, ByRef beta_val As Double, ByVal options As options, Optional fit_MAC As fit_MAC = Nothing)
        'Dim rzm As Double
        'Dim rzx1 As Double
        'Dim phi_rzm As Double
        'Dim phi0 As Double
        'Dim radiation As String = studied_element.line(line_indice).xray_name(0)
        Dim Ec As Double = studied_element.line(line_indice).Ec

        multi_layer(layer_handler, mother_layer_id, studied_element, line_indice, E0, Ec, rzm, rzx1, phi_rzm, phi0)

        alpha_val = alpha(rzx1, rzm, phi_rzm)
        beta_val = beta(rzm, phi_rzm, phi0)

        '************************
        'Integrate the phi(rz) function
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
        'Calculate the 'wt.fraction averaged mass absorption coefficient'
        'for the current layer
        '************************
        Dim mac As Double
        Dim chi As Double
        mac = MAC_calculation(studied_element.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)

        chi = mac / sin_toa_in_rad 'Math.Sin(toa * Math.PI / 180)
        '************************


        Dim lim_min As Double
        Dim lim_max As Double
        If mother_layer_id = 0 Then
            lim_min = 0
        Else
            lim_min = cumulative_mass_depth(mother_layer_id - 1)
        End If
        lim_max = cumulative_mass_depth(mother_layer_id)
        If lim_max > rzx1 Then lim_max = rzx1

        Dim H1 As Double
        Dim H2 As Double

        If lim_min >= rzm Then
            H1 = 0
        Else
            If chi = 0 Then
                H1 = 0.5 * Math.Sqrt(Math.PI) * beta_val * phi_rzm * erf_Bastin((Math.Min(rzm, lim_max) - lim_min) / beta_val)
            Else
                'H1 = 0.5 * Math.Sqrt(Math.PI) * beta_val * phi_rzm * Math.Exp(0.25 * chi * (beta_val ^ 2 * chi - 4 * Math.Min(rzm, lim_max))) *
                '    (erf_Bastin(beta_val * chi / 2) + erf_Bastin((Math.Min(rzm, lim_max) - lim_min) / beta_val - beta_val * chi / 2))
                H1 = Math.Sqrt(Math.PI) / 2 * beta_val * phi_rzm * Math.Exp(0.25 * chi * (beta_val ^ 2 * chi - 4 * rzm)) *
                    (erf_Bastin(beta_val * chi / 2 + (Math.Min(rzm, lim_max) - rzm) / beta_val) - erf_Bastin((lim_min - rzm) / beta_val + beta_val * chi / 2))
            End If
        End If

        If lim_max <= rzm Then
            H2 = 0
        Else
            If chi = 0 Then
                H2 = 0.5 * Math.Sqrt(Math.PI) * alpha_val * phi_rzm * erf_Bastin((lim_max - rzm) / alpha_val)
            Else
                H2 = Math.Sqrt(Math.PI) / 2 * alpha_val * phi_rzm * Math.Exp(0.25 * chi * (alpha_val ^ 2 * chi - 4 * rzm)) *
                    (erf_Bastin(alpha_val * chi / 2 + (Math.Min(rzx1, lim_max) - rzm) / alpha_val) - erf_Bastin((Math.Max(lim_min, rzm) - rzm) / alpha_val + alpha_val * chi / 2))
            End If
        End If

        Dim H1_plus_H2_test As Double = H1 + H2
        If H1_plus_H2_test <= 0 Then
            H1_plus_H2_test = 0.00000000001
        End If
        Dim abs As Double = abs_outer_layers(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, mac, sin_toa_in_rad, fit_MAC, options)
        phi_rz = H1_plus_H2_test * abs


    End Sub

    'Public Sub plot_phi_rz(ByVal rzm As Double, ByVal rzx1 As Double, ByVal phi_rzm As Double, ByVal phi0 As Double, ByVal alpha_val As Double, ByVal beta_val As Double,
    '                       ByVal steps As Integer, ByRef chart1 As DataVisualization.Charting.Chart)
    '    Dim style As DataVisualization.Charting.SeriesChartType = DataVisualization.Charting.SeriesChartType.Line

    '    Dim x(steps) As Double
    '    Dim y(steps) As Double
    '    Dim rz As Double

    '    For i As Integer = 0 To steps
    '        rz = rzx1 / steps * i
    '        x(i) = rz * 10 ^ 3

    '        If rz <= rzm Then
    '            y(i) = phi_rzm * Math.Exp(-(rz - rzm) ^ 2 / beta_val ^ 2)
    '        Else
    '            y(i) = phi_rzm * Math.Exp(-(rz - rzm) ^ 2 / alpha_val ^ 2)
    '        End If
    '    Next

    '    graph_data_simple(x, y, chart1, "0.00", "0.00", 1, False, color_table(chart1.Series.Count Mod color_table.Count), "rz (µg/cm²)", "phi(rz)",
    '        False, "phi", style)

    '    Dim res As String = ""
    '    For i As Integer = 0 To UBound(x)
    '        res = res & x(i) & vbTab & y(i) & vbCrLf
    '    Next
    '    Clipboard.SetText(res)

    'End Sub



    'Public Sub Xphi_bulk(ByVal layer_handler As layer(), ByVal studied_element As element, ByVal E0 As Double, ByVal toa As Double,
    '                     ByRef phi0 As Double, ByRef phim As Double, ByRef rzm As Double, ByRef rzx As Double)

    '    Dim U0 As Double = E0 / studied_element.Ec

    '    'phi0
    '    Dim Z_bar As Double = 0
    '    For i As Integer = 0 To UBound(layer_handler(studied_element.mother_layer_id).element)
    '        Z_bar = Z_bar + layer_handler(studied_element.mother_layer_id).element(i).concentration * layer_handler(studied_element.mother_layer_id).element(i).z
    '    Next

    '    Dim m As Double = 1
    '    If studied_element.xray_name(0) Like "[Mm]" Then m = 0.8

    '    Dim D1 As Double
    '    Dim D2 As Double
    '    Dim D3 As Double
    '    D1 = 0.02 * Z_bar
    '    D2 = 0.1 * Z_bar
    '    D3 = 0.4 * Z_bar - m + 1

    '    Dim b1 As Double
    '    Dim b2 As Double
    '    Dim b3 As Double
    '    b1 = (1 - (1 - 1 / U0 ^ D1) / D1 * Math.Log(U0)) / D1
    '    b2 = (1 - (1 - 1 / U0 ^ D2) / D2 * Math.Log(U0)) / D2
    '    b3 = (1 - (1 - 1 / U0 ^ D3) / D3 * Math.Log(U0)) / D3

    '    Dim P_Z As Double
    '    Dim a_phi0 As Double
    '    P_Z = -0.25 + 0.0787 * Z_bar ^ 0.3
    '    a_phi0 = 1.87 * Z_bar * (-0.00391 * Math.Log(Z_bar + 1) + 0.00721 * (Math.Log(Z_bar + 1)) ^ 2 - 0.001067 * (Math.Log(Z_bar + 1) ^ 3)) * (E0 / 30) ^ P_Z

    '    phi0 = 1 + a_phi0 * (0.27 * b1 + (1.1 + 5 / Z_bar) * (b2 - 1.1 * b3))


    '    'rzx1
    '    Dim f As Double
    '    Dim g As Double
    '    Dim a_rzx1 As Double
    '    f = 1.2 - 0.04 * E0
    '    g = 2.2 - 0.0166 * E0
    '    a_rzx1 = 1.845 * 10 ^ -6 * (2.6 - 0.216 * E0 + 0.015 * E0 ^ 2 - 0.000387 * E0 ^ 3 + 0.00000501 * E0 ^ 4)

    '    Dim rzx1 As Double = 0

    '    For i As Integer = 0 To UBound(layer_handler(studied_element.mother_layer_id).element)
    '        Dim rzx1_j As Double
    '        rzx1_j = a_rzx1 * (E0 ^ g - studied_element.Ec ^ g) / (1.078 - 0.015 * layer_handler(studied_element.mother_layer_id).element(i).z ^ 0.7) ^ f

    '        rzx1 = rzx1 + layer_handler(studied_element.mother_layer_id).element(i).concentration * layer_handler(studied_element.mother_layer_id).element(i).z /
    '            layer_handler(studied_element.mother_layer_id).element(i).a * rzx1_j
    '    Next

    '    Dim multicomp_denominator As Double = 0
    '    For i As Integer = 0 To UBound(layer_handler(studied_element.mother_layer_id).element)
    '        multicomp_denominator = multicomp_denominator + layer_handler(studied_element.mother_layer_id).element(i).concentration * layer_handler(studied_element.mother_layer_id).element(i).z /
    '            layer_handler(studied_element.mother_layer_id).element(i).a
    '    Next

    '    rzx1 = rzx1 / multicomp_denominator

    '    'rzm
    '    rzm = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z_bar)) / (1 + 1000 / (Z_bar * U0 ^ 10))

    '    'rzm_M
    '    Dim rzm_M As Double
    '    rzm_M = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z_bar))




    'End Sub


    'Public Function phi0(ByVal radiation As String, ByVal Z As Double, ByVal E0 As Double, ByVal U0 As Double) As Double
    '    Dim m As Double = 1
    '    If radiation = "M" Then
    '        m = 0.8
    '    End If

    '    Dim a As Double = 1.87 * Z * (-0.00391 * Math.Log(Z + 1) + 0.00721 * (Math.Log(Z + 1) ^ 2) - 0.001067 * (Math.Log(Z + 1) ^ 3)) * (E0 / 30) ^ P(Z)

    '    Dim D1 As Double = 0.02 * Z - m + 1
    '    Dim D2 As Double = 0.1 * Z - m + 1
    '    Dim D3 As Double = 0.4 * Z - m + 1

    '    Dim b1 As Double = (1 - (1 - 1 / U0 ^ D1) / (D1 * Math.Log(U0))) / D1
    '    Dim b2 As Double = (1 - (1 - 1 / U0 ^ D2) / (D2 * Math.Log(U0))) / D2
    '    Dim b3 As Double = (1 - (1 - 1 / U0 ^ D3) / (D3 * Math.Log(U0))) / D3

    '    phi0 = 1 + a * (0.27 * b1 + (1.1 + 5 / Z) * (b2 - 1.1 * b3))

    'End Function

    'Public Function P(ByVal z As Double)
    '    P = -0.25 + 0.0787 * z ^ 0.3
    'End Function

    'Public Function rzm(ByVal Z As Double, ByVal U0 As Double, ByVal rzx1 As Double) As Double
    '    rzm = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z)) / (1 + 1000 / (Z * U0 ^ 10))
    'End Function

    'Public Function rzx1(ByVal element() As element, ByVal E0 As Double, ByVal Ec As Double) As Double

    '    Dim temp1 As Double = 0
    '    For i As Integer = 0 To UBound(element)

    '        temp1 = temp1 + element(i).concentration * element(i).z / element(i).a * rzx1_j(element(i).z, E0, Ec)
    '    Next

    '    Dim temp2 As Double = 0
    '    For i As Integer = 0 To UBound(element)
    '        temp2 = temp2 + element(i).concentration * element(i).z / element(i).a
    '    Next

    '    rzx1 = temp1 / temp2
    'End Function

    'Public Function rzx1_j(ByVal Z As Double, ByVal E0 As Double, ByVal Ec As Double) As Double
    '    Dim a As Double = 1.845 * 10 ^ -6 * (2.6 - 0.216 * E0 + 0.015 * E0 ^ 2 - 0.000387 * E0 ^ 3 + 0.00000501 * E0 ^ 4)
    '    Dim g As Double = 2.2 - 0.0166 * E0
    '    Dim f As Double = 1.2 - 0.04 * E0

    '    rzx1_j = a * (E0 ^ g - Ec ^ g) / (1.078 - 0.015 * Z ^ 0.7) ^ f * (1 + 2 / E0 ^ 2) 'factor (1 + 2 / E0 ^ 2) from Mikrochim. Acta 114/115, 363-376 (1994)
    '    Debug.Print(rzx1_j)
    '    Debug.Print((1 + 2 / E0 ^ 2))
    'End Function

    'Public Function rzm_M(ByVal Z As Double, ByVal rzx1 As Double, ByVal U0 As Double) As Double
    '    rzm_M = rzx1 * (0.1 + 0.35 * Math.Exp(-0.07 * Z))
    'End Function

    'Public Function Q(ByVal cst As Double, ByVal U As Double, ByVal element As element, ByVal radiation As String) As Double
    '    Dim m As Double
    '    If radiation = "K" Then
    '        m = 0.95
    '    Else
    '        m = 0.8 ' m = 0.7 in Mikrochim. Acta 114/115, 363-376 (1994)
    '    End If
    '    Q = cst * Math.Log(U) / U ^ m
    'End Function

    'Public Function phi_rzm(ByVal element() As element, ByVal radiation As String, ByVal E0 As Double, ByVal Ec As Double, ByVal rzm As Double, ByVal rzm_M As Double) As Double

    '    Dim temp1 As Double = 0
    '    For i As Integer = 0 To UBound(element)
    '        Dim phi_j_rzm_val As Double = phi_j_rzm(element(i), radiation, rzm, rzm_M, E0, Ec)
    '        temp1 = temp1 + element(i).concentration * element(i).z / element(i).a * phi_j_rzm_val
    '    Next

    '    Dim temp2 As Double = 0
    '    For i As Integer = 0 To UBound(element)
    '        temp2 = temp2 + element(i).concentration * element(i).z / element(i).a
    '    Next

    '    phi_rzm = temp1 / temp2

    'End Function

    'Public Function phi_j_rzm(ByVal element As element, ByVal radiation As String, ByVal rzm As Double, ByVal rzm_M As Double, ByVal E0 As Double, ByVal Ec As Double) As Double
    '    Dim U0 As Double = E0 / Ec
    '    Dim Z As Double = element.z
    '    Dim A As Double = element.a

    '    Dim Ud_val As Double = Ud(rzm, rzm_M, Z, A, E0, U0)

    '    phi_j_rzm = phi_tau_j(rzm, Ud_val, E0, U0, element, radiation) + phi_eta_j(Ud_val, U0, Ec, element) +
    '        (phi_delta_j(Ud_val, U0, Ec, element, radiation) - phi_eta_j(Ud_val, U0, Ec, element)) * rzm / rzm_M

    'End Function

    'Public Function phi_tau_j(ByVal rzm As Double, ByVal Ud As Double, ByVal E0 As Double, ByVal U0 As Double, ByVal element As element,
    '                          ByVal radiation As String) As Double

    '    Dim Z As Double = element.z
    '    Dim A As Double = element.a

    '    Dim Rx_j As Double = rzx1_j(Z, E0, 0)

    '    Dim s1 As Double = 4.65 + 0.0356 * Z
    '    Dim s2 As Double = 1.112 + 0.00414 * Z ^ 2

    '    Dim tau1 As Double = (1 - rzm / Rx_j) ^ s1
    '    Dim tau2 As Double = (1 - rzm / Rx_j) ^ s2

    '    Dim tau_j As Double = tau1 + 4.65 * (tau2 - tau1) / (3.54 + 0.0356 * Z - 0.00414 * Z ^ 2)

    '    phi_tau_j = tau_j * Q(1, Ud, element, radiation) / Q(1, U0, element, radiation)
    'End Function

    'Public Function Ud(ByVal rzm As Double, ByVal rzm_M As Double, ByVal Z As Double, ByVal A As Double, ByVal E0 As Double, ByVal U0 As Double) As Double
    '    Dim Jj As Double = 0.00929 * (Z + 1.287 * Z ^ 0.33333)

    '    Ud = U0 * (1 - 7.85 * 10 ^ 4 * Z * (1 + 0.15 * rzm / rzm_M) * rzm / (E0 ^ 1.61 * A * Jj ^ 0.3)) ^ 0.61
    'End Function

    'Public Function phi_eta_j(ByVal Ud As Double, ByVal U0 As Double, ByVal Ec As Double, ByVal element As element) As Double
    '    Dim Z As Double = element.z
    '    'Dim Ec As Double = element.Ec

    '    Dim D As Double = Z * (-0.00391 * Math.Log(Z + 1) + 0.00721 * Math.Log(Z + 1) ^ 2 - 0.001067 * Math.Log(Z + 1) ^ 3) * (Ud * Ec / 30) ^ P(Z)

    '    Dim d1 As Double = 0.02 * Z
    '    Dim d2 As Double = 0.1 * Z
    '    Dim d3 As Double = 0.4 * Z

    '    Dim b1 As Double = bi(d1, Ud, U0)
    '    Dim b2 As Double = bi(d2, Ud, U0)
    '    Dim b3 As Double = bi(d3, Ud, U0)

    '    phi_eta_j = D * (0.27 * b1 + (1.1 + 5 / Z) * (b2 - 1.1 * b3))

    'End Function

    'Public Function bi(ByVal di As Double, ByVal Ud As Double, ByVal U0 As Double) As Double
    '    bi = (Ud / U0) ^ di * (Math.Log(Ud) / Math.Log(U0)) * (1 / di) * (1 - (1 - 1 / Ud ^ di) / (di * Math.Log(Ud)))
    'End Function

    'Public Function phi_delta_j(ByVal Ud As Double, ByVal U0 As Double, ByVal Ec As Double, ByVal element As element, ByVal radiation As String) As Double
    '    Dim Z As Double = element.z
    '    'Dim Ec As Double = element.Ec

    '    Dim m As Double
    '    If radiation = "K" Then
    '        m = 0.95
    '    Else
    '        m = 0.8
    '    End If

    '    Dim B As Double = 5 * (1 - 1 / (1 + Z) ^ 0.8) * (Ud * Ec / 30) ^ P(Z)

    '    Dim d1 As Double = 1 - m
    '    Dim d2 As Double = 7 - 4 * Math.Exp(-0.1 * Z) - m

    '    phi_delta_j = B * (0.28 * (1 - 0.5 * Math.Exp(-0.1 * Z)) * bi(d1, Ud, U0) + 0.165 * Z ^ 0.6 * bi(d2, Ud, U0))

    'End Function

    'Public Function beta(ByVal rzm_val As Double, ByVal phi_rzm_val As Double, ByVal phi0_val As Double) As Double

    '    beta = rzm_val / Math.Sqrt(Math.Log(phi_rzm_val / phi0_val))

    '    'beta = rzm_val * Math.Log(phi_rzm_val) / Math.Sqrt(phi0_val)

    'End Function

    'Public Function alpha(ByVal rzx1_val As Double, ByVal rzm_val As Double, ByVal phi_rzm_val As Double) As Double

    '    alpha = 0.46598 * (rzx1_val - rzm_val)

    '    'alpha = (rzx1_val - rzm_val) / Math.Sqrt(Math.Log(phi_rzm_val / 0.01))

    'End Function

    'Public Sub xphi(ByVal element() As element, ByVal radiation As String, ByVal E0 As Double, ByVal Ec As Double, ByRef rzm_val As Double, ByRef rzx1_val As Double,
    '                ByRef phi_rzm_val As Double, ByRef phi0_val As Double, ByRef alpha_val As Double, ByRef beta_val As Double)

    '    Dim Z_bar As Double = 0
    '    For i As Integer = 0 To UBound(element)
    '        Z_bar = Z_bar + element(i).concentration * element(i).z
    '    Next

    '    rzx1_val = rzx1(element, E0, Ec) '* 1.2

    '    rzm_val = rzm(Z_bar, E0 / Ec, rzx1_val)

    '    Dim rzm_M_val As Double = rzm_M(Z_bar, rzx1_val, E0 / Ec)

    '    phi_rzm_val = phi_rzm(element, radiation, E0, Ec, rzm_val, rzm_M_val)

    '    phi0_val = phi0(radiation, Z_bar, E0, E0 / Ec)

    '    beta_val = beta(rzm_val, phi_rzm_val, phi0_val)

    '    rzx1_val = rzx1(element, E0, Ec) * (1 + Z_bar * 2 / 100)
    '    alpha_val = alpha(rzx1_val, rzm_val, phi_rzm_val)


    'End Sub



End Module
