Imports System.IO

Module XPP_module
    Public Function my_xpp(ByRef layer_handler() As layer, ByVal mother_layer_id As Integer, ByVal studied_element As Elt_exp, ByVal line_indice As Integer,
                      ByVal elt_exp_all() As Elt_exp, ByVal E0 As Single, ByVal sin_toa_in_rad As Single, ByRef phi_rz As Single, ByRef F As Single, ByRef phi0 As Single,
                      ByRef R_bar As Double, ByRef P As Single, ByRef A_XPP As Single, ByRef B_XPP As Single, fit_MAC As fit_MAC,
                      ByVal options As options) As Integer
        Try
            If layer_handler.Count > 1 Then
                MsgBox("XPP only works for bulk samples.")
                Return -1
            End If

            Dim El As Single = studied_element.line(line_indice).Ec
            'El = 1.84
            Dim U0 As Single = E0 / El


            '************************
            'Calculate the 'wt.fraction averaged mass absorption coefficient'
            'for the current layer
            '************************
            Dim mac As Single
            Dim chi As Single
            'mac = MAC_calculation(studied_element.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
            mac = MAC_calculation(studied_element, line_indice, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
            'Debug.Print(mac)
            chi = mac / sin_toa_in_rad
            '************************

            Dim Zb_bar As Single = 0
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    Zb_bar = Zb_bar + layer_handler(i).element(k).z ^ 0.5 * CSng(layer_handler(i).element(k).conc_wt)
                Next
            Next
            Zb_bar = Zb_bar ^ 2
            '************************

            '************************
            'Calculate the BSE factor R
            '************************
            Dim eta_bar As Single
            Dim W_bar As Single
            Dim q_ As Single
            Dim J_U0 As Single
            Dim G_U0 As Single
            Dim R As Single
            Dim r_ As Single

            eta_bar = 0.00175 * Zb_bar + 0.37 * (1 - Math.Exp(-0.015 * (Zb_bar ^ 1.3)))
            W_bar = 0.595 + eta_bar / 3.7 + eta_bar ^ 4.55
            q_ = (2 * W_bar - 1) / (1 - W_bar)
            J_U0 = 1 + U0 * (Math.Log(U0) - 1)
            G_U0 = (U0 - 1 - ((1 - (1 / U0 ^ (q_ + 1))) / (1 + q_))) / ((2 + q_) * J_U0)
            R = 1 - eta_bar * W_bar * (1 - G_U0)
            '************************

            '************************
            'Calculate the mean ionization potential J
            'Calculate M
            '************************
            Dim J As Single
            Dim M As Single
            J = 0
            M = 0
            For i As Integer = 0 To UBound(layer_handler)
                For k As Integer = 0 To UBound(layer_handler(i).element)
                    With layer_handler(i).element(k)
                        M = M + .z / .a * .conc_wt
                        J = J + .conc_wt * .z / .a * Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22)))
                    End With
                Next
            Next
            J = Math.Exp(J / M)
            '************************

            '************************
            'Calculate m_
            '************************
            Dim m_ As Single
            If (studied_element.line(line_indice).xray_name(0) = "L") Then
                m_ = 0.82 'corr AMXX
            ElseIf (studied_element.line(line_indice).xray_name(0) = "M") Then
                m_ = 0.78 'corr AMXX
            ElseIf (studied_element.line(line_indice).xray_name(0) = "K") Then
                m_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
            Else
                'Do not handle other X-ray lines.
                phi_rz = 0
                Return -1
            End If
            '************************

            '************************
            'Calculate 1/S (inv_S_total)
            '************************
            Dim Pi(2) As Single
            Dim Di(2) As Single
            Pi(0) = 0.78
            Pi(1) = 0.1
            Pi(2) = -(0.5 - 0.25 * J)
            Di(0) = 6.6 * 10 ^ -6
            Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
            Di(2) = 2.2 * 10 ^ -6 / J

            Dim inv_S As Single = 0
            Dim inv_S_total As Single = 0
            Dim V0 As Single = E0 / J
            For k As Integer = 0 To 2
                Dim Tk As Single = 1 + Pi(k) - m_
                inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
            Next
            inv_S_total = inv_S * U0 / (V0 * M)

            inv_S = 0
            For k As Integer = 0 To 2
                Dim Tk As Single = 1 + Pi(k) - 30
                inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
            Next
            inv_S_total = inv_S_total + 1.5 * inv_S * U0 / (V0 * M)

            inv_S = 0
            For k As Integer = 0 To 2
                Dim Tk As Single = 1 + Pi(k) - 5
                inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
            Next
            inv_S_total = inv_S_total + 0.1 * inv_S * U0 / (V0 * M)

            '************************
            Dim test1 As Single = R * inv_S_total
            '************************
            'Calculate ionization cross section QA_l
            '************************
            Dim QA_l As Single
            QA_l = Math.Log(U0) / El ^ 2 / U0 ^ m_

            '************************
            'Calculate the area F
            '************************
            F = R * inv_S_total / QA_l

            '************************
            'Calculate the surface ionization phi0
            '************************
            r_ = 2 - 2.3 * eta_bar
            phi0 = 1 + 3.3 * (1 - (1 / U0 ^ r_)) * eta_bar ^ 1.2

            '************************
            'Calculate the XPP factors
            '************************

            Dim X As Single = 1 + 1.3 * Math.Log(Zb_bar)
            Dim Y As Single = 0.2 + Zb_bar / 200

            Dim ak1 As Single = 1 + (X * Math.Log(1 + Y * (1 - 1 / U0 ^ 0.42))) / Math.Log(1 + Y) 'ak1=F/R_bar
            R_bar = F / ak1 ' (1 + (X * Math.Log(1 + Y * (1 - 1 / U0 ^ 0.42))) / Math.Log(1 + Y))


            Dim h_xpp As Single = 1 - 10 * (1 - 1 / (1 + U0 / 10)) / Zb_bar ^ 2 'cleg

            If ak1 < phi0 Then
                ak1 = phi0
            End If

            Dim g As Single = 0.22 * Math.Log(4 * Zb_bar) * (1 - 2 * Math.Exp(-Zb_bar * (U0 - 1) / 15))

            Dim lim1 As Single = g * h_xpp ^ 4 'ak2

            Dim a_ As Single
            Dim b_ As Single
            b_ = Math.Sqrt(2) * (1 + Math.Sqrt(1 - R_bar * phi0 / F)) / R_bar 'g=b_*R_bar
            Dim g_ As Single = b_ * R_bar
            Dim lim2 As Single = 0.9 * b_ * R_bar ^ 2 * (b_ - 2 * phi0 / F)

            If lim1 > lim2 Then
                lim1 = lim2
            End If
            P = lim1 * F / R_bar ^ 2

            a_ = (P + b_ * (2 * phi0 - b_ * F)) / (b_ * F * (2 - b_ * R_bar) - phi0)

            Dim epsilon As Single
            Dim sign As Integer = 1
            epsilon = (a_ - b_) / b_
            If epsilon < 0 Then sign = -1
            If Math.Abs(epsilon) < 0.000001 Then
                epsilon = epsilon * sign
                a_ = b_ * (1 + epsilon)
            End If

            B_XPP = (b_ ^ 2 * F * (1 + epsilon) - P - phi0 * b_ * (2 + epsilon)) / epsilon

            'A_XPP = (B_XPP / b_ + phi0 - b_ * F) * (1 + epsilon) / epsilon
            A_XPP = (B_XPP / b_ - P / b_ - phi0) / epsilon
            'chi = 2264

            phi_rz = (phi0 + B_XPP / (b_ + chi) - A_XPP * b_ * epsilon / (b_ * (1 + epsilon) + chi)) / (b_ + chi)

            Return 0

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in my_xpp " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try


        '        '************************
        '        'Normalize the concentrations as if the whole sample was a homogeneous bulk sample.
        '        'Fictitious composition stored in the layer_handler.
        '        '************************
        '        Dim total_concentration As Double = 0
        '        For i As Integer = 0 To UBound(layer_handler)
        '            For k As Integer = 0 To UBound(layer_handler(i).element)
        '                'layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).conc_wt ' line removed AMXX 22/11/2019
        '                total_concentration = total_concentration + layer_handler(i).element(k).conc_wt
        '            Next

        '        Next

        '        For i As Integer = 0 To UBound(layer_handler)
        '            For k As Integer = 0 To UBound(layer_handler(i).element)
        '                layer_handler(i).element(k).fictitious_concentration = layer_handler(i).element(k).conc_wt / total_concentration
        '            Next
        '        Next
        '        '************************

        '        Dim Rx As Double
        '        Dim Rx_old As Double
        '        Const RX_CONVERGENCE As Double = 0.01
        '        Dim flag_first_iter = True

        '        Dim M As Double = 0
        '        Dim Zn_bar As Double = 0
        '        Dim J As Double = 0
        '        Dim Pi(2) As Double
        '        Dim Di(2) As Double

        '        Z_bar = 0

        '        While (Math.Abs(Rx - Rx_old) / Rx > RX_CONVERGENCE Or flag_first_iter = True)
        '            'If flag_first_iter = True And layer_handler.Count > 1 Then
        '            '    flag_first_iter = False
        '            '    Rx = 7 * (E0 ^ 1.7 * El ^ 1.7) * (1 + 3 / El ^ 0.5 / (E0 / El + 0.3) ^ 2) * 10 ^ -6
        '            '    'AM XX 06-09-2018 (Septembre)
        '            '    Call pap_weight(layer_handler, 0.8 * Rx, -0.4 * 0.8 * Rx)
        '            '    Continue While
        '            'End If

        '            flag_first_iter = False
        '            Rx_old = Rx
        '            '************************
        '            'Calculate M
        '            'Calculate Zn bar
        '            'Calculate Z bar
        '            'Calculate J
        '            '************************
        '            M = 0
        '            Zn_bar = 0
        '            Z_bar = 0
        '            J = 0
        '            For i As Integer = 0 To UBound(layer_handler)
        '                For k As Integer = 0 To UBound(layer_handler(i).element)
        '                    With layer_handler(i).element(k)
        '                        M = M + .z / .a * .fictitious_concentration
        '                        Zn_bar = Zn_bar + Math.Log(.z) * .fictitious_concentration
        '                        Z_bar = Z_bar + .z * .fictitious_concentration
        '                        J = J + .fictitious_concentration * .z / .a * Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
        '                    End With
        '                Next
        '            Next
        '            Zn_bar = Math.Exp(Zn_bar)
        '            J = Math.Exp(J / M)
        '            '************************


        '            '************************
        '            'Calculate Pi
        '            'Calculate Di
        '            '************************
        '            Pi(0) = 0.78
        '            Pi(1) = 0.1
        '            Pi(2) = -(0.5 - 0.25 * J)
        '            Di(0) = 6.6 * 10 ^ -6
        '            Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
        '            Di(2) = 2.2 * 10 ^ -6 / J
        '            '************************

        '            '************************
        '            'Calculate R0 from p.61 (green book)
        '            '************************
        '            Dim R0 As Double = 0
        '            For k As Integer = 0 To 2
        '                R0 = R0 + J ^ (1 - Pi(k)) * Di(k) * (E0 ^ (1 + Pi(k)) - El ^ (1 + Pi(k))) * 1 / (1 + Pi(k))
        '            Next
        '            R0 = R0 / M
        '            '************************

        '            '************************
        '            'Calculate Q0
        '            'Calculate b
        '            'Calculate Q
        '            'Calculate h
        '            'Calculate D
        '            'Calculate Rx
        '            '************************
        '            Dim Q0 As Double
        '            Dim b As Double
        '            Dim Q As Double
        '            Dim h As Double
        '            Dim D As Double
        '            Q0 = 1 - 0.535 * Math.Exp(-(21 / Zn_bar) ^ 1.2) - 0.00025 * (Zn_bar / 20) ^ 3.5
        '            b = 40 / Z_bar
        '            Q = Q0 + (1 - Q0) * Math.Exp(-(U0 - 1) / b)
        '            h = Z_bar ^ 0.45
        '            D = 1 + 1 / (U0 ^ h)
        '            Rx = Q * D * R0
        '            '************************

        '            '************************
        '            'Iterate to calculate fictitious concentrations
        '            'in order to determine Rx (only in case of multilayer specimen?).
        '            '************************
        '            'If (Math.Abs(Rx * 1000000.0 - rt * 1000000.0) > 0.02) Then 'not the same condition as described by Pouchou and Pichoir p.54
        '            'Stop
        '            If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
        '                'Call pap_weight(layer_handler, 0.8 * Rx, -0.4 * 0.8 * Rx)
        '                Call pap_weight(layer_handler, Rx, -0.4 * Rx)
        '            Else
        '                Exit While
        '            End If

        '            'rt = Rx
        '            ' flag_first_iter = False
        '            'End If
        '            '************************
        '        End While



        '        '************************
        '        'Calculate other fictitious concentrations
        '        'in order to determine Zb bar (only in case of multilayer specimen).
        '        '************************
        '        If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
        '            'Call pap_weight(layer_handler, 0.5 * Rx, -0.1 * (0.5 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '            Call pap_weight(layer_handler, 0.5 * Rx, -0.4 * (0.5 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        '        End If

        '        Dim Zb_bar As Double = 0
        '200:
        '        For i As Integer = 0 To UBound(layer_handler)
        '            For k As Integer = 0 To UBound(layer_handler(i).element)
        '                Zb_bar = Zb_bar + layer_handler(i).element(k).z ^ 0.5 * layer_handler(i).element(k).fictitious_concentration
        '            Next
        '        Next
        '        Zb_bar = Zb_bar ^ 2
        '        '************************

        '        '************************
        '        'Calculate parameters
        '        'Eta bar
        '        'W bar
        '        'q
        '        'J(U0)
        '        'G(U0)
        '        'R
        '        'r
        '        'phi(0)
        '        '************************
        '        Dim eta_bar As Double
        '        Dim W_bar As Double
        '        Dim q_ As Double
        '        Dim J_U0 As Double
        '        Dim G_U0 As Double
        '        Dim R As Double
        '        Dim r_ As Double
        '        'Dim phi0 As Double

        '        eta_bar = 0.00175 * Zb_bar + 0.37 * (1 - Math.Exp(-0.015 * (Zb_bar ^ 1.3)))
        '        W_bar = 0.595 + eta_bar / 3.7 + eta_bar ^ 4.55
        '        q_ = (2 * W_bar - 1) / (1 - W_bar)
        '        J_U0 = 1 + U0 * (Math.Log(U0) - 1)
        '        G_U0 = (U0 - 1 - ((1 - (1 / U0 ^ (q_ + 1))) / (1 + q_))) / ((2 + q_) * J_U0)
        '        R = 1 - eta_bar * W_bar * (1 - G_U0)
        '        r_ = 2 - 2.3 * eta_bar

        '        phi0 = 1 + 3.3 * (1 - (1 / U0 ^ r_)) * eta_bar ^ 1.2 '* ((1 - 0.6) / (1 - 100) * Zb_bar + (1 * 100 - 0.6) / (100 - 1)) 'AMXXX
        '        '************************

        '        '************************
        '        'Calculate fictitious concentrations
        '        'in order to determine fictitious Z bar (only in case of multilayer specimen?).
        '        '************************
        '        If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
        '            'Call pap_weight(layer_handler, 0.65 * Rx, -0.6 * (0.65 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!0.7
        '            Call pap_weight(layer_handler, 0.7 * Rx, -0.6 * (0.7 * Rx)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!0.7
        '        End If
        '        '************************

        '        '************************
        '        'With the fictitious concentrations
        '        'Calculate M
        '        'Calculate Z bar
        '        'Calculate J
        '        '************************
        '        Z_bar = 0
        '        J = 0
        '        M = 0
        '        For i As Integer = 0 To UBound(layer_handler)
        '            For k As Integer = 0 To UBound(layer_handler(i).element)
        '                With layer_handler(i).element(k)
        '                    Z_bar = Z_bar + .z * .fictitious_concentration
        '                    M = M + .z / .a * .fictitious_concentration
        '                    J = J + .fictitious_concentration * .z / .a * Math.Log(.z * 10 ^ -3 * (10.04 + 8.25 * Math.Exp(- .z / 11.22))) 'change the 10^-3
        '                End With
        '            Next
        '        Next
        '        J = Math.Exp(J / M)
        '        '************************

        '        '************************
        '        'Calculate m_
        '        '************************
        '        Dim m_ As Double
        '        If (studied_element.line(line_indice).xray_name(0) = "L") Then
        '            m_ = 0.82 'corr AMXX
        '        ElseIf (studied_element.line(line_indice).xray_name(0) = "M") Then
        '            m_ = 0.78 'corr AMXX
        '        ElseIf (studied_element.line(line_indice).xray_name(0) = "K") Then
        '            m_ = 0.86 + 0.12 * Math.Exp(-(studied_element.z / 5) ^ 2)
        '        Else
        '            'Do not handle other X-ray lines.
        '            phi_rz = 0
        '            PAP_Rx = 0
        '            PAP_Rm = 0
        '            PAP_Rc = 0
        '            PAP_A1 = 0
        '            PAP_A2 = 0
        '            PAP_B1 = 0
        '            Exit Sub
        '        End If
        '        '************************

        '        '************************
        '        'Because now J can be different
        '        'Recalculate P3
        '        'Recalculate D2
        '        'Recalculate D3
        '        '************************
        '        Pi(2) = -(0.5 - 0.25 * J)
        '        Di(1) = 1.12 * 10 ^ -5 * (1.35 - 0.45 * J ^ 2)
        '        Di(2) = 2.2 * 10 ^ -6 / J
        '        '************************

        '        '************************
        '        'Calculate V0 
        '        'Calculate QA_l
        '        'Calculate 1/S
        '        'Calculate F = R/S * 1/QA_l
        '        '************************
        '        Dim V0 As Double
        '        Dim QA_l As Double
        '        Dim inv_S As Double = 0
        '        'Dim F As Double
        '        V0 = E0 / J

        '        '**************************************************
        '        'Test: tried to replace the ionization cross section used by Pouchou and Pichoir by the model developped by Bote and Salvat.
        '        'Dim indice As Integer = 0
        '        'For i As Integer = 0 To UBound(elt_exp_all)
        '        '    If elt_exp_all(i).z = studied_element.z Then
        '        '        studied_element.el_ion_xs = elt_exp_all(i).el_ion_xs
        '        '        Exit For
        '        '    End If
        '        'Next
        '        'Dim shell1 As Integer
        '        'Dim shell2 As Integer
        '        'Siegbahn_to_transition_num(studied_element.line(line_indice).xray_name, shell1, shell2)
        '        'Dim fc6 As Double = 0
        '        'Dim norm_xs As Double = 1


        '        'If My.Forms.Form1.CheckBox3.Checked = True Then
        '        'fc6 = Xray_production_xs_el_impact(studied_element, shell1, E0)
        '        'fc6 = fc6 * 6.022 * 10 ^ 23 / (4 * Math.PI)
        '        'Else
        '        'norm_xs = Xray_production_xs_el_impact(studied_element, shell1, 15) * 6.022 * 10 ^ 23 / (4 * Math.PI) /
        '        '            qe0(studied_element.line(line_indice).Ec, 15, studied_element.line(line_indice).xray_name, studied_element.z)
        '        '    fc6 = qe0(studied_element.line(line_indice).Ec, E0, studied_element.line(line_indice).xray_name, studied_element.z)
        '        'End If

        '        'QA_l = fc6 * norm_xs
        '        '**************************************************

        '        QA_l = Math.Log(U0) / El ^ 2 / U0 ^ m_

        '        For k As Integer = 0 To 2
        '            Dim Tk As Double = 1 + Pi(k) - m_
        '            inv_S = inv_S + Di(k) * (V0 / U0) ^ Pi(k) * (Tk * U0 ^ Tk * Math.Log(U0) - U0 ^ Tk + 1) / Tk ^ 2
        '        Next
        '        inv_S = inv_S * U0 / (V0 * M)

        '        F = R * inv_S / QA_l 'Is it * QA_l or / QA_l? From equation 13 p.37 it is "*" but from equations 2 and 3 it seems that it is "/".
        '        '************************

        '        '************************
        '        'Calculate fictitious mean atomic number (z)
        '        'in order to determine Rm (dRm) -> in text p.55 (only in case of multilayer specimen?).
        '        '************************
        '        If layer_handler.Count > 1 Then 'AM XX 06-09-2018 (Septembre)
        '            Z_bar = (Z_bar + Zb_bar) / 2
        '        End If
        '        '************************

        '        '************************
        '        'Calculate G1
        '        'Calculate G2
        '        'Calculate G3
        '        'Calculate Rm
        '        '************************
        '        Dim G1 As Double
        '        Dim G2 As Double
        '        Dim G3 As Double
        '        Dim Rm As Double
        '        G1 = 0.11 + 0.41 * Math.Exp(-(Z_bar / 12.75) ^ 0.75)
        '        G2 = 1 - Math.Exp(-(U0 - 1.0) ^ 0.35 / 1.19) 'AM XX 29-01-2018
        '        G3 = 1 - Math.Exp(-(U0 - 0.5) * Z_bar ^ 0.4 / 4)

        '        Rm = G1 * G2 * G3 * Rx '* ((1 - 1.5) / (1 - 100) * Zb_bar + (1 * 100 - 1.5) / (100 - 1)) 'AMXXX   'AMXXX
        '        '************************

        '        '************************
        '        'Calculate d
        '        'Check the consistency of d
        '        '************************
        '        Dim d_ As Double
        '        d_ = (Rx - Rm) * (F - (phi0 * Rx / 3)) * ((Rx - Rm) * F - phi0 * Rx * (Rm + Rx / 3))
        '        If (d_ < 0.0) Then
        '            Dim Rm_temp As Double = Rm
        '            Rm = Rx * (F - phi0 * Rx / 3) / (F + phi0 * Rx)
        '            d_ = 0
        '        End If
        '        '************************

        '        '************************
        '        'Calculate the 'wt.fraction averaged mass absorption coefficient'
        '        'for the current layer
        '        '************************
        '        Dim mac As Double
        '        Dim chi As Double
        '        'mac = MAC_calculation(studied_element.line(line_indice).xray_energy, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
        '        mac = MAC_calculation(studied_element, line_indice, mother_layer_id, layer_handler, elt_exp_all, fit_MAC, options)
        '        'Debug.Print(mac)
        '        chi = mac / sin_toa_in_rad
        '        '************************

        '        '************************
        '        'Calculation of phi_rz
        '        'Handle patologic cases
        '        '************************
        '        Dim Rc As Double
        '        Dim A1 As Double
        '        Dim A2 As Double
        '        Dim B1 As Double

        '        If Rm < 0 Or Rm > Rx Then
        '            Dim s_min As Double = 4.5
        '            If F < phi0 / s_min Then
        '                phi_rz = phi0 * F / (phi0 + chi * F)
        '            Else
        '                Dim phi0_p As Double
        '                phi0_p = 3 * (F * s_min - phi0) / (Rx * s_min - 3)
        '                phi_rz = ((2 / chi ^ 3) * (1 - Math.Exp(-chi * Rx)) - 2 * Rx / chi ^ 2 + Rx ^ 2 / chi) * (phi0_p / Rx ^ 2) + (phi0 - phi0_p) / (s_min + chi)
        '            End If
        '        Else
        '            '************************
        '            'Calculate Rc
        '            'Calculate A1
        '            'Calculate A2
        '            'Calculate B1
        '            '************************
        '            Rc = 1.5 * ((F - phi0 * Rx / 3) / phi0 - Math.Sqrt(d_) / (phi0 * (Rx - Rm)))
        '            If Rc < 0 Then
        '                Rc = 3 * Rm * (F + phi0 * Rx) / (2 * phi0 * Rx) '????
        '            End If
        '            A1 = phi0 / (Rm * (Rc - Rx * (Rc / Rm - 1)))
        '            A2 = A1 * (Rc - Rm) / (Rc - Rx)
        '            B1 = phi0 - A1 * Rm ^ 2
        '            '************************

        '            'If Rc < 0 Then
        '            '    Rc = 0  '???
        '            'End If

        '            '************************
        '            'Integrate the phi(rz) function
        '            'Calculate the cumulative mass depth for each layer
        '            '************************
        '            Dim cumulative_mass_depth() As Double = Nothing
        '            Dim temp As Double = 0
        '            For i As Integer = 0 To UBound(layer_handler)
        '                If cumulative_mass_depth Is Nothing Then
        '                    ReDim cumulative_mass_depth(0)
        '                Else
        '                    ReDim Preserve cumulative_mass_depth(UBound(cumulative_mass_depth) + 1)
        '                End If
        '                temp = temp + layer_handler(i).mass_thickness
        '                cumulative_mass_depth(i) = temp
        '            Next

        '            Dim lim_min As Double
        '            Dim lim_max As Double
        '            If mother_layer_id = 0 Then
        '                lim_min = 0
        '            Else
        '                lim_min = cumulative_mass_depth(mother_layer_id - 1)
        '            End If
        '            lim_max = cumulative_mass_depth(mother_layer_id)
        '            If lim_max > Rx Then lim_max = Rx

        '            Dim H1 As Double
        '            Dim H2 As Double


        '            If lim_min >= Rc Then
        '                H1 = 0
        '            Else
        '                If chi = 0 Then
        '                    H1 = A1 * (ppint2(A1, B1, Rm, chi, Math.Min(lim_max, Rc)) - ppint2(A1, B1, Rm, chi, lim_min))
        '                Else
        '                    H1 = -A1 / chi * (ppint2(A1, B1, Rm, chi, Math.Min(lim_max, Rc)) - ppint2(A1, B1, Rm, chi, lim_min))
        '                End If
        '            End If

        '            If lim_max <= Rc Then
        '                H2 = 0
        '            Else
        '                If chi = 0 Then
        '                    H2 = A2 * (ppint2(A2, 0, Rx, chi, Math.Min(lim_max, Rx)) - ppint2(A2, 0, Rx, chi, Math.Max(Rc, lim_min)))
        '                Else
        '                    H2 = -A2 / chi * (ppint2(A2, 0, Rx, chi, Math.Min(lim_max, Rx)) - ppint2(A2, 0, Rx, chi, Math.Max(Rc, lim_min)))
        '                End If
        '            End If

        '            Dim H1_plus_H2_test As Double = H1 + H2
        '            If H1_plus_H2_test <= 0 Then
        '                H1_plus_H2_test = 0.00000000001
        '            End If
        '            phi_rz = H1_plus_H2_test

        '        End If

        '        Dim abs As Double = abs_outer_layers(layer_handler, mother_layer_id, studied_element, line_indice, elt_exp_all, mac, sin_toa_in_rad, fit_MAC, options)
        '        phi_rz = phi_rz * abs
        '        '************************

        '        PAP_Rx = Rx
        '        PAP_Rm = Rm
        '        PAP_Rc = Rc
        '        PAP_A1 = A1
        '        PAP_A2 = A2
        '        PAP_B1 = B1




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

    End Function
    Public Function Yield(ByVal studied_element As Elt_exp, ByVal line_indice As Integer) As Double
        Try
            Dim Z As Integer = studied_element.z
            ' K lines
            If studied_element.line(line_indice).xray_name(0) = "K" Then
                Yield = 0.06893861 + Z * (0.024152 + Z * (0.0003324179 - Z * 0.00000392704544))

                ' L lines
            ElseIf studied_element.line(line_indice).xray_name(0) = "L" Then
                Yield = -0.111065 + Z * (0.01368 - Z * Z * 0.00000021772)

            ElseIf studied_element.line(line_indice).xray_name(0) = "M" Then
                Yield = -0.00036 + Z * (0.00386 + Z * Z * 0.00000020101)
            End If
            ' M lines


            Yield = Yield ^ 4
            Yield = Yield / (1 + Yield)

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Yield " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

    Public Function Beta_func(ByVal studied_element As Elt_exp, ByVal line_indice As Integer) As Single
        Try
            Dim v(100) As Single
            ' K lines
            If studied_element.line(line_indice).xray_name(0) = "K" Then
                v(1) = 0
                v(2) = 0
                v(3) = 0
                v(4) = 0
                v(5) = 0
                v(6) = 0
                v(7) = 0
                v(8) = 0
                v(9) = 0
                v(10) = 0
                v(11) = 0
                v(12) = 0.013
                v(13) = 0.014
                v(14) = 0.025
                v(15) = 0.042
                v(16) = 0.063
                v(17) = 0.085
                v(18) = 0.11
                v(19) = 0.12
                v(20) = 0.127
                v(21) = 0.131
                v(22) = 0.131
                v(23) = 0.132
                v(24) = 0.133
                v(25) = 0.134
                v(26) = 0.134
                v(27) = 0.135
                v(28) = 0.136
                v(29) = 0.137
                v(30) = 0.139
                v(31) = 0.146
                v(32) = 0.152
                v(33) = 0.156
                v(34) = 0.161
                v(35) = 0.166
                v(36) = 0.17
                v(37) = 0.175
                v(38) = 0.175
                v(39) = 0.183
                v(40) = 0.187
                v(41) = 0.191
                v(42) = 0.195
                v(43) = 0.199
                v(44) = 0.202
                v(45) = 0.206
                v(46) = 0.209
                v(47) = 0.212
                v(48) = 0.216
                v(49) = 0.218
                v(50) = 0.222
                v(51) = 0.224
                v(52) = 0.227
                v(53) = 0.23
                v(54) = 0.233
                v(55) = 0.235
                v(56) = 0.237
                v(57) = 0.24
                v(58) = 0.242
                v(59) = 0.244
                v(60) = 0.246
                v(61) = 0.248
                v(62) = 0.25
                v(63) = 0.252
                v(64) = 0.254
                v(65) = 0.256
                v(66) = 0.258
                v(67) = 0.259
                v(68) = 0.261
                v(69) = 0.262
                v(70) = 0.264
                v(71) = 0.265
                v(72) = 0.266
                v(73) = 0.268
                v(74) = 0.269
                v(75) = 0.27
                v(76) = 0.271
                v(77) = 0.272
                v(78) = 0.273
                v(79) = 0.275
                v(80) = 0.276
                v(81) = 0.277
                v(82) = 0.277
                v(83) = 0.278
                v(84) = 0.279
                v(85) = 0.28
                v(86) = 0.281
                v(87) = 0.282
                v(88) = 0.283
                v(89) = 0.283
                v(90) = 0.284
                v(91) = 0.285
                v(92) = 0.286
                v(93) = 0.286
                v(94) = 0.287
                v(95) = 0.288
                v(96) = 0.289
                v(97) = 0.289
                v(98) = 0.29
                v(99) = 0.29
                v(100) = 0.291

                Beta_func = v(studied_element.z)

                ' L lines
            ElseIf studied_element.line(line_indice).xray_name(0) = "L" Then
                Beta_func = -0.015 + 0.00575 * studied_element.z

                ' M lines
            ElseIf studied_element.line(line_indice).xray_name(0) = "M" Then
                Beta_func = 0.5
            End If

            Return Beta_func

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Beta_func " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function
End Module
