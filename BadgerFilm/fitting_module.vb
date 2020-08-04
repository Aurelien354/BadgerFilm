'Imports System.ComponentModel
Imports System.IO

Public Class fitting_module
    Public Sub fit(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef elt_exp_all() As Elt_exp, ByVal toa As Double,
                   ByVal pen_path As String, ByVal Ec_data() As String, ByRef fit_MAC As fit_MAC, ByVal options As options)

        Dim pactual(UBound(p)) As Double '= {20.0, 1.0}   '/* Actual values used To make data */
        For i As Integer = 0 To UBound(p)
            pactual(i) = p(i)
        Next

        Dim perror(UBound(p)) As Double ' = {0.0, 0.0}                   '/* Returned parameter errors */

        Dim result As MPFitLib.mp_result = New MPFitLib.mp_result(UBound(p) + 1)
        result.xerror = perror

        Dim v As CustomUserVariable = New CustomUserVariable()
        v.X = x
        v.Y = y
        v.Ey = ey
        v.layer_handler = layer_handler
        v.elt_exp_handler = elt_exp_handler
        v.elt_exp_all = elt_exp_all
        v.toa = toa
        v.pen_path = pen_path
        v.Ec_data = Ec_data
        v.fit_MAC = fit_MAC
        v.options = options

        Dim conf As MPFitLib.mp_config = New MPFitLib.mp_config()
        conf.epsfcn = 0.1
        'conf.gtol = 1.0E-22
        'conf.ftol = 1.0E-30
        'conf.covtol = 1.0E-22
        'conf.xtol = 1.0E-35
        conf.maxiter = 5000
        'conf.maxfev = 50000

        'conf.stepfactor = 0.1
        'conf.douserscale = 10.1
        'conf.gtol = 10000000

        '/* Call fitting function */
        Dim status As Integer
        status = MPFitLib.MPFit.Solve(AddressOf ForwardModels.customfunc, x.Count, p.Count, p, pars, conf, v, result)


        'Update the results with the final fitting parameters
        Dim count As Integer = 0
        For i As Integer = 0 To UBound(layer_handler)
            layer_handler(i).thickness = p(count)
            layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            count = count + 1
        Next
        For i As Integer = 0 To UBound(layer_handler)
            For j As Integer = 0 To UBound(layer_handler(i).element)
                layer_handler(i).element(j).conc_wt = p(count)
                count = count + 1
            Next
        Next

        If fit_MAC.activated = True Then
            fit_MAC.MAC = p(UBound(p))
            fit_MAC.scaling_factor = p(UBound(p) - 1)
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                       elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC)

                        Debug.Print(options.phi_rz_mode & vbTab & "Unk: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab & elt_exp_handler(i).line(j).k_ratio(k).elt_intensity)

                        elt_exp_handler(i).line(j).k_ratio(k).theo_value = p(UBound(p) - 1) * elt_exp_handler(i).line(j).k_ratio(k).elt_intensity '/ elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                    Next
                Next
            Next
        Else
            'Calculate the final intensities with the best fit parameters
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                       elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC)

                        Debug.Print(options.phi_rz_mode & vbTab & "Unk: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab & elt_exp_handler(i).line(j).k_ratio(k).elt_intensity)

                        elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                    Next
                Next
            Next
        End If



        'Output the results
        Console.WriteLine("*** Fitting status = {0}", status)
        buffer_text = buffer_text & "*** Fitting status = " & status
        PrintResult(p, pactual, result, buffer_text)

    End Sub

    '/* Simple routine to print the fit results */
    Private Sub PrintResult(ByVal x() As Double, ByVal xact() As Double, ByVal result As MPFitLib.mp_result, ByRef buffer_text As String)

        If (x Is Nothing) Then Exit Sub

        buffer_text = buffer_text & "  CHI-SQUARE = " & result.bestnorm & "    (" & result.nfunc - result.nfree & " DOF)" & vbCrLf
        buffer_text = buffer_text & "        NPAR = " & result.npar & vbCrLf
        buffer_text = buffer_text & "       NFREE = " & result.nfree & vbCrLf
        buffer_text = buffer_text & "     NPEGGED = " & result.npegged & vbCrLf
        buffer_text = buffer_text & "       NITER = " & result.niter & vbCrLf
        buffer_text = buffer_text & "        NFEV = " & result.nfev & vbCrLf
        buffer_text = buffer_text & vbCrLf

        If (xact IsNot Nothing) Then
            For i As Integer = 0 To result.npar - 1
                buffer_text = buffer_text & "  P[" & i & "] = " & x(i) & " +/- " & result.xerror(i) & "     (INITIAL " & xact(i) & ")" & vbCrLf
            Next
        Else
            For i As Integer = 0 To result.npar - 1
                buffer_text = buffer_text & "  P[" & i & "] = " & x(i) & " +/- " & result.xerror(i) & vbCrLf
            Next
        End If
    End Sub


    Public Sub fit_brem_fluo(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef elt_exp_all() As Elt_exp, ByVal toa As Double,
                ByRef elt() As String, ByRef line() As String, ByVal pen_path As String, ByVal Ec_data() As String, ByVal options As options, ByRef results As String) 'ByRef elt() As String, ByRef line() As String)

        Dim pactual(UBound(p)) As Double    '/* Actual values used To make data */
        For i As Integer = 0 To UBound(p)
            pactual(i) = p(i)
        Next

        Dim perror(UBound(p)) As Double     '/* Returned parameter errors */

        Dim result As MPFitLib.mp_result = New MPFitLib.mp_result(UBound(p) + 1)
        result.xerror = perror

        Dim v As CustomUserVariable_brem_fluo = New CustomUserVariable_brem_fluo()
        v.X = x
        v.Y = y
        v.Ey = ey
        v.elt = elt
        v.line = line
        v.layer_handler = layer_handler
        v.elt_exp_handler = elt_exp_handler
        v.elt_exp_all = elt_exp_all
        v.toa = toa
        v.pen_path = pen_path
        v.Ec_data = Ec_data
        v.options = options
        v.results = results

        Dim conf As MPFitLib.mp_config = New MPFitLib.mp_config()
        conf.epsfcn = 0.1
        'conf.stepfactor = 0.1
        conf.gtol = 1.0E-22
        conf.ftol = 1.0E-30
        conf.covtol = 1.0E-22
        conf.xtol = 1.0E-35
        conf.maxiter = 50000
        conf.maxfev = 50000
        'conf.douserscale = 10.1
        'conf.gtol = 10000000

        '/* Call fitting function */
        Dim status As Integer
        status = MPFitLib.MPFit.Solve(AddressOf ForwardModels.customfunc_brem_fluo, x.Count, p.Count, p, pars, conf, v, result)

        results = v.results
        Console.WriteLine("*** TestLinFit status = {0}", status)
        buffer_text = buffer_text & "*** TestLinFit status = {0}"
        PrintResult(p, pactual, result, buffer_text)

    End Sub

    Public Sub fit_coeff1(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef Z() As Double, ByRef wt() As Double, ByRef A() As Double, ByRef Ex() As Double)

        Dim pactual(UBound(p)) As Double    '/* Actual values used To make data */
        For i As Integer = 0 To UBound(p)
            pactual(i) = p(i)
        Next

        Dim perror(UBound(p)) As Double     '/* Returned parameter errors */

        Dim result As MPFitLib.mp_result = New MPFitLib.mp_result(UBound(p) + 1)
        result.xerror = perror

        Dim v As CustomUserVariable_coeff = New CustomUserVariable_coeff()
        v.X = x
        v.Y = y
        v.Ey = ey
        v.Z = Z
        v.wt = wt
        v.A = A
        v.Ex = Ex


        Dim conf As MPFitLib.mp_config = New MPFitLib.mp_config()
        conf.epsfcn = 0.1
        'conf.stepfactor = 0.1
        conf.gtol = 1.0E-22
        conf.ftol = 1.0E-30
        conf.covtol = 1.0E-22
        conf.xtol = 1.0E-35
        conf.maxiter = 50000
        conf.maxfev = 50000
        'conf.douserscale = 10.1
        'conf.gtol = 10000000

        '/* Call fitting function */
        Dim status As Integer
        status = MPFitLib.MPFit.Solve(AddressOf ForwardModels.customfunc_coeff, x.Count, p.Count, p, pars, conf, v, result)

        Console.WriteLine("*** TestLinFit status = {0}", status)
        buffer_text = buffer_text & "*** TestLinFit status = {0}"
        PrintResult(p, pactual, result, buffer_text)

    End Sub
End Class

Public Class CustomUserVariable
    Public X As Double()
    Public Y As Double()
    Public Ey As Double()
    Public layer_handler() As layer
    Public elt_exp_handler() As Elt_exp
    Public elt_exp_all() As Elt_exp
    Public toa As Double
    Public pen_path As String
    Public mode As String
    Public Ec_data() As String
    Public fit_MAC As fit_MAC
    Public options As options
End Class

Public Class CustomUserVariable_brem_fluo
    Public X As Double()
    Public Y As Double()
    Public Ey As Double()
    Public layer_handler() As layer
    Public elt_exp_handler() As Elt_exp
    Public elt_exp_all() As Elt_exp
    Public toa As Double
    Public pen_path As String
    Public mode As String
    Public Ec_data() As String
    Public options As options
    Public elt() As String
    Public line() As String
    Public results As String
End Class

Public Class CustomUserVariable_coeff
    Public X As Double()
    Public Y As Double()
    Public Ey As Double()
    Public Z As Double()
    Public wt As Double()
    Public A As Double()
    Public Ex As Double()

End Class

Public Class ForwardModels

    Public Shared Function customfunc(p() As Double, dy() As Double, dvec() As IList(Of Double), vars As Object) As Integer 'IList<Double>[] dvec
        Dim x(), y(), ey() As Double
        Dim layer_handler() As layer
        Dim elt_exp_handler() As Elt_exp
        Dim elt_exp_all() As Elt_exp

        Dim toa As Double
        Dim pen_path As String
        Dim mode As String
        Dim Ec_data() As String = Nothing
        Dim fit_MAC As fit_MAC
        Dim options As options

        Dim v As CustomUserVariable = vars

        x = v.X
        y = v.Y
        ey = v.Ey
        layer_handler = v.layer_handler
        elt_exp_handler = v.elt_exp_handler
        elt_exp_all = v.elt_exp_all
        toa = v.toa
        pen_path = v.pen_path
        mode = v.mode
        Ec_data = v.Ec_data
        fit_MAC = v.fit_MAC
        options = v.options


        If options.BgWorker.CancellationPending = True Then
            Return -1
        End If

        Dim calculated_y(UBound(y)) As Double
        Dim indice As Integer = 0

        Dim count As Integer = 0
        For i As Integer = 0 To UBound(layer_handler)
            layer_handler(i).thickness = p(count)
            layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
            count = count + 1
        Next
        For i As Integer = 0 To UBound(layer_handler)
            For j As Integer = 0 To UBound(layer_handler(i).element)
                layer_handler(i).element(j).conc_wt = p(count)
                count = count + 1
            Next
        Next

        If fit_MAC.activated = True Then
            fit_MAC.MAC = p(UBound(p))
            Dim max As Double = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                       elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC)
                        'elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                        calculated_y(indice) = p(UBound(p) - 1) * elt_exp_handler(i).line(j).k_ratio(k).elt_intensity  'elt_exp_handler(i).line(j).k_ratio(k).theo_value
                        If calculated_y(indice) > max Then max = calculated_y(indice)
                        indice = indice + 1
                    Next
                Next
            Next

        Else
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                       elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC)
                        elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                        calculated_y(indice) = elt_exp_handler(i).line(j).k_ratio(k).theo_value
                        indice = indice + 1
                    Next
                Next
            Next
        End If
        'For i As Integer = 0 To UBound(layer_handler)
        '    For j As Integer = 0 To UBound(layer_handler(i).element)
        '        If layer_handler(i).element(j).k_ratio IsNot Nothing Then
        '            For k As Integer = 0 To UBound(layer_handler(i).element(j).k_ratio)
        '                layer_handler(i).element(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, layer_handler(i).element(j), layer_handler(i).element(j).k_ratio(k).kv, toa, Ec_data, options, False)
        '                layer_handler(i).element(j).k_ratio(k).theo_value = layer_handler(i).element(j).k_ratio(k).elt_intensity / layer_handler(i).element(j).k_ratio(k).std_intensity
        '                calculated_y(indice) = layer_handler(i).element(j).k_ratio(k).theo_value
        '                indice = indice + 1
        '            Next
        '        End If
        '    Next
        'Next

        For i As Integer = 0 To UBound(calculated_y) - indice
            calculated_y(indice + i) = 0
            For j As Integer = 0 To UBound(layer_handler(i).element)
                calculated_y(indice + i) = calculated_y(indice + i) + layer_handler(i).element(j).conc_wt
            Next
        Next


        For i As Integer = 0 To dy.Length - 1
            dy(i) = (y(i) - calculated_y(i)) / ey(i)
        Next

        Dim print_res_debug As String = ""
        For i As Integer = 0 To dy.Length - 1
            print_res_debug = print_res_debug & y(i) - calculated_y(i) & vbTab
        Next
        Debug.WriteLine(print_res_debug)

        Return 0
    End Function

    Public Shared Function customfunc_brem_fluo(p() As Double, dy() As Double, dvec() As IList(Of Double), vars As Object) As Integer 'IList<Double>[] dvec
        Dim x(), y(), ey() As Double
        Dim elt() As String
        Dim line() As String

        Dim layer_handler() As layer
        Dim elt_exp_handler() As Elt_exp
        Dim elt_exp_all() As Elt_exp

        Dim toa As Double
        Dim pen_path As String
        Dim mode As String
        Dim Ec_data() As String = Nothing
        Dim options As options
        Dim results As String

        Dim v As CustomUserVariable_brem_fluo = vars

        x = v.X
        y = v.Y
        ey = v.Ey
        elt = v.elt
        line = v.line

        layer_handler = v.layer_handler
        elt_exp_handler = v.elt_exp_handler
        elt_exp_all = v.elt_exp_all
        toa = v.toa
        pen_path = v.pen_path
        mode = v.mode
        Ec_data = v.Ec_data
        options = v.options
        results = v.results
        'Form1.TextBox2.Text = p(0)
        'Form1.txt6 = p(1)
        'Form1.TextBox6.Text = p(1)
        'Call Form1.Button6_Click()



        Dim save_results As String = ""
        For ll As Integer = 2 To 30
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)

                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, ll, toa, Ec_data, options, True, save_results, Nothing,, p)

                    Next
                Next
            Next
        Next

        vars.results = p(0) & vbTab & p(1) & vbCrLf & save_results
        'Clipboard.SetText(save_results)
        'Debug.Print(save_results)
        Debug.Print(p(0) & vbTab & p(1))


        Dim calc_res() As String = Split(save_results, vbCrLf)

        Dim calc_kV(UBound(calc_res) - 1) As Double
        Dim calc_elt(UBound(calc_res) - 1) As String
        Dim calc_line(UBound(calc_res) - 1) As String
        Dim calc_brem_val(UBound(calc_res) - 1) As Double

        For i As Integer = 0 To UBound(calc_res) - 1
            Dim tmp2() As String = Split(calc_res(i), vbTab)
            calc_kV(i) = tmp2(0)
            calc_elt(i) = tmp2(1)
            calc_line(i) = tmp2(2)
            calc_brem_val(i) = tmp2(5)
        Next

        Dim count As Integer = 0
        For i As Integer = 0 To UBound(calc_kV)
            If x(count) = calc_kV(i) And elt(count) = calc_elt(i) And line(count) = calc_line(i) Then
                dy(count) = (y(count) - calc_brem_val(i)) / ey(count)

                If Double.IsNaN(dy(count)) Then
                    dy(count) = 0
                End If
                If Double.IsInfinity(dy(count)) Then
                    dy(count) = 0
                End If
                count = count + 1
                If count > UBound(x) Then Exit For
            End If
        Next


        'For i As Integer = 0 To dy.Length - 1
        '    dy(i) = (y(i) - calculated_y(i)) / ey(i)
        'Next

        Return 0
    End Function

    Public Shared Function customfunc_coeff(p() As Double, dy() As Double, dvec() As IList(Of Double), vars As Object) As Integer 'IList<Double>[] dvec
        Dim x(), y(), ey(), Z(), wt(), A(), Ex() As Double

        Dim v As CustomUserVariable_coeff = vars

        x = v.X
        y = v.Y
        ey = v.Ey
        Z = v.Z
        wt = v.wt
        A = v.A
        Ex = v.Ex

        For i As Integer = 0 To UBound(y)
            dy(i) = (y(i) - (Z(i) * p(0) + wt(i) * p(1) + A(i) * p(2) + Ex(i) * p(3) + p(4))) / ey(i)
        Next

        'For i As Integer = 0 To dy.Length - 1
        '    dy(i) = (y(i) - calculated_y(i)) / ey(i)
        'Next

        Return 0
    End Function
End Class


