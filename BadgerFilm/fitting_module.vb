'Imports System.ComponentModel
Imports System.IO

Public Class fitting_module
    Public Sub fit(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef elt_exp_all() As Elt_exp, ByVal toa As Double,
                   ByVal pen_path As String, ByVal Ec_data() As String, ByRef fit_MAC As fit_MAC, ByVal options As options)
        Try
            'Dim O_wt_by_stoichio() As Double
            'Dim Elt_wt_by_stoichio() As Double

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
            'v.O_wt_by_stoichio = O_wt_by_stoichio
            'v.Elt_wt_by_stoichio = Elt_wt_by_stoichio

            Dim conf As MPFitLib.mp_config = New MPFitLib.mp_config()
            conf.epsfcn = 0.1
            'conf.gtol = 1.0E-22
            ' conf.ftol = 1.0E-30
            'conf.covtol = 1.0E-22
            'conf.xtol = 1.0E-35
            conf.maxiter = 5000
            'conf.maxfev = 50000

            'conf.stepfactor = 0.1
            'conf.douserscale = 10.1
            'conf.gtol = 10000000

            '***************************************************
            'Count the number of layers having O definied by stoichiometry.
            'Count the number of layers having an element other than O definied by stoichiometry.
            'Redim v.O_wt_by_stoichio() and v.Elt_wt_by_stoichio() accordingly.
            'Dim num_of_stoichio_O As Integer = 0
            'Dim num_of_stoichio_Elt As Integer = 0
            'For i As Integer = 0 To UBound(layer_handler)
            '    If layer_handler(i).stoichiometry.O_by_stoichio = True Then
            '        'For j As Integer = 0 To UBound(layer_handler(i).element)
            '        '    If layer_handler(i).element(j).elt_name = "O" Then
            '        '        num_of_stoichio_O = num_of_stoichio_O + 1
            '        '    End If
            '        '    If layer_handler(i).element(j).elt_name = "C" Then
            '        '        num_of_stoichio_Elt = num_of_stoichio_Elt + 1
            '        '    End If
            '        'Next
            '        num_of_stoichio_O = num_of_stoichio_O + 1
            '    End If
            '    If layer_handler(i).stoichiometry.Elt_by_stoichio_to_O = True Then
            '        num_of_stoichio_Elt = num_of_stoichio_Elt + 1
            '    End If
            'Next
            'ReDim v.O_wt_by_stoichio(num_of_stoichio_O - 1)
            'ReDim v.Elt_wt_by_stoichio(num_of_stoichio_Elt - 1)
            '***************************************************

            '***************************************************
            '/* Call fitting function */
            Dim status As Integer
            status = MPFitLib.MPFit.Solve(AddressOf ForwardModels.customfunc, x.Count, p.Count, p, pars, conf, v, result)
            '***************************************************

            '***************************************************
            'Update the results with the final fitting parameters and stoichiometry values
            Dim count As Integer = 0
            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).thickness = p(count)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
                count = count + 1
            Next
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    'Re-calculates concentations if defined as fixed and defined in atomic fraction
                    If layer_handler(i).element(j).isConcFixed = True And layer_handler(i).wt_fraction = False Then
                        Dim temp As Double = 0
                        For k As Integer = 0 To UBound(layer_handler(i).element)
                            If k = j Then Continue For
                            temp = temp + layer_handler(i).element(k).conc_wt / zaro(layer_handler(i).element(k).z)(0)
                        Next
                        layer_handler(i).element(j).conc_wt = temp * layer_handler(i).element(j).conc_at_ori / (1 - layer_handler(i).element(j).conc_at_ori) * zaro(layer_handler(i).element(j).z)(0)
                        p(count) = layer_handler(i).element(j).conc_wt
                    Else
                        layer_handler(i).element(j).conc_wt = p(count)
                    End If
                    count = count + 1
                Next
            Next

            Dim index As Integer = 0
            For i As Integer = 0 To UBound(layer_handler)
                If layer_handler(i).stoichiometry.O_by_stoichio = True Then
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        If layer_handler(i).element(j).elt_name = "O" Then
                            layer_handler(i).element(j).conc_wt = layer_handler(i).stoichiometry.O_wt_conc 'v.O_wt_by_stoichio(index)
                        End If
                    Next
                    If layer_handler(i).stoichiometry.Elt_by_stoichio_to_O = True Then
                        For j As Integer = 0 To UBound(layer_handler(i).element)
                            If layer_handler(i).element(j).elt_name = layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_name Then
                                layer_handler(i).element(j).conc_wt = layer_handler(i).stoichiometry.Elt_wt_conc 'v.Elt_wt_by_stoichio(index)
                            End If
                        Next
                    End If
                    index = index + 1
                End If
            Next

            If fit_MAC.activated = True Then
                fit_MAC.MAC = p(UBound(p))
                fit_MAC.scaling_factor = p(UBound(p) - 1)
                'Dim max As Double = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        Dim norm As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, fit_MAC.norm_kV, toa, Ec_data, options, False, "", fit_MAC)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                           elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC) * fit_MAC.scaling_factor '/ norm

                            Debug.Print(options.phi_rz_mode & vbTab & "Unk: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab & elt_exp_handler(i).line(j).k_ratio(k).elt_intensity)

                            elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity '* fit_MAC.scaling_factor
                            'If elt_exp_handler(i).line(j).k_ratio(k).elt_intensity > max Then max = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity
                        Next
                    Next
                Next

                'fit_MAC.scaling_factor = max
                'For i As Integer = 0 To UBound(elt_exp_handler)
                '    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                '        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                '            elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / fit_MAC.scaling_factor  '/ elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                '        Next
                '    Next
                'Next
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
            '***************************************************

            '***************************************************
            'Output the results
            Console.WriteLine("*** Fitting status = {0}", status)
            buffer_text = buffer_text & "*** Fitting status = " & status
            If status = -24 Then
                Dim freepars As Integer = 0
                For i As Integer = 0 To UBound(pars)
                    If pars(i).isFixed = 0 Then freepars = freepars + 1
                Next
                MsgBox("Error: problem with the degrees of freedom. The system has " & freepars & " variables (concentrations, thicknesses) but only " & x.Count &
                       " equations (k-ratios, sum of concentrations = 1)." & vbCrLf &
                       "To correct this, decrease the number of variables (fix the composition or thickness of a layer) or increase the number of equations (use the option ""sum of concentrations = 1"" or enter more k-ratios).")
            End If
            PrintResult(p, pactual, result, buffer_text)
            '***************************************************

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in fit " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    '/* Simple routine to print the fit results */
    Private Sub PrintResult(ByVal x() As Double, ByVal xact() As Double, ByVal result As MPFitLib.mp_result, ByRef buffer_text As String)
        Try
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
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in PrintResult " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub


    Public Sub fit_brem_fluo(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef elt_exp_all() As Elt_exp, ByVal toa As Double,
                ByRef elt() As String, ByRef line() As String, ByVal pen_path As String, ByVal Ec_data() As String, ByVal options As options, ByRef results As String) 'ByRef elt() As String, ByRef line() As String)
        Try
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

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in fit_brem_fluo " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub fit_coeff1(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef Z() As Double, ByRef wt() As Double, ByRef A() As Double, ByRef Ex() As Double)
        Try
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

        Catch exc As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in fit_coeff1 " & exc.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
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
    'Public O_wt_by_stoichio() As Double
    'Public Elt_wt_by_stoichio() As Double
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
        Try
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
            'Dim O_wt_by_stoichio() As Double
            'Dim Elt_wt_by_stoichio() As Double

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

            '******************************************************
            'update the thciknesses and compositions with the fitting parameters (stored in the p array by the fitting algorithm)
            '******************************************************
            Dim count As Integer = 0
            For i As Integer = 0 To UBound(layer_handler)
                layer_handler(i).thickness = p(count)
                layer_handler(i).mass_thickness = layer_handler(i).density * layer_handler(i).thickness * 10 ^ -8
                count = count + 1
            Next
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    'Re-calculates concentations if defined as fixed and defined in atomic fraction
                    If layer_handler(i).element(j).isConcFixed = True And layer_handler(i).wt_fraction = False Then
                        Dim temp As Double = 0
                        For k As Integer = 0 To UBound(layer_handler(i).element)
                            If k = j Then Continue For
                            temp = temp + layer_handler(i).element(k).conc_wt / zaro(layer_handler(i).element(k).z)(0)
                        Next
                        layer_handler(i).element(j).conc_wt = temp * layer_handler(i).element(j).conc_at_ori / (1 - layer_handler(i).element(j).conc_at_ori) * zaro(layer_handler(i).element(j).z)(0)
                        p(count) = layer_handler(i).element(j).conc_wt
                    Else
                        layer_handler(i).element(j).conc_wt = p(count)
                    End If
                    count = count + 1
                Next
            Next
            '******************************************************

            '******************************************************
            ' Calculate O by stoichiometry and another element by stoichiometry relative to O
            '******************************************************
            Try
                For i As Integer = 0 To UBound(layer_handler)
                    If layer_handler(i).stoichiometry.O_by_stoichio = True Then 'if O is defined by stoichiometry on the current layer, then proceed
                        If layer_handler(i).stoichiometry.Elt_by_stoichio_to_O = True Then 'if another element is definied by stoichiometry relative to O on the current layer, then proceed
                            Dim O_wt_conc As Double = 0
                            Dim Elt_wt_conc As Double = 0
                            Dim index_O As Integer = -1
                            Dim index_C As Integer = -1

                            For j As Integer = 0 To UBound(layer_handler(i).element)
                                If layer_handler(i).element(j).elt_name = "O" Then 'find the index of O in the list of elements of the current layer
                                    index_O = j
                                End If
                                If layer_handler(i).element(j).elt_name = layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_name Then 'find the index of the other element defined by stoichiometry in the list of elements of the current layer
                                    index_C = j
                                End If
                            Next

                            Dim O_wt_conc_initial As Double = 0
                            For j As Integer = 0 To UBound(layer_handler(i).element) 'calculate the initial O wt concentration based on the concentration of the other elements (the measured elements)
                                If layer_handler(i).element(j).elt_name <> "O" And layer_handler(i).element(j).elt_name <> layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_name Then
                                    Dim coeff As Double = Convert_wt_to_oxide_wt(layer_handler(i).element(j))
                                    O_wt_conc_initial = O_wt_conc_initial + (coeff - 1) * layer_handler(i).element(j).conc_wt
                                End If
                            Next

                            If index_C <> -1 Then 'if another element defined by stoichiometry has been found
                                'Dim O_wt_conc_from_Elt As Double = 0
                                'Dim old_O_wt_conc As Double = 0
                                'Dim O_wt_variations As Double = Math.Abs(old_O_wt_conc - O_wt_conc_initial)
                                'Dim iter As Integer = 0
                                Dim coeff As Double = Convert_wt_to_oxide_wt(layer_handler(i).element(index_C))
                                Dim B As Double = (coeff - 1) * layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio *
                                    layer_handler(i).element(index_C).a / layer_handler(i).element(index_O).a

                                O_wt_conc = O_wt_conc_initial * 1 / (1 - B)
                                Elt_wt_conc = O_wt_conc * layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio * layer_handler(i).element(index_C).a / layer_handler(i).element(index_O).a

                                'While O_wt_variations > 0.0001 And iter < 500 'loop on the O weight concentration until variations from one iteration to another are smaller than 0.0001 (or 0.01%) or until 500 iterations
                                '    old_O_wt_conc = O_wt_conc
                                '    Elt_wt_conc = (O_wt_conc_initial + O_wt_conc_from_Elt) / 16 * layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio * 12.011 'calculate the weight concentration of this other element based on the oxygen concentration
                                '    layer_handler(i).element(index_C).conc_wt = Elt_wt_conc 'update the weight concentration of this other element
                                '    O_wt_conc_from_Elt = O_by_stochiometry_from_elt_wt(layer_handler(i).element(index_C)) 'calculate the O concentration brought by this other element
                                '    O_wt_conc = O_wt_conc_initial + O_wt_conc_from_Elt 'calculate the new weigth concentration of O based on the newly added weigth concentration of this other element definied by stoichiometry
                                '    O_wt_variations = Math.Abs(old_O_wt_conc - O_wt_conc) 'calculate the variation on the O concentrtation brought by the other element defined by stoichiometry.
                                '    iter = iter + 1
                                'End While

                                layer_handler(i).stoichiometry.Elt_wt_conc = Elt_wt_conc 'update the concentration 

                            Else 'this should not happen but just in case
                                O_wt_conc = O_wt_conc_initial
                                MsgBox("You have indicated than " & layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_name & " was calculated by stoichiometry relative to O but this element is not included in the layer #" & i + 1 & "." & vbCrLf & "Please add this element.")
                                Return -1
                            End If

                            If index_O <> -1 Then 'update the O concentration
                                layer_handler(i).element(index_O).conc_wt = O_wt_conc
                                layer_handler(i).stoichiometry.O_wt_conc = O_wt_conc

                            Else 'this should not happen but just in case
                                MsgBox("You have indicated than O was calculated by stoichiometry but this element is not included in the layer #" & i + 1 & "." & vbCrLf & "Please add this element.")
                                Return -1
                            End If

                        Else 'only O is defined by stoichiometry
                            Dim index_O As Integer = -1
                            For j As Integer = 0 To UBound(layer_handler(i).element) 'find the index of O in the list of elements of the current layer
                                If layer_handler(i).element(j).elt_name = "O" Then
                                    index_O = j
                                End If
                            Next

                            If index_O <> -1 Then 'if O was found, then proceed
                                Dim O_wt_conc As Double = 0
                                For j As Integer = 0 To UBound(layer_handler(i).element)
                                    If layer_handler(i).element(j).elt_name <> "O" Then
                                        Dim coeff As Double = Convert_wt_to_oxide_wt(layer_handler(i).element(j))
                                        O_wt_conc = O_wt_conc + (coeff - 1) * layer_handler(i).element(j).conc_wt 'calculate the O weight concentration based on the weight concentration of the other elements
                                    End If
                                Next

                                layer_handler(i).element(index_O).conc_wt = O_wt_conc
                                layer_handler(i).stoichiometry.O_wt_conc = O_wt_conc

                            Else
                                MsgBox("Oxygen was defined by stoichiometry but no oxygen is present in the layer #" & i + 1 & "." & vbCrLf & "Please add oxygen to this layer.")
                                Return -1
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox("Error in fitting module: O by stoichiometry" & vbCrLf & ex.Message)
                Return -1
            End Try
            '******************************************************

            Dim calculated_y(UBound(y)) As Double 'used to store the calculated kratios or X-ray intensities
            Dim indice As Integer = 0
            '******************************************************
            'Calculates the MAC value of the kratios (with given elemental compositions and film thicknesses)
            '******************************************************
            If fit_MAC.activated = True Then
                'Calculates the MAC
                Try
                    fit_MAC.scaling_factor = p(UBound(p) - 1) 'update the scaling factor
                    fit_MAC.MAC = p(UBound(p)) 'update the MACs with the last value stored in the p array
                    For i As Integer = 0 To UBound(elt_exp_handler) 'loop on all the elements that have experimental data
                        For j As Integer = 0 To UBound(elt_exp_handler(i).line) 'loop on all the X-ray lines that have data for the given element
                            'Dim norm As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, fit_MAC.norm_kV, toa, Ec_data, options, False, "", fit_MAC)
                            For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio) 'calculate the X-ray intensity and scale it to fit the experimental data
                                elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                               elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC) * fit_MAC.scaling_factor '/ norm

                                calculated_y(indice) = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity '* p(UBound(p) - 1) 'store the intensity in the calculated_y array
                                indice = indice + 1
                            Next
                        Next
                    Next
                Catch ex As Exception
                    MsgBox("Error in fitting module: fit MAC" & vbCrLf & ex.Message)
                    Return -1
                End Try

            Else
                'Calculates the k-ratios
                Try
                    For i As Integer = 0 To UBound(elt_exp_handler) 'loop on all the elements that have experimental data
                        For j As Integer = 0 To UBound(elt_exp_handler(i).line) 'loop on all the X-ray lines that have data for the given element
                            For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio) 'calculate the X-ray intensity and scale it to fit the experimental data
                                elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                               elt_exp_handler(i).line(j).k_ratio(k).kv, toa, Ec_data, options, False, "", fit_MAC) 'calculate the X-ray intensity of the unknown
                                If elt_exp_handler(i).line(j).k_ratio(k).elt_intensity < 0 Then Return -1

                                elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / elt_exp_handler(i).line(j).k_ratio(k).std_intensity 'calculate the theoretical k-ratio value
                                calculated_y(indice) = elt_exp_handler(i).line(j).k_ratio(k).theo_value ' store the calculated k-ratio in the calculated_y array
                                indice = indice + 1
                            Next
                        Next
                    Next
                Catch ex As Exception
                    MsgBox("Error in fitting module: calculate k-ratios" & vbCrLf & ex.Message)
                    Return -1
                End Try
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
            '******************************************************

            '******************************************************
            'Re-calculates concentations if fixed and defined in atomic fraction
            '******************************************************
            'Dim xx As Double
            'For i As Integer = 0 To UBound(layer_handler)
            '    For j As Integer = 0 To UBound(layer_handler(i).element)
            '        If layer_handler(i).element(j).isConcFixed = True And layer_handler(i).wt_fraction = False Then
            '            Dim temp As Double = 0
            '            For k As Integer = 0 To UBound(layer_handler(i).element)
            '                If k = j Then Continue For
            '                temp = temp + layer_handler(i).element(k).conc_wt / zaro(layer_handler(i).element(k).z)(0)
            '            Next
            '            layer_handler(i).element(j).conc_wt = temp * layer_handler(i).element(j).conc_at_ori / (1 - layer_handler(i).element(j).conc_at_ori) * zaro(layer_handler(i).element(j).z)(0)
            '        End If
            '    Next
            'Next


            '******************************************************





            '******************************************************
            'Sum of the concentrations in each layer equals 1.
            '******************************************************
            Try
                If options.sum_conc_equals_one = True Then 'check if the constraint "sum of the concentrations should be as close as possible to 1" must be used
                    For i As Integer = 0 To layer_handler.Count - 1 'UBound(calculated_y) - indice
                        calculated_y(indice) = 0
                        For j As Integer = 0 To UBound(layer_handler(i).element)
                            calculated_y(indice) = calculated_y(indice) + layer_handler(i).element(j).conc_wt 'add the weight concentrations of all the elements in the current layer
                        Next
                        indice = indice + 1
                    Next
                End If
            Catch ex As Exception
                MsgBox("Error in fitting module: Sum of the concentrations = 1" & vbCrLf & ex.Message)
                Return -1
            End Try
            '******************************************************

            '******************************************************
            'Compare the calculated k-ratios or X-ray intensities to the experimental data.
            'The difference "guides" the fitting algorithm towards the best solution
            '******************************************************
            Try
                If fit_MAC.activated = True Then
                    For i As Integer = 0 To dy.Length - 1
                        dy(i) = (y(i) - calculated_y(i)) 'for the MAC determination, we only calculate the X-ray intensity difference between experimetal and calculated data
                    Next
                Else
                    For i As Integer = 0 To dy.Length - 1
                        dy(i) = (y(i) - calculated_y(i)) / ey(i) 'calculate the X-ray intensity difference between experimetal and calculated data and divide it by the error on the experimental data (this is supposed to give more weigth to experimental data with a small uncertainty).
                    Next
                End If
            Catch ex As Exception
                MsgBox("Error in fitting module: calculate dy" & vbCrLf & ex.Message)
                Return -1
            End Try
            '******************************************************

            'Some debug tests
#If DEBUG Then
            'Dim print_res_debug As String = ""
            'For i As Integer = 0 To dy.Length - 1
            '    print_res_debug = print_res_debug & y(i) - calculated_y(i) & vbTab
            'Next
            'Debug.WriteLine(print_res_debug)

            'calculate the chi-squared value
            Dim test_sum As Double = 0
            For i As Integer = 0 To dy.Length - 1
                test_sum = test_sum + (y(i) - calculated_y(i)) ^ 2
            Next
            Debug.WriteLine(fit_MAC.MAC & vbTab & Math.Sqrt(test_sum))
#End If

            Return 0
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in customfunc " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Public Shared Function customfunc_brem_fluo(p() As Double, dy() As Double, dvec() As IList(Of Double), vars As Object) As Integer 'IList<Double>[] dvec
        Try
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
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in customfunc_brem_fluo " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Public Shared Function customfunc_coeff(p() As Double, dy() As Double, dvec() As IList(Of Double), vars As Object) As Integer 'IList<Double>[] dvec
        Try
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
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in customfunc_coeff " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function
End Class


