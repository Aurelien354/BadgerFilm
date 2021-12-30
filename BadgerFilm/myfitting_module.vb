Imports System.Drawing.Drawing2D

Module myfitting_module
    Public Sub myfit(ByRef x() As Double, ByRef y() As Double, ByRef ey() As Double, ByRef p() As Double, ByRef pars() As MPFitLib.mp_par,
                ByRef buffer_text As String, ByRef layer_handler() As layer, ByRef elt_exp_handler() As Elt_exp, ByRef elt_exp_all() As Elt_exp, ByVal toa As Double,
                   ByVal pen_path As String, ByVal Ec_data() As String, ByRef fit_MAC As fit_MAC, ByVal options As options)


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

        Dim dy(UBound(y)) As Double
        Dim fit_step As Double = 0.1
        Dim iter As Integer = 0
        'Dim result() As Double
        Dim converged As Boolean = False
        Dim conv_criteria As Double = 0.1


        While converged = False

            Dim jac()() As Double = jacobian(p, fit_step, v)

            Dim jacT()() As Double = MatrixTranspose(jac)

            Dim jacTjac()() As Double = MatrixProduct(jacT, jac)

            Dim jacTjacInv()() As Double = MatrixInverse(jacTjac)

            Dim jacTjacInvjacT()() As Double = MatrixProduct(jacTjacInv, jacT)

            meanLeastSquared(p, dy, v)

            Dim jacTjacInvjacTr() As Double = MatrixProduct(jacTjacInvjacT, dy)

            'Dim new_p(p.Length - 1) As Double

            For i As Integer = 0 To p.Length - 1
                p(i) = p(i) - jacTjacInvjacTr(i)
            Next

            Dim new_dy(dy.Length - 1) As Double
            meanLeastSquared(p, new_dy, v)

            converged = True
            For i As Integer = 0 To UBound(new_dy)
                If Math.Abs(new_dy(i) - dy(i)) > conv_criteria Then
                    converged = False
                    Exit For
                End If
            Next

            iter = iter + 1
        End While


    End Sub

    Public Function jacobian(ByVal p() As Double, ByVal fit_step As Double, ByVal vars As Object) As Double()()
        Dim v As CustomUserVariable = vars
        Dim r(v.Y.Length - 1) As Double

        Dim pInit(p.Length - 1) As Double
        p.CopyTo(pInit, 0)
        Dim rInit(r.Length - 1) As Double
        meanLeastSquared(pInit, rInit, v)

        Dim jac(r.Length - 1)() As Double
        For i As Integer = 0 To r.Length - 1
            ReDim jac(i)(p.Length - 1)
        Next

        For j As Integer = 0 To p.Length - 1
            p(j) = p(j) * (1 + fit_step)
            meanLeastSquared(p, r, v)

            For i As Integer = 0 To r.Length - 1
                jac(i)(j) = (rInit(i) - r(i)) / (pInit(j) - p(j))
            Next

            p(j) = pInit(j)
        Next

        Return jac

    End Function


    Public Function meanLeastSquared(p() As Double, dy() As Double, vars As Object) As Integer 'IList<Double>[] dvec
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
                layer_handler(i).element(j).conc_wt = p(count)
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
                            Dim coeff As Double = Convert_wt_to_oxide_wt(layer_handler(i).element(index_C))
                            Dim B As Double = (coeff - 1) * layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio *
                                    layer_handler(i).element(index_C).a / layer_handler(i).element(index_O).a

                            O_wt_conc = O_wt_conc_initial * 1 / (1 - B)
                            Elt_wt_conc = O_wt_conc * layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio * layer_handler(i).element(index_C).a / layer_handler(i).element(index_O).a


                            'Dim O_wt_conc_from_Elt As Double = 0
                            'Dim old_O_wt_conc As Double = 0
                            'Dim O_wt_variations As Double = Math.Abs(old_O_wt_conc - O_wt_conc_initial)
                            'Dim iter As Integer = 0
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
                    dy(i) = (y(i) - calculated_y(i)) ^ 2 'for the MAC determination, we only calculate the X-ray intensity difference between experimetal and calculated data
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
    End Function


    Public Function MatrixInverse(ByVal matrix As Double()()) As Double()()
        Dim n As Integer = matrix.Length
        Dim result As Double()() = MatrixCreate(n, n)
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                result(i)(j) = matrix(i)(j)
                'j = j + 1
            Next
            'i = i + 1
        Next

        Dim lum()() As Double
        Dim perm() As Integer
        Dim toggle As Integer
        toggle = MatrixDecompose(matrix, lum, perm)

        Dim b(n - 1) As Double
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                If i = perm(j) Then
                    b(j) = 1
                Else
                    b(j) = 0
                End If
            Next

            Dim x() As Double = Helper(lum, b)
            For j As Integer = 0 To n - 1
                result(j)(i) = x(j)
            Next
        Next

        Return result
    End Function

    Public Function MatrixCreate(ByVal rows As Integer, ByVal cols As Integer) As Double()()
        Dim result(rows - 1)() As Double
        For i As Integer = 0 To rows - 1
            ReDim result(i)(cols - 1)
        Next
        Return result
    End Function


    Public Function MatrixDecompose(ByVal m()() As Double, ByRef lum()() As Double, ByRef perm() As Integer) As Integer
        ' Crout's LU decomposition for matrix determinant and inverse
        ' stores combined lower & upper in lum[][]
        ' stores row permuations into perm[]
        ' returns +1 Or -1 according to even Or odd number of row permutations
        ' lower gets dummy 1.0s on diagonal (0.0s above)
        ' upper gets lum values on diagonal (0.0s below)

        Dim toggle As Integer = 1 'even (+1) Or odd (-1) row permutatuions
        Dim n As Integer = m.Length

        ' make a copy of m[][] into result lu[][]
        lum = MatrixCreate(n, n)
        For i As Integer = 0 To n - 1
            For j As Integer = 0 To n - 1
                lum(i)(j) = m(i)(j)
            Next
        Next

        'make perm
        ReDim perm(n - 1)
        For i As Integer = 0 To n - 1
            perm(i) = i
        Next

        For j As Integer = 0 To n - 2 'process by column. note n-2
            Dim max As Double = Math.Abs(lum(j)(j))
            Dim piv As Integer = j
            For i As Integer = j + 1 To n - 1 'find pivot index
                Dim xij As Double = Math.Abs(lum(i)(j))
                If xij > max Then
                    max = xij
                    piv = i

                End If
            Next
            If piv <> j Then
                Dim tmp() As Double = lum(piv) 'swap rows j, piv
                lum(piv) = lum(j)
                lum(j) = tmp

                Dim t As Integer = perm(piv) 'swap perm elements
                perm(piv) = perm(j)
                perm(j) = t

                toggle = -toggle
            End If

            Dim xjj As Double = lum(j)(j)
            If xjj <> 0.0 Then
                For i As Integer = j + 1 To n - 1 'find pivot index
                    Dim xij As Double = lum(i)(j) / xjj
                    lum(i)(j) = xij
                    For k As Integer = j + 1 To n - 1
                        lum(i)(k) -= xij * lum(j)(k)
                    Next
                Next
            End If
        Next

        Return toggle
    End Function

    Public Function Helper(ByVal luMatrix()() As Double, ByRef b() As Double) As Double()
        Dim n As Integer = luMatrix.Length
        Dim x(n - 1) As Double
        b.CopyTo(x, 0)

        For i As Integer = 1 To n - 1
            Dim sum As Double = x(i)
            For j As Integer = 0 To i - 1
                sum -= luMatrix(i)(j) * x(j)
            Next
            x(i) = sum
        Next

        x(n - 1) /= luMatrix(n - 1)(n - 1)
        For i As Integer = n - 2 To 0 Step -1
            Dim sum As Double = x(i)
            For j As Integer = i + 1 To n - 1
                sum -= luMatrix(i)(j) * x(j)
            Next
            x(i) = sum / luMatrix(i)(i)
        Next

        Return x
    End Function

    Public Function MatrixDeterminant(ByVal matrix()() As Double) As Double
        Dim lum()() As Double
        Dim perm() As Integer
        Dim toggle As Integer = MatrixDecompose(matrix, lum, perm)
        Dim result As Double = toggle
        For i As Integer = 0 To lum.Length - 1
            result = result * lum(i)(i)
        Next
        Return result
    End Function

    Public Function MatrixProduct(ByVal matrixA()() As Double, ByVal matrixB()() As Double) As Double()()
        Dim aRows As Integer = matrixA.Length
        Dim aCols As Integer = matrixA(0).Length
        Dim bRows As Integer = matrixB.Length
        Dim bCols As Integer = matrixB(0).Length
        If aCols <> bRows Then
            Throw New Exception("Non-conformable matrices")
        End If

        Dim result()() As Double = MatrixCreate(aRows, bCols)

        For i As Integer = 0 To aRows - 1
            For j As Integer = 0 To bCols - 1
                For k As Integer = 0 To aCols - 1 'could use k < bRows
                    result(i)(j) += matrixA(i)(k) * matrixB(k)(j)
                Next
            Next
        Next

        Return result
    End Function

    Public Function MatrixProduct(ByVal matrixA()() As Double, ByVal matrixB() As Double) As Double()
        Dim aRows As Integer = matrixA.Length
        Dim aCols As Integer = matrixA(0).Length
        Dim bRows As Integer = matrixB.Length

        If aCols <> bRows Then
            Throw New Exception("Non-conformable matrices")
        End If

        Dim result(aRows - 1) As Double

        For i As Integer = 0 To aRows - 1
            For k As Integer = 0 To aCols - 1 'could use k < bRows
                result(i) += matrixA(i)(k) * matrixB(k)
            Next
        Next

        Return result
    End Function

    Public Function MatrixTranspose(ByVal matrix()() As Double) As Double()()
        Dim rows As Integer = matrix.Length
        Dim cols As Integer = matrix(0).Length

        Dim transpose(cols - 1)() As Double
        For i As Integer = 0 To cols - 1
            ReDim transpose(i)(rows - 1)
        Next

        For i As Integer = 0 To rows - 1
            For j As Integer = 0 To cols - 1
                transpose(j)(i) = matrix(i)(j)
            Next
        Next

        Return transpose
    End Function

    Public Sub test_matrix_inverse()
        Dim m()() As Double = MatrixCreate(4, 4)
        m(0)(0) = 3.0
        m(0)(1) = 7.0
        m(0)(2) = 2.0
        m(0)(3) = 5.0

        m(1)(0) = 1.0
        m(1)(1) = 8.0
        m(1)(2) = 4.0
        m(1)(3) = 2.0

        m(2)(0) = 2.0
        m(2)(1) = 1.0
        m(2)(2) = 9.0
        m(2)(3) = 3.0

        m(3)(0) = 5.0
        m(3)(1) = 4.0
        m(3)(2) = 7.0
        m(3)(3) = 1.0


        Dim inv()() As Double = MatrixInverse(m)

        Dim prod()() As Double = MatrixProduct(m, inv)

        Stop
    End Sub


End Module
