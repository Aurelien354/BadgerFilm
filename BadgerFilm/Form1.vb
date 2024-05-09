Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Structure layer
    Dim id As Integer 'identificator of the layer. Usually, top layer = 0.
    Dim thickness As Single 'in Angstrom
    Dim density As Single 'in g/cm3
    Dim mass_thickness As Double 'in g/cm²
    Dim element() As Elt_layer
    Public wt_fraction As Boolean
    Dim stoichiometry As stoichiometry
    Public isfix As Boolean
End Structure

Public Structure data_xs
    Dim Z As Integer
    Dim energy() As Double
    Dim cross_section(,) As Double
End Structure

Public Structure data_atomic_parameters
    Dim Z As Integer
    Dim shell1() As Integer
    Dim shell2() As Integer
    Dim shell3() As Integer
    Dim energy() As Double
    Dim transition_probability() As Double
End Structure

Public Structure k_ratio
    Public experimental_value As Double
    Public err_experimental_value As Double
    Public theo_value As Double
    Public kv As Double
    Public elt_intensity As Double
    Public std_intensity As Double
End Structure

Public Structure options
    Dim phi_rz_mode As String
    Dim MAC_mode As String
    Dim ionizationXS_mode As String
    Dim char_fluo_flag As Boolean
    Dim brem_fluo_flag As Boolean
    Dim sum_conc_equals_one As Boolean
    Dim experimental_MAC As experimental_MAC
    Dim BgWorker As BackgroundWorker
End Structure

Public Structure experimental_MAC
    Dim experimental_MAC_enabled As Boolean
    Dim emitter() As String
    Dim xray_line() As String
    Dim absorber() As String
    Dim exp_MAC() As Double
End Structure

Public Structure stoichiometry
    Dim O_by_stoichio_name As String
    Dim O_by_stoichio As Boolean
    Dim O_wt_conc As Double
    Dim Elt_by_stoichio_to_O As Boolean
    Dim Elt_by_stoichio_to_O_name As String
    Dim Elt_by_stoichio_to_O_ratio As Double
    Dim Elt_wt_conc As Double
    Dim stoichio_table(,) As Integer
End Structure

Public Structure Elt_exp
    Public z As Integer
    Public a As Single
    Public elt_name As String
    Public line() As Line
    Public el_ion_xs As data_xs
    Public ph_ion_xs As data_xs
    Public mac_data As data_xs
    Public at_data As data_atomic_parameters
    Public Ec_data() As String
End Structure

Public Structure Line
    Public xray_name As String
    Public xray_energy As Single
    Public Ec As Double
    Public k_ratio() As k_ratio
    Public std_filename As String
    Public std As String
End Structure

Public Structure Elt_layer
    Public mother_layer_id As Integer
    Public z As Integer
    Public a As Single
    'Public concentration As Single
    Public normalized_concentration As Double
    Public fictitious_concentration As Double
    Public elt_name As String
    Public conc_wt As Double
    Public conc_at As Double
    Public conc_at_ori As Double 'original atomic concentration in case the layer is defined by atomic fraction and it is fixed.
    Public isConcFixed As Boolean
End Structure

Public Structure data_to_plot
    Public energy() As Double
    Public k_ratio(,) As Double
    Public elts_name() As String
End Structure

Public Structure fit_MAC
    Public absorbed_elt As String
    Public X_ray As String
    Public X_ray_energy As Double
    Public MAC As Double
    Public scaling_factor As Double
    Public absorber_elt As String
    Public activated As Boolean
    Public norm_kV As Double
    Public compound_MAC As Boolean
End Structure

Public Class Form1
    Public VERSION As String = "v.1.2.28"
    Public options As options
    Dim pen_path As String = Application.StartupPath() & "\PenelopeData" '"D:\Travail\Penelope"
    Dim eadl_path As String = Application.StartupPath() & "\EADL" '"D:\Travail\Penelope"
    Dim ffast_path As String = Application.StartupPath() & "\FFAST" '"D:\Travail\Penelope"
    Dim epdl23_path As String = Application.StartupPath() & "\EPDL23"
    Dim experimental_MAC_path As String = Application.StartupPath & "\Experimental_MACs.txt"
    '
    Public color_table() As String = {"Red", "Blue", "Green", "Orange", "Purple", "Pink", "Black", "Gray"}

    Public Shared layer_handler() As layer
    Public Shared elt_exp_handler() As Elt_exp
    Dim elt_exp_all() As Elt_exp = Nothing

    Public Const DEFAULT_THICKNESS As Double = 20
    Public loaded As Boolean = False

    Public save_results As String = ""

    Dim chart1_width As Integer
    Dim chart1_height As Integer
    Dim chart1_location As Point
    Public graph_limits(3) As Double

    Public toa As Double = 40

    Public at_data() As data_atomic_parameters
    Public el_ion_xs()() As String
    Public ph_ion_xs()() As String
    Public MAC_data_PEN14()() As String
    Public MAC_data_PEN18()() As String
    Public MAC_data_FFAST()() As String
    Public MAC_data_EPDL23()() As String
    Public MAC_data()() As String

    Public Ec_data() As String = Nothing

    Public stoichio_O(,) As Integer = Nothing
    Public stoichio_N(,) As Integer = Nothing

    'Public WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker


    Public Sub plot_kratio(ByVal xmin As Double, ByVal xmax As Double, ByVal istep As Integer, ByVal layer_handler() As layer, ByVal toa As Double,
                           ByVal Ec_data() As String, ByVal pen_path As String, ByVal print_res As Boolean, ByRef save_results As String, ByVal fit_MAC As fit_MAC,
                           ByRef data_to_plot As data_to_plot)
        Try
            Dim total_elts As Integer = 0
            'For i As Integer = 0 To UBound(layer_handler)
            '    For j As Integer = 0 To UBound(layer_handler(i).element)
            '        If layer_handler(i).element(j).k_ratio IsNot Nothing Then
            '            total_elts = total_elts + 1
            '        End If
            '    Next
            'Next
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    If elt_exp_handler(i).line(j).k_ratio IsNot Nothing Then
                        total_elts = total_elts + 1
                    End If
                Next
            Next


            ReDim data_to_plot.energy(istep - 1)
            ReDim data_to_plot.k_ratio(total_elts - 1, istep - 1)
            ReDim data_to_plot.elts_name(total_elts - 1)

            Dim _x As Double
            'Dim curr_element As Integer = 0

            Dim cnt As Integer = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)

                    Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(i).line(j).std_filename) 'AMXXXXXXXXXXXXXXXXXX

                    If fit_MAC.activated = True Then
                        Dim norm As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, fit_MAC.norm_kV, toa, Ec_data, options, print_res, save_results, fit_MAC)
                        If norm < 0 Then Exit Sub

                        For ll As Integer = 0 To istep - 1
                            _x = xmin + (xmax - xmin) / (istep - 1) * ll
                            'If _x = 10 Then
                            '    Stop
                            'End If
                            '_x = 19
                            data_to_plot.energy(ll) = _x

                            Dim Ix_unk As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, _x, toa, Ec_data, options, print_res, save_results, fit_MAC)
                            If Ix_unk < 0 Then Exit Sub

                            'data_to_plot.k_ratio(cnt, ll) = Ix_unk / norm
                            data_to_plot.k_ratio(cnt, ll) = Ix_unk * fit_MAC.scaling_factor '* Ix_unk
                        Next

                    ElseIf IsNothing(elt_exp_handler(i).line(j).std_filename) = True Or flag_file_exists = False Then
                        If elt_exp_handler(i).line(j).std_filename <> "" And flag_file_exists = False Then
                            MsgBox("Standard file not found for: " & elt_exp_handler(i).elt_name & vbCrLf & "Calculating intensities relative to pure standards.")
                        End If
                        For ll As Integer = 0 To istep - 1
                            _x = xmin + (xmax - xmin) / (istep - 1) * ll
                            If _x = 9.75 Then
                                Stop
                            End If
                            '_x = 19
                            data_to_plot.energy(ll) = _x
                            Dim Ix_std As Double = init_pure_std(elt_exp_handler(i), j, _x, toa, Ec_data)
                            'Debug.Print(options.phi_rz_mode & vbTab & "Std: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab & elt_exp_handler(i).line(j).k_ratio(kk).std_intensity)

                            Dim Ix_unk As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, _x, toa, Ec_data, options, print_res, save_results, fit_MAC)
                            If Ix_unk < 0 Then Exit Sub
                            If Ix_std = 0 Then
                                data_to_plot.k_ratio(cnt, ll) = 0
                            Else
                                data_to_plot.k_ratio(cnt, ll) = Ix_unk / Ix_std
                            End If
                        Next
                    Else
                        Dim layer_handler_std() As layer = Nothing
                        Dim elt_exp_handler_std() As Elt_exp = Nothing
                        Dim toa_std As Double
                        load_data(elt_exp_handler(i).line(j).std_filename, layer_handler_std, elt_exp_handler_std, toa_std)
                        Dim elt_exp_all_std() As Elt_exp = Nothing
                        init_elt_exp_all(elt_exp_all_std, layer_handler_std, Ec_data, pen_path)
                        'Normalize the concentrations of the standard
                        For ll As Integer = 0 To UBound(layer_handler_std)
                            Dim sum_conc_wt As Double = 0
                            For jj As Integer = 0 To UBound(layer_handler_std(ll).element)
                                sum_conc_wt = sum_conc_wt + layer_handler_std(ll).element(jj).conc_wt
                            Next
                            For jj As Integer = 0 To UBound(layer_handler_std(ll).element)
                                layer_handler_std(ll).element(jj).conc_wt = layer_handler_std(ll).element(jj).conc_wt / sum_conc_wt
                            Next
                        Next
                        For ll As Integer = 0 To istep - 1
                            _x = xmin + (xmax - xmin) / (istep - 1) * ll
                            data_to_plot.energy(ll) = _x
                            'If _x = 15 And elt_exp_handler(i).z = 22 Then
                            '    Stop
                            'End If

                            Dim Ix_std As Double = pre_auto(layer_handler_std, elt_exp_handler(i), j, elt_exp_all_std, _x, toa, Ec_data, options, False, save_results, fit_MAC) 'Oct 21.2021 AMXX changed toa_std to toa
                            If Ix_std < 0 Then Exit Sub
                            Dim Ix_unk As Double = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all, _x, toa, Ec_data, options, print_res, save_results, fit_MAC)
                            If Ix_unk < 0 Then Exit Sub
                            If Ix_std = 0 Then
                                data_to_plot.k_ratio(cnt, ll) = 0
                            Else
                                data_to_plot.k_ratio(cnt, ll) = Ix_unk / Ix_std
                            End If

                        Next
                    End If

                    data_to_plot.elts_name(cnt) = elt_exp_handler(i).elt_name & " " & elt_exp_handler(i).line(j).xray_name

                    cnt = cnt + 1
                Next
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in plot-kratio " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub


    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Try
            If IsNumeric(TextBox3.Text) = False Then Exit Sub
            Dim number_layers As Integer = CInt(TextBox3.Text)
            Dim old_number_layers As Integer = ListBox1.Items.Count
            ListBox1.Items.Clear()

            If number_layers < 1 Then
                MsgBox("The number of layers (with the substrate) must be at least 1.")
                Exit Sub
            End If

            If number_layers > 1 Then
                CheckBox22.Enabled = False
                CheckBox22.Checked = False
                CheckBox10.Checked = True
            Else
                CheckBox22.Enabled = True
            End If

            For i As Integer = 1 To number_layers - 1
                ListBox1.Items.Add("Layer " & i)
            Next
            ListBox1.Items.Add("Substrate")


            If layer_handler Is Nothing Then
                ReDim layer_handler(number_layers - 1)
            Else
                'Update the display after a loading.
                If layer_handler.Count = TextBox3.Text Then
                    ListBox1.SelectedIndex = 0
                    Exit Sub
                End If

                'Creates new layers (old layers are kept).
                Dim layer_handler_tmp(UBound(layer_handler)) As layer
                For i As Integer = 0 To UBound(layer_handler)
                    layer_handler_tmp(UBound(layer_handler) - i) = layer_handler(i)
                Next

                ReDim Preserve layer_handler_tmp(number_layers - 1)
                ReDim layer_handler(number_layers - 1)

                For i As Integer = 0 To UBound(layer_handler)
                    layer_handler(UBound(layer_handler) - i) = layer_handler_tmp(i)

                Next

                For i As Integer = 0 To UBound(layer_handler)
                    If layer_handler(i).element Is Nothing Then Continue For
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        layer_handler(i).element(j).mother_layer_id = i
                    Next
                Next
            End If

            'Populates the new layers with default values.
            For i As Integer = 0 To number_layers - old_number_layers - 1
                If IsNumeric(TextBox4.Text) = True Then
                    layer_handler(i).density = TextBox4.Text
                Else
                    layer_handler(i).density = 2.2
                End If
                layer_handler(i).wt_fraction = True
                layer_handler(i).thickness = DEFAULT_THICKNESS
                layer_handler(i).stoichiometry.Elt_by_stoichio_to_O = False
                layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_name = ""
                layer_handler(i).stoichiometry.Elt_by_stoichio_to_O_ratio = 0
                layer_handler(i).stoichiometry.Elt_wt_conc = 0
                layer_handler(i).stoichiometry.O_by_stoichio = False
                layer_handler(i).stoichiometry.O_by_stoichio_name = ""
                layer_handler(i).stoichiometry.O_wt_conc = 0
            Next

            ListBox1.SelectedIndex = 0
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox3_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'On récupère le séparateur qui est utilisé sur la station de travail
            Dim oldDecimalSeparator As String = Application.CurrentCulture.NumberFormat.NumberDecimalSeparator

            'On compare le séparateur instancié avec le point
            If oldDecimalSeparator = "." Then
                'Le séparateur instancié dans le panneau de configuration est le point : "."
            Else
                'Le séparateur instancié dans le panneau de configuration est la virgule : ","
                Dim forceDotCulture As CultureInfo

                'Code un peu louche il faut avouer, mais il faut faire avec car le framework pose problème
                'ici; en effet, il faut cloner la culture pour pouvoir modifier les paramètres de l'application
                'car sinon la culture de base est en lecture seule.
                forceDotCulture = Application.CurrentCulture.Clone()

                'On affecte le point : "." comme paramètre de séparateur décimal
                forceDotCulture.NumberFormat.NumberDecimalSeparator = "."

                'Là, on affecte l'application cloné à celle où l'on travaille 
                'C'est un passage flou car en fait, l'appli est en mode readonly et l'on ne peut pas
                'la modifier directement, d'où cette affectation
                Application.CurrentCulture = forceDotCulture

            End If


            Me.FormBorderStyle = FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False

            options.phi_rz_mode = "PAP"
            options.MAC_mode = "PENELOPE2018"
            options.ionizationXS_mode = "Bote" '"OriPAP"
            options.brem_fluo_flag = True
            options.char_fluo_flag = True
            options.sum_conc_equals_one = True
            options.experimental_MAC.experimental_MAC_enabled = True
            options.BgWorker = BackgroundWorker1

            Me.ListBox1.Items.Clear()
            Me.ListBox1.Items.Add("Substrate")
            CheckBox18.Checked = True
            creat_mendeleiev_table()

            GroupBox9.Visible = False

            loaded = True
            ListBox1.SelectedIndex = 0


            Dim files_EADL As String = "EADL"
            Dim fit_dll_file As String = "MPFitLib.dll"
            Dim experimental_MAC_file As String = "Experimental_MACs.txt"
            Dim stoichio_O_file As String = "stoichiometry_O.txt"
            Dim stoichio_N_file As String = "stoichiometry_N.txt"
            'Dim list_files_PEN() As String = {"pdatconf.p14", "pdesi", "phmaxs", "phpixs"}
            Dim flag_file_exists As Boolean = True

            If check_Pen_files("PENELOPE2018") = False Then flag_file_exists = False

            If check_Pen_files("PENELOPE2014") = False Then flag_file_exists = False


            For j As Integer = 1 To 99
                If My.Computer.FileSystem.FileExists(eadl_path & "\" & files_EADL & Format(j, "00") & ".txt") = False Then
                    flag_file_exists = False
                    Exit For
                End If
            Next

            If flag_file_exists = False Then
                MsgBox("File missing in the database!" & vbCrLf & "Please reinstall BadgerFilm.")
            End If


            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\" & fit_dll_file) = False Then
                MsgBox("Fitting DLL missing!" & vbCrLf & "Please reinstall BadgerFilm.")
            End If

            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\" & experimental_MAC_file) = False Then
                MsgBox("The file Experimental_MACs.txt is missing!" & vbCrLf & "Please reinstall BadgerFilm or create the file manually in BadgerFilm's folder.")
            End If

            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\" & stoichio_O_file) = False Then
                MsgBox("The file stoichiometry_O.txt is missing!" & vbCrLf & "Please reinstall BadgerFilm or create the file manually in BadgerFilm's folder.")
            End If

            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\" & stoichio_N_file) = False Then
                MsgBox("The file stoichiometry_N.txt is missing!" & vbCrLf & "Please reinstall BadgerFilm or create the file manually in BadgerFilm's folder.")
            End If


            chart1_height = Chart1.Height
            chart1_width = Chart1.Width
            chart1_location = Chart1.Location

            graph_limits(0) = 0
            graph_limits(1) = 30
            graph_limits(2) = 0
            graph_limits(3) = 1

            Me.Text = "BadgerFilm " & VERSION

            init_atomic_parameters(pen_path, eadl_path, ffast_path, epdl23_path, at_data, el_ion_xs, ph_ion_xs, MAC_data_PEN14, MAC_data_PEN18, MAC_data_FFAST, MAC_data_EPDL23, options)

            init_Ec(Ec_data, pen_path)

            init_stoichio(stoichio_O, stoichio_O_file)
            init_stoichio(stoichio_N, stoichio_N_file)

            If ComboBox1.Items.Count > 0 Then
                ComboBox1.SelectedIndex = 0    ' The first item has index 0 '
            End If

            '***************************************************
            ' This block must be called last. Thanks yueyinqiu for finding this bug.
            If HaveInternetConnection() = False Then Exit Sub

            Dim version_check As String = Nothing
            update_get_version(version_check)
            'If version_check = "" Then Exit Sub
            If compare_version(version_check, VERSION) Then
                Label1.Text = "Status:  New version available (" & version_check & ")" '. Please click the Update button to download the latest version."
            End If
            '***************************************************

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Form1_Load " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Function check_Pen_files(ByVal sub_folder As String) As Boolean
        Try
            Dim list_files_PEN() As String = {"pdatconf.p14", "pdesi", "phmaxs", "phpixs"}
            'check_Pen_files = True

            If sub_folder.Last <> "\" Then sub_folder = sub_folder & "\"

            If My.Computer.FileSystem.FileExists(pen_path & "\" & sub_folder & list_files_PEN(0)) = False Then
                'check_Pen_files = False
                Return False
            End If

            For j As Integer = 1 To 99
                If My.Computer.FileSystem.FileExists(pen_path & "\" & sub_folder & list_files_PEN(1) & Format(j, "00") & ".p14") = False Then
                    'check_Pen_files = False
                    Return False
                End If
            Next

            For j As Integer = 1 To 99
                If My.Computer.FileSystem.FileExists(pen_path & "\" & sub_folder & list_files_PEN(2) & Format(j, "00")) = False Then
                    'check_Pen_files = False
                    Return False
                End If
            Next

            For j As Integer = 1 To 99
                If My.Computer.FileSystem.FileExists(pen_path & "\" & sub_folder & list_files_PEN(3) & Format(j, "00")) = False Then
                    'check_Pen_files = False
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in check_Pen_files " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            display_grid_layer()
            display_grid()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in ListBox1_SelectedIndexChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub display_grid_layer()
        Try
            If check_valid_layer_selected() = False Then Exit Sub

            DataGridView2.Rows.Clear()
            TextBox4.Text = layer_handler(ListBox1.SelectedIndex).density
            CheckBox18.Checked = layer_handler(ListBox1.SelectedIndex).isfix
            CheckBox20.Checked = layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio
            CheckBox21.Checked = layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O

            If layer_handler(ListBox1.SelectedIndex).isfix = True Then
                TextBox5.BackColor = Color.LightBlue
            End If

            If layer_handler(ListBox1.SelectedIndex).wt_fraction = True Then
                CheckBox12.Checked = True
                CheckBox17.Checked = False
                DataGridView2.Columns(1).HeaderText = "conc (wt.)"
            Else
                CheckBox17.Checked = True
                CheckBox12.Checked = False
                DataGridView2.Columns(1).HeaderText = "conc (at.)"
            End If

            If ListBox1.SelectedItem = "Substrate" Then
                TextBox5.Text = "Inf"
                layer_handler(ListBox1.SelectedIndex).thickness = 1000000000.0
            Else
                If CheckBox15.Checked = True Then
                    TextBox5.Text = layer_handler(ListBox1.SelectedIndex).thickness
                Else
                    TextBox5.Text = layer_handler(ListBox1.SelectedIndex).density * layer_handler(ListBox1.SelectedIndex).thickness * 10 ^ -8 * 10 ^ 6
                End If
            End If

            If layer_handler(ListBox1.SelectedIndex).element Is Nothing Then
                select_elt()
                Exit Sub
            End If

            Dim ind_color_cell As Integer = 0
            For i As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                With layer_handler(ListBox1.SelectedIndex).element(i)
                    If layer_handler(ListBox1.SelectedIndex).wt_fraction = True Then
                        DataGridView2.Rows.Add(.elt_name, Format(.conc_wt, "0.0000"))
                    Else
                        DataGridView2.Rows.Add(.elt_name, Format(.conc_at, "0.0000"))
                        Dim tot_at As Double = 0
                        For nn As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                            tot_at = tot_at + layer_handler(ListBox1.SelectedIndex).element(nn).conc_at * zaro(symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(nn).elt_name))(0)
                        Next
                        .conc_wt = .conc_at * zaro(symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(i).elt_name))(0) / tot_at


                    End If
                    If .isConcFixed = True Then
                        DataGridView2.Item(1, ind_color_cell).Style.BackColor = Color.LightBlue
                    End If
                    ind_color_cell = ind_color_cell + 1
                End With
            Next

            select_elt()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in display_grid_layer " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub
    Public Sub display_grid()
        Try
            If elt_exp_handler Is Nothing Then Exit Sub
            DataGridView1.Rows.Clear()
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(2).DefaultCellStyle.BackColor = Color.Gray

            For i As Integer = 0 To UBound(elt_exp_handler)
                With elt_exp_handler(i)
                    If elt_exp_handler(i).line IsNot Nothing Then
                        For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                            For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                                DataGridView1.Rows.Add(.elt_name, .line(j).xray_name, Format(.line(j).k_ratio(k).theo_value, "0.0000"), .line(j).k_ratio(k).experimental_value,
                                                       .line(j).k_ratio(k).err_experimental_value, .line(j).k_ratio(k).kv, Split(.line(j).std_filename, "\").Last)
                            Next
                        Next
                    End If
                End With
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in display_grid " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub select_elt()
        Try
            For Each b As Control In Me.Controls
                If TypeOf (b) IsNot Class1.TestB Then
                    Continue For
                End If
                b.BackColor = Color.White
            Next

            If check_valid_layer_selected() = False Then Exit Sub
            If layer_handler(ListBox1.SelectedIndex).element Is Nothing Then Exit Sub

            For i As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                Dim name As String = "Element" & symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(i).elt_name)
                Dim b As Control() = Me.Controls.Find(name, True)
                b(0).BackColor = Color.Gray
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in select_elt " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Public Sub creat_mendeleiev_table()
        Try
            For Each b As Control In Me.Controls
                If TypeOf (b) Is Class1.TestB Then
                    AddHandler b.Click, AddressOf ButtonTableau_Click
                End If
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in creat_mendeleiev_table " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub ButtonTableau_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles Element1.Click
        Try
            If check_valid_layer_selected() = False Then Exit Sub

            Dim tmp() As String

            tmp = Split(sender.name, "t")
            'Stop
            'If tmp(0) = "xray" Then
            Dim selected_elt_name As String = elt_list(tmp(1) - 1)

            'Else
            If (DirectCast(sender, Button).BackColor = Color.White) Then
                DirectCast(sender, Button).BackColor = Color.Gray
                'DataGridView1.Rows.Add(elt_list(tmp(1) - 1), Nothing, Nothing, Nothing, TextBox2.Text)
                If layer_handler(ListBox1.SelectedIndex).element Is Nothing Then
                    ReDim layer_handler(ListBox1.SelectedIndex).element(0)
                Else
                    ReDim Preserve layer_handler(ListBox1.SelectedIndex).element(UBound(layer_handler(ListBox1.SelectedIndex).element) + 1)
                End If


                Dim current_indice As Integer = UBound(layer_handler(ListBox1.SelectedIndex).element)
                layer_handler(ListBox1.SelectedIndex).element(current_indice).elt_name = selected_elt_name
                layer_handler(ListBox1.SelectedIndex).element(current_indice).conc_wt = 1
                layer_handler(ListBox1.SelectedIndex).element(current_indice).conc_at = 1
                layer_handler(ListBox1.SelectedIndex).element(current_indice).isConcFixed = False
                layer_handler(ListBox1.SelectedIndex).element(current_indice).z = symbol_to_Z(selected_elt_name)
                layer_handler(ListBox1.SelectedIndex).element(current_indice).a = zaro(layer_handler(ListBox1.SelectedIndex).element(current_indice).z)(0)
                layer_handler(ListBox1.SelectedIndex).element(current_indice).mother_layer_id = ListBox1.SelectedIndex


                Dim flag_found_elt As Boolean = False
                If elt_exp_handler Is Nothing Then
                    ReDim elt_exp_handler(0)
                    elt_exp_handler(0).elt_name = selected_elt_name
                    ReDim elt_exp_handler(0).line(0)
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                Else
                    For i As Integer = 0 To UBound(elt_exp_handler)
                        If selected_elt_name = elt_exp_handler(i).elt_name Then
                            flag_found_elt = True
                            Exit For
                        End If
                    Next
                    If flag_found_elt = False Then
                        ReDim Preserve elt_exp_handler(UBound(elt_exp_handler) + 1)
                        elt_exp_handler(UBound(elt_exp_handler)).elt_name = selected_elt_name
                        ReDim elt_exp_handler(UBound(elt_exp_handler)).line(0)
                        ReDim elt_exp_handler(UBound(elt_exp_handler)).line(0).k_ratio(0)
                    End If
                End If

                If flag_found_elt = False Then
                    Dim z As Integer = symbol_to_Z(selected_elt_name)
                    If z < 37 Then
                        elt_exp_handler(UBound(elt_exp_handler)).line(0).xray_name = "Ka"
                    ElseIf z < 87 Then
                        elt_exp_handler(UBound(elt_exp_handler)).line(0).xray_name = "La"
                    Else
                        elt_exp_handler(UBound(elt_exp_handler)).line(0).xray_name = "Ma"
                    End If
                    elt_exp_handler(UBound(elt_exp_handler)).line(0).k_ratio(0).kv = 15
                End If

            Else
                DirectCast(sender, Button).BackColor = Color.White
                For i As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                    If selected_elt_name = layer_handler(ListBox1.SelectedIndex).element(i).elt_name Then
                        For j As Integer = i To UBound(layer_handler(ListBox1.SelectedIndex).element) - 1
                            layer_handler(ListBox1.SelectedIndex).element(j) = layer_handler(ListBox1.SelectedIndex).element(j + 1)
                        Next
                        ReDim Preserve layer_handler(ListBox1.SelectedIndex).element(UBound(layer_handler(ListBox1.SelectedIndex).element) - 1)

                        'remove_elt(analysis_cond_handler(ListBox1.SelectedIndex).elts, i)
                        Exit For
                    End If
                Next

                Dim flag_found_elt As Boolean = False
                For i As Integer = 0 To UBound(layer_handler)
                    If layer_handler(i).element Is Nothing Then Continue For
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        If layer_handler(i).element(j).elt_name = selected_elt_name Then
                            flag_found_elt = True
                            Exit For
                        End If
                    Next
                    If flag_found_elt = True Then Exit For
                Next

                If flag_found_elt = False Then
                    For i As Integer = 0 To UBound(elt_exp_handler)
                        If elt_exp_handler(i).elt_name = selected_elt_name Then
                            For j As Integer = i To UBound(elt_exp_handler) - 1
                                elt_exp_handler(j) = elt_exp_handler(j + 1)
                            Next
                            ReDim Preserve elt_exp_handler(UBound(elt_exp_handler) - 1)
                            Exit For
                        End If
                    Next
                End If

                'display_grid_layer()
                'display_grid()

            End If

            display_grid_layer()
            display_grid()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in ButtonTableau_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    'Public Function check_dbl_elt(ByVal analysis_cond_handler() As analysis_conditions, ByVal analysed_elt As analyzed_element, ByVal pos() As Integer) As String
    '    If analysis_cond_handler Is Nothing Then Exit Function
    '    For i As Integer = 0 To UBound(analysis_cond_handler)
    '        If analysis_cond_handler(i).elts Is Nothing Then Continue For
    '        For j As Integer = 0 To UBound(analysis_cond_handler(i).elts)
    '            If i = pos(0) And j = pos(1) Then
    '                Continue For
    '            End If
    '            If analysis_cond_handler(i).elts(j).name = analysed_elt.name And analysis_cond_handler(i).elts(j).line = analysed_elt.line Then
    '                check_dbl_elt = i & " " & j
    '                Exit Function
    '            End If
    '        Next
    '    Next
    '    check_dbl_elt = ""
    'End Function

    'Public Sub remove_elt(ByRef elts() As analyzed_element, ByVal ind As Integer)
    '    For i As Integer = ind To UBound(elts) - 1
    '        elts(i) = elts(i + 1)
    '    Next
    '    ReDim Preserve elts(UBound(elts) - 1)
    '    'Stop
    'End Sub

    Private Sub DataGridView1_CellClicked(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            If elt_exp_handler Is Nothing Or e.ColumnIndex < 0 Or e.RowIndex < 0 Or ListBox1.SelectedIndex < 0 Then
                Exit Sub
            End If

            'Button5.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value

            If e.ColumnIndex = 6 Then
                If DataGridView1.SelectedCells.Count = 0 Then Exit Sub
                Dim OpenFileDialog1 As New OpenFileDialog()
                Dim data_file As String = Nothing

                OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                OpenFileDialog1.FilterIndex = 1
                OpenFileDialog1.RestoreDirectory = True
                OpenFileDialog1.Title = "Load data"
                OpenFileDialog1.AddExtension = True
                OpenFileDialog1.DefaultExt = ".txt"

                Dim filename As String = ""
                If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    filename = OpenFileDialog1.FileName
                Else
                    filename = ""
                End If

                Dim indice As Integer = 0
                'For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '    For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis)
                '        If DataGridView1.SelectedCells(0).RowIndex = indice Then
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(j).std_filename(k) = filename
                '            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = filename 'Split(filename, "\").Last
                '        End If
                '        indice = indice + 1
                '    Next
                'Next

                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).std_filename = filename
                                DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = filename
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next


            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in DataGridView1_CellClicked " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub DataGridView2_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        Try
            If layer_handler Is Nothing Or e.ColumnIndex <= 0 Or e.RowIndex < 0 Or ListBox1.SelectedIndex < 0 Then
                'no valid layer selected
                Exit Sub
            End If

            If layer_handler(ListBox1.SelectedIndex).wt_fraction = False Then
                layer_handler(ListBox1.SelectedIndex).element(e.RowIndex).conc_at = DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Dim tot_at As Double = 0
                For nn As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                    tot_at = tot_at + layer_handler(ListBox1.SelectedIndex).element(nn).conc_at * zaro(symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(nn).elt_name))(0)
                Next
                For nn As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                    layer_handler(ListBox1.SelectedIndex).element(nn).conc_wt = layer_handler(ListBox1.SelectedIndex).element(nn).conc_at *
                        zaro(symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(nn).elt_name))(0) / tot_at
                Next
                'layer_handler(ListBox1.SelectedIndex).element(e.RowIndex).conc_wt = layer_handler(ListBox1.SelectedIndex).element(e.RowIndex).conc_at *
                '    zaro(symbol_to_Z(layer_handler(ListBox1.SelectedIndex).element(e.RowIndex).elt_name))(0) / tot_at
            Else
                layer_handler(ListBox1.SelectedIndex).element(e.RowIndex).conc_wt = DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                convert_wt_to_at(layer_handler, ListBox1.SelectedIndex)
            End If

            DataGridView2.CurrentCell = DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in DataGridView2_CellValueChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Try
            If elt_exp_handler Is Nothing Or e.ColumnIndex < 0 Or e.RowIndex < 0 Or ListBox1.SelectedIndex = -1 Then
                Exit Sub
            End If
            'If e.ColumnIndex = 0 Then
            '    For i As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
            '        If analysis_cond_handler(ListBox1.SelectedIndex).elts(i).name = DataGridView1.Rows(e.RowIndex).Cells(0).Value Then
            '            analysis_cond_handler(ListBox1.SelectedIndex).elts(i).name = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '            display_grid()
            '            Exit Sub
            '        End If
            '    Next
            'End If

            If e.ColumnIndex = 1 Then
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).xray_name = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next
            End If
            If e.ColumnIndex = 2 Then
                'MsgBox("Should not happen!")
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).k_ratio(k).theo_value = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next

                'For i As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '    If analysis_cond_handler(ListBox1.SelectedIndex).elts(i).name = DataGridView1.Rows(e.RowIndex).Cells(0).Value Then

                '        'Dim concentration_in_at As Double = 0
                '        If analysis_cond_handler(ListBox1.SelectedIndex).wt_fraction = False Then
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(i).conc_at = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                '            Dim tot_at As Double = 0
                '            For nn As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '                tot_at = tot_at + analysis_cond_handler(ListBox1.SelectedIndex).elts(nn).conc_at * zaro(symbol_to_Z(analysis_cond_handler(ListBox1.SelectedIndex).elts(nn).name))(0)
                '            Next
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(i).conc_wt = analysis_cond_handler(ListBox1.SelectedIndex).elts(i).conc_at *
                '                zaro(symbol_to_Z(analysis_cond_handler(ListBox1.SelectedIndex).elts(i).name))(0) / tot_at
                '        Else
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(i).conc_wt = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                '        End If
                '        display_grid()
                '        Exit Sub
                '    End If
                'Next
            End If

            If e.ColumnIndex = 3 Then
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).k_ratio(k).experimental_value = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next
                'If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then Exit Sub
                'Dim indice As Integer = 0
                'For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '    For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio)
                '        If e.RowIndex = indice Then
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio(k) = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                '        End If
                '        indice = indice + 1
                '    Next
                'Next
            End If

            If e.ColumnIndex = 4 Then
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next
            End If

            If e.ColumnIndex = 5 Then
                If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then Exit Sub
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                elt_exp_handler(i).line(j).k_ratio(k).kv = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next
                'Dim indice As Integer = 0
                'If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False And DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value <> Nothing Then
                '    MsgBox("Non numeric value!")
                '    For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '        For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured)
                '            If e.RowIndex = indice Then
                '                DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured(k)
                '            End If
                '            indice = indice + 1
                '        Next
                '    Next
                '    Exit Sub
                'End If

                'For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '    For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured)
                '        If e.RowIndex = indice Then
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured(k) = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                '            chg_meas_kr_for_duplicate_elt(analysis_cond_handler, analysis_cond_handler(ListBox1.SelectedIndex).elts(j),
                '                                          DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value,
                '                                          analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis(k))
                '        End If
                '        indice = indice + 1
                '    Next
                'Next
            End If

            If e.ColumnIndex = 6 Then
                Dim indice As Integer = 0
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            If e.RowIndex = indice Then
                                DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Split(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "\").Last
                                display_grid()
                                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                                Exit Sub
                            End If
                            indice = indice + 1
                        Next
                    Next
                Next
                'If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then Exit Sub
                'Dim indice As Integer = 0
                'For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
                '    For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis)
                '        If e.RowIndex = indice Then
                '            analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis(k) = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                '        End If
                '        indice = indice + 1
                '    Next
                'Next
            End If

            'If e.ColumnIndex = 6 Then
            '    Dim indice As Integer = 0
            '    For j As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts)
            '        For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis)
            '            'DataGridView1.SelectedRows(0).Index
            '            If DataGridView1.SelectedCells(0).RowIndex = indice Then
            '                'analysis_cond_handler(ListBox1.SelectedIndex).elts(j).std_filename(k) = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            '                DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Split(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "\").Last
            '            End If
            '            indice = indice + 1
            '        Next
            '    Next
            'End If

            'Stop
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in DataGridView1_CellValueChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    'Public Sub chg_meas_kr_for_duplicate_elt(ByRef analysis_cond_handler() As analysis_conditions, ByRef analysed_elt As analyzed_element, ByVal meas_kratio As Double,
    '                                         ByVal kv As Double)
    '    If analysis_cond_handler Is Nothing Then Exit Sub

    '    For i As Integer = 0 To UBound(analysis_cond_handler)
    '        If analysis_cond_handler(i).elts Is Nothing Then Continue For
    '        For j As Integer = 0 To UBound(analysis_cond_handler(i).elts)
    '            If analysis_cond_handler(i).elts(j).name = analysed_elt.name And analysis_cond_handler(i).elts(j).line = analysed_elt.line Then
    '                For k As Integer = 0 To UBound(analysis_cond_handler(i).elts(j).energy_analysis)
    '                    If analysis_cond_handler(i).elts(j).energy_analysis(k) = kv Then
    '                        analysis_cond_handler(i).elts(j).kratio_measured(k) = meas_kratio
    '                    End If
    '                Next
    '            End If
    '        Next
    '    Next


    'End Sub


    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Try
            If check_valid_layer_selected() = False Then Exit Sub

            If IsNumeric(TextBox4.Text) = True Then
                layer_handler(ListBox1.SelectedIndex).density = TextBox4.Text
                layer_handler(ListBox1.SelectedIndex).mass_thickness = layer_handler(ListBox1.SelectedIndex).density * layer_handler(ListBox1.SelectedIndex).thickness * 10 ^ -8
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox4_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        Try
            If layer_handler Is Nothing Then Exit Sub
            If check_valid_layer_selected() = False Then Exit Sub
            'If ListBox1.Items.Count - 1 = ListBox1.SelectedIndex Then
            '    analysis_cond_handler(ListBox1.SelectedIndex).isfix = True
            '    CheckBox18.Checked = True
            'End If

            layer_handler(ListBox1.SelectedIndex).isfix = CheckBox18.Checked
            If CheckBox18.Checked = True Then
                TextBox5.BackColor = Color.LightBlue
            Else
                TextBox5.BackColor = Color.White
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox18_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        Try
            If loaded = False Then Exit Sub
            If CheckBox15.Checked = True Then
                CheckBox16.Checked = False
                Label4.Text = "Å"
                If check_valid_layer_selected() = False Then Exit Sub
                If ListBox1.SelectedItem = "Substrate" Then
                    TextBox5.Text = "Inf"
                    layer_handler(ListBox1.SelectedIndex).thickness = 1000000000.0
                Else
                    TextBox5.Text = layer_handler(ListBox1.SelectedIndex).thickness
                End If
                'analysis_cond_handler(ListBox1.SelectedIndex).label = "Angstrom"
                'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
            End If
            If CheckBox15.Checked = False Then
                CheckBox16.Checked = True
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox15_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        Try
            If loaded = False Then Exit Sub
            If CheckBox16.Checked = True Then
                CheckBox15.Checked = False
                Label4.Text = "µg/cm²"
                If check_valid_layer_selected() = False Then Exit Sub
                If ListBox1.SelectedItem = "Substrate" Then
                    TextBox5.Text = "Inf"
                    layer_handler(ListBox1.SelectedIndex).thickness = 1000000000.0
                Else
                    TextBox5.Text = layer_handler(ListBox1.SelectedIndex).density * layer_handler(ListBox1.SelectedIndex).thickness * 10 ^ -8 * 10 ^ 6
                End If
                'analysis_cond_handler(ListBox1.SelectedIndex).label = "ug/cm2"
                'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
            End If
            If CheckBox16.Checked = False Then
                CheckBox15.Checked = True
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox16_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            If check_valid_layer_selected() = False Then Exit Sub
            If IsNumeric(TextBox5.Text) Then
                If CheckBox15.Checked = True Then
                    layer_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
                Else
                    layer_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text / layer_handler(ListBox1.SelectedIndex).density * 10 ^ 2
                End If
                layer_handler(ListBox1.SelectedIndex).mass_thickness = layer_handler(ListBox1.SelectedIndex).density * layer_handler(ListBox1.SelectedIndex).thickness * 10 ^ -8
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox5_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    'Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
    '    If loaded = False Then Exit Sub
    '    If CheckBox12.Checked = True Then
    '        CheckBox17.Checked = False
    '        DataGridView2.Columns(1).HeaderText = "conc (wt)"
    '        If ListBox1.SelectedIndex < 0 Then Exit Sub
    '        layer_handler(ListBox1.SelectedIndex).wt_fraction = True
    '        'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
    '    End If
    '    If CheckBox12.Checked = False Then
    '        CheckBox17.Checked = True
    '    End If
    '    display_grid_layer()
    'End Sub

    'Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
    '    If loaded = False Then Exit Sub
    '    If CheckBox17.Checked = True Then
    '        CheckBox12.Checked = False
    '        DataGridView2.Columns(1).HeaderText = "conc (at.)"
    '        If ListBox1.SelectedIndex < 0 Then Exit Sub
    '        layer_handler(ListBox1.SelectedIndex).wt_fraction = False
    '        'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
    '    End If
    '    If CheckBox17.Checked = False Then
    '        CheckBox12.Checked = True
    '    End If
    '    display_grid_layer()
    'End Sub

    'Private Sub CheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox12.CheckedChanged, CheckBox17.CheckedChanged
    '    If loaded = False Then Exit Sub
    '    'cast sender
    '    Dim senderCheck As CheckBox = DirectCast(sender, CheckBox)

    '    'loop through all checkboxes
    '    For Each checkbox In {CheckBox12, CheckBox17}

    '        'only apply changes to non-sender  boxes
    '        If checkbox IsNot senderCheck Then

    '            'set property to opposite of sender so you can renable when unchecked
    '            checkbox.Checked = Not senderCheck.Checked
    '        Else
    '            If checkbox Is CheckBox12 Then
    '                DataGridView2.Columns(1).HeaderText = "conc (wt)"
    '                If ListBox1.SelectedIndex < 0 Then Exit Sub
    '                layer_handler(ListBox1.SelectedIndex).wt_fraction = True
    '            Else
    '                DataGridView2.Columns(1).HeaderText = "conc (at.)"
    '                If ListBox1.SelectedIndex < 0 Then Exit Sub
    '                layer_handler(ListBox1.SelectedIndex).wt_fraction = False
    '            End If
    '        End If

    '    Next
    '    display_grid_layer()
    'End Sub

    'Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
    '    'If CheckBox10.Checked = False Then
    '    '    CheckBox10.Checked = True
    '    '    Exit Sub
    '    'End If
    '    If CheckBox10.Checked = True Then
    '        CheckBox8.Checked = False
    '        CheckBox9.Checked = False
    '        CheckBox11.Checked = False
    '    End If

    'End Sub

    'Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
    '    'If CheckBox8.Checked = False Then
    '    '    CheckBox8.Checked = True
    '    '    Exit Sub
    '    'End If

    '    If CheckBox8.Checked = True Then
    '        CheckBox10.Checked = False
    '        CheckBox9.Checked = False
    '        CheckBox11.Checked = False
    '    End If
    'End Sub

    'Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
    '    'If CheckBox9.Checked = False Then
    '    '    CheckBox9.Checked = True
    '    '    Exit Sub
    '    'End If

    '    If CheckBox9.Checked = True Then
    '        CheckBox10.Checked = False
    '        CheckBox8.Checked = False
    '        CheckBox11.Checked = False
    '    End If
    'End Sub

    'Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
    '    'If CheckBox11.Checked = False Then
    '    '    CheckBox11.Checked = True
    '    '    Exit Sub
    '    'End If
    '    If CheckBox11.Checked = True Then
    '        CheckBox10.Checked = False
    '        CheckBox9.Checked = False
    '        CheckBox8.Checked = False
    '    End If
    'End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            If DataGridView1.CurrentRow Is Nothing Then Exit Sub

            Dim rowIndex As Integer = DataGridView1.CurrentRow.Index
            Dim columnIndex As Integer = DataGridView1.CurrentCell.ColumnIndex
            Dim indice As Integer = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If DataGridView1.CurrentRow.Index = indice Then
                            ReDim Preserve elt_exp_handler(i).line(UBound(elt_exp_handler(i).line) + 1)
                            ReDim elt_exp_handler(i).line(UBound(elt_exp_handler(i).line)).k_ratio(0)
                            display_grid()
                            DataGridView1.CurrentCell = DataGridView1.Rows(rowIndex).Cells(columnIndex)
                            Exit Sub
                        End If
                        indice = indice + 1
                    Next
                Next
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button12_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If DataGridView1.CurrentRow Is Nothing Then Exit Sub

            Dim rowIndex As Integer = DataGridView1.CurrentRow.Index
            Dim columnIndex As Integer = DataGridView1.CurrentCell.ColumnIndex
            Dim indice As Integer = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If DataGridView1.CurrentRow.Index = indice Then
                            ReDim Preserve elt_exp_handler(i).line(j).k_ratio(UBound(elt_exp_handler(i).line(j).k_ratio) + 1)
                            display_grid()
                            DataGridView1.CurrentCell = DataGridView1.Rows(rowIndex).Cells(columnIndex)
                            Exit Sub
                        End If
                        indice = indice + 1
                    Next
                Next
            Next

            'Dim indice As Integer = 0
            'For j As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
            '    For k As Integer = 0 To UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis)
            '        If DataGridView1.CurrentRow.Index = indice Then
            '            ReDim Preserve analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis(UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).energy_analysis) + 1)
            '            ReDim Preserve analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio(UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio) + 1)
            '            ReDim Preserve analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured(UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).kratio_measured) + 1)
            '            ReDim Preserve analysis_cond_handler(ListBox1.SelectedIndex).elts(j).std_filename(UBound(analysis_cond_handler(ListBox1.SelectedIndex).elts(j).std_filename) + 1)
            '            display_grid()
            '            Exit Sub
            '        End If
            '        indice = indice + 1
            '    Next
            'Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button2_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Chart1_DoubleClick(sender As Object, e As EventArgs) Handles Chart1.DoubleClick
        Try
            Dim OFFSET As Integer = 40
            If Chart1.Height = Me.Height - OFFSET Then
                Chart1.Height = chart1_height
                Chart1.Width = chart1_width
                Chart1.Location = chart1_location
                'Chart1.Height = 252 '206
                'Chart1.Width = 503 '448
                'Chart1.Location = New Point(775, 364) '830, 410
            Else
                Chart1.Height = Me.Height - OFFSET
                Chart1.Width = Me.Width - 16
                Chart1.Location = New Point(0, 0)
                Chart1.BringToFront()
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Chart1_DoubleClick " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub DataGridView2_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseClick
        Try
            If e.Button = MouseButtons.Right Then
                Dim ConcFixedValue As Boolean
                If DataGridView2.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.LightBlue Then
                    DataGridView2.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.White
                    ConcFixedValue = False
                Else
                    DataGridView2.Item(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.LightBlue
                    ConcFixedValue = True
                End If

                If check_valid_layer_selected() = False Then Exit Sub
                Dim indice As Integer = 0
                For j As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                    If e.RowIndex = indice Then
                        layer_handler(ListBox1.SelectedIndex).element(j).isConcFixed = ConcFixedValue
                        Exit For
                    End If
                    indice = indice + 1
                Next
                display_grid_layer()
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in DataGridView2_CellMouseClick " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            '***********************************************
            'Remove the selected line in the data grid
            '***********************************************
            remove_selected_line()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button3_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub remove_selected_line()
        Try
            Dim indice As Integer = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If DataGridView1.CurrentRow.Index = indice Then
                            If elt_exp_handler(i).line(j).k_ratio.Count > 1 Then
                                For l As Integer = k To UBound(elt_exp_handler(i).line(j).k_ratio) - 1
                                    elt_exp_handler(i).line(j).k_ratio(l) = elt_exp_handler(i).line(j).k_ratio(l + 1)
                                Next
                                ReDim Preserve elt_exp_handler(i).line(j).k_ratio(UBound(elt_exp_handler(i).line(j).k_ratio) - 1)
                                display_grid()
                                Exit Sub
                            Else
                                If elt_exp_handler(i).line.Count > 1 Then
                                    For l As Integer = j To UBound(elt_exp_handler(i).line) - 1
                                        elt_exp_handler(i).line(l) = elt_exp_handler(i).line(l + 1)
                                    Next
                                    ReDim Preserve elt_exp_handler(i).line(UBound(elt_exp_handler(i).line) - 1)
                                    display_grid()
                                    Exit Sub
                                Else
                                    For l As Integer = i To UBound(elt_exp_handler) - 1
                                        elt_exp_handler(l) = elt_exp_handler(l + 1)
                                    Next
                                    ReDim Preserve elt_exp_handler(UBound(elt_exp_handler) - 1)
                                    display_grid()
                                    Exit Sub
                                End If

                            End If
                        End If
                        indice = indice + 1
                    Next
                Next
            Next

            'Dim selected_index As Integer = ListBox1.SelectedIndex
            'Dim tmp As Integer = 0
            'For i As Integer = 0 To UBound(analysis_cond_handler(selected_index).elts)
            '    For j As Integer = 0 To UBound(analysis_cond_handler(selected_index).elts(i).energy_analysis)
            '        If DataGridView1.SelectedCells.Item(0).RowIndex = tmp Then
            '            If analysis_cond_handler(selected_index).elts(i).energy_analysis.Count = 1 Then
            '                remove_elt(analysis_cond_handler(selected_index).elts, i)
            '                display_grid()
            '                Exit Sub
            '            Else
            '                For k As Integer = j To UBound(analysis_cond_handler(selected_index).elts(i).energy_analysis) - 1
            '                    analysis_cond_handler(selected_index).elts(i).energy_analysis(k) = analysis_cond_handler(selected_index).elts(i).energy_analysis(k + 1)
            '                    analysis_cond_handler(selected_index).elts(i).kratio(k) = analysis_cond_handler(selected_index).elts(i).kratio(k + 1)
            '                    analysis_cond_handler(selected_index).elts(i).kratio_measured(k) = analysis_cond_handler(selected_index).elts(i).kratio_measured(k + 1)
            '                    analysis_cond_handler(selected_index).elts(i).std_filename(k) = analysis_cond_handler(selected_index).elts(i).std_filename(k + 1)

            '                Next
            '                ReDim Preserve analysis_cond_handler(selected_index).elts(i).energy_analysis(UBound(analysis_cond_handler(selected_index).elts(i).energy_analysis) - 1)
            '                ReDim Preserve analysis_cond_handler(selected_index).elts(i).kratio(UBound(analysis_cond_handler(selected_index).elts(i).kratio) - 1)
            '                ReDim Preserve analysis_cond_handler(selected_index).elts(i).kratio_measured(UBound(analysis_cond_handler(selected_index).elts(i).kratio_measured) - 1)
            '                ReDim Preserve analysis_cond_handler(selected_index).elts(i).std_filename(UBound(analysis_cond_handler(selected_index).elts(i).std_filename) - 1)
            '                display_grid()
            '                Exit Sub
            '            End If

            '        End If
            '        tmp = tmp + 1

            '    Next
            'Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in remove_selected_line " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub check_graph_limits(ByRef graph_limits() As Double)
        Try
            If graph_limits Is Nothing Or graph_limits.Count <> 4 Then
                MsgBox("Error in definition of graph limits")
                ReDim graph_limits(3)
                graph_limits(0) = 0
                graph_limits(1) = 30
                graph_limits(2) = 0
                graph_limits(3) = 1
                Exit Sub
            End If

            Dim xmin As Double = graph_limits(0)
            Dim xmax As Double = graph_limits(1)
            Dim ymin As Double = graph_limits(2)
            Dim ymax As Double = graph_limits(3)

            If IsNumeric(xmin) = False Then
                xmin = 0
                MsgBox("Wrong X min value")
            End If
            If IsNumeric(xmax) = False Then
                xmax = 30
                MsgBox("Wrong X max value")
            End If
            If IsNumeric(ymin) = False Then
                ymin = 0
                MsgBox("Wrong Y min value")
            End If
            If IsNumeric(ymax) = False Then
                ymax = 1
                MsgBox("Wrong Y max value")
            End If

            If xmin > xmax Then
                Dim tmp As Double = xmin
                xmin = xmax
                xmax = xmin
            End If
            If ymin > ymax Then
                Dim tmp As Double = ymin
                ymin = ymax
                ymax = ymin
            End If
            If xmin = xmax Then
                xmax = xmin * 1.1
            End If
            If ymin = ymax Then
                ymax = ymin * 1.1
            End If

            graph_limits(0) = xmin
            graph_limits(1) = xmax
            graph_limits(2) = ymin
            graph_limits(3) = ymax
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in check_graph_limits " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            graph_limits(0) = TextBox7.Text
            graph_limits(1) = TextBox8.Text
            graph_limits(2) = TextBox9.Text
            graph_limits(3) = TextBox10.Text

            check_graph_limits(graph_limits)

            Chart1.ChartAreas(0).AxisX.Minimum = graph_limits(0)
            Chart1.ChartAreas(0).AxisX.Maximum = graph_limits(1)
            Chart1.ChartAreas(0).AxisY.Minimum = graph_limits(2)
            Chart1.ChartAreas(0).AxisY.Maximum = graph_limits(3)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button4_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Function total_number_of_kratios(ByVal elt_exp_handler() As Elt_exp) As Integer
        Try
            Dim total_kratios As Integer = 0

            For i As Integer = 0 To UBound(elt_exp_handler)
                'total_elts = total_elts + 1
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        'If elt_exp_handler(i).line(j).k_ratio(k).experimental_value = 0 Then
                        '    Continue For
                        'End If
                        total_kratios = total_kratios + 1
                    Next
                Next
            Next

            Return total_kratios
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in total_number_of_kratios " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Public Function kVs_extractor(elt_exp_handler() As Elt_exp) As Double()
        Try
            Dim kVs() As Double = Nothing
            Dim flag_isNew As Boolean = True
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If kVs Is Nothing Then
                            ReDim kVs(0)
                            kVs(0) = elt_exp_handler(i).line(j).k_ratio(k).kv
                        Else
                            For l As Integer = 0 To UBound(kVs)
                                If kVs(l) = elt_exp_handler(i).line(j).k_ratio(k).kv Then
                                    flag_isNew = False
                                    Exit For
                                End If
                            Next
                            If flag_isNew = True Then
                                ReDim Preserve kVs(UBound(kVs) + 1)
                                kVs(UBound(kVs)) = elt_exp_handler(i).line(j).k_ratio(k).kv
                            Else
                                flag_isNew = True
                            End If
                        End If
                    Next
                Next
            Next
            Return kVs
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in kVs_extractor " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function


    Public Sub init_standard_intensities(ByRef elt_exp_handler() As Elt_exp)
        Try
            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    If elt_exp_handler(i).line(j).k_ratio IsNot Nothing Then
                        'Reads the name of the file of the standard
                        Dim file_name As String = Split(elt_exp_handler(i).line(j).std_filename, "\").Last
                        Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(i).line(j).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                        If flag_file_exists = False Then
                            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                                elt_exp_handler(i).line(j).std_filename = Application.StartupPath & "\Examples\" & file_name
                                flag_file_exists = True
                            ElseIf My.Computer.FileSystem.FileExists(Application.StartupPath & "\" & file_name) = True Then
                                elt_exp_handler(i).line(j).std_filename = Application.StartupPath & "\" & file_name
                                flag_file_exists = True
                            End If
                        End If
                        If IsNothing(elt_exp_handler(i).line(j).std_filename) = True Or flag_file_exists = False Then
                            If elt_exp_handler(i).line(j).std_filename <> "" And flag_file_exists = False Then
                                MsgBox("Standard file not found for: " & elt_exp_handler(i).elt_name & vbCrLf & "Please reselect or recreate file: " & file_name & vbCrLf & "Calculating intensity for pure standard.")
                            End If
                            For kk As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                                elt_exp_handler(i).line(j).k_ratio(kk).std_intensity = init_pure_std(elt_exp_handler(i), j, elt_exp_handler(i).line(j).k_ratio(kk).kv, toa, Ec_data)
                                Debug.Print(options.phi_rz_mode & vbTab & "Std: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab & elt_exp_handler(i).line(j).k_ratio(kk).std_intensity)
                            Next
                        Else
                            Dim layer_handler_std() As layer = Nothing
                            Dim elt_exp_handler_std() As Elt_exp = Nothing
                            Dim toa_std As Double
                            load_data(elt_exp_handler(i).line(j).std_filename, layer_handler_std, elt_exp_handler_std, toa_std)
                            Dim elt_exp_all_std() As Elt_exp = Nothing
                            init_elt_exp_all(elt_exp_all_std, layer_handler_std, Ec_data, pen_path)
                            '*****************
                            'We must normalize the standard concentrations to 1
                            For ll As Integer = 0 To UBound(layer_handler_std)
                                Dim sum_conc_wt As Double = 0
                                For jj As Integer = 0 To UBound(layer_handler_std(ll).element)
                                    sum_conc_wt = sum_conc_wt + layer_handler_std(ll).element(jj).conc_wt
                                Next
                                For jj As Integer = 0 To UBound(layer_handler_std(ll).element)
                                    layer_handler_std(ll).element(jj).conc_wt = layer_handler_std(ll).element(jj).conc_wt / sum_conc_wt
                                Next
                            Next
                            '*****************
                            For kk As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                                elt_exp_handler(i).line(j).k_ratio(kk).std_intensity = pre_auto(layer_handler_std, elt_exp_handler(i), j, elt_exp_all_std, elt_exp_handler(i).line(j).k_ratio(kk).kv,
                                                                                                toa, Ec_data, options, False, save_results, Nothing)
                                Debug.Print(options.phi_rz_mode & vbTab & "Std: " & vbTab & elt_exp_handler(i).elt_name & vbTab & elt_exp_handler(i).line(j).xray_name & vbTab &
                                            elt_exp_handler(i).line(j).k_ratio(kk).std_intensity)

                            Next
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in init_standard_intensities " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    'sender As Object, e As EventArgs
    Public Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim calculate_MAC_flag As Boolean = CheckBox13.Checked
            calculate(calculate_MAC_flag)

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button6_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Public Sub calculate(ByVal calculate_MAC_flag As Boolean)
        Try
            If IsNothing(options) Then
                Dim err As String = Date.Now.ToString & vbTab & "Error in calculate: option is nothing."
                MsgBox(err)
            End If

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Defining and loading MACs.")
            End Using

            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            ElseIf options.MAC_mode = "PENELOPE2018" Then
                MAC_data = MAC_data_PEN18
            ElseIf options.MAC_mode = "FFAST" Then
                MAC_data = MAC_data_FFAST
            ElseIf options.MAC_mode = "EPDL23" Then
                MAC_data = MAC_data_EPDL23
            Else
                MAC_data = MAC_data_PEN14
            End If

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("MACs loaded: " & options.MAC_mode)
            End Using

            'Check if a calculation is already running.
            'If yes, abort it.
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Checkinng if a calculation is already running.")
            End Using

            If BackgroundWorker1.IsBusy = True Then
                BackgroundWorker1.CancelAsync()
                Button6.Text = "Calculate"
                Using err As StreamWriter = New StreamWriter("log.txt", True)
                    err.WriteLine("Calculation already running.")
                End Using
                Exit Sub
            End If

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("No other running calculations.")
            End Using

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Loading experimental MACs.")
            End Using

            'Load the current path. A folder with the atomic data extracted from PENELOPE should be present.
            'Load the experimental MACs if the option is enabled
            If options.experimental_MAC.experimental_MAC_enabled = True Then
                load_experimental_MAC(experimental_MAC_path, options.experimental_MAC)
            End If

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Experimental MACs loaded.")
            End Using

            '*******************************************
            'Check the definition of the system
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Check the definition of the system.")
            End Using

            If layer_handler Is Nothing Then Exit Sub
            For i As Integer = 0 To UBound(layer_handler)
                If layer_handler(i).element Is Nothing Then Exit Sub
                If layer_handler(i).element.Length = 0 Then Exit Sub
            Next
            'If layer_handler(0).element Is Nothing Then Exit Sub

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Layer_handler correctly definied.")
            End Using

            If elt_exp_handler Is Nothing Then Exit Sub

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("elt_exp_handler correctly definied.")
            End Using
            '*******************************************

            Label1.Text = "Status: Initialization"
            'Delete previous results
            TextBox12.Text = ""

            '*******************************************
            'Total number of k-ratios and total number of elts
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Loading total number of k-ratios and total number of elts.")
            End Using

            Dim total_kratios As Integer = total_number_of_kratios(elt_exp_handler)

            '*******************************************

            '*******************************************
            'Calculate the number of elements definied by stoichiometry
            'Dim num_of_elt_def_by_stoichio As Integer = 0
            'For i As Integer = 0 To UBound(layer_handler)
            '    If layer_handler(i).O_by_stochio = True Then
            '        num_of_elt_def_by_stoichio = num_of_elt_def_by_stoichio + 1
            '    End If
            'Next
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating x().")
            End Using
            '*******************************************
            'Arbitrary x values: 1 2 3 4 5 6 7... used for the fitting
            Dim x() As Double
            If CheckBox14.Checked = True Then
                options.sum_conc_equals_one = False
                ReDim x(total_kratios - 1)   'x
            Else
                options.sum_conc_equals_one = True
                ReDim x(total_kratios + layer_handler.Count - 1)  'x
            End If

            For i As Integer = 0 To UBound(x)
                x(i) = i + 1
            Next
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating k_ratio_measured().")
            End Using

            '*******************************************
            'Array k_ratio_measured
            'Used as y for fitting
            Dim k_ratio_measured(UBound(x)) As Double
            Dim tmp As Integer = 0
            For i As Integer = 0 To UBound(elt_exp_handler)
                'total_elts = total_elts + 1
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If elt_exp_handler(i).line(j).k_ratio(k).experimental_value = 0 Then
                            'Stop
                            Continue For
                        End If
                        k_ratio_measured(tmp) = elt_exp_handler(i).line(j).k_ratio(k).experimental_value
                        tmp = tmp + 1
                    Next
                Next
            Next

            'k_ratio_measured also contains sum(ci)=1 for each layer
            'for each layer the total concentration should be as close as possible to 1
            'If CheckBox14.Checked = False Then
            'ReDim Preserve k_ratio_measured(UBound(k_ratio_measured) + layer_handler.Count)
            If CheckBox14.Checked = False Then
                For i As Integer = 0 To layer_handler.Count - 1
                    k_ratio_measured(tmp) = 1
                    tmp = tmp + 1
                Next
            End If

            'If num_of_elt_def_by_stoichio <> 0 Then
            '    For i As Integer = 0 To num_of_elt_def_by_stoichio - 1
            '        k_ratio_measured(tmp) = 0
            '        tmp = tmp + 1
            '    Next
            'End If

            'End If
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating ey().")
            End Using

            '*******************************************
            'Uncertainty on the measured k-ratios
            tmp = 0
            Dim ey(UBound(k_ratio_measured)) As Double '5% of uncertainty on the y values by default
            For i As Integer = 0 To UBound(elt_exp_handler)
                'ey(i) = 0.05 * k_ratio_measured(i)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        If elt_exp_handler(i).line(j).k_ratio(k).experimental_value = 0 Then
                            Continue For
                        End If
                        If elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value = 0 Then
                            ey(tmp) = 0.05 * k_ratio_measured(tmp)
                        Else
                            ey(tmp) = elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value
                            'ey(tmp) = 0.05 * k_ratio_measured(tmp)
                        End If
                        tmp = tmp + 1
                    Next
                Next
            Next

            'Uncertainty on the total concentration of each layer
            If CheckBox14.Checked = False Then
                For i As Integer = 0 To layer_handler.Count - 1
                    ey(tmp) = 0.05 * k_ratio_measured(tmp)
                    tmp = tmp + 1
                Next
            End If

            'If num_of_elt_def_by_stoichio <> 0 Then
            '    For i As Integer = 0 To num_of_elt_def_by_stoichio - 1
            '        ey(tmp) = 0.005
            '        tmp = tmp + 1
            '    Next
            'End If

            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating thicknesses().")
            End Using

            '*******************************************
            'Array for the layer thicknesses
            'Used as parameter p for the fit
            Dim thicknesses(UBound(layer_handler)) As Double
            For i As Integer = 0 To UBound(layer_handler)
                thicknesses(i) = layer_handler(i).thickness
            Next
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating concentration().")
            End Using

            '*******************************************
            'Array concentration
            'Used as the second part of the parameter array p for the fit

            'Convertion factor: wt% to wt fraction
            Dim norm_conc_factor As Double = 1
            Dim total_conc As Double = 0
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    total_conc = total_conc + layer_handler(i).element(j).conc_wt
                Next
                If total_conc > 3.0 Then
                    norm_conc_factor = 100
                End If
                total_conc = 0
            Next


            Dim num_concentration As Integer = 0
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    num_concentration = num_concentration + 1
                Next
            Next

            Dim concentration(num_concentration - 1) As Double
            Dim isConFixedArray(num_concentration - 1) As Boolean
            tmp = 0
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    concentration(tmp) = layer_handler(i).element(j).conc_wt / norm_conc_factor
                    isConFixedArray(tmp) = layer_handler(i).element(j).isConcFixed
                    tmp = tmp + 1
                Next
            Next
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating kVs().")
            End Using

            '*******************************************
            'Array with the kVs used
            Dim kVs() As Double = kVs_extractor(elt_exp_handler)
            '*******************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating p().")
            End Using

            '*******************************************
            'Array of parameters
            ' p() is an array of doubles containing the fitting parameters
            '(i.e., the elemental concentrations and the film thicknesses).
            Dim p(UBound(thicknesses) + 1 + UBound(concentration)) As Double
            For i As Integer = 0 To UBound(thicknesses)
                p(i) = thicknesses(i)
            Next
            For i As Integer = 0 To UBound(concentration)
                p(UBound(thicknesses) + 1 + i) = concentration(i)
            Next
            '*******************************************

            If calculate_MAC_flag Then
                ReDim Preserve p(UBound(p) + 2)
                p(UBound(p) - 1) = 1 'scaling factor
                p(UBound(p)) = TextBox15.Text 'Initial MAC value to be fitted
            End If

            '*******************************************
            'Bounding parameters
            'pars is an array of MPFitLib.mp_par containing the constraints on the fitting parameters.
            'pars(i).limited = {X,X} defines if their is a constrain on the min and max values that the associated fitting parapmeters p(i) can have.
            'X=0: no constrain
            'X=1: constrain
            'pars(i).limits(0) = X

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating pars().")
            End Using

            Dim pars(UBound(thicknesses) + 1 + UBound(concentration)) As MPFitLib.mp_par
            For i As Integer = 0 To UBound(thicknesses)
                pars(i) = New MPFitLib.mp_par
                pars(i).limited = {1, 0}
                pars(i).limits(0) = 0
                If layer_handler(i).isfix = True Then
                    pars(i).isFixed = 1
                End If
            Next
            pars(UBound(thicknesses)).isFixed = 1 'this corresponds to the substrate. Its thickness is fixed.

            For i As Integer = 0 To UBound(concentration)
                pars(UBound(thicknesses) + 1 + i) = New MPFitLib.mp_par
                pars(UBound(thicknesses) + 1 + i).limited = {1, 1}
                pars(UBound(thicknesses) + 1 + i).limits(0) = 0
                pars(UBound(thicknesses) + 1 + i).limits(1) = 1
                If isConFixedArray(i) = True Then
                    pars(UBound(thicknesses) + 1 + i).isFixed = 1
                End If
            Next
            '*******************************************

            If calculate_MAC_flag Then
                ReDim Preserve pars(UBound(pars) + 2)
                pars(UBound(pars) - 1) = New MPFitLib.mp_par
                pars(UBound(pars) - 1).limited = {1, 0}
                pars(UBound(pars) - 1).limits(0) = 0
                pars(UBound(pars) - 1).isFixed = 0

                pars(UBound(pars)) = New MPFitLib.mp_par
                pars(UBound(pars)).limited = {1, 0}
                pars(UBound(pars)).limits(0) = 30
                pars(UBound(pars)).isFixed = 0
                pars(UBound(pars)).relstep = 0.01
                For i As Integer = 0 To UBound(pars) - 2
                    pars(i).isFixed = 1 'Fix everything except pars(ubound(pars)) that corresponds to the MAC
                Next
            End If

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Creating fit_MAC().")
            End Using

            Dim fit_MAC As New fit_MAC
            If calculate_MAC_flag Then

                fit_MAC.absorbed_elt = TextBox17.Text
                fit_MAC.X_ray = TextBox14.Text
                fit_MAC.MAC = TextBox15.Text

                fit_MAC.absorber_elt = correct_symbol(TextBox16.Text)
                If fit_MAC.absorber_elt = "" Then
                    fit_MAC.compound_MAC = True
                Else
                    fit_MAC.compound_MAC = False
                End If

                fit_MAC.norm_kV = 15
                Dim shell1, shell2 As Integer
                Siegbahn_to_transition_num(fit_MAC.X_ray, shell1, shell2, fit_MAC.absorbed_elt)
                Dim z As Integer = symbol_to_Z(fit_MAC.absorbed_elt)
                Dim Ec_shell2 As Double = find_Ec(z, shell2, Ec_data)
                fit_MAC.X_ray_energy = find_Ec(z, shell1, Ec_data) - Ec_shell2 'in keV 
                fit_MAC.activated = True

                'Dim max As Double = k_ratio_measured.Max
                'For i As Integer = 0 To UBound(k_ratio_measured)
                '    k_ratio_measured(i) = k_ratio_measured(i) / max
                'Next
            End If
            '*******************************************

            '*******************************************
            ' Save the atomic fraction if it is fixed and the layer is defined by atomic fraction.
            '*******************************************
            Dim only_one_fixed_at_frac As Boolean = False
            Dim sum As Double
            For i As Integer = 0 To UBound(layer_handler)
                sum = 0
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    sum = sum + layer_handler(i).element(j).conc_at
                Next
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    layer_handler(i).element(j).conc_at = layer_handler(i).element(j).conc_at / sum
                Next
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    If layer_handler(i).element(j).isConcFixed = True And layer_handler(i).wt_fraction = False Then
                        If only_one_fixed_at_frac = True Then
                            MsgBox("Warning: when defining the concentration by atomic fraction, only one element can have its value fixed." & vbCrLf & "If there are more than one element in this situation, please try to click multiple times on the Calculate button until the atomic fractions converge to steady values.")
                        End If
                        layer_handler(i).element(j).conc_at_ori = layer_handler(i).element(j).conc_at
                        only_one_fixed_at_frac = True
                    End If
                Next
                only_one_fixed_at_frac = False
            Next





            Dim buffer_text As String = Nothing

            'Dim Ec_data() As String = Nothing
            'init_Ec(Ec_data, pen_path)
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Initialization of the elements: init_element.")
            End Using

            For i As Integer = 0 To UBound(elt_exp_handler)
                init_element(elt_exp_handler(i).elt_name, "", vbNull, Ec_data, elt_exp_handler(i), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            Next

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Initialization of the elements: init_elt_exp_all.")
            End Using

            elt_exp_all = Nothing
            init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)

            '*******************************************************
            ' Load the standard intensities
            '*******************************************************
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Initialization of the standards: init_standard_intensities.")
            End Using
            init_standard_intensities(elt_exp_handler)
            '*******************************************************

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Start the BackgroundWorker.")
            End Using

            'Dim worker = New Worker_class
            'worker.backgroundWorker1.RunWorkerAsync()
            'Me.backgroundWorker1 = New System.ComponentModel.BackgroundWorker
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.WorkerReportsProgress = True
            Button6.Text = "Abort"
            BackgroundWorker1.RunWorkerAsync({x, k_ratio_measured, ey, p, pars, buffer_text, Ec_data, fit_MAC})

            '*******************************************

            'If save_results = "" Then save_results = " "
            'My.Computer.Clipboard.SetText(save_results)
            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Calculation Done & Ready.")
            End Using

            Label1.Text = "Status: Done & Ready"

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in calculate " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub


    Public Sub init_elt_exp_all(ByRef elt_exp_all() As Elt_exp, ByVal layer_handler() As layer, ByVal Ec_data() As String, ByVal pen_path As String)
        Try
            'Dim indice As Integer = 0
            For i As Integer = 0 To UBound(layer_handler)
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    If elt_exp_all Is Nothing Then
                        ReDim elt_exp_all(0)
                        init_element(layer_handler(i).element(j).elt_name, "", vbNull, Ec_data, elt_exp_all(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    Else
                        Dim isunic As Boolean = True
                        For k As Integer = 0 To UBound(elt_exp_all)
                            If layer_handler(i).element(j).elt_name = elt_exp_all(k).elt_name Then
                                isunic = False
                                Exit For
                            End If
                        Next
                        If isunic = True Then
                            ReDim Preserve elt_exp_all(UBound(elt_exp_all) + 1)
                            init_element(layer_handler(i).element(j).elt_name, "", vbNull, Ec_data, elt_exp_all(UBound(elt_exp_all)), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in init_elt_exp_all " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub



    Public Function init_pure_std(ByVal elt_exp As Elt_exp, ByVal line_indice As Integer, ByVal kv As Double, ByVal toa As Double, ByVal Ec_data() As String) As Double
        Try
            '*****************************************************
            'Calculate X-ray intensity for a pure element
            '*****************************************************
            Dim layer_handler(0) As layer
            'Dim elt_exp_handler(0) As Elt_exp

            Dim elt_layer As Elt_layer = Nothing
            init_element_layer(elt_exp.elt_name, vbNull, elt_layer)

            ReDim layer_handler(0).element(0)
            layer_handler(0).element(0) = elt_layer
            layer_handler(0).element(0).z = elt_exp.z
            layer_handler(0).element(0).mother_layer_id = 0
            layer_handler(0).element(0).conc_wt = 1
            layer_handler(0).density = zaro(elt_layer.z)(1)
            layer_handler(0).id = 0
            layer_handler(0).thickness = 100000000.0 'in Angstrom
            layer_handler(0).mass_thickness = layer_handler(0).density * layer_handler(0).thickness * 10 ^ -8

            Return pre_auto(layer_handler, elt_exp, line_indice, elt_exp_all, kv, toa, Ec_data, options, False, save_results, Nothing)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in init_pure_std " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            '*******************************************************
            ' Save the data.
            '*******************************************************
            SaveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            SaveFileDialog1.FilterIndex = 1
            SaveFileDialog1.RestoreDirectory = True
            SaveFileDialog1.Title = "Save"
            SaveFileDialog1.AddExtension = True
            SaveFileDialog1.DefaultExt = ".txt"

            'Display a dialogbox for the user to chose the location and name of the saved file.
            If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                'Call the export function that will save the data.
                Call export(SaveFileDialog1.FileName, layer_handler, elt_exp_handler, toa, VERSION)

                'Change the text at the top of the windiws with the new name
                Me.Text = "BadgerFilm " & VERSION & "  " & SaveFileDialog1.FileName
                Label1.Text = "Status: Saved"
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button7_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            '*******************************************************
            ' Load the data
            '*******************************************************
            Dim OpenFileDialog1 As New OpenFileDialog()
            Dim data_file As String = Nothing

            OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            OpenFileDialog1.FilterIndex = 1
            OpenFileDialog1.RestoreDirectory = True
            OpenFileDialog1.Title = "Load data"
            OpenFileDialog1.AddExtension = True
            OpenFileDialog1.DefaultExt = ".txt"

            'Open a dialog box for the user.
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                'If the user has selected a file, copy its name in data_file
                data_file = OpenFileDialog1.FileName
            Else
                'Else abord loding data
                Label1.Text = "Status: Loading canceled"
                Exit Sub
            End If

            'Change the text at the top of the window.
            Me.Text = "BadgerFilm " & VERSION & "  " & data_file

            'Clear the graphs
            Chart1.Series.Clear()

            'Call the function to load the data
            load_data(data_file, layer_handler, elt_exp_handler, toa)

            'Update the fields
            update_form_fields()
            TextBox12.Text = ""

            Label1.Text = "Status: Loaded"
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button8_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub


    Public Sub update_form_fields()
        Try
            If layer_handler Is Nothing Then Exit Sub
            If check_valid_layer_selected() = False Then Exit Sub

            Dim retrieve_selected_layer As Integer = ListBox1.SelectedIndex
            'TextBox3.Text = ""
            TextBox3.Text = layer_handler.Count
            If retrieve_selected_layer + 1 > layer_handler.Count Then retrieve_selected_layer = 0
            ListBox1.SelectedIndex = retrieve_selected_layer
            display_grid()
            display_grid_layer()
            TextBox1.Text = toa
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in update_form_fields " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Try
            '******************
            ' Enable or disable secondary fluorescence by characteristic X-rays.
            '******************
            If CheckBox1.Checked = True Then
                options.char_fluo_flag = True
            Else
                options.char_fluo_flag = False
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox1_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Try
            '******************
            ' Enable or disable secondary fluorescence by characteristic bremsstrahlung.
            '******************
            If CheckBox2.Checked = True Then
                options.brem_fluo_flag = True
            Else
                options.brem_fluo_flag = False
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox2_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub













    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim penepma_data As String = TextBox11.Text
            Dim tmp() As String = Split(penepma_data, vbCrLf)
            Dim kV(UBound(tmp)) As Double
            Dim elt(UBound(tmp)) As String
            Dim line(UBound(tmp)) As String
            Dim brem_val(UBound(tmp)) As Double
            Dim brem_err(UBound(tmp)) As Double

            TextBox6.Text = "1"
            TextBox2.Text = "1"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            For i As Integer = 0 To UBound(tmp)
                Dim tmp2() As String = Split(tmp(i), vbTab)
                kV(i) = tmp2(0)
                elt(i) = tmp2(1)
                line(i) = tmp2(2)
                brem_val(i) = tmp2(7)
                brem_err(i) = tmp2(8)
            Next

            '*******************************************
            'Array of parameters
            Dim p(1) As Double
            p(0) = TextBox2.Text
            p(1) = TextBox6.Text
            '*******************************************
            '*******************************************
            'Bounding parameters
            Dim pars(1) As MPFitLib.mp_par

            pars(0) = New MPFitLib.mp_par
            pars(0).limited = {1, 0}
            pars(0).limits(0) = 0
            'pars(0).limits(1) = 25

            pars(1) = New MPFitLib.mp_par
            pars(1).limited = {1, 1}
            pars(1).limits(0) = 0
            pars(1).limits(1) = 100
            'If symbol_to_Z(elt(UBound(tmp))) > 14 Then
            '    pars(1).isFixed = True
            'End If

            '*******************************************

            Dim buffer_text As String = Nothing
            'Dim Ec_data() As String = Nothing
            'init_Ec(Ec_data, pen_path)

            For i As Integer = 0 To UBound(elt_exp_handler)
                init_element(elt_exp_handler(i).elt_name, "", vbNull, Ec_data, elt_exp_handler(i), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            Next

            elt_exp_all = Nothing
            init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)

            Dim results As String = ""
            Dim fitting_methode As New fitting_module

            fitting_methode.fit_brem_fluo(kV, brem_val, brem_err, p, pars, buffer_text, layer_handler, elt_exp_handler, elt_exp_all, toa, elt, line, pen_path, Ec_data, options, results) ' elt, line)

            TextBox13.Text = TextBox13.Text & symbol_to_Z(elt(0)) & vbTab & elt(0) & vbTab & line(0) & vbTab & p(0) & vbTab & p(1) & vbCrLf

            Dim sw As StreamWriter = New StreamWriter(elt(0) & "_" & line(0) & "lol.txt")
            sw.WriteLine(results)
            sw.Close()

            Dim sw_2 As StreamWriter = File.AppendText("Coeffslol.txt")
            sw_2.WriteLine(elt(0) & vbTab & line(0) & vbTab & p(0) & vbTab & p(1))
            sw_2.Close()

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button1_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim NUM As Integer = 13

            Dim Z() As Double = {12, 13, 14, 26, 26, 14, 26, 26, 8, 8, 8, 8, 8, 8}
            Dim wt() As Double = {60.3, 52.93, 46.74, 77.73, 69.94, 33.46, 66.54, 72.36, 39.7, 47.07, 53.26, 22.27, 22.27, 25.97}
            Dim A() As Double = {24.305, 26.9815, 28.0855, 55.845, 55.845, 28.0855, 55.845, 55.845, 15.9994, 15.9994, 15.9994, 15.9994, 15.9994, 15.9994}
            Dim Ex() As Double = {1.254, 1.487, 1.74, 6.4, 6.4, 1.74, 6.4, 6.4, 0.523, 0.523, 0.523, 0.523, 0.523, 0.523}
            Dim x() As Double = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}
            Dim y() As Double = {47.96875221, 54.2724196, 61.50978353, 15.77062702, 15.83049081, 19.29115266, 15.75235591, 16.03285884, 18.46107413, 16.85118903, 17.01812111,
            22.50557219, 22.38993055, 25.54272408}

            Dim y_err(UBound(y)) As Double
            For i As Integer = 0 To UBound(y)
                y_err(i) = y(i) * 0.05
            Next

            '*******************************************
            'Array of parameters
            Dim p(4) As Double
            For i As Integer = 0 To UBound(p)
                p(i) = 1
            Next
            '*******************************************

            '*******************************************
            'Bounding parameters
            Dim pars(4) As MPFitLib.mp_par
            For i As Integer = 0 To UBound(pars)
                pars(i) = New MPFitLib.mp_par
                pars(i).limited = {1, 1}
                pars(i).limits(0) = -100
                pars(i).limits(1) = 100
            Next

            '*******************************************

            Dim buffer_text As String = Nothing

            Dim fitting_methode As New fitting_module
            fitting_methode.fit_coeff1(x, y, y_err, p, pars, buffer_text, Z, wt, A, Ex)
            Console.WriteLine(buffer_text)

            Dim tmp As String = ""
            For i As Integer = 0 To UBound(p)
                tmp = tmp & p(i) & vbTab
            Next
            TextBox12.Text = tmp
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button5_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            If IsNumeric(TextBox1.Text) Then
                toa = TextBox1.Text
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox1_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub


    'Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
    '    Dim path As String = Application.StartupPath & "\PenelopeData"
    '    Dim dir As DirectoryInfo = New DirectoryInfo(path)
    '    For Each file As FileInfo In dir.GetFiles
    '        crypt(path, file.Name)
    '    Next

    'End Sub

    'Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
    '    Dim path As String = Application.StartupPath & "\PenelopeData"
    '    Dim filename As String = "test.txt"
    '    decrypt(path, filename)
    'End Sub


    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            If layer_handler Is Nothing Then Exit Sub

            Dim str As String = ""
            Dim lol() As String = Regex.Split(Me.Text, "v.[0-9]{1,4}.[0-9]{1,4}.[0-9]{1,4}")
            str = lol(0) & Regex.Matches(Me.Text, "v.[0-9]{1,4}.[0-9]{1,4}.[0-9]{1,4}").Item(0).Value & vbCrLf
            str = str & Trim(lol(1)) & vbCrLf

            str = str & "Phi(rho*z):" & vbTab & options.phi_rz_mode & vbCrLf
            str = str & "MAC:" & vbTab & options.MAC_mode & vbCrLf
            str = str & "Electron impact ionization XS:" & vbTab & options.ionizationXS_mode & vbCrLf
            str = str & "SF by characteristic X-rays:" & vbTab & options.char_fluo_flag & vbCrLf
            str = str & "SF by bremsstrahlung:" & vbTab & options.brem_fluo_flag & vbCrLf
            str = str & vbCrLf

            For i As Integer = 0 To UBound(layer_handler)
                If i = UBound(layer_handler) Then
                    str = str & "Substrate" & vbCrLf
                Else
                    str = str & "Layer #" & i + 1 & "(nm)" & vbTab & Format(layer_handler(i).thickness / 10, "0.0") & vbCrLf
                End If
                str = str & "Density (g/cm3)" & vbTab & Format(layer_handler(i).density, "0.000") & vbCrLf
                str = str & "Element" & vbTab & "wt." & vbTab & "at." & vbCrLf

                Dim tot_wt As Double = 0
                Dim tot_at As Double = 0
                If layer_handler(i).element Is Nothing Then
                    MsgBox("Error: System not correctly defined. No element in layer #" & i + 1 & ".")
                    Exit Sub
                End If
                For j As Integer = 0 To UBound(layer_handler(i).element)
                    str = str & layer_handler(i).element(j).elt_name & vbTab & Format(layer_handler(i).element(j).conc_wt, "0.0000") & vbTab &
                        Format(layer_handler(i).element(j).conc_at, "0.0000") & vbCrLf
                    tot_wt = tot_wt + layer_handler(i).element(j).conc_wt
                    tot_at = tot_at + layer_handler(i).element(j).conc_at
                Next
                str = str & "Total" & vbTab & Format(tot_wt, "0.0000") & vbTab & Format(tot_at, "0.0000") & vbCrLf
                str = str & vbCrLf
            Next



            If Chart1.Series IsNot Nothing Then
                Dim max As Integer = 0
                For i As Integer = 0 To Chart1.Series.Count - 1
                    If Chart1.Series(i).Points.Count > max Then
                        max = Chart1.Series(i).Points.Count
                    End If
                Next

                'Dim res(max) As String
                str = str & vbTab
                For i As Integer = 0 To Chart1.Series.Count - 1
                    str = str & Chart1.Series(i).LegendText & vbTab & vbTab
                Next
                str = str & vbCrLf
                str = str & vbTab
                For i As Integer = 0 To Chart1.Series.Count - 1
                    str = str & "E (kV)" & vbTab & "k-ratio" & vbTab
                Next
                str = str & vbCrLf

                For i As Integer = 0 To max - 1
                    str = str & vbTab
                    For j As Integer = 0 To Chart1.Series.Count - 1
                        If Chart1.Series(j).Points.Count > i Then
                            str = str & Chart1.Series(j).Points(i).XValue & vbTab & Chart1.Series(j).Points(i).YValues(0) & vbTab
                        Else
                            str = str & vbTab & vbTab
                        End If
                    Next
                    str = str & vbCrLf
                Next
                Try
                    My.Computer.Clipboard.Clear()
                    If str <> "" Then My.Computer.Clipboard.SetText(str)
                    Label1.Text = "Status: Data copied to the clipboard."
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button9_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    'Private Sub CheckBox17_Click(sender As Object, e As EventArgs) Handles CheckBox17.Click
    '    If loaded = False Then Exit Sub
    '    If CheckBox17.Checked = True Then
    '        CheckBox12.Checked = False
    '        DataGridView2.Columns(1).HeaderText = "conc (at.)"
    '        If ListBox1.SelectedIndex < 0 Then Exit Sub
    '        layer_handler(ListBox1.SelectedIndex).wt_fraction = False
    '        'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
    '    End If
    '    If CheckBox17.Checked = False Then
    '        CheckBox12.Checked = True
    '    End If
    '    display_grid_layer()
    'End Sub

    'Private Sub CheckBox12_Click(sender As Object, e As EventArgs) Handles CheckBox12.Click
    '    If loaded = False Then Exit Sub
    '    If CheckBox12.Checked = False Then
    '        CheckBox12.Checked = True
    '        CheckBox17.Checked = False
    '        DataGridView2.Columns(1).HeaderText = "conc (wt)"
    '        If ListBox1.SelectedIndex < 0 Then Exit Sub
    '        layer_handler(ListBox1.SelectedIndex).wt_fraction = True
    '        'analysis_cond_handler(ListBox1.SelectedIndex).thickness = TextBox5.Text
    '    End If
    '    If CheckBox12.Checked = True Then
    '        CheckBox12.Checked = False
    '        CheckBox17.Checked = True
    '    End If
    '    display_grid_layer()
    'End Sub

    Private Sub CheckBox_Click_Conc(sender As Object, e As EventArgs) Handles CheckBox12.Click, CheckBox17.Click
        Try
            If loaded = False Then Exit Sub
            Dim senderCheck As CheckBox = DirectCast(sender, CheckBox)
            If senderCheck Is CheckBox12 Then
                CheckBox17.Checked = Not CheckBox17.Checked
            Else
                CheckBox12.Checked = Not CheckBox12.Checked
            End If

            'CheckBox12.Checked = Not CheckBox12.Checked
            'CheckBox17.Checked = Not CheckBox17.Checked
            If CheckBox12.Checked = True Then
                CheckBox17.Checked = False
                DataGridView2.Columns(1).HeaderText = "conc (wt.)"
                If check_valid_layer_selected() = False Then Exit Sub
                layer_handler(ListBox1.SelectedIndex).wt_fraction = True
            ElseIf CheckBox17.Checked = True Then
                CheckBox12.Checked = False
                DataGridView2.Columns(1).HeaderText = "conc (at.)"
                If check_valid_layer_selected() = False Then Exit Sub
                layer_handler(ListBox1.SelectedIndex).wt_fraction = False
            End If
            display_grid_layer()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox_Click_Conc " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox_Click_MAC_Model(sender As Object, e As EventArgs) Handles CheckBox4.Click, CheckBox5.Click, CheckBox3.Click, CheckBox19.Click, CheckBox23.Click
        Try
            If loaded = False Then Exit Sub
            Dim senderCheck As CheckBox = DirectCast(sender, CheckBox)
            If senderCheck Is CheckBox4 Then
                CheckBox5.Checked = False 'Not CheckBox5.Checked
                CheckBox3.Checked = False
                CheckBox19.Checked = False
                CheckBox4.Checked = True
                CheckBox23.Checked = False
            ElseIf senderCheck Is CheckBox5 Then
                CheckBox4.Checked = False 'Not CheckBox4.Checked
                CheckBox3.Checked = False
                CheckBox19.Checked = False
                CheckBox5.Checked = True
                CheckBox23.Checked = False
            ElseIf senderCheck Is CheckBox3 Then
                CheckBox4.Checked = False 'Not CheckBox4.Checked
                CheckBox5.Checked = False
                CheckBox19.Checked = False
                CheckBox3.Checked = True
                CheckBox23.Checked = False
            ElseIf senderCheck Is CheckBox19 Then
                CheckBox4.Checked = False 'Not CheckBox4.Checked
                CheckBox5.Checked = False
                CheckBox4.Checked = False
                CheckBox3.Checked = False
                CheckBox19.Checked = True
                CheckBox23.Checked = False
            ElseIf senderCheck Is CheckBox23 Then
                CheckBox4.Checked = False 'Not CheckBox4.Checked
                CheckBox5.Checked = False
                CheckBox4.Checked = False
                CheckBox3.Checked = False
                CheckBox19.Checked = False
                CheckBox23.Checked = True
            End If

            'CheckBox12.Checked = Not CheckBox12.Checked
            'CheckBox17.Checked = Not CheckBox17.Checked
            If CheckBox4.Checked = True Then
                'CheckBox5.Checked = False
                options.MAC_mode = "PENELOPE2018"
            ElseIf CheckBox5.Checked = True Then
                'CheckBox4.Checked = False
                options.MAC_mode = "MAC30"
            ElseIf CheckBox3.Checked = True Then
                'CheckBox4.Checked = False
                options.MAC_mode = "PENELOPE2014"
            ElseIf CheckBox19.Checked = True Then
                'CheckBox4.Checked = False
                options.MAC_mode = "FFAST"
            ElseIf CheckBox23.Checked = True Then
                options.MAC_mode = "EPDL23"
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox_Click_MAC_Model " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox_Click_phi_rho_z_model(sender As Object, e As EventArgs) Handles CheckBox9.Click, CheckBox10.Click, CheckBox11.Click, CheckBox22.Click
        Try
            If loaded = False Then Exit Sub
            Dim senderCheck As CheckBox = DirectCast(sender, CheckBox)
            If senderCheck Is CheckBox10 Then
                CheckBox9.Checked = False
                CheckBox10.Checked = True
                CheckBox11.Checked = False
                CheckBox22.Checked = False
            ElseIf senderCheck Is CheckBox11 Then
                CheckBox9.Checked = False
                CheckBox10.Checked = False
                CheckBox11.Checked = True
                CheckBox22.Checked = False
            ElseIf senderCheck Is CheckBox9 Then
                CheckBox9.Checked = True
                CheckBox10.Checked = False
                CheckBox11.Checked = False
                CheckBox22.Checked = False
            ElseIf senderCheck Is CheckBox22 Then
                CheckBox9.Checked = False
                CheckBox10.Checked = False
                CheckBox11.Checked = False
                CheckBox22.Checked = True
            End If

            'CheckBox12.Checked = Not CheckBox12.Checked
            'CheckBox17.Checked = Not CheckBox17.Checked
            If CheckBox10.Checked = True Then
                'CheckBox5.Checked = False
                options.phi_rz_mode = "PAP"
            ElseIf CheckBox11.Checked = True Then
                'CheckBox4.Checked = False
                options.phi_rz_mode = "PROZA96"
            ElseIf CheckBox9.Checked = True Then
                options.phi_rz_mode = "XPHI"
            ElseIf CheckBox22.Checked = True Then
                options.phi_rz_mode = "XPP"
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox_Click_phi_rho_z_model " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox_Click_ionizXS_Model(sender As Object, e As EventArgs) Handles CheckBox6.Click, CheckBox7.Click
        Try
            If loaded = False Then Exit Sub
            Dim senderCheck As CheckBox = DirectCast(sender, CheckBox)
            If senderCheck Is CheckBox6 Then
                CheckBox7.Checked = Not CheckBox7.Checked
            Else
                CheckBox6.Checked = Not CheckBox6.Checked
            End If

            'CheckBox12.Checked = Not CheckBox12.Checked
            'CheckBox17.Checked = Not CheckBox17.Checked
            If CheckBox6.Checked = True Then
                CheckBox7.Checked = False
                options.ionizationXS_mode = "Bote"
            ElseIf CheckBox7.Checked = True Then
                CheckBox6.Checked = False
                options.ionizationXS_mode = "OriPAP"
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox_Click_ionizXS_Model " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If HaveInternetConnection() = False Then
            MsgBox("No internet connection.")
            Exit Sub
        End If
        Try
            Dim version_check As String = Nothing
            Dim urls() As String = Nothing

            update_get_version_and_url(version_check, urls)

            If compare_version(version_check, VERSION) Then
                If urls.Count <> 0 Then
                    For i As Integer = 0 To UBound(urls)
                        update_download_file(urls(i))
                    Next
                    MsgBox("BadgerFilm updated to version: " & version_check & vbCrLf & "New files downloaded in " & Application.StartupPath & vbCrLf & "You can now close BadgerFilm and launch the new version.")
                Else
                    Debug.WriteLine("No new file but new version?")
                End If
            Else
                MsgBox("Your version of BadgerFilm is up-to-date. (" & VERSION & ")")
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button13_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub


    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'Dim Ec_data() As String = Nothing
        'init_Ec(Ec_data, pen_path)
        ''*******************************************
        ''Plot the k-ratio curves
        'plot_kratio(2, 30, 29, layer_handler, toa, Ec_data, pen_path, True)
        ''*******************************************
        If save_results = "" Then
            save_results = " "
            MsgBox("Nothing to be exported.")
        End If


        save_results = "E (kV)" & vbTab & "Elt" & vbTab & "X-ray line" & vbTab & "Characteristic X-ray intensity (ph/e-/sr)" & vbTab & "SF by characteristic X-rays (ph/e-/sr)" & vbTab &
            "SF by Bremsstrahlung (ph/e-/sr)" & vbTab & "SF char %" & vbTab & "SF brem %" & vbCrLf & save_results
        'Dim save_results_formatted As String = ""

        'Dim lines() As String = Split(save_results, vbCrLf)
        'Dim elt As String = Split(lines(0), vbTab)(0) & Split(lines(0), vbTab)(1)
        'For i As Integer = 1 To UBound(lines)
        '    Dim cnt As Integer = 0
        '    While elt = Split(lines(i), vbTab)(0) & Split(lines(i), vbTab)(1)
        '        cnt = cnt + 1
        '    End While
        '    Stop
        'Next
        Try
            My.Computer.Clipboard.Clear()
            If save_results <> "" Then My.Computer.Clipboard.SetText(save_results)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button14_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Try
            Dim OpenFileDialog1 As New OpenFileDialog()
            Dim data_files() As String = Nothing

            OpenFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            OpenFileDialog1.FilterIndex = 1
            OpenFileDialog1.RestoreDirectory = True
            OpenFileDialog1.Title = "Load data"
            OpenFileDialog1.AddExtension = True
            OpenFileDialog1.DefaultExt = ".txt"
            OpenFileDialog1.Multiselect = True

            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                data_files = OpenFileDialog1.FileNames
            Else
                Label1.Text = "Status: Loading canceled"
                Exit Sub
            End If

            If data_files.Count = 1 Then
                layer_handler = Nothing
                elt_exp_handler = Nothing
                import_Stratagem(data_files(0), layer_handler, elt_exp_handler, toa)
                Me.Text = "BadgerFilm " & VERSION & "  " & data_files(0)
            Else
                multi_import(data_files)
            End If

            update_form_fields()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button15_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            'Changed with version 1.2.15 (August 23 2021)
            'Added the following code to take into account the decimal separator (dot vs comma)
            Dim oldDecimalSeparator As String = Application.CurrentCulture.NumberFormat.NumberDecimalSeparator
            If oldDecimalSeparator = "." Then
            Else
                Dim forceDotCulture As CultureInfo
                forceDotCulture = Application.CurrentCulture.Clone()
                forceDotCulture.NumberFormat.NumberDecimalSeparator = "."
                Application.CurrentCulture = forceDotCulture
            End If

            ' Get the BackgroundWorker object that raised this event.
            'Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

            ' Assign the result of the computation
            ' to the Result property of the DoWorkEventArgs
            ' object. This is will be available to the 
            ' RunWorkerCompleted eventhandler.
            'e.Result = ComputeFibonacci(e.Argument, worker, e)
            BackgroundWorker1.ReportProgress(0, vbNull)

            Dim fitting_methode As New fitting_module
            'fitting_methode.fit(x, k_ratio_measured, ey, p, pars, buffer_text, layer_handler, elt_exp_handler, elt_exp_all, toa, pen_path, Ec_data, options)

            fitting_methode.fit(e.Argument(0), e.Argument(1), e.Argument(2), e.Argument(3), e.Argument(4), e.Argument(5), layer_handler, elt_exp_handler,
                                elt_exp_all, toa, pen_path, e.Argument(6), e.Argument(7), options)

            'myfit(e.Argument(0), e.Argument(1), e.Argument(2), e.Argument(3), e.Argument(4), e.Argument(5), layer_handler, elt_exp_handler,
            '                    elt_exp_all, toa, pen_path, e.Argument(6), e.Argument(7), options)

            'Force update of the theoretical k-ratios in the case where compositions and thicknesses are all fixed (fitting_status=-19).
            Dim fitting_status As Integer = 0
            fitting_status = Trim(Split(Split(e.Argument(5), "=")(1), "CHI")(0))
            If fitting_status = -19 Then
                For i As Integer = 0 To UBound(elt_exp_handler)
                    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                            elt_exp_handler(i).line(j).k_ratio(k).elt_intensity = pre_auto(layer_handler, elt_exp_handler(i), j, elt_exp_all,
                                                                           elt_exp_handler(i).line(j).k_ratio(k).kv, toa, e.Argument(6), options, False, "", e.Argument(7))
                            elt_exp_handler(i).line(j).k_ratio(k).theo_value = elt_exp_handler(i).line(j).k_ratio(k).elt_intensity / elt_exp_handler(i).line(j).k_ratio(k).std_intensity
                        Next
                    Next
                Next
            End If

            If BackgroundWorker1.CancellationPending = True Then
                e.Cancel = True
                Exit Sub
            End If

            BackgroundWorker1.ReportProgress(50, e.Argument(5))


            For i As Integer = 0 To UBound(layer_handler)
                convert_wt_to_at(layer_handler, i)
            Next

            Dim save_results As String = ""
            ''*******************************************
            ''Plot the k-ratio curves
            Dim data_to_plot As data_to_plot = Nothing
            plot_kratio(1, 40, 39, layer_handler, toa, e.Argument(6), pen_path, True, save_results, e.Argument(7), data_to_plot)
            ''*******************************************
            BackgroundWorker1.ReportProgress(55, save_results)
            BackgroundWorker1.ReportProgress(60, data_to_plot)

            '*******************************************
            'Plot_kratio the measured k-ratios
            BackgroundWorker1.ReportProgress(70, e.Argument(7)) '"Plot experimental k-ratios")

            If TypeOf e.Argument(7) Is fit_MAC Then
                Dim fit_MAC As fit_MAC = CType(e.Argument(7), fit_MAC)
                If fit_MAC.activated = True Then
                    BackgroundWorker1.ReportProgress(80, e.Argument(3))
                End If
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in backgroundWorker1_DoWork " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub 'backgroundWorker1_DoWork

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            '' This event is fired when you call the ReportProgress method from inside your DoWork.
            '' Any visual indicators about the progress should go here.
            'Label1.Text = CType(e.UserState, String)

            If e.ProgressPercentage = 0 Then
                Label1.Text = "Status: Calculating"

            ElseIf e.ProgressPercentage = 50 Then
                'TextBox12.Text = CType(e.UserState, String)
                'update_form_fields()

                Dim str As String = ""

                For i As Integer = 0 To UBound(layer_handler)
                    If i = UBound(layer_handler) Then
                        str = str & "Substrate" & vbCrLf
                    Else
                        str = str & "Layer #" & i + 1 & "(nm)" & vbTab & Format(layer_handler(i).thickness / 10, "0.0") & vbCrLf
                    End If
                    str = str & "Density (g/cm3)" & vbTab & Format(layer_handler(i).density, "0.000") & vbCrLf
                    str = str & "Element" & vbTab & "wt." & vbTab & "at." & vbCrLf

                    Dim tot_wt As Double = 0
                    Dim tot_at As Double = 0
                    If layer_handler(i).element Is Nothing Then
                        MsgBox("Error: System not correctly defined. No element in layer #" & i + 1 & ".")
                        Exit Sub
                    End If
                    For j As Integer = 0 To UBound(layer_handler(i).element)
                        str = str & layer_handler(i).element(j).elt_name & vbTab & Format(layer_handler(i).element(j).conc_wt, "0.0000") & vbTab &
                        Format(layer_handler(i).element(j).conc_at, "0.0000") & vbCrLf
                        tot_wt = tot_wt + layer_handler(i).element(j).conc_wt
                        tot_at = tot_at + layer_handler(i).element(j).conc_at
                    Next
                    str = str & "Total" & vbTab & Format(tot_wt, "0.0000") & vbTab & Format(tot_at, "0.0000") & vbCrLf
                    str = str & vbCrLf
                Next
                TextBox12.Text = str & vbCrLf & vbCrLf & CType(e.UserState, String)

                'update_form_fields()

                Label1.Text = "Status: Plotting results"

            ElseIf e.ProgressPercentage = 55 Then
                save_results = e.UserState

            ElseIf e.ProgressPercentage = 60 Then
                If TypeOf e.UserState Is data_to_plot Then
                    Dim data_to_plot As data_to_plot = CType(e.UserState, data_to_plot)
                    'Dim istep As Integer = 29
                    Dim reset As Boolean = True
                    Dim style As DataVisualization.Charting.SeriesChartType = DataVisualization.Charting.SeriesChartType.Line
                    For j As Integer = 0 To UBound(data_to_plot.elts_name)
                        Dim temp(UBound(data_to_plot.k_ratio, 2)) As Double
                        For i As Integer = 0 To UBound(data_to_plot.k_ratio, 2)
                            temp(i) = data_to_plot.k_ratio(j, i)
                        Next

                        graph_data_simple(data_to_plot.energy, temp, Chart1, "0.0", "0.00", graph_limits, reset, color_table(j Mod color_table.Count),
                                          "Acc. V. (kV)", "k-ratio", True, data_to_plot.elts_name(j), style)
                        reset = False

                    Next

                End If
                update_form_fields()

            ElseIf e.ProgressPercentage = 70 Then
                'Dim fit_MAC As fit_MAC
                'If TypeOf e.UserState Is fit_MAC Then
                '    fit_MAC = CType(e.UserState, fit_MAC)
                'End If
                'Dim max As Double = 0
                'For i As Integer = 0 To UBound(elt_exp_handler)
                '    For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                '        For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                '            If elt_exp_handler(i).line(j).k_ratio(k).experimental_value > max Then max = elt_exp_handler(i).line(j).k_ratio(k).experimental_value
                '        Next
                '    Next
                'Next
                Dim tmp As Integer = 0
                For j As Integer = 0 To UBound(elt_exp_handler)
                    For k As Integer = 0 To UBound(elt_exp_handler(j).line)
                        Dim kratio_meas() As Double = Nothing
                        Dim Acc_Volt() As Double = Nothing
                        For l As Integer = 0 To UBound(elt_exp_handler(j).line(k).k_ratio)
                            If kratio_meas Is Nothing Then
                                ReDim kratio_meas(0)
                                ReDim Acc_Volt(0)
                            Else
                                ReDim Preserve kratio_meas(UBound(kratio_meas) + 1)
                                ReDim Preserve Acc_Volt(UBound(Acc_Volt) + 1)
                            End If
                            kratio_meas(UBound(kratio_meas)) = elt_exp_handler(j).line(k).k_ratio(l).experimental_value
                            Acc_Volt(UBound(Acc_Volt)) = elt_exp_handler(j).line(k).k_ratio(l).kv
                        Next

                        'If fit_MAC.activated = True Then
                        '    For i As Integer = 0 To UBound(kratio_meas)
                        '        kratio_meas(i) = kratio_meas(i) / max
                        '    Next
                        'End If
                        Dim style As DataVisualization.Charting.SeriesChartType = DataVisualization.Charting.SeriesChartType.Point
                        graph_data_simple(Acc_Volt, kratio_meas, Chart1, "0.0", "0.00", graph_limits, False, color_table(tmp Mod color_table.Count), "Acc. V. (kV)", "k-ratio",
                            False, elt_exp_handler(j).elt_name & " " & elt_exp_handler(j).line(k).xray_name, style)
                        tmp = tmp + 1
                    Next
                Next
                Button6.Text = "Calculate"

            ElseIf e.ProgressPercentage = 80 Then
                If TypeOf e.UserState Is Double() Then
                    Dim p() As Double = CType(e.UserState, Double())
                    TextBox15.Text = Format(p.Last, "0.0")
                End If

            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in BackgroundWorker1_ProgressChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            '' This event is fired when your BackgroundWorker exits.
            '' It may have exitted Normally after completing its task, 
            '' or because of Cancellation, or due to any Error.

            If e.Error IsNot Nothing Then
                '' if BackgroundWorker terminated due to error
                MessageBox.Show(e.Error.Message)
                Label1.Text = "Status: Error occurred!"

            ElseIf e.Cancelled Then
                '' otherwise if it was cancelled
                'MessageBox.Show("Task canceled!")
                Label1.Text = "Status: Task canceled!"

            Else
                '' otherwise it completed normally
                'MessageBox.Show("Task completed!")
                Label1.Text = "Status: Task completed!"
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in BackgroundWorker1_RunWorkerCompleted " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        Try
            If e.Control AndAlso e.KeyCode = Keys.V Then
                Try
                    Dim col_index As Integer = DataGridView1.SelectedCells.Item(0).ColumnIndex
                    Dim row_index As Integer = DataGridView1.SelectedCells.Item(0).RowIndex
                    Dim lines() As String = Clipboard.GetText.Split(vbNewLine)
                    For i As Integer = 0 To UBound(lines)
                        If Not lines(i).Trim.ToString = "" Then
                            Dim item() As String = lines(i).Split(vbTab)
                            For j As Integer = 0 To UBound(item)
                                DataGridView1.Item(col_index + j, row_index + i).Value = item(j)
                            Next
                        End If
                    Next

                Catch ex As Exception
                    Debug.WriteLine(ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in DataGridView1_KeyDown " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Try
            Dim str As String = "BadgerFilm " & VERSION & vbCrLf & "Developed by Aurélien Moy and John Fournelle," & vbCrLf
            str = str & "Department of Geoscience, University of Wisconsin-Madison." & vbCrLf & vbCrLf
            str = str & "Contact:" & vbCrLf & "amoy6@wisc.edu" & vbCrLf & "johnf@geology.wisc.edu" & vbCrLf & vbCrLf
            str = str & "Support for this research came from the National Science Foundation: EAR-1554269 (JHF), EAR-1849386 (JHF)." & vbCrLf & vbCrLf
            str = str & "Atomic data extracted from the PENELOPE and the EADL databases:" & vbCrLf
            str = str & "- F. Salvat, PENELOPE-2014: a code system for Monte Carlo simulation of electron and photon transport, Issy-les-Moulineaux, France: OECD/NEA Data Bank, 2015 (available from http://www.nea.fr/lists/penelope.html)." & vbCrLf
            str = str & "- D.E. Cullen, et al., Tables and Graphs of Atomic Subshell and Relaxation Data Derived from the LLNL Evaluated Atomic Data Library (EADL), Z = 1 - 100, Lawrence Livermore National Laboratory, UCRL-50400, Vol. 30, October 1991." & vbCrLf & vbCrLf

            str = str & "To cite BadgerFilm:" & vbCrLf
            str = str & "- A. Moy & J. Fournelle, ϕ(ρz) Distributions in Bulk and Thin Film Samples for EPMA. Part 1: A Modified ϕ(ρz) Distribution for Bulk Materials, including Characteristic and Bremsstrahlung Fluorescence. Microscopy and Microanalysis, 2021, 27(2), 266–283." & vbCrLf & vbCrLf
            str = str & "- A. Moy & J. Fournelle, ϕ(ρz) Distributions in Bulk and Thin-Film Samples for EPMA. Part 2: BadgerFilm: A New Thin-Film Analysis Program. Microscopy and Microanalysis, 2021, 27(2), 284–296." & vbCrLf & vbCrLf
            'str = str & "IN NO EVENT SHALL BADGERFILM BE LIABLE TO ANY PARTY FOR DIRECT, INDIRECT, SPECIAL, INICIDENTAL, OR CONSEQUENTIAL DAMAGES, INCLUDING LOST PROFITS, ARISING OUT OF THE USE OF THIS SOFTWARE AND ITS DOCUMENTATION, EVEN IF BADGERFILM HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. BADGERFILM SPECIFICALLY DISCLAIMS ANY WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. THE SOFTWARE PROVIDED HEREUNDER IS ON AN AS IS BASIS, AND BADGERFILM HAVE NO OBLIGATIONS TO PROVIDE MAINTENANCE, SUPPORT, UPDATES, ENHANCEMENTS, OR MODIFICATIONS."
            str = str & "BADGERFILM TERMS OF USE

The following terms of use apply to the BadgerFilm software and its associated documentation ("“Badgerfilm”"). If you do not agree to these terms of use, do not use Badgerfilm. Badgerfilm is copyrighted by the Board of Regents of the University of Wisconsin System ("“UW”"), and may be downloaded, modified, and distributed for research and educational purposes; it may not be included in any product offered for sale. Should you modify and distribute Badgerfilm, you must not remove any copyright notices, you must cause any modified files to carry prominent notices that you changed the files, and you must ensure that any other recipients of Badgerfilm or any modified version of Badgerfilm receive a copy of this license. Use of Badgerfilm may require additional licenses to third party owned intellectual property contained within Badgerfilm, and obtaining any such license and complying with its terms is your responsibility. Badgerfilm is provided as a resource in connection with UW’s outreach mission, and is provided on an "“as is”" basis.

UW MAKES NO REPRESENTATIONS OR WARRANTIES CONCERNING BADGERFILM OR ANY OUTCOME THAT MAY BE OBTAINED BY USING OR MODIFYING BADGERFILM, AND EXPRESSLY DISCLAIMS ALL SUCH WARRANTIES, INCLUDING WITHOUT LIMITATION ANY EXPRESS OR IMPLIED WARRANTY OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT OF INTELLECTUAL PROPERTY RIGHTS. UW MAKES NO WARRANTY OR REPRESENTATION THAT BADGERFILM WILL OPERATE ERROR FREE OR UNINTERRUPTED.
TO THE FULLEST EXTENT PERMITTED BY LAW, IN NO EVENT SHALL UW OR THE AUTHORS BE LIABLE TO YOU (OR ANY PERSON, INSTITUTION, OR BUSINESS WITH WHICH YOU ARE AFFILIATED) FOR ANY LOST PROFITS OR ANY DIRECT, INDIRECT, EXEMPLARY, PUNITIVE, INCIDENTAL, SPECIAL, OR CONSEQUENTIAL DAMAGES ARISING FROM BADGERFILM OR ITS USE OR MODIFICATION. NEITHER UW NOR THE AUTHORS HAVE ANY LIABILITY FOR ANY DECISION, ACT OR OMISSION MADE BY YOU AS A RESULT OF USE OF BADGERFILM."
            MessageBox.Show(str)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Label2_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub


    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Try
            Dim path As String = "C:\Users\Aurélien\Desktop\BadgerFilm Releases\Thin film prog test\Fit Brem Batch\Ka\"
            Dim list_file() As String = {"C_Ka", "N_Ka", "O_Ka", "F_Ka", "Na_Ka", "Mg_Ka", "Al_Ka", "Si_Ka", "Cl_Ka", "Ca_Ka", "Fe_Ka", "Cu_Ka", "Ge_Ka",
                "Se_Ka", "Zr_Ka", "Mo_Ka"}
            'Dim list_BadgerFilm() As String

            For i As Integer = 0 To UBound(list_file)
                Dim data_file_Badger As String = path & list_file(i) & ".txt"

                Me.Text = "BadgerFilm " & VERSION & "  " & data_file_Badger
                'Dim analysis_cond_handler() As analysis_conditions = Nothing
                Chart1.Series.Clear()
                load_data(data_file_Badger, layer_handler, elt_exp_handler, toa)
                update_form_fields()
                TextBox12.Text = ""
                Label1.Text = "Status: Loaded"

                Dim data_file_PEN As String = path & list_file(i) & "_PEN.txt"
                Dim sr As StreamReader = New StreamReader(data_file_PEN)
                TextBox11.Text = sr.ReadToEnd()
                sr.Close()

                Button1_Click(sender, e)

            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button17_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Try
            Dim path As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Fit Brem Batch\La\"
            Dim list_file() As String = {"Fe_La", "Cu_La", "Ge_La", "Se_La", "Zr_La", "Mo_La", "Ag_La", "Sn_La", "Sb_La", "Te_La", "Cs_La", "Ba_La", "Ce_La", "Nd_La", "Gd_La", "Er_La", "Tm_La", "Yb_La",
               "Lu_La", "Hf_La", "W_La", "Pt_La", "Pb_La", "Po_La", "Rn_La", "Ra_La", "Th_La", "U_La", "Pu_La", "Am_La"}
            'Dim list_BadgerFilm() As String

            For i As Integer = 0 To UBound(list_file)
                Dim data_file_Badger As String = path & list_file(i) & ".txt"

                Me.Text = "BadgerFilm " & VERSION & "  " & data_file_Badger
                'Dim analysis_cond_handler() As analysis_conditions = Nothing
                Chart1.Series.Clear()
                load_data(data_file_Badger, layer_handler, elt_exp_handler, toa)
                update_form_fields()
                TextBox12.Text = ""
                Label1.Text = "Status: Loaded"

                Dim data_file_PEN As String = path & list_file(i) & "_PEN.txt"
                Dim sr As StreamReader = New StreamReader(data_file_PEN)
                TextBox11.Text = sr.ReadToEnd()
                sr.Close()

                Button1_Click(sender, e)


            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button19_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Try
            Dim path_data_file As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Pouchou and Pichoir data 1991.txt"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            Dim sr As StreamReader = New StreamReader(path_data_file)
            Dim text As String = sr.ReadToEnd()
            sr.Close()

            Dim lines() As String = Split(text, vbCrLf)

            Dim analyzed_elt(UBound(lines)) As String
            Dim X_ray_line(UBound(lines)) As String
            Dim companion(UBound(lines)) As String
            Dim kV(UBound(lines)) As Double
            Dim weight_frac(UBound(lines)) As Double
            Dim k_ratio(UBound(lines)) As Double
            Dim toas(UBound(lines)) As Double

            For i As Integer = 0 To UBound(lines)
                Dim tmp() As String = Split(lines(i), vbTab)
                analyzed_elt(i) = Z_to_symbol(tmp(1))
                X_ray_line(i) = num_to_Xray(tmp(2))
                companion(i) = Z_to_symbol(tmp(3))
                kV(i) = tmp(4)
                weight_frac(i) = tmp(5)
                k_ratio(i) = tmp(6)
                toas(i) = tmp(7)
            Next

            For i As Integer = 0 To UBound(lines)
                'If i = 467 Then
                '    Stop
                'End If
                Debug.Print(i)
                TextBox1.Text = toas(i)

                ReDim layer_handler(0)
                layer_handler(0).density = 3
                layer_handler(0).isfix = True
                layer_handler(0).thickness = 1000000.0
                layer_handler(0).wt_fraction = True

                layer_handler(0).id = 0

                Dim num_elt As Integer = 2
                ReDim layer_handler(0).element(num_elt - 1)

                'For j As Integer = 0 To num_elt - 1
                layer_handler(0).element(0).elt_name = analyzed_elt(i)
                layer_handler(0).element(0).isConcFixed = True
                layer_handler(0).element(0).conc_wt = weight_frac(i)
                layer_handler(0).element(0).mother_layer_id = 0

                layer_handler(0).element(1).elt_name = companion(i)
                layer_handler(0).element(1).isConcFixed = True
                layer_handler(0).element(1).conc_wt = 1 - weight_frac(i)
                layer_handler(0).element(1).mother_layer_id = 0
                'Next

                convert_wt_to_at(layer_handler, 0)

                init_element_layer(layer_handler(0).element(0).elt_name, vbNull, layer_handler(0).element(0))
                init_element_layer(layer_handler(0).element(1).elt_name, vbNull, layer_handler(0).element(1))
                layer_handler(0).mass_thickness = layer_handler(0).density * layer_handler(0).thickness * 10 ^ -8


                'Dim Ec_data() As String = Nothing
                'init_Ec(Ec_data, pen_path)


                ReDim elt_exp_handler(0)
                elt_exp_handler(0).elt_name = analyzed_elt(i)
                elt_exp_handler(0).z = symbol_to_Z(analyzed_elt(i))
                elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)

                ' Dim num_lines As Integer = 1
                ReDim elt_exp_handler(0).line(0)


                'elt_exp_handler(j).line(k).Ec = Split(Line(indice), vbTab).Last
                'elt_exp_handler(j).line(k).xray_energy = Split(Line(indice), vbTab).Last
                elt_exp_handler(0).line(0).xray_name = X_ray_line(i)
                elt_exp_handler(0).line(0).std = ""
                elt_exp_handler(0).line(0).std_filename = ""


                ' Dim num_kratios As Integer = 1
                ReDim elt_exp_handler(0).line(0).k_ratio(0)

                'For l As Integer = 0 To num_kratios - 1
                'elt_exp_handler(j).line(k).k_ratio(l).elt_intensity = Split(Line(indice), vbTab).Last
                elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                elt_exp_handler(0).line(0).k_ratio(0).experimental_value = k_ratio(i)
                'If VERSION <> "" Then
                'elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = Split(Line(indice), vbTab).Last
                'Else
                elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0
                'End If
                'elt_exp_handler(j).line(k).k_ratio(l).std_intensity = Split(Line(indice), vbTab).Last
                'elt_exp_handler(j).line(k).k_ratio(l).theo_value = Split(Line(indice), vbTab).Last
                ' Next

                init_element(elt_exp_handler(0).elt_name, X_ray_line(i), vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                elt_exp_all = Nothing
                init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)



                'For j As Integer = 0 To UBound(elt_exp_handler(0).line)
                If elt_exp_handler(0).line(0).k_ratio IsNot Nothing Then

                    Dim file_name As String = Split(elt_exp_handler(0).line(0).std_filename, "\").Last
                    Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(0).line(0).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                    If flag_file_exists = False Then
                        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                            elt_exp_handler(0).line(0).std_filename = Application.StartupPath & "\Examples\" & file_name
                            flag_file_exists = True
                        End If
                    End If
                    If IsNothing(elt_exp_handler(0).line(0).std_filename) = True Or flag_file_exists = False Then
                        'For kk As Integer = 0 To UBound(elt_exp_handler(0).line(0).k_ratio)
                        elt_exp_handler(0).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(0), 0, elt_exp_handler(0).line(0).k_ratio(0).kv, toa, Ec_data)
                        'Next

                    End If
                End If
                'Next

                elt_exp_handler(0).line(0).k_ratio(0).elt_intensity = pre_auto(layer_handler, elt_exp_handler(0), 0, elt_exp_all, elt_exp_handler(0).line(0).k_ratio(0).kv,
                                                                               toa, Ec_data, options, False, "", Nothing)
                elt_exp_handler(0).line(0).k_ratio(0).theo_value = elt_exp_handler(0).line(0).k_ratio(0).elt_intensity / elt_exp_handler(0).line(0).k_ratio(0).std_intensity

                If elt_exp_handler(0).line(0).k_ratio(0).theo_value > 1 Or Double.IsNaN(elt_exp_handler(0).line(0).k_ratio(0).theo_value) Then
                    Stop
                End If
                'update_form_fields()
                'Exit For
                TextBox13.Text = TextBox13.Text & elt_exp_handler(0).line(0).k_ratio(0).theo_value & vbCrLf
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button18_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub



    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Try
            Dim path_data_file As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Bastin and Heijligers 2000a.txt"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            Dim sr As StreamReader = New StreamReader(path_data_file)
            Dim text As String = sr.ReadToEnd()
            sr.Close()

            Dim lines() As String = Split(text, vbCrLf)

            Dim substrate_elt(UBound(lines)) As String
            Dim mass_thickness(UBound(lines)) As String
            Dim kratio_Al_Ka(UBound(lines)) As Double
            Dim kratio_substrate(UBound(lines)) As Double
            Dim kV(UBound(lines)) As Double
            Dim X_ray_line_substrate(UBound(lines)) As String

            Dim tmp_results As String = ""

            For i As Integer = 0 To UBound(lines)
                Dim tmp() As String = Split(lines(i), vbTab)
                substrate_elt(i) = Z_to_symbol(tmp(1))
                mass_thickness(i) = tmp(2)
                If IsNumeric(tmp(3)) Then
                    kratio_Al_Ka(i) = tmp(3)
                Else
                    kratio_Al_Ka(i) = 0
                End If
                If IsNumeric(tmp(4)) Then
                    kratio_substrate(i) = tmp(4)
                Else
                    kratio_substrate(i) = 0
                End If
                kV(i) = tmp(5)
                X_ray_line_substrate(i) = tmp(6)
            Next


            'Dim Ec_data() As String = Nothing
            'init_Ec(Ec_data, pen_path)

            For i As Integer = 0 To UBound(lines)
                If i = 466 Then
                    'Stop
                End If
                'Debug.Print(i + 1)
                'TextBox1.Text = toas(i)

                ReDim layer_handler(1)
                '********************************************************************
                layer_handler(1).density = zaro(symbol_to_Z(substrate_elt(i)))(1)
                'layer_handler(1).mass_thickness = mass_thickness(i)
                'layer_handler(1).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8 * 10 ^ 6)
                layer_handler(1).isfix = True
                layer_handler(1).thickness = 1000000.0
                layer_handler(1).mass_thickness = layer_handler(1).density * layer_handler(1).thickness * 10 ^ -8
                layer_handler(1).wt_fraction = True
                layer_handler(1).id = 1

                Dim num_elt As Integer = 1
                ReDim layer_handler(1).element(num_elt - 1)

                'For j As Integer = 0 To num_elt - 1
                layer_handler(1).element(0).elt_name = substrate_elt(i)
                layer_handler(1).element(0).isConcFixed = True
                layer_handler(1).element(0).conc_wt = 1
                layer_handler(1).element(0).mother_layer_id = 1
                convert_wt_to_at(layer_handler, 1)
                init_element_layer(layer_handler(1).element(0).elt_name, vbNull, layer_handler(1).element(0))
                'Next
                '********************************************************************

                '********************************************************************
                layer_handler(0).density = 2.7
                layer_handler(0).mass_thickness = mass_thickness(i) / 10 ^ 6
                layer_handler(0).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8)
                layer_handler(0).isfix = True
                'layer_handler(0).thickness = 1000000.0
                layer_handler(0).wt_fraction = True
                layer_handler(0).id = 0

                num_elt = 1
                ReDim layer_handler(0).element(num_elt - 1)

                layer_handler(0).element(0).elt_name = "Al"
                layer_handler(0).element(0).isConcFixed = True
                layer_handler(0).element(0).conc_wt = 1
                layer_handler(0).element(0).mother_layer_id = 0
                convert_wt_to_at(layer_handler, 0)
                init_element_layer(layer_handler(0).element(0).elt_name, vbNull, layer_handler(0).element(0))
                '********************************************************************





                If kratio_Al_Ka(i) <> 0 And kratio_substrate(i) <> 0 Then
                    ReDim elt_exp_handler(1)
                    elt_exp_handler(0).elt_name = "Al"
                    elt_exp_handler(0).z = 13
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, "Ka", vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    ' elt_exp_handler(0).line(0).xray_name = "Ka"
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_Al_Ka(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0

                    elt_exp_handler(1).elt_name = substrate_elt(i)
                    elt_exp_handler(1).z = symbol_to_Z(substrate_elt(i))
                    elt_exp_handler(1).a = zaro(elt_exp_handler(1).z)(0)
                    init_element(elt_exp_handler(1).elt_name, num_to_Xray_Bastin(X_ray_line_substrate(i)), vbNull, Ec_data, elt_exp_handler(1), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(1).line(0)
                    'elt_exp_handler(1).line(0).xray_name = num_to_Xray_Bastin(X_ray_line_substrate(i))
                    elt_exp_handler(1).line(0).std = ""
                    elt_exp_handler(1).line(0).std_filename = ""
                    ReDim elt_exp_handler(1).line(0).k_ratio(0)
                    elt_exp_handler(1).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(1).line(0).k_ratio(0).experimental_value = kratio_substrate(i)
                    elt_exp_handler(1).line(0).k_ratio(0).err_experimental_value = 0

                ElseIf kratio_Al_Ka(i) <> 0 Then
                    ReDim elt_exp_handler(0)
                    elt_exp_handler(0).elt_name = "Al"
                    elt_exp_handler(0).z = 13
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, "Ka", vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    'elt_exp_handler(0).line(0).xray_name = "Ka"
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_Al_Ka(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0

                Else
                    ReDim elt_exp_handler(0)
                    elt_exp_handler(0).elt_name = substrate_elt(i)
                    elt_exp_handler(0).z = symbol_to_Z(substrate_elt(i))
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, num_to_Xray_Bastin(X_ray_line_substrate(i)), vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    'elt_exp_handler(0).line(0).xray_name = num_to_Xray_Bastin(X_ray_line_substrate(i))
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_substrate(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0
                End If


                'For j As Integer = 0 To UBound(elt_exp_handler)
                '    init_element(elt_exp_handler(j).elt_name, elt_exp_handler(j).line(0).xray_name, vbNull, Ec_data, elt_exp_handler(j), pen_path, eadl_path, options)
                'Next
                elt_exp_all = Nothing
                init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)



                For j As Integer = 0 To UBound(elt_exp_handler)
                    If elt_exp_handler(j).line(0).k_ratio IsNot Nothing Then

                        Dim file_name As String = Split(elt_exp_handler(j).line(0).std_filename, "\").Last
                        Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(j).line(0).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                        If flag_file_exists = False Then
                            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                                elt_exp_handler(j).line(0).std_filename = Application.StartupPath & "\Examples\" & file_name
                                flag_file_exists = True
                            End If
                        End If
                        If IsNothing(elt_exp_handler(j).line(0).std_filename) = True Or flag_file_exists = False Then
                            'For kk As Integer = 0 To UBound(elt_exp_handler(0).line(0).k_ratio)
                            elt_exp_handler(j).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(j), 0, elt_exp_handler(j).line(0).k_ratio(0).kv, toa, Ec_data)
                            'Next

                        End If
                    End If


                    elt_exp_handler(j).line(0).k_ratio(0).elt_intensity = pre_auto(layer_handler, elt_exp_handler(j), 0, elt_exp_all, elt_exp_handler(j).line(0).k_ratio(0).kv,
                                                                               toa, Ec_data, options, False, "", Nothing)
                    elt_exp_handler(j).line(0).k_ratio(0).theo_value = elt_exp_handler(j).line(0).k_ratio(0).elt_intensity / elt_exp_handler(j).line(0).k_ratio(0).std_intensity

                    'If elt_exp_handler(0).line(0).k_ratio(0).theo_value > 1 Or Double.IsNaN(elt_exp_handler(0).line(0).k_ratio(0).theo_value) Then
                    'Stop
                    'End If
                    If elt_exp_handler(j).line(0).xray_name = "Ka" And elt_exp_handler(j).elt_name = "Al" Then
                        tmp_results = tmp_results & i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                            elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbCrLf

                        Debug.Print(i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                                    elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab &
                                        (1 - elt_exp_handler(j).line(0).k_ratio(0).theo_value / elt_exp_handler(j).line(0).k_ratio(0).experimental_value) * 100 & vbCrLf)
                    Else
                        Debug.Print(i + 1)
                    End If
                Next
                'TextBox13.Text = TextBox13.Text & vbCrLf
                'update_form_fields()
                'Exit For

            Next

            TextBox13.Text = tmp_results
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button20_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Try
            Dim path_data_file As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Bastin and Heijligers 2000a Pd.txt"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            Dim sr As StreamReader = New StreamReader(path_data_file)
            Dim text As String = sr.ReadToEnd()
            sr.Close()

            Dim lines() As String = Split(text, vbCrLf)

            Dim substrate_elt(UBound(lines)) As String
            Dim mass_thickness(UBound(lines)) As String
            Dim kratio_Al_Ka(UBound(lines)) As Double
            Dim kratio_substrate(UBound(lines)) As Double
            Dim kV(UBound(lines)) As Double
            Dim X_ray_line_substrate(UBound(lines)) As String

            Dim tmp_results As String = ""

            For i As Integer = 0 To UBound(lines)
                Dim tmp() As String = Split(lines(i), vbTab)
                substrate_elt(i) = Z_to_symbol(tmp(1))
                mass_thickness(i) = tmp(2)
                If IsNumeric(tmp(3)) Then
                    kratio_Al_Ka(i) = tmp(3)
                Else
                    kratio_Al_Ka(i) = 0
                End If
                If IsNumeric(tmp(4)) Then
                    kratio_substrate(i) = tmp(4)
                Else
                    kratio_substrate(i) = 0
                End If
                kV(i) = tmp(5)
                X_ray_line_substrate(i) = tmp(6)
            Next


            'Dim Ec_data() As String = Nothing
            'init_Ec(Ec_data, pen_path)

            For i As Integer = 0 To UBound(lines)
                If i = 466 Then
                    'Stop
                End If
                'Debug.Print(i + 1)
                'TextBox1.Text = toas(i)

                ReDim layer_handler(1)
                '********************************************************************
                layer_handler(1).density = zaro(symbol_to_Z(substrate_elt(i)))(1)
                'layer_handler(1).mass_thickness = mass_thickness(i)
                'layer_handler(1).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8 * 10 ^ 6)
                layer_handler(1).isfix = True
                layer_handler(1).thickness = 1000000.0
                layer_handler(1).mass_thickness = layer_handler(1).density * layer_handler(1).thickness * 10 ^ -8
                layer_handler(1).wt_fraction = True
                layer_handler(1).id = 1

                Dim num_elt As Integer = 1
                ReDim layer_handler(1).element(num_elt - 1)

                'For j As Integer = 0 To num_elt - 1
                layer_handler(1).element(0).elt_name = substrate_elt(i)
                layer_handler(1).element(0).isConcFixed = True
                layer_handler(1).element(0).conc_wt = 1
                layer_handler(1).element(0).mother_layer_id = 1
                convert_wt_to_at(layer_handler, 1)
                init_element_layer(layer_handler(1).element(0).elt_name, vbNull, layer_handler(1).element(0))
                'Next
                '********************************************************************

                '********************************************************************
                layer_handler(0).density = 12.0
                layer_handler(0).mass_thickness = mass_thickness(i) / 10 ^ 6
                layer_handler(0).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8)
                layer_handler(0).isfix = True
                'layer_handler(0).thickness = 1000000.0
                layer_handler(0).wt_fraction = True
                layer_handler(0).id = 0

                num_elt = 1
                ReDim layer_handler(0).element(num_elt - 1)

                layer_handler(0).element(0).elt_name = "Pd"
                layer_handler(0).element(0).isConcFixed = True
                layer_handler(0).element(0).conc_wt = 1
                layer_handler(0).element(0).mother_layer_id = 0
                convert_wt_to_at(layer_handler, 0)
                init_element_layer(layer_handler(0).element(0).elt_name, vbNull, layer_handler(0).element(0))
                '********************************************************************





                If kratio_Al_Ka(i) <> 0 And kratio_substrate(i) <> 0 Then
                    ReDim elt_exp_handler(1)
                    elt_exp_handler(0).elt_name = "Pd"
                    elt_exp_handler(0).z = 46
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, "La", vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    ' elt_exp_handler(0).line(0).xray_name = "Ka"
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_Al_Ka(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0

                    elt_exp_handler(1).elt_name = substrate_elt(i)
                    elt_exp_handler(1).z = symbol_to_Z(substrate_elt(i))
                    elt_exp_handler(1).a = zaro(elt_exp_handler(1).z)(0)
                    init_element(elt_exp_handler(1).elt_name, num_to_Xray_Bastin(X_ray_line_substrate(i)), vbNull, Ec_data, elt_exp_handler(1), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(1).line(0)
                    'elt_exp_handler(1).line(0).xray_name = num_to_Xray_Bastin(X_ray_line_substrate(i))
                    elt_exp_handler(1).line(0).std = ""
                    elt_exp_handler(1).line(0).std_filename = ""
                    ReDim elt_exp_handler(1).line(0).k_ratio(0)
                    elt_exp_handler(1).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(1).line(0).k_ratio(0).experimental_value = kratio_substrate(i)
                    elt_exp_handler(1).line(0).k_ratio(0).err_experimental_value = 0

                ElseIf kratio_Al_Ka(i) <> 0 Then
                    ReDim elt_exp_handler(0)
                    elt_exp_handler(0).elt_name = "Pd"
                    elt_exp_handler(0).z = 46
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, "La", vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    'elt_exp_handler(0).line(0).xray_name = "Ka"
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_Al_Ka(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0

                Else
                    ReDim elt_exp_handler(0)
                    elt_exp_handler(0).elt_name = substrate_elt(i)
                    elt_exp_handler(0).z = symbol_to_Z(substrate_elt(i))
                    elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                    init_element(elt_exp_handler(0).elt_name, num_to_Xray_Bastin(X_ray_line_substrate(i)), vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                    'ReDim elt_exp_handler(0).line(0)
                    'elt_exp_handler(0).line(0).xray_name = num_to_Xray_Bastin(X_ray_line_substrate(i))
                    elt_exp_handler(0).line(0).std = ""
                    elt_exp_handler(0).line(0).std_filename = ""
                    ReDim elt_exp_handler(0).line(0).k_ratio(0)
                    elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                    elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_substrate(i)
                    elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0
                End If


                'For j As Integer = 0 To UBound(elt_exp_handler)
                '    init_element(elt_exp_handler(j).elt_name, elt_exp_handler(j).line(0).xray_name, vbNull, Ec_data, elt_exp_handler(j), pen_path, eadl_path, options)
                'Next

                elt_exp_all = Nothing
                init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)



                For j As Integer = 0 To UBound(elt_exp_handler)
                    If elt_exp_handler(j).line(0).k_ratio IsNot Nothing Then

                        Dim file_name As String = Split(elt_exp_handler(j).line(0).std_filename, "\").Last
                        If elt_exp_handler(j).line(0).std_filename = "" Then
                            elt_exp_handler(j).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(j), 0, elt_exp_handler(j).line(0).k_ratio(0).kv, toa, Ec_data)
                        Else

                            Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(j).line(0).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                            If flag_file_exists = False Then
                                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                                    elt_exp_handler(j).line(0).std_filename = Application.StartupPath & "\Examples\" & file_name
                                    flag_file_exists = True
                                End If
                            End If
                            If IsNothing(elt_exp_handler(j).line(0).std_filename) = True Or flag_file_exists = False Then
                                'For kk As Integer = 0 To UBound(elt_exp_handler(0).line(0).k_ratio)
                                elt_exp_handler(j).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(j), 0, elt_exp_handler(j).line(0).k_ratio(0).kv, toa, Ec_data)
                                'Next

                            End If
                        End If
                    End If


                    elt_exp_handler(j).line(0).k_ratio(0).elt_intensity = pre_auto(layer_handler, elt_exp_handler(j), 0, elt_exp_all, elt_exp_handler(j).line(0).k_ratio(0).kv,
                                                                               toa, Ec_data, options, False, "", Nothing)
                    elt_exp_handler(j).line(0).k_ratio(0).theo_value = elt_exp_handler(j).line(0).k_ratio(0).elt_intensity / elt_exp_handler(j).line(0).k_ratio(0).std_intensity

                    'If elt_exp_handler(0).line(0).k_ratio(0).theo_value > 1 Or Double.IsNaN(elt_exp_handler(0).line(0).k_ratio(0).theo_value) Then
                    'Stop
                    'End If
                    If elt_exp_handler(j).line(0).xray_name = "La" And elt_exp_handler(j).elt_name = "Pd" Then
                        tmp_results = tmp_results & i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                            elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbCrLf

                        Debug.Print(i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                                    elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab &
                                        (1 - elt_exp_handler(j).line(0).k_ratio(0).theo_value / elt_exp_handler(j).line(0).k_ratio(0).experimental_value) * 100 & vbCrLf)
                    Else
                        Debug.Print(i + 1)
                    End If
                Next
                'TextBox13.Text = TextBox13.Text & vbCrLf
                'update_form_fields()
                'Exit For

            Next

            TextBox13.Text = tmp_results
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button21_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Try
            Dim path_data_file As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Bastin and Heijligers 2000a Pd v2.txt"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            Dim sr As StreamReader = New StreamReader(path_data_file)
            Dim text As String = sr.ReadToEnd()
            sr.Close()

            Dim lines() As String = Split(text, vbCrLf)

            Dim substrate_elt(UBound(lines)) As String
            Dim mass_thickness(UBound(lines)) As String
            Dim kratio_Al_Ka(UBound(lines)) As Double
            Dim kratio_substrate(UBound(lines)) As Double
            Dim kV(UBound(lines)) As Double
            Dim X_ray_line_substrate(UBound(lines)) As String

            Dim tmp_results As String = ""

            For i As Integer = 0 To UBound(lines)
                Dim tmp() As String = Split(lines(i), vbTab)
                substrate_elt(i) = Z_to_symbol(tmp(1))
                mass_thickness(i) = tmp(2)
                If IsNumeric(tmp(3)) Then
                    kratio_Al_Ka(i) = tmp(3)
                Else
                    kratio_Al_Ka(i) = 0
                End If
                If IsNumeric(tmp(4)) Then
                    kratio_substrate(i) = tmp(4)
                Else
                    kratio_substrate(i) = 0
                End If
                kV(i) = tmp(5)
                X_ray_line_substrate(i) = tmp(6)
            Next


            'Dim Ec_data() As String = Nothing
            'init_Ec(Ec_data, pen_path)

            For i As Integer = 0 To UBound(lines)
                If i = 466 Then
                    'Stop
                End If
                'Debug.Print(i + 1)
                'TextBox1.Text = toas(i)

                ReDim layer_handler(1)
                '********************************************************************
                layer_handler(1).density = zaro(symbol_to_Z(substrate_elt(i)))(1)
                'layer_handler(1).mass_thickness = mass_thickness(i)
                'layer_handler(1).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8 * 10 ^ 6)
                layer_handler(1).isfix = True
                layer_handler(1).thickness = 1000000.0
                layer_handler(1).mass_thickness = layer_handler(1).density * layer_handler(1).thickness * 10 ^ -8
                layer_handler(1).wt_fraction = True
                layer_handler(1).id = 1

                Dim num_elt As Integer = 1
                ReDim layer_handler(1).element(num_elt - 1)

                'For j As Integer = 0 To num_elt - 1
                layer_handler(1).element(0).elt_name = substrate_elt(i)
                layer_handler(1).element(0).isConcFixed = True
                layer_handler(1).element(0).conc_wt = 1
                layer_handler(1).element(0).mother_layer_id = 1
                convert_wt_to_at(layer_handler, 1)
                init_element_layer(layer_handler(1).element(0).elt_name, vbNull, layer_handler(1).element(0))
                'Next
                '********************************************************************

                '********************************************************************
                layer_handler(0).density = 12.0
                layer_handler(0).mass_thickness = mass_thickness(i) / 10 ^ 6
                layer_handler(0).thickness = layer_handler(0).mass_thickness / (layer_handler(0).density * 10 ^ -8)
                layer_handler(0).isfix = False
                'layer_handler(0).thickness = 1000000.0
                layer_handler(0).wt_fraction = True
                layer_handler(0).id = 0

                num_elt = 1
                ReDim layer_handler(0).element(num_elt - 1)

                layer_handler(0).element(0).elt_name = "Pd"
                layer_handler(0).element(0).isConcFixed = True
                layer_handler(0).element(0).conc_wt = 1
                layer_handler(0).element(0).mother_layer_id = 0
                convert_wt_to_at(layer_handler, 0)
                init_element_layer(layer_handler(0).element(0).elt_name, vbNull, layer_handler(0).element(0))
                '********************************************************************

                Dim start_indice As Integer = i
                elt_exp_handler = Nothing
                ReDim elt_exp_handler(1)
                While substrate_elt(i) = substrate_elt(start_indice) And mass_thickness(i) = mass_thickness(start_indice) And X_ray_line_substrate(i) = X_ray_line_substrate(start_indice)

                    If kratio_Al_Ka(i) <> 0 Then
                        If elt_exp_handler(0).line Is Nothing Then
                            elt_exp_handler(0).elt_name = "Pd"
                            elt_exp_handler(0).z = 46
                            elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)
                            init_element(elt_exp_handler(0).elt_name, "La", vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                            elt_exp_handler(0).line(0).std = ""
                            elt_exp_handler(0).line(0).std_filename = ""
                            ReDim elt_exp_handler(0).line(0).k_ratio(0)
                            elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                            elt_exp_handler(0).line(0).k_ratio(0).experimental_value = kratio_Al_Ka(i)
                            elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0
                        Else

                            ReDim Preserve elt_exp_handler(0).line(0).k_ratio(UBound(elt_exp_handler(0).line(0).k_ratio) + 1)
                            elt_exp_handler(0).line(0).k_ratio(UBound(elt_exp_handler(0).line(0).k_ratio)).kv = kV(i)
                            elt_exp_handler(0).line(0).k_ratio(UBound(elt_exp_handler(0).line(0).k_ratio)).experimental_value = kratio_Al_Ka(i)
                            elt_exp_handler(0).line(0).k_ratio(UBound(elt_exp_handler(0).line(0).k_ratio)).err_experimental_value = 0

                        End If
                    End If

                    If kratio_substrate(i) <> 0 Then
                        If elt_exp_handler(1).line Is Nothing Then
                            elt_exp_handler(1).elt_name = substrate_elt(i)
                            elt_exp_handler(1).z = symbol_to_Z(substrate_elt(i))
                            elt_exp_handler(1).a = zaro(elt_exp_handler(1).z)(0)
                            init_element(elt_exp_handler(1).elt_name, num_to_Xray_Bastin(X_ray_line_substrate(i)), vbNull, Ec_data, elt_exp_handler(1), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
                            elt_exp_handler(1).line(0).std = ""
                            elt_exp_handler(1).line(0).std_filename = ""
                            ReDim elt_exp_handler(1).line(0).k_ratio(0)
                            elt_exp_handler(1).line(0).k_ratio(0).kv = kV(i)
                            elt_exp_handler(1).line(0).k_ratio(0).experimental_value = kratio_substrate(i)
                            elt_exp_handler(1).line(0).k_ratio(0).err_experimental_value = 0
                        Else

                            ReDim Preserve elt_exp_handler(1).line(0).k_ratio(UBound(elt_exp_handler(1).line(0).k_ratio) + 1)
                            elt_exp_handler(1).line(0).k_ratio(UBound(elt_exp_handler(1).line(0).k_ratio)).kv = kV(i)
                            elt_exp_handler(1).line(0).k_ratio(UBound(elt_exp_handler(1).line(0).k_ratio)).experimental_value = kratio_substrate(i)
                            elt_exp_handler(1).line(0).k_ratio(UBound(elt_exp_handler(1).line(0).k_ratio)).err_experimental_value = 0

                        End If
                    End If

                    i = i + 1
                    If i > UBound(substrate_elt) Then
                        Exit While
                    End If
                End While
                i = i - 1

                If elt_exp_handler(0).line Is Nothing And elt_exp_handler(1).line Is Nothing Then
                    Continue For
                End If

                If elt_exp_handler(1).line Is Nothing Then
                    ReDim Preserve elt_exp_handler(0)
                End If

                If elt_exp_handler(0).line Is Nothing Then
                    elt_exp_handler(0) = elt_exp_handler(1)
                    ReDim Preserve elt_exp_handler(0)
                End If

                elt_exp_all = Nothing
                init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)

                TextBox12.Text = ""
                calculate(False)

                While True
                    If options.BgWorker.IsBusy = False Then
                        Exit While
                    End If
                    Application.DoEvents()
                End While

                'For j As Integer = 0 To UBound(elt_exp_handler)
                '    If elt_exp_handler(j).line(0).k_ratio IsNot Nothing Then

                '        Dim file_name As String = Split(elt_exp_handler(j).line(0).std_filename, "\").Last
                '        If elt_exp_handler(j).line(0).std_filename = "" Then
                '            elt_exp_handler(j).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(j), 0, elt_exp_handler(j).line(0).k_ratio(0).kv, toa, Ec_data)
                '        Else

                '            Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(j).line(0).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                '            If flag_file_exists = False Then
                '                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                '                    elt_exp_handler(j).line(0).std_filename = Application.StartupPath & "\Examples\" & file_name
                '                    flag_file_exists = True
                '                End If
                '            End If
                '            If IsNothing(elt_exp_handler(j).line(0).std_filename) = True Or flag_file_exists = False Then
                '                'For kk As Integer = 0 To UBound(elt_exp_handler(0).line(0).k_ratio)
                '                elt_exp_handler(j).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(j), 0, elt_exp_handler(j).line(0).k_ratio(0).kv, toa, Ec_data)
                '                'Next

                '            End If
                '        End If
                '    End If


                '    elt_exp_handler(j).line(0).k_ratio(0).elt_intensity = pre_auto(layer_handler, elt_exp_handler(j), 0, elt_exp_all, elt_exp_handler(j).line(0).k_ratio(0).kv,
                '                                                                   toa, Ec_data, options, False, "")
                '    elt_exp_handler(j).line(0).k_ratio(0).theo_value = elt_exp_handler(j).line(0).k_ratio(0).elt_intensity / elt_exp_handler(j).line(0).k_ratio(0).std_intensity

                '    'If elt_exp_handler(0).line(0).k_ratio(0).theo_value > 1 Or Double.IsNaN(elt_exp_handler(0).line(0).k_ratio(0).theo_value) Then
                '    'Stop
                '    'End If
                '    If elt_exp_handler(j).line(0).xray_name = "La" And elt_exp_handler(j).elt_name = "Pd" Then
                '        tmp_results = tmp_results & i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                '                elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbCrLf

                '        Debug.Print(i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & elt_exp_handler(j).z & vbTab & layer_handler(0).mass_thickness & vbTab & elt_exp_handler(j).line(0).k_ratio(0).kv & vbTab &
                '                        elt_exp_handler(j).line(0).k_ratio(0).theo_value & vbTab & elt_exp_handler(j).line(0).k_ratio(0).experimental_value & vbTab &
                '                            (1 - elt_exp_handler(j).line(0).k_ratio(0).theo_value / elt_exp_handler(j).line(0).k_ratio(0).experimental_value) * 100 & vbCrLf)
                '    Else
                '        Debug.Print(i + 1)
                '    End If
                'Next
                Debug.Print(TextBox12.Text)

                tmp_results = tmp_results & i + 1 & vbTab & layer_handler(0).element(0).z & " on " & layer_handler(1).element(0).z & vbTab & layer_handler(0).mass_thickness & vbCrLf

            Next

            TextBox13.Text = tmp_results
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button22_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Try
            Dim path As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Fit Brem Batch\Ma\"
            Dim list_file() As String = {"Nd_Ma", "Gd_Ma", "Er_Ma", "Tm_Ma", "Yb_Ma", "Lu_Ma", "Hf_Ma", "W_Ma", "Pt_Ma", "Pb_Ma", "Po_Ma", "Rn_Ma", "Ra_Ma", "Th_Ma", "U_Ma", "Pu_Ma", "Am_Ma"}
            'Dim list_BadgerFilm() As String

            For i As Integer = 0 To UBound(list_file)
                Dim data_file_Badger As String = path & list_file(i) & ".txt"

                Me.Text = "BadgerFilm " & VERSION & "  " & data_file_Badger
                'Dim analysis_cond_handler() As analysis_conditions = Nothing
                Chart1.Series.Clear()
                load_data(data_file_Badger, layer_handler, elt_exp_handler, toa)
                update_form_fields()
                TextBox12.Text = ""
                Label1.Text = "Status: Loaded"

                Dim data_file_PEN As String = path & list_file(i) & "_PEN.txt"
                Dim sr As StreamReader = New StreamReader(data_file_PEN)
                TextBox11.Text = sr.ReadToEnd()
                sr.Close()

                Button1_Click(sender, e)


            Next
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button23_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Sub multi_import(ByVal data_files() As String)
        Try
            layer_handler = Nothing
            elt_exp_handler = Nothing
            Dim layer_handler_tmp()() As layer
            ReDim layer_handler_tmp(UBound(data_files))
            Dim elt_exp_handler_tmp()() As Elt_exp
            ReDim elt_exp_handler_tmp(UBound(data_files))

            For i As Integer = 0 To UBound(data_files)
                Me.Text = "BadgerFilm " & VERSION & "  " & data_files(i)
                import_Stratagem(data_files(i), layer_handler_tmp(i), elt_exp_handler_tmp(i), toa)
            Next

            layer_handler = layer_handler_tmp(0)
            elt_exp_handler = elt_exp_handler_tmp(0)

            For i As Integer = 1 To UBound(elt_exp_handler_tmp)
                For j As Integer = 0 To UBound(elt_exp_handler_tmp(i))
                    If elt_exp_handler(j).elt_name = elt_exp_handler_tmp(i)(j).elt_name And elt_exp_handler(j).line.Count = elt_exp_handler_tmp(i)(j).line.Count Then
                        For k As Integer = 0 To UBound(elt_exp_handler(j).line)
                            If elt_exp_handler(j).line(k).xray_name = elt_exp_handler_tmp(i)(j).line(k).xray_name And
                                elt_exp_handler(j).line(k).xray_energy = elt_exp_handler_tmp(i)(j).line(k).xray_energy And
                                elt_exp_handler(j).line(k).k_ratio.Count = elt_exp_handler_tmp(i)(j).line(k).k_ratio.Count Then
                                For l As Integer = 0 To UBound(elt_exp_handler(j).line(k).k_ratio)
                                    If elt_exp_handler(j).line(k).k_ratio(l).kv = elt_exp_handler_tmp(i)(j).line(k).k_ratio(l).kv Then
                                        elt_exp_handler(j).line(k).k_ratio(l).experimental_value += elt_exp_handler_tmp(i)(j).line(k).k_ratio(l).experimental_value
                                        elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = Math.Sqrt((elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value) ^ 2 + (elt_exp_handler_tmp(i)(j).line(k).k_ratio(l).err_experimental_value) ^ 2)

                                    Else
                                        MsgBox("Error: Imported files have different electron beam accelerating voltage.")
                                        Exit Sub
                                    End If
                                Next
                            Else
                                MsgBox("Error: Imported files have different X-ray lines.")
                                Exit Sub
                            End If
                        Next
                    Else
                        MsgBox("Error: Imported files have different elements.")
                        Exit Sub
                    End If
                Next
            Next

            For i As Integer = 0 To UBound(elt_exp_handler)
                For j As Integer = 0 To UBound(elt_exp_handler(i).line)
                    For k As Integer = 0 To UBound(elt_exp_handler(i).line(j).k_ratio)
                        elt_exp_handler(i).line(j).k_ratio(k).experimental_value = elt_exp_handler(i).line(j).k_ratio(k).experimental_value / elt_exp_handler_tmp.Count
                        elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value = elt_exp_handler(i).line(j).k_ratio(k).err_experimental_value / elt_exp_handler_tmp.Count
                    Next
                Next
            Next

            'import_Stratagem(data_file, layer_handler, elt_exp_handler, toa)
            'Me.Text = "BadgerFilm " & VERSION & "  " & data_file
            'update_form_fields()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in multi_import " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub


    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Try
            Dim path_data_file As String = "F:\Work\BadgerFilm\BadgerFilm Releases\Thin film prog test\Heinrich_All_Data.txt"

            'init_atomic_parameters(pen_path, eadl_path, at_data, el_ion_xs, ph_ion_xs, MAC_data, options)
            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            Else
                MAC_data = MAC_data_PEN18
            End If

            Dim sr As StreamReader = New StreamReader(path_data_file)
            Dim text As String = sr.ReadToEnd()
            sr.Close()

            Dim lines() As String = Split(text, vbCrLf)

            Dim analyzed_elt(UBound(lines)) As String
            Dim X_ray_line(UBound(lines)) As String
            Dim companion(UBound(lines)) As String
            Dim kV(UBound(lines)) As Double
            Dim weight_frac(UBound(lines)) As Double
            Dim k_ratio(UBound(lines)) As Double
            Dim toas(UBound(lines)) As Double

            For i As Integer = 0 To UBound(lines)
                Dim tmp() As String = Split(lines(i), vbTab)
                X_ray_line(i) = num_to_Xray_Heinrich(tmp(1))
                analyzed_elt(i) = Z_to_symbol(tmp(2))
                companion(i) = Z_to_symbol(tmp(3))
                weight_frac(i) = tmp(4)
                kV(i) = tmp(5) / 1000
                toas(i) = tmp(7)
                k_ratio(i) = tmp(8)
            Next

            For i As Integer = 0 To UBound(lines)
                'If i = 467 Then
                '    Stop
                'End If
                Debug.Print(i)
                TextBox1.Text = toas(i)

                ReDim layer_handler(0)
                layer_handler(0).density = 3
                layer_handler(0).isfix = True
                layer_handler(0).thickness = 1000000.0
                layer_handler(0).wt_fraction = True

                layer_handler(0).id = 0

                Dim num_elt As Integer = 2
                ReDim layer_handler(0).element(num_elt - 1)

                'For j As Integer = 0 To num_elt - 1
                layer_handler(0).element(0).elt_name = analyzed_elt(i)
                layer_handler(0).element(0).isConcFixed = True
                layer_handler(0).element(0).conc_wt = weight_frac(i)
                layer_handler(0).element(0).mother_layer_id = 0

                layer_handler(0).element(1).elt_name = companion(i)
                layer_handler(0).element(1).isConcFixed = True
                layer_handler(0).element(1).conc_wt = 1 - weight_frac(i)
                layer_handler(0).element(1).mother_layer_id = 0
                'Next

                convert_wt_to_at(layer_handler, 0)

                init_element_layer(layer_handler(0).element(0).elt_name, vbNull, layer_handler(0).element(0))
                init_element_layer(layer_handler(0).element(1).elt_name, vbNull, layer_handler(0).element(1))
                layer_handler(0).mass_thickness = layer_handler(0).density * layer_handler(0).thickness * 10 ^ -8


                'Dim Ec_data() As String = Nothing
                'init_Ec(Ec_data, pen_path)


                ReDim elt_exp_handler(0)
                elt_exp_handler(0).elt_name = analyzed_elt(i)
                elt_exp_handler(0).z = symbol_to_Z(analyzed_elt(i))
                elt_exp_handler(0).a = zaro(elt_exp_handler(0).z)(0)

                ' Dim num_lines As Integer = 1
                'ReDim elt_exp_handler(0).line(0)


                init_element(elt_exp_handler(0).elt_name, X_ray_line(i), vbNull, Ec_data, elt_exp_handler(0), at_data, el_ion_xs, ph_ion_xs, MAC_data, options)

                'elt_exp_handler(j).line(k).Ec = Split(Line(indice), vbTab).Last
                'elt_exp_handler(j).line(k).xray_energy = Split(Line(indice), vbTab).Last
                'elt_exp_handler(0).line(0).xray_name = X_ray_line(i)
                elt_exp_handler(0).line(0).std = ""
                elt_exp_handler(0).line(0).std_filename = ""


                ' Dim num_kratios As Integer = 1
                ReDim elt_exp_handler(0).line(0).k_ratio(0)

                'For l As Integer = 0 To num_kratios - 1
                'elt_exp_handler(j).line(k).k_ratio(l).elt_intensity = Split(Line(indice), vbTab).Last
                elt_exp_handler(0).line(0).k_ratio(0).kv = kV(i)
                elt_exp_handler(0).line(0).k_ratio(0).experimental_value = k_ratio(i)
                'If VERSION <> "" Then
                'elt_exp_handler(j).line(k).k_ratio(l).err_experimental_value = Split(Line(indice), vbTab).Last
                'Else
                elt_exp_handler(0).line(0).k_ratio(0).err_experimental_value = 0
                'End If
                'elt_exp_handler(j).line(k).k_ratio(l).std_intensity = Split(Line(indice), vbTab).Last
                'elt_exp_handler(j).line(k).k_ratio(l).theo_value = Split(Line(indice), vbTab).Last
                ' Next

                elt_exp_all = Nothing
                init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)



                'For j As Integer = 0 To UBound(elt_exp_handler(0).line)
                If elt_exp_handler(0).line(0).k_ratio IsNot Nothing Then

                    Dim file_name As String = Split(elt_exp_handler(0).line(0).std_filename, "\").Last
                    Dim flag_file_exists As Boolean = My.Computer.FileSystem.FileExists(elt_exp_handler(0).line(0).std_filename) 'AMXXXXXXXXXXXXXXXXXX
                    If flag_file_exists = False Then
                        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Examples\" & file_name) = True Then
                            elt_exp_handler(0).line(0).std_filename = Application.StartupPath & "\Examples\" & file_name
                            flag_file_exists = True
                        End If
                    End If
                    If IsNothing(elt_exp_handler(0).line(0).std_filename) = True Or flag_file_exists = False Then
                        'For kk As Integer = 0 To UBound(elt_exp_handler(0).line(0).k_ratio)
                        elt_exp_handler(0).line(0).k_ratio(0).std_intensity = init_pure_std(elt_exp_handler(0), 0, elt_exp_handler(0).line(0).k_ratio(0).kv, toa, Ec_data)
                        'Next

                    End If
                End If
                'Next

                elt_exp_handler(0).line(0).k_ratio(0).elt_intensity = pre_auto(layer_handler, elt_exp_handler(0), 0, elt_exp_all, elt_exp_handler(0).line(0).k_ratio(0).kv,
                                                                               toa, Ec_data, options, False, "", Nothing)
                elt_exp_handler(0).line(0).k_ratio(0).theo_value = elt_exp_handler(0).line(0).k_ratio(0).elt_intensity / elt_exp_handler(0).line(0).k_ratio(0).std_intensity

                If elt_exp_handler(0).line(0).k_ratio(0).theo_value > 1 Or Double.IsNaN(elt_exp_handler(0).line(0).k_ratio(0).theo_value) Then
                    Stop
                End If
                'update_form_fields()
                'Exit For
                TextBox13.Text = TextBox13.Text & elt_exp_handler(0).line(0).k_ratio(0).theo_value & vbCrLf
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button24_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        Try
            If CheckBox13.Checked = True Then
                CheckBox7.Checked = True
                CheckBox6.Checked = False
                options.ionizationXS_mode = "OriPAP"
                GroupBox8.Visible = False
                GroupBox9.Visible = True
            Else
                GroupBox9.Visible = False
                GroupBox8.Visible = True
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox13_CheckedChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub select_elt_by_stoichio()
        Try
            If layer_handler Is Nothing Then Exit Sub
            If check_valid_layer_selected() = False Then Exit Sub

            If CheckBox20.Checked = False Then
                'Add O by stoichiometry to the system.
                'Check if O is already an element of the current layer.
                Dim index As Integer = -1
                If layer_handler(ListBox1.SelectedIndex).element IsNot Nothing Then
                    For i As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                        If layer_handler(ListBox1.SelectedIndex).element(i).elt_name = ComboBox1.SelectedItem Then '"O" Then
                            index = i
                            Exit For
                        End If
                    Next
                End If

                'If no O is present, add it to the system.
                If index = -1 Then
                    If layer_handler(ListBox1.SelectedIndex).element Is Nothing Then
                        ReDim layer_handler(ListBox1.SelectedIndex).element(0)
                    Else
                        ReDim Preserve layer_handler(ListBox1.SelectedIndex).element(UBound(layer_handler(ListBox1.SelectedIndex).element) + 1)
                    End If

                    Dim current_indice As Integer = UBound(layer_handler(ListBox1.SelectedIndex).element)
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).elt_name = ComboBox1.SelectedItem '"O"
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).conc_wt = 1
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).conc_at = 1
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).isConcFixed = True
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).z = symbol_to_Z(ComboBox1.SelectedItem)
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).a = zaro(layer_handler(ListBox1.SelectedIndex).element(current_indice).z)(0)
                    layer_handler(ListBox1.SelectedIndex).element(current_indice).mother_layer_id = ListBox1.SelectedIndex

                    'If O is already present, fix its concentration.
                Else
                    layer_handler(ListBox1.SelectedIndex).element(index).isConcFixed = True
                End If

                'Remove O from the list of experimental k-ratios if present (an element cannot be quantified by k-ratio and by stoichiometry at the same time).
                If elt_exp_handler IsNot Nothing Then
                    For i As Integer = 0 To UBound(elt_exp_handler)
                        If elt_exp_handler(i).elt_name = ComboBox1.SelectedItem Then '"O" Then
                            For j As Integer = i To UBound(elt_exp_handler) - 1
                                elt_exp_handler(j) = elt_exp_handler(j + 1)
                            Next
                            ReDim Preserve elt_exp_handler(UBound(elt_exp_handler) - 1)
                            Exit For
                        End If
                    Next
                End If
                'Mark the button corresponding to O in the periodic table as clicked.
                'Element8.BackColor = Color.Gray
                Dim Element As Class1.TestB = CType(Me.Controls("Element" & symbol_to_Z(ComboBox1.SelectedItem)), Class1.TestB)
                Element.BackColor = Color.Gray
                'Check the checkbox.
                CheckBox20.Checked = True
                CheckBox21.Enabled = True

            Else
                'Remove O by stoichiometry from the system.
                For i As Integer = 0 To UBound(layer_handler(ListBox1.SelectedIndex).element)
                    If layer_handler(ListBox1.SelectedIndex).element(i).elt_name = layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio_name Then '"O" Then
                        For j As Integer = i To UBound(layer_handler(ListBox1.SelectedIndex).element) - 1
                            layer_handler(ListBox1.SelectedIndex).element(j) = layer_handler(ListBox1.SelectedIndex).element(j + 1)
                        Next
                        ReDim Preserve layer_handler(ListBox1.SelectedIndex).element(UBound(layer_handler(ListBox1.SelectedIndex).element) - 1)
                        Exit For
                    End If
                Next
                'Mark the button corresponding to O in the periodic table as not clicked.
                'Element8.BackColor = Color.White
                Dim Element As Class1.TestB = CType(Me.Controls("Element" & symbol_to_Z(layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio_name)), Class1.TestB)
                Element.BackColor = Color.White
                'Uncheck the checkbox and unvlaidate Stoichiometry to O checkbox.
                CheckBox20.Checked = False
                CheckBox21.Checked = False
                CheckBox21.Enabled = False
                TextBox18.Enabled = False
                TextBox19.Enabled = False
            End If

            'Tells whether the selected layer is defining O by stoichiometry.
            layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio = CheckBox20.Checked
            layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio_name = ComboBox1.SelectedItem
            layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O = CheckBox21.Checked
            init_stoichio(layer_handler(ListBox1.SelectedIndex).stoichiometry.stoichio_table, "stoichiometry_" & layer_handler(ListBox1.SelectedIndex).stoichiometry.O_by_stoichio_name & ".txt")
            display_grid()
            display_grid_layer()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox20_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub CheckBox20_Click(sender As Object, e As EventArgs) Handles CheckBox20.Click
        select_elt_by_stoichio()
    End Sub


    Private Sub CheckBox21_Click(sender As Object, e As EventArgs) Handles CheckBox21.Click
        Try
            If check_valid_layer_selected() = False Then Exit Sub
            CheckBox21.Checked = Not (CheckBox21.Checked)
            If CheckBox21.Checked = True Then
                TextBox18.Enabled = True
                TextBox19.Enabled = True
            Else
                TextBox18.Enabled = False
                TextBox19.Enabled = False
            End If
            layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O = CheckBox21.Checked
            layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O_name = correct_symbol(TextBox19.Text)
            layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O_ratio = TextBox18.Text

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in CheckBox21_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        Try
            If check_valid_layer_selected() = False Then Exit Sub
            If layer_handler Is Nothing Then Exit Sub
            If IsNumeric(TextBox18.Text) Then
                layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O_ratio = TextBox18.Text
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox18_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        Try
            If check_valid_layer_selected() = False Then Exit Sub
            If layer_handler Is Nothing Then Exit Sub
            If TextBox19.Text = "" Then Exit Sub
            layer_handler(ListBox1.SelectedIndex).stoichiometry.Elt_by_stoichio_to_O_name = correct_symbol(TextBox19.Text)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in TextBox19_TextChanged " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Public Function check_valid_layer_selected() As Boolean
        Try
            If ListBox1.SelectedIndex < 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in check_valid_layer_selected " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Try
            test_matrix_inverse()
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button25_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Try
            Dim shell1 As Integer = 4
            'Dim shell2 As Integer = 9
            Dim En As Double = 15.0

            If options.MAC_mode = "PENELOPE2014" Then
                MAC_data = MAC_data_PEN14
            ElseIf options.MAC_mode = "PENELOPE2018" Then
                MAC_data = MAC_data_PEN18
            ElseIf options.MAC_mode = "FFAST" Then
                MAC_data = MAC_data_FFAST
            ElseIf options.MAC_mode = "EPDL23" Then
                MAC_data = MAC_data_EPDL23
            End If

            init_elt_exp_all(elt_exp_all, layer_handler, Ec_data, pen_path)


            Dim sigma As Double = Xray_production_xs_el_impact(elt_exp_all(0), shell1, En)

            Dim sigma_el_xs() As Double = interpol_log_log(elt_exp_all(0).el_ion_xs, En)

            Dim one_plus_TCK As Double = sigma / sigma_el_xs(shell1 - 1)

            MsgBox("1+TCK factor: " & one_plus_TCK)
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in Button26_Click " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Dim Z As Integer = 50
        Dim A As Double = zaro(Z)(0)
        Dim E_start As Double = 400
        Dim E_end As Double = 900

        init_Ec(Ec_data, pen_path)

        Dim text As New StringBuilder
        For i As Integer = 0 To 500
            text.AppendLine((E_end - E_start) * i / 500 + E_start & vbTab & Heinrich_MAC30(Z, (E_end - E_start) * i / 500 + E_start, A, Ec_data))
        Next

        Dim res As String = text.ToString

        Clipboard.SetText(res)

    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        select_elt_by_stoichio()
        select_elt_by_stoichio()
        CheckBox21.Text = "Stoichiometry to " & ComboBox1.SelectedItem & ":"
        Label19.Text = "to " & ComboBox1.SelectedItem
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged

    End Sub
End Class
