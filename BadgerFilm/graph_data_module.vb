Imports System.IO

Module graph_data_module
    Public Sub graph_data_simple(ByVal x() As Double, ByVal y() As Double, ByRef chart1 As DataVisualization.Charting.Chart,
                                 ByVal precisionX As String, ByVal precisionY As String, ByVal graph_limits() As Double, ByVal reset As Boolean,
                                 ByVal color As String, ByVal AxisX_Title As String, ByVal AxisY_Title As String,
                                 ByVal IsVisibleInLegend As Boolean, ByVal legendText As String, ByVal style As DataVisualization.Charting.SeriesChartType)
        Try

            If (x Is Nothing) Then
                MsgBox("Empty data in image_extractor->graph_data")
                Exit Sub
            End If
            If (y Is Nothing) Then
                MsgBox("Empty data rows in image_extractor->graph_data")
                Exit Sub
            End If

            If reset = True Then
                chart1.Series.Clear()
            End If

            If chart1.Legends.Count = 0 Then
                chart1.Legends.Add(New DataVisualization.Charting.Legend("Legend"))
            End If

            chart1.Series.Add("Plot_" & chart1.Series.Count)

            With chart1.Series("Plot_" & chart1.Series.Count - 1)
                .Points.DataBindXY(x, y)
                .ChartType = style 'DataVisualization.Charting.SeriesChartType.Line
                If style = DataVisualization.Charting.SeriesChartType.Line Then
                    .BorderWidth = 2
                Else
                    .MarkerSize = 7
                End If

                If color IsNot Nothing Then
                    .Color = Drawing.Color.FromName(color)
                End If
                .IsVisibleInLegend = IsVisibleInLegend
                .LegendText = legendText
                .Legend = "Legend"
            End With
            '

            chart1.ChartAreas(0).RecalculateAxesScale()
            chart1.Series("Plot_" & chart1.Series.Count - 1).ChartArea = chart1.ChartAreas(0).Name

            With chart1.ChartAreas(0)
                .AxisX.Title = AxisX_Title
                .AxisY.Title = AxisY_Title
                .AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
                .AxisY.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.NotSet
                .AxisX.LabelStyle.Format = precisionX
                .AxisY.LabelStyle.Format = precisionY
                .AxisX.Minimum = graph_limits(0)
                .AxisX.Maximum = graph_limits(1)
                .AxisY.Minimum = graph_limits(2)
                .AxisY.Maximum = graph_limits(3)
            End With


        Catch Ex As Exception
            MessageBox.Show("Error in graph_data_module->graph_data_simple: " & Ex.Message)

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine("Error in graph_data_module->graph_data_simple: " & Ex.Message)
            End Using
        End Try


    End Sub
End Module
