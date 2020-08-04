Module Bastin_test_module
    Public Function num_to_Xray(ByVal num As Integer) As String
        If num = 1 Then Return "Ka"
        If num = 2 Then Return "Kb"
        If num = 3 Then Return "La"
        If num = 4 Then Return "Lb"
        If num = 5 Then Return "Ma"
        If num = 6 Then Return "Mb"
        Stop
        Return ""
    End Function

    Public Function num_to_Xray_Bastin(ByVal num As Integer) As String
        If num = 0 Then Return "Ka"
        If num = 1 Then Return "La"
        If num = 2 Then Return "Ma"
        Stop
        Return ""
    End Function

    Public Function num_to_Xray_Heinrich(ByVal num As Integer) As String
        If num = 1 Then Return "Ka"
        If num = 2 Then Return "Kb"
        If num = 3 Then Return "La"
        If num = 4 Then Return "Lb"
        If num = 8 Then Return "Ma"
        If num = 9 Then Return "Mb"
        Stop
        Return ""
    End Function

End Module
