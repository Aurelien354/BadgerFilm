Imports System.IO

Module Bastin_test_module
    Public Function num_to_Xray(ByVal num As Integer) As String
        Try
            If num = 1 Then Return "Ka"
            If num = 2 Then Return "Kb"
            If num = 3 Then Return "La"
            If num = 4 Then Return "Lb"
            If num = 5 Then Return "Ma"
            If num = 6 Then Return "Mb"
            Stop
            Return ""
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in num_to_Xray " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Public Function num_to_Xray_Bastin(ByVal num As Integer) As String
        Try
            If num = 0 Then Return "Ka"
            If num = 1 Then Return "La"
            If num = 2 Then Return "Ma"
            Stop
            Return ""
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in num_to_Xray_Bastin " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

    Public Function num_to_Xray_Heinrich(ByVal num As Integer) As String
        Try
            If num = 1 Then Return "Ka"
            If num = 2 Then Return "Kb"
            If num = 3 Then Return "La"
            If num = 4 Then Return "Lb"
            If num = 8 Then Return "Ma"
            If num = 9 Then Return "Mb"
            Stop
            Return ""
        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in num_to_Xray_Heinrich " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try
    End Function

End Module
