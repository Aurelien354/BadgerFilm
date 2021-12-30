Imports System.IO

Module expint_module
    '     Exponential Integral

    Public Function expint(ByVal x1 As Double) As Double
        Try
            If x1 < -30 Then Return 0
            If Double.IsNaN(x1) Then Return 0

            Dim x As Double = x1
            Dim xiter As Double = 1.0
            Dim t As Double = x

            expint = 0.5772156649 + Math.Log(Math.Abs(x)) + x '0153287

            While True
                xiter = xiter + 1.0
                t = x * t * (xiter - 1) / xiter / xiter
                expint = expint + t
                If (Math.Abs(t / expint) < 0.000000000001) Then Return expint
                If Double.IsInfinity(expint) Then Return expint
            End While

            'If x1 < -30 Then Return 0

            'Dim x, sign, xiter, t As Double
            'If Double.IsNaN(x1) Then Return 0

            'x = x1
            'sign = 1.0
            'If (x < 0.0) Then sign = -sign
            'If (x < 0.0) Then x = -x

            'xiter = 1.0
            't = sign * x
            'eidp = 0.5772156649 + Math.Log(x) + sign * x

            'While True
            '    xiter = xiter + 1.0
            '    t = sign * x * t * (xiter - 1) / xiter / xiter
            '    eidp = eidp + t
            '    If (Math.Abs(t / eidp) < 0.000000000001) Then Return eidp
            '    If Double.IsInfinity(eidp) Then Return eidp
            'End While

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in expint " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MsgBox(tmp)
        End Try

    End Function
End Module
