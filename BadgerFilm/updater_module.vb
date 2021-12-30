Imports System.IO
Imports System.Net
Imports System.Net.Http

Module updater_module
    Public RemoteUri As String = "https://drive.google.com/uc?export=download&id=1eYNqOaroLFk8jEVGqWiDctdUn6oPLRfa"
    Public AternateUri As String = "https://drive.google.com/uc?export=download&id=1eYNqOaroLFk8jEVGqWiDctdUn6oPLRfa"

    Public Function HaveInternetConnection() As Boolean
        'If My.Computer.Network.IsAvailable Then
        '    'MsgBox("Computer is connected.")
        '    Return True
        'Else
        '    'MsgBox("Computer is not connected.")
        '    Return False
        'End If
        Try
            Return My.Computer.Network.Ping("www.google.com")
        Catch
            Return False
        End Try
    End Function

    Public Sub update_get_version(ByRef version As String)
        Try

            Dim client As HttpClient = New HttpClient()
            Dim Response As HttpResponseMessage = client.GetAsync(RemoteUri).Result
            If Response.IsSuccessStatusCode = False Then
                RemoteUri = AternateUri
            End If

            Dim myWebClient As New WebClient
            Dim file As New System.IO.StreamReader(myWebClient.OpenRead(RemoteUri))
            Dim Contents As String = file.ReadToEnd()
            file.Close()

            Dim lines() As String = Split(Contents, vbCrLf)
            version = lines(0)
            'Catch ex As WebException
            '    Dim errorResponse As HttpWebResponse = ex.Response 'As HttpWebResponse
            '    If (errorResponse.StatusCode = HttpStatusCode.NotFound) Then

            '    End If
        Catch
            version = ""
        End Try
    End Sub

    Public Sub update_get_version_and_url(ByRef version As String, ByRef urls() As String)
        Try
            Dim myWebClient As New WebClient

            Dim file As New System.IO.StreamReader(myWebClient.OpenRead(RemoteUri))
            Dim Contents As String = file.ReadToEnd()
            file.Close()

            Dim lines() As String = Split(Contents, vbCrLf)
            version = lines(0)
            ReDim urls(UBound(lines) - 1)
            For i As Integer = 1 To UBound(lines)
                urls(i - 1) = lines(i)
            Next

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in update_get_version_and_url " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Sub

    Public Sub update_download_file(ByVal line As String)
        Try
            Dim myWebClient As New WebClient
            'myWebClient.DownloadFile(RemoteUri & Files(i),  Application.StartupPath & "\" & Files(i))

            Dim parse() As String = Split(line, vbTab)

            If parse.Count = 2 Then
                If parse(0) = "BadgerFilm.exe" Then
                    My.Computer.FileSystem.RenameFile("BadgerFilm.exe", "BadgerFilm_" & Form1.VERSION & ".exe.bak")
                End If
                myWebClient.DownloadFile(parse(1), Application.StartupPath & "\" & parse(0))
            Else
                MsgBox("Error in update function.")
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in update_download_file " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Sub

    Public Function compare_version(ByVal internet_version As String, ByVal current_version As String) As Boolean
        Try
            Dim internet_version_int() As String = Split(internet_version, ".")
            Dim current_version_int() As String = Split(current_version, ".")

            Dim internet_version_tot As Integer = 0
            For i As Integer = 1 To UBound(internet_version_int)
                internet_version_tot = internet_version_tot + internet_version_int(i) * 10 ^ (internet_version_int.Count - i - 1)
            Next
            Dim current_version_tot As Integer = 0
            For i As Integer = 1 To UBound(current_version_int)
                current_version_tot = current_version_tot + current_version_int(i) * 10 ^ (current_version_int.Count - i - 1)
            Next

            If internet_version_tot > current_version_tot Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Dim tmp As String = Date.Now.ToString & vbTab & "Error in compare_version " & ex.Message

            Using err As StreamWriter = New StreamWriter("log.txt", True)
                err.WriteLine(tmp)
            End Using
            MessageBox.Show(tmp)
        End Try
    End Function

End Module
