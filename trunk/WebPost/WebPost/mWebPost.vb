Imports System.Net
Imports System.IO
Imports System.Text

Module mWebPost

    Sub Main(ByVal cmdArgs() As String)
        Dim method As String = "post"

        If cmdArgs.Count > 0 Then
            For i As Integer = 0 To cmdArgs.Length - 1
                ' parse args:
                '   -XXX changes http method
                '   other values are urls that will be requested
                Dim arg As String = cmdArgs(i)

                If arg.StartsWith("-") Then
                    method = arg.Substring(1)
                    Debug.Print(method)
                Else
                    Dim URL As String = System.Uri.EscapeUriString(arg)
                    Debug.Print(URL)
                    Dim ret As String = getWebSite(URL, method)
                    Debug.Print(ret)
                End If
            Next
        Else
            ' default behaviour when no params are given
            Dim URL As String = "http://192.168.178.10:8081/control.cgi?command=play"
            Dim ret As String = getWebSite(URL, "post")
            Debug.Print(ret)
        End If
    End Sub

    Public Function getWebSite(ByVal url As String, ByVal method As String) As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim readStream As StreamReader = Nothing
        Dim htmlContent As String = ""

        Try
            ' Create the web request
            request = DirectCast(WebRequest.Create(url), HttpWebRequest)
            ' Set credentials to use for this request.
            request.Credentials = CredentialCache.DefaultCredentials
            ' Get response
            request.Method = method.ToUpper()
            Debug.Print(request.Method)
            request.ContentType = "application/octet-stream"
            request.ContentLength = 0
            response = DirectCast(request.GetResponse(), HttpWebResponse)
            If Not response Is Nothing Then response.Close()

            Dim receiveStream As Stream = response.GetResponseStream()
            readStream = New StreamReader(receiveStream, Encoding.UTF8) 'Default, Encoding.GetEncoding(1252)
            Dim s As Encoding = Encoding.Default
            htmlContent = readStream.ReadToEnd()
        Catch ex As Exception
            If Err.Number = 5 Then
                Debug.Print(">>>> " & Err.Description)
            Else
                Debug.Print(Err.Number)
                Debug.Print("caught Exception: getWebSite")
                Debug.Print("----------------------------------------------------------")
                Debug.Print("Message:" & ex.Message)
                Debug.Print("StackTrace:" & ex.StackTrace)
            End If

        Finally
            If Not response Is Nothing Then response.Close()
        End Try

        If Not response Is Nothing Then response.Close()
        If Not readStream Is Nothing Then readStream.Close()

        Return htmlContent
    End Function

End Module
