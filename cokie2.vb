    Private Sub main_proc()
        Dim ck As String
        Dim ret As Boolean
        Dim writer As System.IO.StreamWriter
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim enc2 As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim url As String

        Dim cc As CookieContainer = New CookieContainer()
        url = "http://xxxx"
        ck = getCookie(url)
        ck = ck.Replace(";", ",")

        url = "http://xxxxxx"
        cc.SetCookies(New Uri(url), ck)

        Dim req As HttpWebRequest _
          = CType(WebRequest.Create(url), HttpWebRequest)
        req.CookieContainer = cc

        Dim res As WebResponse = req.GetResponse()

        ' レスポンスの読み取り
        Dim resStream As Stream = res.GetResponseStream()
        Dim sr As StreamReader = New StreamReader(resStream, enc2)
        Dim result As String = sr.ReadToEnd()
        sr.Close()
        resStream.Close()

        writer = New System.IO.StreamWriter("aaa.htm", False, enc2)
        writer.Write(result)
        writer.Close()
        MsgBox("main_proc end")

    End Sub
