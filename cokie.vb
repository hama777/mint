    Private Sub cokie()
        'Dim enc As Encoding = Encoding.UTF8
        Dim writer As StreamWriter

        Dim url As String = "https://xxxx"
        Dim param As String = ""
        Dim cc As CookieContainer = New CookieContainer()

        Dim ht As Hashtable = New Hashtable()

        ht("u") = "userid"
        ht("p") = "pass"
        ht("service_id") = "n16"
        ht("return_url") = "action/my/MRF07_008_01.do"
        ht("submit") = "%a5%ed%a5%b0%a5%a4%a5%f3"

        For Each k As String In ht.Keys
            param = param & String.Format("{0}={1}&", k, ht(k))
        Next
        Dim data As Byte() = Encoding.ASCII.GetBytes(param)

        Dim req As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/x-www-form-urlencoded"
        req.ContentLength = data.Length
        req.CookieContainer = cc


        Dim reqStream As Stream = req.GetRequestStream()
        reqStream.Write(data, 0, data.Length)
        reqStream.Close()


        Dim res As WebResponse = req.GetResponse()
        Dim resStream As Stream = res.GetResponseStream()
        Dim sr As StreamReader = New StreamReader(resStream, System.Text.Encoding.GetEncoding("euc-jp"))
        Dim html As String = sr.ReadToEnd()
        sr.Close()
        resStream.Close()

        writer = New StreamWriter("aaa.htm", False, System.Text.Encoding.GetEncoding("euc-jp"))
        'writer = New StreamWriter("data.htm", False)
        writer.WriteLine(html)
        MessageBox.Show("post end ")

        url = "https://xxxxxxxxx"
        req = CType(WebRequest.Create(url), HttpWebRequest)
        req.CookieContainer = cc
        req.Method = "GET"
        res = req.GetResponse()

        resStream = res.GetResponseStream()
        sr = New StreamReader(resStream, System.Text.Encoding.GetEncoding("euc-jp"))
        Dim result As String = sr.ReadToEnd()
        sr.Close()
        resStream.Close()

        writer = New StreamWriter("bbb.htm", False, System.Text.Encoding.GetEncoding("euc-jp"))
        'writer = New StreamWriter("data.htm", False)
        writer.WriteLine(html)
        MessageBox.Show("cc  end ")



    End Sub
