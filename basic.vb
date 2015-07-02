    Private Sub basic()
        Dim url As String
        Dim writer As StreamWriter

        url = "https://xxxxx"
        Dim webreq As System.Net.HttpWebRequest = CType(System.Net.WebRequest.Create(url), System.Net.HttpWebRequest)

        webreq.Credentials = New System.Net.NetworkCredential("userid", "pass")
        Dim webres As System.Net.HttpWebResponse = CType(webreq.GetResponse(), System.Net.HttpWebResponse)
        Dim st As System.IO.Stream = webres.GetResponseStream()
        Dim sr As New System.IO.StreamReader(st, System.Text.Encoding.GetEncoding("euc-jp"))

        writer = New StreamWriter("bbb.htm", False, System.Text.Encoding.GetEncoding("shift_jis"))
        writer.WriteLine(sr.ReadToEnd())

        sr.Close()
        st.Close()

        writer.Close()

    End Sub
