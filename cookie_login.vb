    '  クッキーを使ってアクセスする  その2
    '  参考  

    Shared encoder As Encoding = Encoding.GetEncoding("utf-8")

    Shared Sub main()
        Dim writer As System.IO.StreamWriter
        Dim id As String = "xxxx"
        Dim password As String = "passwd"
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

        Dim cc As CookieContainer = New CookieContainer()

        Login(id, password, cc)
        Dim result As String = ReadLog(cc)

        writer = New System.IO.StreamWriter("aaa.htm", False, enc)
        writer.Write(result)
        writer.Close()
        MsgBox("end")



    End Sub

    Shared Function HttpGet(url As String, cc As CookieContainer) As String

        ' リクエストの作成
        Dim req As HttpWebRequest _
          = CType(WebRequest.Create(url), HttpWebRequest)
        req.CookieContainer = cc

        Dim res As WebResponse = req.GetResponse()

        ' レスポンスの読み取り
        Dim resStream As Stream = res.GetResponseStream()
        Dim sr As StreamReader = New StreamReader(resStream, encoder)
        Dim result As String = sr.ReadToEnd()
        sr.Close()
        resStream.Close()

        Return result
    End Function

    Shared Function HttpPost(url As String, vals As Hashtable, cc As CookieContainer) As String
        Dim writer As System.IO.StreamWriter
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

        Dim param As String = ""
        For Each k As String In vals.Keys
            param += String.Format("{0}={1}&", k, vals(k))
        Next
        Dim data As Byte() = Encoding.ASCII.GetBytes(param)

        ' リクエストの作成
        Dim req As HttpWebRequest _
          = CType(WebRequest.Create(url), HttpWebRequest)
        req.Method = "POST"
        req.ContentType = "application/x-www-form-urlencoded"
        req.ContentLength = data.Length
        req.CookieContainer = cc

        ' ポスト・データの書き込み
        Dim reqStream As Stream = req.GetRequestStream()
        reqStream.Write(data, 0, data.Length)
        reqStream.Close()

        Dim res As WebResponse = req.GetResponse()

        ' レスポンスの読み取り
        Dim resStream As Stream = res.GetResponseStream()
        Dim sr As StreamReader = New StreamReader(resStream, encoder)
        Dim result As String = sr.ReadToEnd()

        writer = New System.IO.StreamWriter("bbb.htm", False, enc)
        writer.Write(result)
        writer.Close()



        sr.Close()
        resStream.Close()

        Return result
    End Function

    Shared Function Login(id As String, password As String, cc As CookieContainer) As String
        ' ログイン・ページへのアクセス
        Dim vals As Hashtable = New Hashtable()
        vals("login") = id
        vals("passwd") = password
        vals(".chkP") = "Y"
        vals(".src") = "auc"
        vals(".albatross") = "-100"
        vals(".tries") = "1"
        vals(".intl") = "jp"
        vals("hasMsgr") = "0"
        vals(".intl") = "jp"
        vals(".nojs") = "1"
        vals("done") = "http://openwatchlist.auctions.yahoo.co.jp/jp/show/mystatus?select=watchlist"

        Dim loginUrl As String = "https://login.yahoo.co.jp/config/login"
        Return HttpPost(loginUrl, vals, cc)
    End Function

    Shared Function ReadLog(cc As CookieContainer) As String
        ' 足あとページへのアクセス
        Dim showlog As String = "http://openwatchlist.auctions.yahoo.co.jp/jp/show/mystatus?select=watchlist"
        Return HttpGet(showlog, cc)
    End Function
