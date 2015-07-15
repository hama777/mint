    '  IEのクッキーを使用してアクセス
    ' 参考 http://d.hatena.ne.jp/Kazzz/20061020/p2

    Private Declare Function InternetSetCookie Lib "wininet.dll" _
      Alias "InternetSetCookieA" (ByVal lpszUrlName As String, _
      ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Boolean

    Private Declare Function InternetGetCookie Lib "wininet.dll" _
      Alias "InternetGetCookieA" (ByVal lpszUrlName As String, _
      ByVal lpszCookieName As String, ByVal lpszCookieData As StringBuilder, _
     ByRef lpdwSize As Long) As Boolean
    Shared encoder As Encoding = Encoding.GetEncoding("utf-8")

    Private Sub access()
        Dim writer As System.IO.StreamWriter
        Dim s, ret As String
        Dim data As String
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

        s = "http://openwatchlist.auctions.yahoo.co.jp/jp/show/mystatus?select=watchlist"
        ret = getCookie(s)
        MsgBox(ret)

        data = HttpGet(s, ret)

        writer = New System.IO.StreamWriter("aaa.htm", False, enc)
        writer.Write(data)
        writer.Close()

        MsgBox("end")

    End Sub

    Private Function getCookie(ByVal szUrlName As String) As String
        Dim sCookieVal As New StringBuilder(2048)
        Dim lpLength As Long
        Dim bRet As Boolean
        lpLength = sCookieVal.Capacity
        bRet = InternetGetCookie(szUrlName, _
         vbNull, sCookieVal, lpLength)
        If bRet = True Then Return sCookieVal.ToString
        Return ""
    End Function

    Private Function setCookie(ByVal szUrlName As String, ByVal szCookieName As String, _
       ByVal szCookieData As String) As Boolean
        Return InternetSetCookie(szUrlName, _
         szCookieName, szCookieData)
    End Function

    Shared Function HttpGet(url As String, ccstr As String) As String

        Dim cc As CookieContainer = New CookieContainer()
        Dim conurl As Uri
        ' リクエストの作成
        Dim req As HttpWebRequest _
          = CType(WebRequest.Create(url), HttpWebRequest)
        conurl = New Uri(url)
        cc.SetCookies(conurl, ccstr)

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
