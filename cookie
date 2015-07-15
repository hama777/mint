' クッキーを使ってアクセスする
' 参考 http://dobon.net/vb/dotnet/internet/usecookie.html

    Private Declare Function InternetSetCookie Lib "wininet.dll" _
      Alias "InternetSetCookieA" (ByVal lpszUrlName As String, _
      ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Boolean

    Private Declare Function InternetGetCookie Lib "wininet.dll" _
      Alias "InternetGetCookieA" (ByVal lpszUrlName As String, _
      ByVal lpszCookieName As String, ByVal lpszCookieData As StringBuilder, _
     ByRef lpdwSize As Long) As Boolean


    Private Shared cContainer As New System.Net.CookieContainer


    Private Sub access()
        Dim writer As System.IO.StreamWriter
        Dim url As String
        Dim res As String
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
        url = "http://openwatchlist.auctions.yahoo.co.jp/jp/show/mystatus?select=watchlist"
        res = GetHtml(url)
        writer = New System.IO.StreamWriter("aaa.htm", False, enc)
        writer.Write(res)
        writer.Close()

        res = GetHtml(url)
        writer = New System.IO.StreamWriter("bbb.htm", False, enc)
        writer.Write(res)
        writer.Close()
        MsgBox("end")

    End Sub

    Private Shared Function GetHtml(ByVal url As String) As String
        '文字コードを指定する
        Dim enc As System.Text.Encoding = _
            System.Text.Encoding.GetEncoding("utf-8")

        'WebRequestの作成
        Dim webreq As System.Net.HttpWebRequest = _
            CType(System.Net.WebRequest.Create(url),  _
            System.Net.HttpWebRequest)

        'CookieContainerプロパティを設定する
        webreq.CookieContainer = New System.Net.CookieContainer
        '要求元のURIに関連したCookieを追加し、要求に使用する
        webreq.CookieContainer.Add( _
            cContainer.GetCookies(webreq.RequestUri))

        'サーバーからの応答を受信するためのWebResponseを取得
        Dim webres As System.Net.HttpWebResponse = _
            CType(webreq.GetResponse(), System.Net.HttpWebResponse)

        '受信したCookieのコレクションを取得する
        Dim cookies As System.Net.CookieCollection = _
            webreq.CookieContainer.GetCookies(webreq.RequestUri)
        'Cookie名と値を列挙する
        Dim cook As System.Net.Cookie
        For Each cook In cookies
            Console.WriteLine("{0}={1}", cook.Name, cook.Value)
        Next cook
        '取得したCookieを保存しておく
        cContainer.Add(cookies)

        '応答データを受信するためのStreamを取得
        Dim st As System.IO.Stream = webres.GetResponseStream()
        Dim sr As New System.IO.StreamReader(st, enc)
        '受信して表示
        Dim html As String = sr.ReadToEnd()
        '閉じる
        sr.Close()

        Return html
    End Function
