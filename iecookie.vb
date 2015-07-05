Imports System.Text
Imports System.IO
Imports System.Net


Public Class Form1
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



    Private Sub test_WebRequest(geturl As String)
        Dim writer As StreamWriter
        'HttpWebRequestの作成
        '        Dim webreq As System.Net.HttpWebRequest = _
        'CType(System.Net.WebRequest.Create("http://openwatchlist.auctions.yahoo.co.jp/jp/show/mystatus?select=watchlist"),  _
        ' (   System.Net.HttpWebRequest)
        'または、


        Dim webreq As System.Net.WebRequest = _
            System.Net.WebRequest.Create(geturl)

        'サーバーからの応答を受信するためのHttpWebResponseを取得
        Dim webres As System.Net.HttpWebResponse = _
            CType(webreq.GetResponse(), System.Net.HttpWebResponse)
        'または、
        'Dim webres As System.Net.WebResponse = webreq.GetResponse()

        '文字コード(EUC)を指定する
        Dim enc As System.Text.Encoding = _
            System.Text.Encoding.GetEncoding("utf-8")
        '応答データを受信するためのStreamを取得
        Dim st As System.IO.Stream = webres.GetResponseStream()
        Dim sr As New System.IO.StreamReader(st, enc)
        '受信して表示
        Dim html As String = sr.ReadToEnd()

        writer = New StreamWriter("data", False, Encoding.Default)
        writer.Write(html)
        writer.Close()

        Console.WriteLine(html)

        '閉じる
        'webres.Close()でもよい
        sr.Close()


    End Sub


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call access()
    End Sub
End Class
