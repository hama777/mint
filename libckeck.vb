Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1
    Public Const VERSION As String = "1.42"
    Public Const FINDURL = "https://www.lib.city.kobe.jp/opac/opacs/find_books?kanname[all-pub]=1&title="
    Public Const FINDPARM = "&btype=B&searchmode=syosai"
    Public Const LIBTOPURL = "https://www.lib.city.kobe.jp"
    Public Const BOOKURL1 = "https://www.lib.city.kobe.jp/opac/opacs/find_detailbook?type=CtlgBook&pvolid=PV%3A"
    Public Const BOOKURL2 = "&mode=one_line&kobeid=PV%3A"
    Public Const MAGURL = "https://www.lib.city.kobe.jp/opac/opacs/magazine_detail?type=CtlgMagazine&kobeid=CT%3A"
    Public Const RENTALLISTURL = "https://www.lib.city.kobe.jp/opac/opacs/lending_display"


    Public Const BUFSIZE As Integer = 4096
    Public Const BOROWSER As String = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    '    Public Const BOROWSER As String = "D:\ols\Sleipnir\bin\Sleipnir.exe"
    Public Const EDITOR As String = "d:\ols\hide\hidemaru.exe"
    Public Const BOOKDB As String = "D:\doc\図書.mdb"
    'Public Const BOOKDB As String = "D:\doc\VB2010\libcheck\bin\Debug\図書test.mdb"

    'foundstate の状態
    Public Const CNOTFOUND As Integer = 0
    Public Const CFOUND As Integer = 1
    Public Const CFOUNDDUP As Integer = 2
    Public Const CERROR As Integer = 3

    Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
    Public Const INTERNET_OPEN_TYPE_DIRECT = 1
    Public Const INTERNET_OPEN_TYPE_PROXY = 3

    Public Const INTERNET_FLAG_RELOAD = &H80000000

    Public Declare Function InternetOpen Lib "wininet.dll" _
        Alias "InternetOpenA" (ByVal sAgent As String, _
                               ByVal lAccessType As Integer, _
                               ByVal sProxyName As String, _
                               ByVal sProxyBypass As String, _
                               ByVal lFlags As Integer) As Integer
    Public Declare Function InternetOpenUrl Lib "wininet.dll" _
        Alias "InternetOpenUrlA" (ByVal hInternetSession As Integer, _
                                  ByVal sUrl As String, _
                                  ByVal sHeaders As String, _
                                  ByVal lHeadersLength As Integer, _
                                  ByVal lFlags As Integer, _
                                  ByVal lContext As Integer) As Integer

    '第2引数をバイト型 (Byte)の参照渡しにしておく
    Public Declare Function InternetReadFile Lib "wininet.dll" ( _
        ByVal hFile As Integer, _
        ByRef sBuffer As Byte, _
        ByVal lNumBytesToRead As Integer, _
        ByRef lNumberOfBytesRead As Integer) As Integer
    Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Integer

    Public Structure reserve_t
        Public name As String         '  タイトル
        Public res_date As String     ' 日付
        Public state As Integer       ' 予約中 0 /受取可  1
        Public period As String
        Public rent As Integer        ' 0 在庫   1 貸し出し中 2 受入中  -1 エラー
        Public numbook As Integer     ' 冊数
        Public reserve As Integer     ' 予約数
        Public code As String         ' 本コード
        Public acc_date As String     ' 取り置き日
        Public url As String
        Public res_id As String       ' 予約番号  予約順位の取得で使用
        Public order As Integer       ' 予約順位
    End Structure

    Public Structure wish_t
        Public name As String         '  タイトル
        Public rentstate As Integer        ' 0 在庫   1 貸し出し中  -1 エラー
        Public numbook As Integer     ' 冊数
        Public reserve As Integer     ' 予約数
        Public code As String         ' PVコード
        Public url As String
        Public rank As String          ' ランク
        Public add_date As String      ' 追加日
    End Structure

    Public Structure rental_t
        Public name As String         '  タイトル
        Public limit As String         '  貸出期限
        Public reserve As Boolean      ' 予約
        Public extend As Boolean      ' 延長
    End Structure
    Public Structure mag_t
        Public name As String
        Public code As String
    End Structure

    Public Structure user_t
        Public id As String
        Public pass As String
    End Structure

    Public Structure bookurl_t
        Public name As String
        Public url As String
    End Structure

    Public bookurl As bookurl_t

    Dim logdatetbl As Hashtable = New Hashtable


    Public mag As mag_t
    Public user As user_t
    Public res As reserve_t    ' 予約リスト構造体
    Public wish As wish_t   ' 希望本リスト
    Public rental As rental_t
    Public reslist As New ArrayList
    Public bookurllist As New ArrayList
    Public userlist As New ArrayList
    Public rentallist As New ArrayList


    Public namelist As New ArrayList
    Public maglist As New ArrayList

    Public wishlist As New ArrayList
    Public ineterr As Integer
    Public apppath As String
    Public cnt_rental_ok As Integer
    Public checkstartdate As Integer   ' お気に入りチェック開始年月

    ' ********************************************************************
    '                貸出状況リスト
    ' ********************************************************************

    Private Sub displayRentalList()
        Dim u As user_t
        Call OutputRentalHeader()
        For Each u In userlist
            Call RentalListCommon(u.id, u.pass)
        Next
        Call outputRentalList()
        lbl_state.Text = "終了"
    End Sub

    Private Sub OutputRentalHeader()
        Dim writer As StreamWriter
        Dim reader As StreamReader
        Dim s, logdate As String

        logdate = Format(System.DateTime.Now, "yy/MM/dd HH:mm ")


        reader = New StreamReader("rentalheader.htm", Encoding.GetEncoding("Shift_JIS"))
        writer = New StreamWriter("rental_list.htm", False, Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            writer.WriteLine(s)
        Loop
        writer.WriteLine("&nbsp;&nbsp;&nbsp;<span style=font-size: 11pt >" & logdate & "現在</span><br><br>")
        reader.Close()
        writer.Close()
    End Sub

    Private Sub RentalListCommon(userid As String, pass As String)

        Dim ret As Integer

        lbl_state.Text = "ログイン中"
        Me.Refresh()
        Call login(userid, pass)

        ret = siteAccess(RENTALLISTURL)
        lbl_state.Text = "分析中"
        Me.Refresh()
        Call AnalizeRenatlList()


        System.IO.File.Delete("www.htm")
        'Call getReserveOrder()
        Call OutputRenatalList()

    End Sub

    Private Sub AnalizeRenatlList()
        Dim reader As StreamReader
        Dim s, strreserve As String
        Dim r_td, r_element As Regex
        Dim m As Match
        Dim st As Integer
        Dim r_tag As New Regex("<.*>")

        rentallist.Clear()

        r_td = New Regex("<td>(?<1>.+)</td>")
        r_element = New Regex(">(?<1>.+)<")
        reader = New StreamReader("www.htm", Encoding.GetEncoding("utf-8"))

        st = 0
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            If s.IndexOf("<td") >= 0 Then
                st = st + 1
                Select Case st
                    Case 3
                        m = r_element.Match(s)
                        rental.name = m.Groups(1).Value    ' 書名
              
                    Case 4
                        m = r_element.Match(s)
                        rental.limit = m.Groups(1).Value   ' 期限

                    Case 6
                        rental.reserve = False
                        m = r_element.Match(s)
                        strreserve = m.Groups(1).Value
                        If strreserve.IndexOf("あり") <> -1 Then
                            rental.reserve = True
                        End If

                    Case 9
                        rental.extend = False
                        m = r_element.Match(s)
                        If m.Groups(1).Value = "延長済" Then
                            rental.extend = True
                        End If
                        rentallist.Add(rental)

                End Select

            End If
            If s.IndexOf("</tr>") >= 0 Then
                st = 0
            End If

        Loop
        reader.Close()

    End Sub

    Private Sub OutputRenatalList()
        Dim writer As StreamWriter
        Dim rr As rental_t
        Dim i As Integer
        Dim s, strreserve, strexpand, limit As String
        Dim limit_date As DateTime

        writer = New StreamWriter("rental_list.htm", True, Encoding.GetEncoding("Shift_JIS"))
        writer.WriteLine("<table>")
        writer.WriteLine("<tr><th>No.</th><th width=450>書名</td><th>期限</th><th>予約</th><th width=80>備考</th></tr>")

        i = 0
        For Each rr In rentallist
            i = i + 1
            If i Mod 2 = 0 Then
                writer.WriteLine("<tr bgcolor=#f9f9e2>")
            Else
                writer.WriteLine("<tr bgcolor=#ffffff>")
            End If

            limit = rr.limit
            s = limit.Substring(0, 4) & "/" & limit.Substring(4, 2) & "/" & limit.Substring(6, 2)
            limit_date = DateTime.Parse(s)
            limit = limit_date.ToString("MM/dd (ddd)")



            strreserve = ""
            If rr.reserve = True Then
                strreserve = "あり"

            End If
            strexpand = ""
            If rr.extend = True Then
                strexpand = "延長済み"
            End If

            s = "<td>" & i & "</td><td width=450>" _
                & rr.name & "</a></td><td>" & limit & "</td><td>" & strreserve & "</td><td>" _
                & strexpand & "</td></tr>"
            writer.WriteLine(s)

        Next
        writer.WriteLine("</table><br>")
        writer.Close()
    End Sub

    Private Sub outputRentalList()
        Dim writer As StreamWriter

        writer = New StreamWriter("rental_list.htm", True, Encoding.GetEncoding("Shift_JIS"))
        writer.WriteLine("</center></body></html>")
        writer.Close()

        Process.Start(BOROWSER, apppath & "rental_list.htm")


    End Sub

    Private Sub displayRentalListLog()
        Process.Start(BOROWSER, apppath & "rental_list.htm")
    End Sub


    ' ********************************************************************
    '                お気に入りリストのチェック  蔵書一覧
    ' ********************************************************************
    '  DB の希望リストについて図書館にある本の冊数、予約数を検索し、表示する

    Private Sub checkWishlist()
        'Call createBookList()     ' テスト用
        checkstartdate = 0
        If ck_term.Checked = True Then
            Call setCheckStartDate()
        End If

        Call accessBookDB(1)
        Call checkBookStatus()
        If checkstartdate = 0 Then  ' 短縮検索ではbookurlを更新しない
            Call updateBookUrlList()
        End If

        Call outputWishlist()
        lbl_state.Text = "終了"

    End Sub

    ' DBのいつから検索するかを決定する
    Private Sub setCheckStartDate()
        Dim sdate As Date
        sdate = System.DateTime.Now.AddDays(-180)    '  半年前
        'sdate = System.DateTime.Now.AddDays(-30)    '  テスト用
        checkstartdate = (sdate.Year - 2000) * 100 + sdate.Month

    End Sub


    ' 貸出状況、冊数、予約数を取得する

    ' 基本は書名で検索し、見つかったURLから詳細をアクセスする
    ' PVコードがあるものはPVコードから詳細をアクセスする
    ' 一度、詳細をアクセスした本はそのURLをファイルに保存し、次回からはそのURLを
    ' 使用することによって検索を省く

    Private Sub checkBookStatus()
        Dim ww As wish_t
        Dim detailurl As String
        Dim st, n, r, cnt, i, foundstate As Integer
        Dim tmpurl As String

        cnt_rental_ok = 0
        cnt = wishlist.Count
        For i = 0 To cnt - 1
            System.Windows.Forms.Application.DoEvents()
            lbl_state.Text = (i + 1) & "/" & cnt
            Me.Refresh()

            ww = wishlist(i)
            If ww.code <> "" Then   ' PVコード指定があればそれを優先
                detailurl = pvcodeToUrl(ww.code)   ' PVコードからURL作成
                foundstate = CFOUND

            Else
                detailurl = searchBookUrl(ww.name)    ' 以前のURLが存在するか検索
                If detailurl = "" Then                ' なければ書名から検索して本のURLを得る
                    Call searchByBookName(ww.name)
                    If ineterr = 0 Then
                        Call analizeFirst(foundstate, detailurl) '第1レベル  HTMLの分析
                        detailurl = LIBTOPURL & detailurl
                        tmpurl = detailurl
                    Else
                        foundstate = CERROR
                    End If
                Else
                    Call siteAccess(detailurl)
                    foundstate = CFOUND
                End If
            End If
            If foundstate = CFOUND Then
                ww.url = detailurl
                Call analizeSecond(detailurl, st, n, r, 0)  ' 第2レベル  HTMLの分析
                If st = 0 Then
                    ww.rentstate = 0    '  在庫
                    cnt_rental_ok = cnt_rental_ok + 1
                Else
                    ww.rentstate = 1    '  貸出中
                End If
                ww.numbook = n
                ww.reserve = r
            Else
                If foundstate = CNOTFOUND Then
                    ww.rentstate = -1
                End If
                If foundstate = CFOUNDDUP Then
                    ww.rentstate = -2
                End If
                If foundstate = CERROR Then
                    ww.rentstate = -3
                End If
                ww.numbook = -1
                ww.reserve = -1
            End If
            wishlist(i) = ww

        Next

    End Sub

    ' boolurl.txt ファイルを更新する
    Private Sub updateBookUrlList()
        Dim writer As StreamWriter
        Dim ww As wish_t

        writer = New StreamWriter("bookurl.txt", False, Encoding.GetEncoding("Shift_JIS"))

        For Each ww In wishlist
            If ww.rentstate >= 0 Then
                writer.WriteLine(ww.name & vbTab & ww.url)
            End If
        Next
        writer.Close()

    End Sub

    Private Sub outputWishlist()
        Dim ww As wish_t
        Dim writer As StreamWriter
        Dim reader As StreamReader
        Dim i As Integer
        Dim s, st, s_num, s_resv As String
        Dim outfile, logdate As String

        outfile = "wishlist.htm"
        logdate = Format(System.DateTime.Now, "yy/MM/dd HH:mm ")
        If checkstartdate <> 0 Then    ' 期間短縮版
            outfile = "wishlist2.htm"
            logdatetbl("wishshort") = logdate

        Else
            logdatetbl("wishlong") = logdate
        End If

        writer = New StreamWriter(outfile, False, Encoding.GetEncoding("Shift_JIS"))
        reader = New StreamReader("wishheader.htm", Encoding.GetEncoding("Shift_JIS"))

        Do Until reader.EndOfStream
            s = reader.ReadLine()
            writer.WriteLine(s)
        Loop
        reader.Close()

        writer.WriteLine("<center><font color=blue>登録数</font> " & wishlist.Count)
        writer.WriteLine("&nbsp;&nbsp;<font color=blue>在庫</font> " & cnt_rental_ok)
        writer.WriteLine("&nbsp;&nbsp;<font color=blue>OK率</font> " & Format(cnt_rental_ok / wishlist.Count, " ###.# %"))
        writer.WriteLine("&nbsp;&nbsp;&nbsp;" & logdate & "現在</center><br>")
        writer.WriteLine("<center>")
        writer.WriteLine("<table>")
        writer.WriteLine("<tr><th width=450>書名</th><th>状況</th><th>冊数</th><th>予約</th><th>ランク</th><th>登録</th></tr>")

        i = 0
        For Each ww In wishlist
            i = i + 1
            If i Mod 2 = 0 Then
                writer.WriteLine("<tr bgcolor=#f9f9e2>")
            Else
                writer.WriteLine("<tr bgcolor=#ffffff>")
            End If
            s = "<a href=""" & ww.url & """ target=""_blank"">" & ww.name & "</a>"

            s_num = ww.numbook
            s_resv = ww.reserve
            Select Case ww.rentstate
                Case 0
                    st = "<font color=blue>在庫</font>"
                Case 1
                    st = "<font color=red>貸出中</font>"
                Case 2
                    st = "<font color=red>受入中</font>"
                Case -1
                    st = "NF"
                    s_num = ""
                    s_resv = ""
                Case -2
                    st = "DUP"
                    s_num = ""
                    s_resv = ""
                Case -3
                    st = "ERR"
                    s_num = ""
                    s_resv = ""
            End Select

            writer.WriteLine("<td width=450>" & s & "</td><td>" & st & "</td><td align=right>" _
                & s_num & "</td><td align=right>" & s_resv & "</td><td>" _
                & ww.rank & "</td><td>" & ww.add_date & "</td></tr>")
        Next
        writer.WriteLine("</table></center></body></html>")
        writer.Close()
        Process.Start(BOROWSER, apppath & outfile)
        Call displayLogDate()
        Call writeLogDateList()
    End Sub

    ' 結果の表示
    Private Sub displayWishLog()
        If ck_term.Checked = True Then
            Process.Start(BOROWSER, apppath & "wishlist2.htm")   ' 短縮モード
        Else
            Process.Start(BOROWSER, apppath & "wishlist.htm")
        End If

    End Sub

    ' 書名に対応する本のurl を返す
    Private Function searchBookUrl(bookname As String)
        Dim b As bookurl_t
        For Each b In bookurllist
            If b.name = bookname Then
                searchBookUrl = b.url
                Exit Function
            End If
        Next
        searchBookUrl = ""
    End Function
    Private Function pvcodeToUrl(c As String)
        pvcodeToUrl = "https://www.lib.city.kobe.jp/opac/opacs/find_detailbook?kobeid=PV%3A" & c & _
            "&amp;type=CtlgBook&amp;pvolid=PV%3A" & c & _
            "&amp;mode=one_line"

    End Function

    ' ********************************************************************
    '                予約リスト
    ' ********************************************************************


    ' 予約リスト
    '  ReserveList
    '    OutputHeader  ヘッダー出力
    '    ReserveListCommon
    '      OutputReserveList   予約一覧


    Private Sub ReserveList()
        Dim u As user_t
        Call OutputHeader()
        For Each u In userlist
            Call ReserveListCommon(u.id, u.pass)
        Next
        Call DisplayReserveList()
        lbl_state.Text = "終了"

    End Sub
    Private Sub ReserveListCommon(userid As String, pass As String)
        Dim ckurl As String
        Dim ret As Integer

        lbl_state.Text = "ログイン中"
        Me.Refresh()
        Call login(userid, pass)
        ckurl = "https://www.lib.city.kobe.jp/opac/opacs/reservation_display"
        ret = siteAccess(ckurl)
        lbl_state.Text = "分析中"
        Me.Refresh()
        Call AnalizeReserveList()
        Call SearchReserveCount()

        System.IO.File.Delete("www.htm")
        Call getReserveOrder()
        Call OutputReserveList()

    End Sub
    ' 予約リストのhtml 解析
    Private Sub AnalizeReserveList()
        Dim reader As StreamReader
        Dim s, t As String
        Dim r_td, r_element As Regex
        Dim m As Match
        Dim st As Integer
        Dim r_tag As New Regex("<.*>")

        reslist.Clear()

        r_td = New Regex("<td>(?<1>.+)</td>")
        r_element = New Regex(">(?<1>.+)<")
        reader = New StreamReader("www.htm", Encoding.GetEncoding("utf-8"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            If s.IndexOf("reservation_cancel_confirmation") <> -1 Then Exit Do
        Loop

        st = 0
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            If s.IndexOf("<td") >= 0 Then
                st = st + 1
                Select Case st
                    Case 3
                        m = r_element.Match(s)
                        res.res_date = m.Groups(1).Value    ' 予約日
                        t = reader.ReadLine()   ' 予約番号
                        res.res_id = r_tag.Replace(t, "").Trim
                        s = reader.ReadLine()   ' <br>
                        s = reader.ReadLine()   ' PVコード
                        s = s.Replace("PV:", "")
                        res.code = s.Trim

                    Case 4
                        m = r_element.Match(s)
                        res.name = m.Groups(1).Value   ' 書名
                    Case 5
                        res.state = 0
                        If s.IndexOf("受取可") <> -1 Then
                            res.state = 1
                        End If
                    Case 6
                        s = reader.ReadLine()   ' 取り置き日
                        If s.IndexOf("未定") >= 0 Then
                            res.acc_date = ""
                        Else
                            s = s.Replace("〜", "")
                            s = s.Trim
                            res.acc_date = s

                        End If
                        reslist.Add(res)

                End Select

            End If
            If s.IndexOf("</tr>") >= 0 Then
                st = 0
            End If

        Loop
        reader.Close()

    End Sub

    ' 予約リストの詳細(冊数、予約数)を調べる
    Private Sub SearchReserveCount()
        Dim rr As reserve_t
        Dim burl, mcode As String
        Dim st, n, r, cnt, i, mag As Integer

        cnt = reslist.Count

        For i = 0 To cnt - 1
            mag = 0     ' 雑誌の時、 1
            rr = reslist(i)
            If rr.name.IndexOf("月号") <> -1 Or rr.name.IndexOf("月新年号") <> -1 Or rr.name.IndexOf("月新春号") <> -1 Then
                mcode = searchMagCode(rr.name)
                If mcode = "" Then
                    'MsgBox("雑誌コード未登録 " & rr.name)
                End If
                burl = MAGURL & mcode & "&pvolid=PV%3A" & rr.code
                mag = 1
            Else
                burl = BOOKURL1 & rr.code & BOOKURL2 & rr.code
            End If
            Call analizeSecond(burl, st, n, r, mag)
            rr.url = burl
            rr.rent = st
            rr.numbook = n
            rr.reserve = r



            reslist(i) = rr
            lbl_state.Text = "検索中 " & i + 1 & " / " & cnt
            Me.Refresh()
        Next

    End Sub

    ' 予約順位を取得する     2015/08/19
    Private Sub getReserveOrder()
        Dim rr As reserve_t
        Dim burl As String
        Dim cnt, i, checkok, ret As Integer
        Dim urltop As String = "https://www.lib.city.kobe.jp/opac/opacs/reservation_cancel_confirmation?reservation_order_confirmation=%e9%a0%86%e4%bd%8d%e7%a2%ba%e8%aa%8d"


        lbl_state.Text = "予約順位取得中 "
        Me.Refresh()

        cnt = reslist.Count
        burl = urltop
        For i = 0 To cnt - 1
            rr = reslist(i)
            burl = burl & "&cancel[" & rr.res_id & "]=1"
        Next

        ret = siteAccess(burl)
        If ret = -1 Then
            MsgBox("予約順位アクセスエラー ")
            Exit Sub
        End If
        System.Threading.Thread.Sleep(10000)
        checkok = 0
        For i = 1 To 20
            If System.IO.File.Exists("www.htm") Then
                checkok = 1
                Exit For
            End If

            System.Threading.Thread.Sleep(10000)
        Next
        If checkok = 1 Then
            Call analizeOrderHtml()
            lbl_state.Text = "予約順位完了 "
        Else
            MsgBox("予約順位取得エラー ")
        End If
        Me.Refresh()

    End Sub

    Private Sub analizeOrderHtml()

        ' td は 連番、予約番号、書名、順位となっているので 4行目の順位だけ取り出す

        Dim reader As StreamReader
        Dim rr As reserve_t
        Dim s, item As String
        Dim r, r_url As Regex
        Dim m As Match
        Dim line, cnt As Integer


        r = New Regex("<td>(.*?)</td>")

        line = 0
        cnt = 0
        reader = New StreamReader("www.htm", Encoding.GetEncoding("utf-8"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            m = r.Match(s)
            If m.Success Then
                line += 1
                If line = 4 Then
                    rr = reslist(cnt)
                    item = m.Groups(1).Value
                    If item = "受取可" Then
                        rr.order = 0
                    ElseIf item = "-" Then    '  詳細不明  確認必要
                        rr.order = -1
                    Else
                        rr.order = item
                    End If
                    reslist(cnt) = rr
                    cnt += 1
                    line = 0
                End If

            End If

        Loop
        reader.Close()

    End Sub


    ' 書名から雑誌コードを検索する
    Private Function searchMagCode(name As String)
        Dim mm As mag_t

        For Each mm In maglist
            If name.IndexOf(mm.name) <> -1 Then
                searchMagCode = mm.code
                Exit Function
            End If
        Next
        searchMagCode = ""

    End Function

    Private Sub OutputHeader()
        Dim writer As StreamWriter
        Dim reader As StreamReader
        Dim s, logdate As String

        logdate = Format(System.DateTime.Now, "yy/MM/dd HH:mm ")
        logdatetbl("resv") = logdate


        reader = New StreamReader("header.htm", Encoding.GetEncoding("Shift_JIS"))
        writer = New StreamWriter("resv_list.htm", False, Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            writer.WriteLine(s)
        Loop
        writer.WriteLine("&nbsp;&nbsp;&nbsp;<span style=font-size: 11pt >" & logdate & "現在</span><br><br>")
        reader.Close()
        writer.Close()


    End Sub


    ' 予約リストの結果をhtmlに出力する
    Private Sub OutputReserveList()
        Dim writer As StreamWriter
        Dim rr As reserve_t
        Dim i As Integer
        Dim bname, state, rt, s, strorder As String

        writer = New StreamWriter("resv_list.htm", True, Encoding.GetEncoding("Shift_JIS"))
        'writer.WriteLine("<html><head><title>予約一覧</title></head><body>")
        'writer.WriteLine("<style type=""text/css""><!-- A { text-decoration: none; } -->")
        'writer.WriteLine("</style></head><body link=#000000 vlink=#000000>")
        'writer.WriteLine("<center><b>予約一覧</b>")
        'writer.WriteLine("&nbsp;&nbsp;&nbsp;" & logdate & "現在<br><br>")
        writer.WriteLine("<table>")
        'writer.WriteLine("<table cellspacing=1 border=0 cellpadding=3><tr bgcolor=#fbf703 >")
        writer.WriteLine("<tr><th>No.</th><th width=450>書名</td><th>予約日</th><th>状況</th><th>貸出</th><th>冊数</th><th>予約数</th><th>順位</th><th>取置日</th></tr>")

        i = 0
        For Each rr In reslist
            i = i + 1
            If i Mod 2 = 0 Then
                writer.WriteLine("<tr bgcolor=#f9f9e2>")
            Else
                writer.WriteLine("<tr bgcolor=#ffffff>")
            End If
            bname = rr.name
            bname = bname.Replace("/", " ")
            state = "予約中"
            If rr.state = 1 Then
                state = "<b><font color=blue>受取可</font></b>"
            End If
            rt = rentstateToString(rr.rent)

            If rr.order = -1 Then
                strorder = "-"
            Else
                strorder = rr.order
            End If

            s = "<td>" & i & "</td><td width=430><a href=""" & rr.url & """ target=""_blank"">" _
                & bname & "</a></td><td>" & rr.res_date & "</td><td>" & state & "</td><td>" _
                & rt & "</td><td align=right>" & rr.numbook & "</td><td align=right>" _
                & rr.reserve & "</td><td align=right>" & strorder & "</td><td align=right>" & rr.acc_date & "</td></tr>"
            writer.WriteLine(s)

        Next
        writer.WriteLine("</table><br>")
        writer.Close()
        'Process.Start(BOROWSER, apppath & "resv_list.htm")
        'Call displayLogDate()
        'Call writeLogDateList()
    End Sub

    Private Sub DisplayReserveList()
        Dim writer As StreamWriter

        writer = New StreamWriter("resv_list.htm", True, Encoding.GetEncoding("Shift_JIS"))
        writer.WriteLine("</center></body></html>")
        writer.Close()

        Process.Start(BOROWSER, apppath & "resv_list.htm")
        Call displayLogDate()
        Call writeLogDateList()
    End Sub


    Private Function rentstateToString(rst As Integer)
        Select Case rst
            Case 0
                rentstateToString = "在庫"
            Case 1
                rentstateToString = "貸出中"
            Case 2
                rentstateToString = "受入中"
            Case Else
                rentstateToString = "エラー"

        End Select


    End Function

    ' ********************************************************************
    '                蔵書有無チェック
    ' ********************************************************************

    Private Sub checkStok()
        Dim bname As String
        Dim status, cnt, i, hit, dup As Integer
        Dim bookurl, s As String
        Dim writer As StreamWriter
        Dim logdate As String

        writer = New StreamWriter("libreport.txt", False, Encoding.GetEncoding("Shift_JIS"))

        logdate = Format(System.DateTime.Now, "yy/MM/dd HH:mm ")
        logdatetbl("checkstok") = logdate

        ' 蔵書にないものを選択

        Call accessBookDB(0)

        'Call createBookList()
        cnt = namelist.Count
        i = 0
        hit = 0
        dup = 0
        For Each bname In namelist
            System.Windows.Forms.Application.DoEvents()
            i = i + 1
            lbl_state.Text = "検索中 " & i & "/" & cnt
            Me.Refresh()


            Call searchByBookName(bname)

            If ineterr = 1 Then
                status = -1
            Else
                Call analizeFirst(status, bookurl)
            End If
            s = ""
            Select Case status
                Case -1
                    s = "ERROR  " & bname
                Case 1
                    s = "OK     " & bname & vbCrLf & LIBTOPURL & bookurl
                    hit = hit + 1
                Case 2
                    s = "DUP    " & bname
                    dup = dup + 1

            End Select
            If s <> "" Then
                writer.WriteLine(s)
            End If
        Next
        writer.Close()
        Call displayLogDate()
        Call writeLogDateList()
        lbl_state.Text = "終了 hit= " & hit & "  dup= " & dup
    End Sub
    Private Sub displayStokLog()
        Process.Start(EDITOR, apppath & "libreport.txt")
    End Sub
    Private Sub displayResvLog()
        Process.Start(BOROWSER, apppath & "resv_list.htm")

    End Sub
    ' テスト用
    Private Sub createBookList()
        Dim reader As StreamReader
        Dim s, sql As String
        Dim ww As wish_t

        reader = New StreamReader("test.txt", Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            namelist.Add(s)
            ww.name = s
            ww.code = ""
            wishlist.Add(ww)
        Loop
        reader.Close()
    End Sub
    '==============================================================================
    '      汎用処理
    '==============================================================================

    Private Sub accessBookDB(flg As Integer)
        Dim i As Integer
        Dim sql As String
        Dim a_date As Integer
        Dim cn As New System.Data.OleDb.OleDbConnection
        Dim command As System.Data.OleDb.OleDbCommand

        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim dbname As String = BOOKDB
        Dim ww As wish_t

        If flg = 0 Then
            sql = "SELECT * FROM [main] WHERE [図] is null;"
            namelist.Clear()
        Else
            sql = "SELECT * FROM [main] WHERE [図] = 'Z' order by [登録日];"
            wishlist.Clear()

        End If

        cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & dbname & ";"
        command = cn.CreateCommand
        command.CommandText = sql
        'command.CommandText = "SELECT * FROM [main]  WHERE [図] = 'Z' order by [登録日];"
        'command.CommandText = "SELECT * FROM [main] ;"

        cn.Open()
        reader = command.ExecuteReader()
        i = 0
        While reader.Read() = True
            i = i + 1
            Debug.WriteLine(reader(1))

            If flg = 0 Then
                namelist.Add(reader(1))
            Else
                ww.name = reader(1)
                a_date = reader(9)
                If a_date < checkstartdate Then
                    Continue While
                End If

                ww.code = ""
                If IsDBNull(reader(6)) Then
                    ww.rank = ""
                Else
                    ww.rank = reader(6)
                End If

                ww.add_date = reader(9)
                If IsDBNull(reader(12)) Then
                    ww.code = ""
                Else
                    ww.code = reader(12)
                End If

                wishlist.Add(ww)
            End If
        End While

        cn.Close()
        command.Dispose()
        cn.Dispose()


    End Sub

    ' 書名を検索し、結果をwww.htmファイルに出力する
    Private Sub searchByBookName(bname As String)
        Dim encodedString, s As String
        Dim ret As Integer

        encodedString = Uri.EscapeDataString(bname)
        s = FINDURL & encodedString & FINDPARM

        ret = siteAccess(s)
        ineterr = 0
        If ret <> 0 Then
            ineterr = 1
        End If

    End Sub
    ' 第1レベル(検索結果)  HTMLの分析
    Private Sub analizeFirst(ByRef foundstate As Integer, ByRef bookurl As String)
        ' 出力
        ' status  0 ... 見つからない  1  .... 見つかる   2 ... 複数みつかる
        ' burl    本詳細へのurl

        Dim reader As StreamReader
        Dim s, nb As String
        Dim r, r_url As Regex
        Dim m As Match

        Dim first As Integer

        first = 0
        '        r = New Regex("<B>(?<1>[0-9]+)</B>&nbsp;件")
        r = New Regex("結果件数：図書 (?<1>[0-9]+) 件")
        r_url = New Regex("<a href=""(.*?)"">")
        bookurl = ""
        reader = New StreamReader("www.htm", Encoding.GetEncoding("utf-8"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()

            m = r.Match(s)
            If m.Success Then
                nb = m.Groups(1).Value
                Select Case nb
                    Case "0"
                        foundstate = CNOTFOUND
                        Exit Do
                    Case "1"
                        foundstate = CFOUND
                        first = 1
                    Case Else
                        foundstate = CFOUNDDUP
                        Exit Do
                End Select
            End If
            If first = 1 Then
                m = r_url.Match(s)
                If m.Success Then
                    bookurl = m.Groups(1).Value
                    Exit Do
                End If
            End If
        Loop
        reader.Close()

    End Sub
    ' 第2レベル(検索結果)  HTMLの分析
    Private Sub analizeSecond(burl As String, ByRef status As Integer, ByRef n As Integer, ByRef r As Integer, mag As Integer)
        ' 引数
        ' 入力   burl ....  本詳細のurl
        ' 出力   status  ...  0  在庫   1  貸出中  2 受入中
        '        n   ...... 蔵書冊数
        '        r   ...... 予約件数

        Dim reader As StreamReader
        Dim r_nbook, r_nresv, r_nresv2 As Regex
        Dim m As Match
        Dim s As String
        Dim zaiko, ukeire, ret, firstflg As Integer

        ret = siteAccess(burl)
        If ret <> 0 Then
            Exit Sub
        End If

        firstflg = 0
        r = 0   '  予約件数
        n = 0   '  所蔵冊数
        r_nbook = New Regex("所蔵冊数.*;(?<1>.+)冊")
        r_nresv = New Regex("昨日までの予約件数:&nbsp;(?<1>.+)件")
        r_nresv2 = New Regex("昨日までの予約件数(?<1>.+)件")     '  雑誌用
        reader = New StreamReader("www.htm", Encoding.GetEncoding("utf-8"))
        zaiko = ukeire = 0
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            m = r_nbook.Match(s)
            If m.Success Then
                If firstflg = 1 Then Exit Do
                n = m.Groups(1).Value
                firstflg = 1
            End If
            If mag = 0 Then
                m = r_nresv.Match(s)
                If m.Success Then
                    r = m.Groups(1).Value
                    firstflg = 1
                End If
            Else
                m = r_nresv2.Match(s)  '  雑誌用
                If m.Success Then
                    r = m.Groups(1).Value
                    firstflg = 1
                End If

            End If
            If s.IndexOf("<td>看護大</td>") <> -1 Then
                Exit Do
            End If
            If s.IndexOf("<td>外大図書館</td>") <> -1 Then
                Exit Do
            End If
            If s.IndexOf("    &nbsp;") <> -1 Then    ' 在庫あり
                zaiko = 1
            End If
            If s.IndexOf("<td>入荷待ち</td>") <> -1 Then
                ukeire = 1
            End If

        Loop
        reader.Close()
        status = 1            ' すべて貸出中の時のみ  貸出中
        If zaiko = 1 Then     ' 1カ所でも在庫があれば  在庫
            status = 0
        Else
            If ukeire = 1 Then   ' 1カ所でも受け入れ中があれば  受入中
                status = 2
            End If
        End If
    End Sub

    Private Function siteAccess(url As String)

        Dim hOpen As Integer
        Dim hConnect As Integer
        Dim buf(BUFSIZE), catbuf(0) As Byte
        Dim dwSize, totalsize, cnt, len, len2 As Integer
        Dim ret As Boolean
        Dim writer As StreamWriter
        Dim str As String
        Dim enc2 As System.Text.Encoding = System.Text.Encoding.GetEncoding("SJIS")

        totalsize = 0
        cnt = 0

        hOpen = InternetOpen("LibTest", INTERNET_OPEN_TYPE_DIRECT, _
                                 vbNullString, vbNullString, 0)
        'URLを開く
        If hOpen = 0 Then
            siteAccess = 1
            Exit Function
        End If
        hConnect = InternetOpenUrl(hOpen, url, vbNullString, 0, _
                                   INTERNET_FLAG_RELOAD, 0)

        If hConnect = 0 Then
            siteAccess = 1
            Exit Function
        End If

        str = ""
        Do
            ret = InternetReadFile(hConnect, buf(0), BUFSIZE, dwSize)
            If ret = False Then
                siteAccess = -1
                Exit Function
            End If
            If dwSize = 0 Then Exit Do
            cnt = cnt + 1

            ReDim Preserve buf(dwSize - 1)

            If cnt = 1 Then
                catbuf = buf.Clone
                len2 = catbuf.Length
            Else
                len = buf.Length
                len2 = catbuf.Length
                ReDim Preserve catbuf(len + len2 - 1)
                Array.Copy(buf, 0, catbuf, len2, len)
            End If
            ReDim buf(BUFSIZE)

        Loop
        str = System.Text.Encoding.UTF8.GetString(catbuf)

        InternetCloseHandle(hConnect)
        InternetCloseHandle(hOpen)

        writer = New StreamWriter("www.htm", False)
        writer.WriteLine(str)
        writer.Close()
        siteAccess = 0
    End Function

    '   ログイン処理
    Private Sub login(userid As String, pass As String)
        Dim ckurl As String

        ckurl = "https://www.lib.city.kobe.jp/opac/opacs/login?user[login]=" & userid _
            & "&user[passwd]=" & pass _
            & "&act_login=%E3%83%AD%E3%82%B0%E3%82%A4%E3%83%B3&nextAction=mypage_display&prevAction=find_request"

        siteAccess(ckurl)

    End Sub

    '==============================================================================
    '            初期化処理
    '==============================================================================
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Text = "libcheck " & VERSION
        apppath = Application.StartupPath() & "\"
        lbl_state.Text = ""

        Call readMcode()
        Call readBookUrl()
        Call readLogDateList()
        Call displayLogDate()
        Call readUserID()
    End Sub
    ' ID読み込み
    Private Sub readUserID()
        Dim reader As StreamReader
        Dim s As String
        Dim dt() As String

        reader = New StreamReader("user.txt", Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            dt = s.Split(";")
            user.id = dt(0)
            user.pass = dt(1)
            userlist.Add(user)
        Loop
        reader.Close()

    End Sub

    ' 雑誌コード読み込み
    Private Sub readMcode()
        Dim reader As StreamReader
        Dim s As String
        Dim dt() As String

        reader = New StreamReader("mcode.txt", Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            dt = s.Split(";")
            mag.name = dt(0)
            mag.code = dt(1)
            maglist.Add(mag)
        Loop
        reader.Close()

    End Sub

    '  書名 url 対応ファイル読み込み
    Private Sub readBookUrl()
        Dim reader As StreamReader
        Dim s As String
        Dim dt() As String

        reader = New StreamReader("bookurl.txt", Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            dt = s.Split(vbTab)
            bookurl.name = dt(0)
            bookurl.url = dt(1)

            bookurllist.Add(bookurl)
        Loop
        reader.Close()

    End Sub

    '  ログ日付データ書き込み
    Private Sub writeLogDateList()
        Dim writer As StreamWriter
        writer = New StreamWriter("logdate.txt", False, Encoding.GetEncoding("Shift_JIS"))

        For Each de As DictionaryEntry In logdatetbl
            writer.WriteLine(de.Key & vbTab & de.Value)
        Next
        writer.Close()
    End Sub

    '  ログ日付データ読み込み
    Private Sub readLogDateList()
        Dim reader As StreamReader
        Dim s As String
        Dim dt() As String

        reader = New StreamReader("logdate.txt", Encoding.GetEncoding("Shift_JIS"))
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            dt = s.Split(vbTab)
            logdatetbl.Add(dt(0), dt(1))
        Loop
        reader.Close()

    End Sub

    '  ログ日付データ表示
    Private Sub displayLogDate()
        Dim s As String

        s = "お気に入りチェックL   " & vbTab & CType(logdatetbl("wishlong"), String) & vbCrLf
        s = s & "お気に入りチェックS   " & vbTab & CType(logdatetbl("wishshort"), String) & vbCrLf
        s = s & "予約一覧              " & vbTab & CType(logdatetbl("resv"), String) & vbCrLf
        s = s & "蔵書チェック          " & vbTab & CType(logdatetbl("checkstok"), String)
        lbl_logdate.Text = s
    End Sub

    '==============================================================================
    '  ボタン
    '==============================================================================

    '  蔵書一覧   希望本のチェック
    Private Sub cmd_wishlist_Click(sender As System.Object, e As System.EventArgs) Handles cmd_wishlist.Click
        Call checkWishlist()
    End Sub
    ' 予約一覧
    Private Sub cmd_mylist_Click(sender As System.Object, e As System.EventArgs) Handles cmd_mylist.Click
        Call ReserveList()
    End Sub
    ' 蔵書チェック
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles cmd_stok.Click
        Call checkStok()
    End Sub

    Private Sub cmd_test_Click(sender As System.Object, e As System.EventArgs)
        Call accessBookDB(1)
    End Sub
    ' 蔵書チェックログ表示
    Private Sub cmd_stok_log_Click(sender As System.Object, e As System.EventArgs) Handles cmd_stok_log.Click
        Call displayStokLog()
    End Sub
    ' 予約一覧ログ表示
    Private Sub cmd_resv_log_Click(sender As System.Object, e As System.EventArgs) Handles cmd_resv_log.Click
        Call displayResvLog()

    End Sub

    Private Sub cmd_display_stok_Click(sender As System.Object, e As System.EventArgs) Handles cmd_display_wish.Click
        Call displayWishLog()
    End Sub
    ' 貸出状況チェック
    Private Sub cmd_rental_Click(sender As System.Object, e As System.EventArgs) Handles cmd_rental.Click
        Call displayRentalList()
    End Sub
    ' 貸出状況ログ表示
    Private Sub cmd_rental_log_Click(sender As System.Object, e As System.EventArgs) Handles cmd_rental_log.Click
        Call displayRentalListLog()
    End Sub
End Class