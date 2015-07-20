    Private Sub webAccess(url As String)
        Dim retry As Integer

        flg_access_complete = 0
        retry = 0
        web.Navigate(New Uri(url))
        While True
            Application.DoEvents()
            System.Threading.Thread.Sleep(10)

            retry = retry + 1
            If retry > 12000 Then
                'MsgBox("time out")
                Exit While ' 120秒待つ
            End If
        End While
    End Sub

    Private Sub web_DocumentCompleted(sender As Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles web.DocumentCompleted
        Dim i, len As Integer
        Dim buf As String

        ' これがないと他のurlを取ってきてしまう場合もある
        If Not (sender.url.ToString = e.Url.ToString) Then  
            Exit Sub
        End If

        len = web.DocumentStream.Length
        Dim fb(len) As Byte

        web.DocumentStream.Read(fb, 0, len)
        buf = System.Text.Encoding.GetEncoding(932).GetString(fb)   ' SJIS

        flg_access_complete = flg_access_complete + 1 ' 読み込み完了
        writer = New StreamWriter(webtmp, False, System.Text.Encoding.GetEncoding("shift_jis"))
        writer.WriteLine(buf)
        writer.Close()
        ReDim fb(0)
    End Sub

