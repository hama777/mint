Public Class frm_main
    Public exemin As Integer
    Private bCloseFlag As Boolean

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim min, hh, ret, starttime, endtime, interval, hhmm As Integer
        Dim nowTime As String


        nowTime = Format(Now, "HH:mm")

        min = Minute(Now)
        hh = Hour(Now)
        hhmm = hh * 100 + min

        starttime = 1120    ' 11時20分
        endtime = 1220      ' 
        interval = 5        ' 間隔

  
        If hhmm >= starttime And hhmm <= endtime And min <> exemin Then
            If min Mod interval = 0 Then
                Call execute(ret)
                exemin = min         ' 実行した分 同じ分内に2度実行されないように
                If ret = 0 Then
                    lbl_state.Text = nowTime & " OK"
                Else
                    lbl_state.Text = nowTime & " NG"
                End If
            End If
        End If
        If hh = 13 Then Me.Close() ' 13時にクローズ
    End Sub

    Private Sub execute(ByVal ret As Integer)
        Dim proc As Process

        '起動
        Try
            proc = Process.Start("D:\ols\discas\discas.exe", "/m")
        Catch
            MsgBox("起動できません")
            ret = 1
            Exit Sub
        End Try
        ret = 0

        '終了するまで待つ
        '        Do While True
        '            Application.DoEvents()           '画面乱れ防止のため
        '            If proc.HasExited = True Then Exit Do
        '        Loop

    End Sub

    Private Sub frm_main_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Timer1.Enabled = True
        exemin = -1
        'Me.Text = "dissch 1.10"
    End Sub

    Private Sub open_frm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles open_frm.Click
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.ShowInTaskbar = True
    End Sub

    Private Sub close_frm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles close_frm.Click
        bCloseFlag = True
        Me.Close()

    End Sub
End Class
