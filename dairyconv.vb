Imports System.IO


Public Class Form1

    '  作成  2015/01/12
    '  変換完了  2015/03/15
    '  祝日の色には対応していない

    Const version As String = "DiaryConv Ver 1.00"
    Const outputdir As String = "D:\misc\Dropbox\doc\diary\"
    Const inputdir As String = "d:\ols\diary\"
    Public writer As StreamWriter
    Public reader As StreamReader
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call main_proc()

    End Sub
    Private Sub main_proc()
        Dim startyy, startmm As Integer
        Dim endyy, endmm As Integer
        Dim yy, mm As Integer

        startyy = 2000
        startmm = 1
        endyy = 2016
        endmm = 10

        yy = startyy
        mm = startmm
        Do While True
            Call monthly(yy, mm)
            If yy = endyy And mm = endmm Then Exit Do
            mm = mm + 1
            If mm > 12 Then
                mm = 1
                yy = yy + 1
            End If
        Loop

    End Sub
    Private Sub monthly(yy As Integer, mm As Integer)
        Dim inputfile As String
        Dim outputfile As String
        Dim yydir As String
        Dim body As String
        Dim curdate As Date
        Dim week As String
        Dim dt() As String
        Dim s As String
        Dim i As Integer
        Dim holicolor As String

        yydir = outputdir & yy & "\"
        If System.IO.Directory.Exists(yydir) = False Then
            System.IO.Directory.CreateDirectory(yydir)
        End If

        outputfile = yydir & yy & mm.ToString("0#") & ".htm"
        writer = New StreamWriter(outputfile, False, System.Text.Encoding.Default)
        Call output_header()

        inputfile = inputdir & "NIK" & yy & mm.ToString("0#") & ".txt"

        Try
            reader = New StreamReader(inputfile, System.Text.Encoding.Default)
        Catch E As Exception
            Exit Sub
        End Try

        writer.WriteLine("<title>" & yy & "年" & mm & "月</title>")
        writer.WriteLine("</head><body>")
        writer.WriteLine("<table border=0 bgcolor=""#0000ff"" cellspacing=0 cellpadding=0><tr><td>")
        writer.WriteLine("<table cellspacing=1 border=0 cellpadding=1>")
        i = 0
        Do Until reader.EndOfStream
            s = reader.ReadLine()

            dt = Split(s, vbTab)
            body = dt(1)
            i = i + 1
            Try
                week = DateTime.Parse(yy & "/" & mm & "/" & i).ToString("ddd")
            Catch e As Exception
                Exit Do
            End Try

            If week = "日" Or week = "土" Then
                holicolor = "bgcolor=#ffd8dd"
            Else
                holicolor = ""
            End If
            curdate.ToString("ddd")
            writer.WriteLine("<tr bgcolor=""#ffffff""><td " & holicolor & " >" & i & "</td><td " & holicolor & ">" & week & "</td><td>" & body & "</td></tr>")
        Loop
        reader.Close()
        writer.WriteLine("</table></td></tr></table></html>")
        writer.Close()
    End Sub
    Private Sub output_header()
        Dim headerfile As String
        Dim s As String

        headerfile = "header.htm"
        reader = New StreamReader(headerfile, System.Text.Encoding.Default)
        Do Until reader.EndOfStream
            s = reader.ReadLine()
            writer.WriteLine(s)
        Loop
        reader.Close()

    End Sub
End Class
