Imports System.IO
Imports System.Text

Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Call main_proc()

    End Sub
    Private Sub main_proc()

        Dim s As String
        Dim c As Char
        Dim i As Integer
        s = TextBox1.Text
        For i = 0 To s.Length - 1
            c = s.Substring(i, 1)  ' index は 0始まり
        Next

        Debug.WriteLine(s)


    End Sub
End Class
