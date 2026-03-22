Imports System.Reflection.Emit

Public Class About
    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = Application.ProductName
        Label2.Text = Application.ProductVersion
        Label3.Text = "Sterlingstar Programs"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class