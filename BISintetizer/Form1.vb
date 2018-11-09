
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("comienza")
        For Each arg In My.Application.CommandLineArgs.ToArray
            If Microsoft.VisualBasic.Right(arg, 5) = ".xlsx" Or
            Microsoft.VisualBasic.Right(arg, 5) = ".xlsm" Or
            Microsoft.VisualBasic.Right(arg, 4) = ".xls" Then

                carga_Excel(arg)


            End If

        Next


        MsgBox("Listo")

        Me.Close()
    End Sub
End Class
