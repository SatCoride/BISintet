Imports System.ComponentModel

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim arrFHoja As DataTable
        'BorrarTablaQlikViewReport()
        'CrearTablaQlikViewReport()
        arrFHoja = ArrReporteCliqView(loadArr())
        LlenarTablaQlickViewReport(arrFHoja)
        'MsgBox(String.Join(Environment.NewLine, arrFHoja(1).ToArray()))

        Me.Close()
    End Sub

End Class
