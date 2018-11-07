
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Module controlExcel
    Function ArrReporteCliqView(ByRef marray(,) As Object)
        Dim pm As String = "Total", tmp As String = "Temp", tmpd As String = "Temp"
        Dim tmparray As New List(Of List(Of String))
        tmparray.Add(New List(Of String))
        Dim cx As Integer = 0
        'Recorre desde eje x
        For x = 3 To marray.GetUpperBound(0)
            'Filtro RC x
            If Not IsNothing(marray(x, 2)) And
               Not IsNothing(marray(x, 1)) Then
                If IsNothing(marray(x, 3)) Then tmp = "" Else tmp = marray(x, 3)
                If IsNothing(marray(x, 2)) Then tmpd = "" Else tmpd = marray(x, 2)

                If tmp <> pm And
                   tmpd <> pm Then
                    'Recorre desde eje y
                    For y = 3 To marray.GetUpperBound(1)
                        If marray(2, y).ToString <> pm And
                           marray(3, y).ToString <> pm Then
                            'Filtro RC y
                            tmparray.Add(New List(Of String))
                            tmparray(cx).Add(Trim(marray(x, 1)).ToString)
                            tmparray(cx).Add(Trim(marray(x, 2).ToString))
                            tmparray(cx).Add(CDate(Trim(marray(2, y))).ToString("dd/MM/yyyy"))
                            tmparray(cx).Add(Trim(marray(1, y).ToString).ToString)
                            tmparray(cx).Add(Math.Round(Val(Trim(marray(x, y))), 2))
                            cx += 1
                        End If
                    Next y
                End If
            End If
        Next x
        Return tmparray
    End Function
    Function loadArr()
        'carga hoja de excel en 
        Dim xlApp = New Excel.Application
        Dim wb As Excel.Workbook = xlApp.Workbooks.Open("D:\Documentos\Desktop\Análisis de Procesos CEVA\Desarrollo\BISintetizer\BISintet\BISintetizer\BISintetizer\tmp\report.xlsx")
        Dim ws As Excel.Worksheet = wb.ActiveSheet
        Dim arr(,) As Object

        'Entorno
        arr = ws.UsedRange.Value

        NAR(ws)
        ws = Nothing


        wb.Saved = True
        wb.Close()
        NAR(wb)
        wb = Nothing

        xlApp.Workbooks.Close()
        NAR(xlApp.Workbooks)

        xlApp.Quit()
        NAR(xlApp)
        xlApp = Nothing


        GC.Collect()
        GC.WaitForPendingFinalizers()
        loadArr = arr
    End Function
    Private Sub NAR(ByRef o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

End Module
