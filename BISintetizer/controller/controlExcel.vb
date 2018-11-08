
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Module controlExcel
    Function ArrReporteCliqView(ByRef marray(,) As Object)
        Dim pm As String = "Total", tmp As String = "Temp", tmpd As String = "Temp"
        Dim cus, pro, ind As String
        Dim dat As Date
        Dim valu As Integer
        Dim tmparray As New DataTable
        tmparray.Columns.Add("customer", GetType(String))
        tmparray.Columns.Add("product", GetType(String))
        tmparray.Columns.Add("date", GetType(DateTime))
        tmparray.Columns.Add("indicator", GetType(String))
        tmparray.Columns.Add("value", GetType(Integer))
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


                            cus = Trim(marray(x, 1)).ToString
                            pro = Trim(marray(x, 2).ToString)
                            dat = CDate(Trim(marray(2, y))).ToString("dd/MM/yyyy")
                            ind = Trim(marray(1, y).ToString)
                            valu = Math.Round(Val(Trim(marray(x, y))), 2)
                            tmparray.Rows.Add(cus, pro, dat, ind, valu)

                        End If
                    Next y
                End If
            End If
        Next x
        tmparray.AcceptChanges()
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
