
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Module controlExcel

    Function ArrFixes(ByRef marray(,) As Object)

        Dim tmparray As New DataTable
        Dim where, fix, fixto As String
        tmparray.Columns.Add("Obj", GetType(String))
        tmparray.Columns.Add("Fix", GetType(String))
        tmparray.Columns.Add("FixTo", GetType(String))
        For x = 1 To marray.GetUpperBound(0)
            'recorre
            where = Trim(marray(x, 1)).ToString
            fix = Trim(marray(x, 2).ToString)
            fixto = Trim(marray(x, 3).ToString)
            tmparray.Rows.Add(where, fix, fixto)
        Next x
        tmparray.AcceptChanges()
        Return tmparray
    End Function
    Function ArrReporteCliqView(ByRef marray(,) As Object)
        Dim pm As String = "Total", tmp As String = "Temp", tmpd As String = "Temp"
        Dim cus, pro, ind As String
        Dim dat As Date
        Dim valu As Decimal
        Dim tmparray As New DataTable
        tmparray.Columns.Add("customer", GetType(String))
        tmparray.Columns.Add("product", GetType(String))
        tmparray.Columns.Add("date", GetType(DateTime))
        tmparray.Columns.Add("indicator", GetType(String))
        tmparray.Columns.Add("value", GetType(Decimal))
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
    Function loadArr(ByVal arch As String)
        'carga hoja de excel en 
        Dim WM_QUIT As UInteger = &H12
        Dim WM_CLOSE As UInteger = &H10

        Dim xlApp As Excel.Application

        Dim wbs As Excel.Workbooks
        Dim wb As Excel.Workbook
        Dim DictHojas As New Dictionary(Of Integer, Object)
        Try
            xlApp = New Excel.Application
            Dim hawnd = xlApp.Hwnd
            wbs = xlApp.Workbooks
            wb = wbs.Open(arch)

            For i = 1 To wb.Sheets.Count

                DictHojas.Add(i, wb.Sheets(i).UsedRange.Value)

            Next

            wb.Saved = True
            wb.Close()
            xlApp.Quit()
            PostMessage(hawnd, WM_CLOSE, 0, 0)
            PostMessage(hawnd, WM_QUIT, 0, 0)
            If (Not IsNothing(wb)) Then Marshal.ReleaseComObject(wb)
            If (Not IsNothing(wbs)) Then Marshal.ReleaseComObject(wb)
            If (Not IsNothing(xlApp)) Then Marshal.ReleaseComObject(wb)

            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally

        End Try

        loadArr = DictHojas
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




    '........................COntrol Ventanas
    Private Declare Auto Function FindWindowEx Lib "user32" (ByVal parentHandle As Integer,
                                                  ByVal childAfter As Integer,
                                                  ByVal lclassName As String,
                                                  ByVal windowTitle As String) As Integer

    Private Declare Auto Function PostMessage Lib "user32" (ByVal hwnd As Integer,
                                                            ByVal message As UInteger,
                                                            ByVal wParam As Integer,
                                                            ByVal lParam As Integer) As Boolean



End Module
