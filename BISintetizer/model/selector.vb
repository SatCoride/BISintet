Module selector
    Dim arrFHoja As DataTable

    Sub carga_Excel(ByRef dire As String)
        Dim dicarrH As New Dictionary(Of Integer, Object)
        dicarrH = loadArr(dire)
        'dicarrH = loadArr("D:\Documentos\Desktop\Análisis de Procesos CEVA\Desarrollo\BISintetizer\BISintet\BISintetizer\BISintetizer\tmp\report.xlsx")
        For Each arrhoj As KeyValuePair(Of Integer, Object) In dicarrH
            Dim v1 As Integer = arrhoj.Key
            Dim v2 As Object = arrhoj.Value
            '------------ARCHIVO QLIKVIEW
            If v2(3, 1).ToString = "Reporting Customer" And
               v2(3, 2).ToString = "Product" And
               v2(3, 3).ToString = "Total" Then

                ejeQlikViewReport(v2)

            End If
            '------------ARCHIVO QLIKVIEW
            '------------ARCHIVO FIX
            If v2(1, 1).ToString = "BD" Then
                For I = 1 To UBound(v2, 1)
                    If v2(I, 1).ToString = "CUSTOMERS" Then
                        ejeFixes(v2)
                        Exit Sub
                    End If
                Next
            End If
            '------------ARCHIVO FIX

        Next

        'arrH = loadArr("D:\Documentos\Desktop\Análisis de Procesos CEVA\Desarrollo\BISintetizer\BISintet\BISintetizer\BISintetizer\tmp\fixes.xlsx")
        'Condicionales 
        'si es qv
        'ejeQlikViewReport(arrH)
        'ejeFixes(arrH)
    End Sub
    Sub ejeFixes(ByRef arrH As Object)
        arrFHoja = ArrFixes(arrH)
        'BorrarTablaFixes()
        'CrearTablaFixes()
        VaciarTablaFixes()
        LlenarTablaFixes(arrFHoja)
        fixTablaQliqview()
    End Sub
    Sub ejeQlikViewReport(ByRef arrH As Object)
        arrFHoja = ArrReporteCliqView(arrH)
        'BorrarTablaQlikViewReport()
        'CrearTablaQlikViewReport()
        VaciarTablaQlickViewReport()
        LlenarTablaQlickViewReport(arrFHoja)
        fixTablaQliqview()
    End Sub
End Module
