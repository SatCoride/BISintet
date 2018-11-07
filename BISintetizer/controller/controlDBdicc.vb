Module controlDBdicc

    Sub LlenarTablaQlickViewReport(ByRef listlimp As List(Of List(Of String)))
        connectDB()

        For Each reg In listlimp
            nonqueryDB("insert into qlickviewreport (customer,product,date,indicator,value) values ('" & String.Join("','", reg).ToArray() & "')")
        Next

        disconnectDB()
    End Sub
    Sub CrearTablaQlikViewReport()
        Dim str = "CREATE TABLE [dbo].[qlickViewReport]
            (
	            [Id] INT NOT NULL PRIMARY KEY IDENTITY, 
                [customer] VARCHAR(50) NULL, 
                [product] VARCHAR(20) NULL, 
                [date] DATE NULL, 
                [indicator] VARCHAR(20) NULL, 
                [value] VARCHAR(20) NULL
            )
            "
        connectDB()
        nonqueryDB(str)
        disconnectDB()
    End Sub
    Sub BorrarTablaQlikViewReport()
        Dim str = "DROP TABLE [dbo].[qlickViewReport];"
        connectDB()
        nonqueryDB(str)
        disconnectDB()
    End Sub
End Module
