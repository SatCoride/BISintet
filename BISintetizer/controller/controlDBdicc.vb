Module controlDBdicc

    '                           FIXES
    Sub VaciarTablaFixes()
        Dim str = "DELETE From [dbo].[fixes];"
        connectDB()
        nonqueryDB(str)
        str = "DBCC CHECKIDENT('[dbo].[fixes]', RESEED, 0)"
        nonqueryDB(str)
        disconnectDB()
    End Sub

    Sub CrearTablaFixes()
        Dim str = "CREATE TABLE [dbo].[fixes]
            (
	            [Id] INT NOT NULL PRIMARY KEY IDENTITY, 
                [obj] VARCHAR(10), 
                [fix] VARCHAR(120), 
                [fixto] VARCHAR(120)
            )
            "
        connectDB()
        nonqueryDB(str)
        disconnectDB()
    End Sub
    Sub BorrarTablaFixes()
        Dim str = "DROP TABLE [dbo].[fixes];"
        connectDB()
        nonqueryDB(str)
        disconnectDB()
    End Sub
    '                           FIXES FIN






    '                           QlikViewReport
    Sub fixTablaQliqview()
        Dim str = "UPDATE
        qlickViewReport
            SET
                qlickViewReport.customer = RAN.fixto
                    FROM
                    qlickViewReport SI
            INNER Join
                fixes RAN
            On 
                SI.customer = RAN.fix
        where ran.obj = 'BD';"
        connectDB()
        nonqueryDB(str)
        disconnectDB()
    End Sub

    Sub VaciarTablaQlickViewReport()
        Dim str = "DELETE From [dbo].[qlickViewReport];"
        connectDB()
        nonqueryDB(str)
        str = "DBCC CHECKIDENT('[dbo].[qlickViewReport]', RESEED, 0)"
        nonqueryDB(str)
        disconnectDB()
    End Sub

    Sub CrearTablaQlikViewReport()
        Dim str = "CREATE TABLE [dbo].[qlickViewReport]
            (
	            [Id] INT NOT NULL PRIMARY KEY IDENTITY, 
                [customer] VARCHAR(50), 
                [product] VARCHAR(20), 
                [date] DATE, 
                [indicator] VARCHAR(20), 
                [value] DECIMAL(20,2) NULL
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
    '                           QlikViewReport FIN
End Module
