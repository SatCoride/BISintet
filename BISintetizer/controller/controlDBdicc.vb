Module controlDBdicc


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
