Imports System.Data.SqlClient
Module controlDBfunc
    Private myConn As SqlConnection
    Private myCmd As SqlCommand
    Private myReader As SqlDataReader
    Private results As String


    'TEMP
    Function LlenarTablaFixes(ByRef listlimp As DataTable)

        Try
            connectDB() 'Function For opening connection

            myCmd = myConn.CreateCommand
            myCmd.CommandText = "fillFixes"
            myCmd.CommandType = CommandType.StoredProcedure
            myCmd.Parameters.AddWithValue("@listlimp", listlimp)
            myCmd.ExecuteNonQuery()
            Return True
        Catch e As SqlException
            ''strMsg = "Data not saved successfully.";
            MsgBox(e.Message.ToString)
            Return False
        Finally

            disconnectDB() 'Function For closing connection
        End Try


    End Function

    Function LlenarTablaQlickViewReport(ByRef listlimp As DataTable)

        Try
            connectDB() 'Function For opening connection

            myCmd = myConn.CreateCommand
            myCmd.CommandText = "fillQliqViewReport"
            myCmd.CommandType = CommandType.StoredProcedure
            myCmd.Parameters.AddWithValue("@listlimp", listlimp)
            myCmd.ExecuteNonQuery()
            Return True
        Catch e As SqlException
            ''strMsg = "Data not saved successfully.";
            MsgBox(e.Message.ToString)
            Return False
        Finally

            disconnectDB() 'Function For closing connection
        End Try


    End Function
    'TEMP
    Sub connectDB()
        Try
            Dim server As String = "(localdb)\MSSQLLocalDB"
            Dim database As String = "BIDB"
            myConn = New SqlConnection("Data Source=" & server & ";Initial Catalog=" & database & ";Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
            myConn.Open()
        Catch ex As Exception
            MessageBox.Show("Error while conecting into db..." & ex.Message, "Connection")
        End Try
    End Sub
    Sub disconnectDB()
        Try
            myConn.Close()
            myConn = Nothing
        Catch ex As Exception
            MessageBox.Show("Error while disconecting into db..." & ex.Message, "Disconnection")
        End Try
    End Sub
    Sub nonqueryDB(ByVal cmd As String)
        Try
            myCmd = myConn.CreateCommand
            'myCmd = New SqlCommand(cmd)
            myCmd.CommandText = cmd
            'myCmd.CommandText = cmd
            myCmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Error while executing query..." & ex.Message, "noQuery")
        End Try
    End Sub

End Module
