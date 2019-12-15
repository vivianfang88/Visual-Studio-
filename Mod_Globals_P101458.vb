Module Mod_Globals_P101458



    Public myconnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Db_TickTockStore_P101458.accdb;Persist Security Info=False;"

    Public myconnection2 As New OleDb.OleDbConnection(myconnection)

    Public Function run_sql_query(mysql As String) As DataTable

        Dim mydata As New DataTable

        Dim myreader As New OleDb.OleDbDataAdapter(mysql, myconnection)

        Try

            myreader.Fill(mydata)


        Catch ex As Exception

            Beep()
            MsgBox("Oops the is an error in the SELECT statement: " & vbCrLf & vbCrLf & ex.Message)
        End Try

        Return mydata


    End Function

    Public Sub run_command(mysql As String)

        Dim mywriter As New OleDb.OleDbCommand(mysql, myconnection2)

        Try

            mywriter.Connection.Open()
            mywriter.ExecuteNonQuery()
            mywriter.Connection.Close()

            MsgBox("Change successful!",, "SUCCESSFUL!")

        Catch ex As Exception

            Beep()
            MsgBox("Oops! The is an error when changing data: " & vbCrLf & vbCrLf & ex.Message)

            mywriter.Connection.Close()

        End Try

    End Sub


End Module




