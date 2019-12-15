Public Class Frm_InsertCustomers_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer

    Private Sub Pict_InsertCustClose_Click(sender As Object, e As EventArgs) Handles Pict_InsertCustClose.Click

        Me.Close()
        MsgBox("Wish you have a nice day ahead",, "GoodBye")

    End Sub

    Private Sub Pict_InsertCustHome_Click(sender As Object, e As EventArgs) Handles Pict_InsertCustHome.Click

        Frm_MainMenu.Show()
        Me.Close()

    End Sub

    Private Sub Pict_InsertCustBack_Click(sender As Object, e As EventArgs) Handles Pict_InsertCustBack.Click

        Dim ReturnPreviousForm = MsgBox("Return to Customer Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Customers_P101458.Show()
            Me.Close()

        End If
    End Sub

    Private Sub Pnl_InsertCustomer_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_InsertCustomer.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Frm_InsertCustomers_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Refresh_Grid()
        Txt_CustId.Text = GetNewCustId()

    End Sub

    Private Sub Refresh_Grid()

        Grd_InsertCustomers.DataSource = run_sql_query("select * from Tbl_Customers_P101458")

        Grd_InsertCustomers.Columns(0).HeaderText = "Customer Id"
        Grd_InsertCustomers.Columns(1).HeaderText = "Customer Name"
        Grd_InsertCustomers.Columns(2).HeaderText = "Customer Phone No "
        Grd_InsertCustomers.Columns(3).HeaderText = "Customer Address"
        Grd_InsertCustomers.Columns(4).HeaderText = "Customer Email"

    End Sub

    Private Function GetNewCustId() As String

        Dim LastCustMatric As String = run_sql_query("select Max(Fld_CustId) as MaxCustId from Tbl_Customers_P101458").Rows(0).Item("MaxCustId")

        Dim NewCustId As String = "TTC" & Format((Mid(LastCustMatric, 4) + 1), "000")

        Return NewCustId

    End Function

    Private Sub Btn_InsertCustomers_Click(sender As Object, e As EventArgs) Handles Btn_InsertCustomers.Click

        Dim mysql As String = "Insert into Tbl_Customers_P101458 values ('" & Txt_CustId.Text & "','" & Txt_CustName.Text & "' , " & Txt_CustPhone.Text & " ,'" & Txt_CustAddress.Text & "', '" & Txt_CustEmail.Text & " ')"



        Dim mywriter As New OleDb.OleDbCommand(mysql, myconnection2)

        Try
            mywriter.Connection.Open()
            mywriter.ExecuteNonQuery()
            mywriter.Connection.Close()

            MsgBox("INSERT successful!",, "SUCCESS")

            Refresh_Grid()
            Txt_CustId.Text = GetNewCustId()
            Txt_CustName.Text = ""
            Txt_CustAddress.Text = ""
            Txt_CustPhone.Text = ""
            Txt_CustEmail.Text = ""


        Catch ex As Exception

            Beep()
            MsgBox("Oops! The is an mistake in the INSERT command: " & ex.Message)


            mywriter.Connection.Close()

        End Try



    End Sub

    Private Sub Pnl_InsertCustomer_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_InsertCustomer.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class