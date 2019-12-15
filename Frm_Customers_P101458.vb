Public Class Frm_Customers_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer

    Private Sub Frm_Customers_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Grd_Customers.DataSource = run_sql_query("select * from Tbl_Customers_P101458")


    End Sub

    Private Sub refresh_grid()




        Grd_Customers.DataSource = run_sql_query("select * from Tbl_Customers_P101458 where Fld_CustName like '%" & Txt_Search.Text & "%' Or Fld_CustId Like '%" & Txt_Search.Text & "%' Or Fld_CustEmail like '%" & Txt_Search.Text & "%' Or Fld_CustPhone Like '%" & Txt_Search.Text & "%' Or Fld_CustAddress Like '%" & Txt_Search.Text & "%' order by Fld_CustId")

        If Grd_Customers.Rows.Count > 0 Then

            Grd_Customers.Columns(0).HeaderText = "Customer Id"
            Grd_Customers.Columns(1).HeaderText = "Customer Name"
            Grd_Customers.Columns(2).HeaderText = "Customer Phone No "
            Grd_Customers.Columns(3).HeaderText = "Customer Address"
            Grd_Customers.Columns(4).HeaderText = "Customer Email"


        End If







    End Sub

    Private Sub Txt_search_TextChanged(sender As Object, e As EventArgs) Handles Txt_Search.TextChanged

        refresh_grid()

    End Sub



    Private Sub Pnl_WindowCustomers_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_WindowCustomers.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pict_CustClose_Click(sender As Object, e As EventArgs) Handles Pict_CustClose.Click

        MsgBox("Guess you have the needed information. Wish you have a nice day ahead!",, "ByeBye")
        Me.Close()
        'MsgBox("Guess you have the needed information. Wish you have a nice day ahead!")


    End Sub

    Private Sub Pict_CustHome_Click(sender As Object, e As EventArgs) Handles Pict_CustHome.Click

        Frm_MainMenu.Show()
        Me.Close()

    End Sub

    Private Sub Btn_InsertCustomer_Click(sender As Object, e As EventArgs) Handles Btn_InsertCustomer.Click

        Frm_InsertCustomers_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Btn_UpdateCustomer_Click(sender As Object, e As EventArgs) Handles Btn_UpdateCustomer.Click

        Frm_UpdateDeleteCustomers_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Pnl_WindowCustomers_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_WindowCustomers.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class