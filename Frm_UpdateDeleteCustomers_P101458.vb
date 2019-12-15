Public Class Frm_UpdateDeleteCustomers_P101458

    Dim CurrentCustId As String
    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer

    Private Sub Frm_UpdateDeleteCustomers_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Refresh_Grid()
        Clear_Fields()
    End Sub

    Private Sub Refresh_Grid()

        Grd_UpdateDeleteCustomers.DataSource = run_sql_query("select * from Tbl_Customers_P101458")
        Grd_UpdateDeleteCustomers.Columns(0).HeaderText = "Customer Id"
        Grd_UpdateDeleteCustomers.Columns(1).HeaderText = "Customer Name"
        Grd_UpdateDeleteCustomers.Columns(2).HeaderText = "Customer Phone"
        Grd_UpdateDeleteCustomers.Columns(3).HeaderText = "Customer Address"
        Grd_UpdateDeleteCustomers.Columns(4).HeaderText = "Customer Email"

    End Sub

    Private Sub Clear_Fields()

        Txt_CustId.Text = ""
        Txt_CustName.Text = ""
        Txt_CustPhone.Text = ""
        Txt_CustEmail.Text = ""
        Txt_CustAddress.Text = ""

    End Sub

    Private Sub Btn_Update_Click(sender As Object, e As EventArgs) Handles Btn_Update.Click

        run_command("Update Tbl_Customers_P101458 set Fld_CustId='" & Txt_CustId.Text & "', Fld_CustName='" & Txt_CustName.Text & "', Fld_CustAddress='" & Txt_CustAddress.Text & "', Fld_CustPhone=" & Txt_CustPhone.Text & ", Fld_CustEmail='" & Txt_CustEmail.Text & "' where Fld_CustId = '" & CurrentCustId & "'")


        Refresh_Grid()
        Clear_Fields()

    End Sub

    Private Sub Grd_UpdateDeleteCustomers_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_UpdateDeleteCustomers.CellClick

        Dim CurrentRow As Integer = e.RowIndex
        CurrentCustId = Grd_UpdateDeleteCustomers(0, CurrentRow).Value


        Txt_CustId.Text = CurrentCustId
        Txt_CustName.Text = Grd_UpdateDeleteCustomers(1, CurrentRow).Value
        Txt_CustPhone.Text = Grd_UpdateDeleteCustomers(2, CurrentRow).Value
        Txt_CustAddress.Text = Grd_UpdateDeleteCustomers(3, CurrentRow).Value
        Txt_CustEmail.Text = Grd_UpdateDeleteCustomers(4, CurrentRow).Value

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_UpdateDeleteCust.Click


        Dim delete_confirmation = MsgBox("Are you sure you want to delete the course " & CurrentCustId & "?", MsgBoxStyle.YesNo, "WARNING!")

        If delete_confirmation = MsgBoxResult.Yes Then

            run_command("delete from Tbl_Customers_P101458 where Fld_CustId = '" & CurrentCustId & "'")

            Beep()
            MsgBox("The Customer '" & CurrentCustId & "' has been deleted successfully.",, "Notice")

            Refresh_Grid()
            Clear_Fields()


        End If

    End Sub

    Private Sub Pnl_UpdateDeleteCust_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_UpdateDeleteCust.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y


    End Sub

    Private Sub Pict_UpdateDeleteCustBack_Click(sender As Object, e As EventArgs) Handles Pict_UpdateDeleteCustBack.Click

        Dim ReturnPreviousForm = MsgBox("Return to Customer Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Customers_P101458.Show()
            Me.Close()

        End If

    End Sub

    Private Sub Pict_UpdateDeleteCustHome_Click(sender As Object, e As EventArgs) Handles Pict_UpdateDeleteCustHome.Click

        Frm_MainMenu.Show()
        Me.Close()

    End Sub

    Private Sub Pict_UpdateDeleteCustClose_Click(sender As Object, e As EventArgs) Handles Pict_UpdateDeleteCustClose.Click

        MsgBox("Wish you have a nice day ahead! =D",, "ByeBye")
        Me.Close()


    End Sub

    Private Sub Pnl_UpdateDeleteCust_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_UpdateDeleteCust.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class