Public Class Frm_UpdateDeleteProducts_P101458


    Dim CurrentCode As String
    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer
    'Dim DefaultPicture As String = Application.StartupPath & "\Pictures\'" & CurrentCode & "'.jpg"

    Private Sub Pnl_UpdateProducts_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_UpdateProducts.MouseDown


        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pict_UpdateProBack_Click(sender As Object, e As EventArgs) Handles Pict_UpdateProBack.Click

        Dim ReturnPreviousForm = MsgBox("Return to Product Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Products_P101458.Show()
            Me.Close()
        End If
    End Sub

    Private Sub Pcit_UpdateProHome_Click(sender As Object, e As EventArgs) Handles Pcit_UpdateProHome.Click

        Frm_MainMenu.Show()
        Me.Close()


    End Sub

    Private Sub Pict_UpdateProClose_Click(sender As Object, e As EventArgs) Handles Pict_UpdateProClose.Click

        MsgBox("Wish you have a nice day",, "Good Bye")
        Me.Close()


    End Sub


    Private Sub Pnl_UpdateProducts_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_UpdateProducts.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If


    End Sub

    Private Sub Refresh_Grid()

        Grd_UpdateProducts.DataSource = run_sql_query("Select * from Tbl_Products_P101458")

        Grd_UpdateProducts.Columns(0).HeaderText = "Clock Id"
        Grd_UpdateProducts.Columns(1).HeaderText = "Clock Name"
        Grd_UpdateProducts.Columns(2).HeaderText = "Price"
        Grd_UpdateProducts.Columns(3).HeaderText = "Clock Category"
        Grd_UpdateProducts.Columns(4).HeaderText = "Clock Dimension"
        Grd_UpdateProducts.Columns(5).HeaderText = "Stock Quantity"
        Grd_UpdateProducts.Columns(6).HeaderText = "Clock Material"
        Grd_UpdateProducts.Columns(7).HeaderText = "Customer Description"

    End Sub

    Private Sub Frm_UpdateProducts_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Cmb_Category.DataSource = run_sql_query("select Distinct Fld_ClockCategory from Tbl_Products_P101458")
        Cmb_Category.DisplayMember = "Fld_ClockCategory"



        Refresh_Grid()

    End Sub

    Private Sub Clear_Fields()


        Txt_ClockId.Text = ""
        Txt_ClockName.Text = ""
        Txt_Price.Text = ""
        Cmb_Category.Text = ""
        Txt_Dimension.Text = ""
        Txt_StockQuantity.Text = ""
        Txt_Material.Text = ""
        Txt_Description.Text = ""

    End Sub

    Private Sub Btn_Update_Click(sender As Object, e As EventArgs) Handles Btn_Update.Click

        run_command("Update Tbl_Products_P101458 set Fld_ClockId='" & Txt_ClockId.Text & "', Fld_ClockName='" & Txt_ClockName.Text & "', Fld_Price=" & Txt_Price.Text & ", Fld_ClockCategory='" & Cmb_Category.Text & "', Fld_Dimension='" & Txt_Dimension.Text & "', Fld_StockQuantity=" & Txt_StockQuantity.Text & ", Fld_Material='" & Txt_Material.Text & "', Fld_Description='" & Txt_Description.Text & "' where Fld_ClockId = '" & CurrentCode & "'")




        Refresh_Grid()
        Clear_Fields()



    End Sub

    Private Sub Btn_Delete_Click(sender As Object, e As EventArgs) Handles Btn_Delete.Click



        Dim delete_confirmation = MsgBox("Are you sure you want to delete the course " & CurrentCode & "?", MsgBoxStyle.YesNo, "WARNING!")

        If delete_confirmation = MsgBoxResult.Yes Then

            run_command("delete from Tbl_Products_P101458 where Fld_ClockId = '" & CurrentCode & "'")

            Beep()
            MsgBox("The Products '" & CurrentCode & "' has been deleted successfully.",, "Notice")

            Refresh_Grid()
            Clear_Fields()


        End If

    End Sub



    Private Sub Grd_UpdateProducts_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_UpdateProducts.CellClick

        Dim CurrentRow As Integer = e.RowIndex
        CurrentCode = Grd_UpdateProducts(0, CurrentRow).Value


        Txt_ClockId.Text = CurrentCode
        Txt_ClockName.Text = Grd_UpdateProducts(1, CurrentRow).Value
        Txt_Price.Text = Grd_UpdateProducts(2, CurrentRow).Value
        Cmb_Category.Text = Grd_UpdateProducts(3, CurrentRow).Value
        Txt_Dimension.Text = Grd_UpdateProducts(4, CurrentRow).Value
        Txt_StockQuantity.Text = Grd_UpdateProducts(5, CurrentRow).Value
        Txt_Material.Text = Grd_UpdateProducts(6, CurrentRow).Value
        Txt_Description.Text = Grd_UpdateProducts(7, CurrentRow).Value



    End Sub
End Class