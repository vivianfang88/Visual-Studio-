Public Class Frm_UpdateDeleteStaffs_P101458

    Dim CurrentCode As String
    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer


    Private Sub Refresh_Grid()

        Grd_UpdateStaffs.DataSource = run_sql_query("select * from Tbl_Staffs_P101458")

        Grd_UpdateStaffs.Columns(0).HeaderText = "Staff Id"
        Grd_UpdateStaffs.Columns(1).HeaderText = "Staff Name"
        Grd_UpdateStaffs.Columns(2).HeaderText = "Staff Position "
        Grd_UpdateStaffs.Columns(3).HeaderText = "Staff Qualification"
        Grd_UpdateStaffs.Columns(4).HeaderText = "Staff Department"




    End Sub

    Private Sub PictClose_Click(sender As Object, e As EventArgs) Handles PictClose.Click

        MsgBox("Wish you have a nice day ahead! =D")
        Me.Close()


    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles Pict_Home.Click

        Frm_MainMenu.Show()
        Me.Close()

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles Pict_Back.Click

        Dim ReturnPreviousForm = MsgBox("Return to Staffs Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Staffs_P101458.Show()
            Me.Close()

        End If

    End Sub

    Private Sub Pnl_InsertDeleteStaffs_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_InsertDeleteStaffs.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Frm_UpdateDeleteStaffs_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Cmb_Position.DataSource = run_sql_query("select distinct Fld_StaffPosition from Tbl_Staffs_P101458")
        Cmb_Position.DisplayMember = "Fld_StaffPosition"

        Cmb_Department.DataSource = run_sql_query("select distinct Fld_Department from Tbl_Staffs_P101458")
        Cmb_Department.DisplayMember = "Fld_Department"

        Refresh_Grid()


    End Sub

    Private Sub Btn_Update_Click(sender As Object, e As EventArgs) Handles Btn_Update.Click

        run_command("Update Tbl_Staffs_P101458 set Fld_StaffId='" & Txt_StaffId.Text & "', Fld_StaffName='" & Txt_StaffName.Text & "', Fld_StaffPosition='" & Cmb_Position.Text & "', Fld_Department='" & Cmb_Department.Text & "', Fld_StaffQualification='" & Txt_Qualification.Text & "' where Fld_StaffId = '" & CurrentCode & "'")

        Refresh_Grid()
        Clear_Fields()

    End Sub

    Private Sub Clear_Fields()

        Txt_StaffId.Text = ""
        Txt_StaffName.Text = ""
        Cmb_Position.Text = ""
        Cmb_Department.Text = ""
        Txt_Qualification.Text = ""

    End Sub


    Private Sub Pnl_InsertDeleteStaffs_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_InsertDeleteStaffs.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub

    Private Sub Btn_Delete_Click(sender As Object, e As EventArgs) Handles Btn_Delete.Click

        Dim delete_confirmation = MsgBox("Are you sure you want to delete the course " & CurrentCode & "?", MsgBoxStyle.YesNo, "WARNING!")

        If delete_confirmation = MsgBoxResult.Yes Then

            run_command("delete from Tbl_Staffs_P101458 where Fld_StaffId = '" & CurrentCode & "'")

            Beep()
            MsgBox("The Staffs '" & CurrentCode & "' has been deleted successfully.",, "Notice")

            Refresh_Grid()
            Clear_Fields()


        End If

    End Sub

    Private Sub Grd_UpdateStaffs_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_UpdateStaffs.CellClick

        Dim CurrentRow As Integer = e.RowIndex
        CurrentCode = Grd_UpdateStaffs(0, CurrentRow).Value


        Txt_StaffId.Text = CurrentCode
        Txt_StaffName.Text = Grd_UpdateStaffs(1, CurrentRow).Value
        Cmb_Position.Text = Grd_UpdateStaffs(2, CurrentRow).Value
        Cmb_Department.Text = Grd_UpdateStaffs(3, CurrentRow).Value
        Txt_Qualification.Text = Grd_UpdateStaffs(4, CurrentRow).Value


    End Sub
End Class