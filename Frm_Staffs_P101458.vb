Public Class Frm_Staffs_P101458


    Dim current_code As String
    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer



    Private Sub Pnl_Staffs_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_Staffs.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pnl_Staffs_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_Staffs.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub

    Private Sub Frm_Staffs_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Refresh_Grid()
        Get_Current_Code()

    End Sub

    Private Sub Txt_StaffsSearch_TextChanged(sender As Object, e As EventArgs) Handles Txt_StaffsSearch.TextChanged

        Refresh_Grid()
        Get_Current_Code()


    End Sub

    Private Sub Refresh_Grid()

        Grd_Staffs.DataSource = run_sql_query("select * from Tbl_Staffs_P101458 where Fld_StaffId like '%" & Txt_StaffsSearch.Text & "%' Or Fld_StaffName Like '%" & Txt_StaffsSearch.Text & "%' Or Fld_StaffPosition like '%" & Txt_StaffsSearch.Text & "%' Or Fld_StaffQualification Like '%" & Txt_StaffsSearch.Text & "%' Or Fld_Department Like '%" & Txt_StaffsSearch.Text & "%' order by Fld_StaffName")

        Grd_Staffs.Columns(0).HeaderText = "Staff Id"
        Grd_Staffs.Columns(1).HeaderText = "Staff Name"
        Grd_Staffs.Columns(2).HeaderText = "Staff Position "
        Grd_Staffs.Columns(3).HeaderText = "Staff Qualification"
        Grd_Staffs.Columns(4).HeaderText = "Staff Department"




    End Sub

    Private Sub Get_Current_Code()

        If Grd_Staffs.SelectedRows.Count > 0 Then
            Dim Current_row As Integer = Grd_Staffs.CurrentRow.Index
            current_code = Grd_Staffs.CurrentRow.Cells(0).Value

            Try
                Pict_Staffs.BackgroundImage = Image.FromFile("StaffsPictures/" & current_code & ".jpg")
            Catch ex As Exception
                Pict_Staffs.BackgroundImage = Image.FromFile("StaffsPictures/NoPict.jpg")
            End Try

        End If
    End Sub

    Private Sub Pict_StaffsClose_MouseClick(sender As Object, e As MouseEventArgs) Handles Pict_StaffsClose.MouseClick

        Me.Close()
        MsgBox("Wish you have a nice day ahead!",, "ByeBye")

    End Sub



    Private Sub Pict_StaffsHome_MouseClick(sender As Object, e As MouseEventArgs) Handles Pict_StaffsHome.MouseClick

        Frm_MainMenu.Show()
        Me.Close()


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Frm_InsertStaffs_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Btn_StaffsInsert_Click(sender As Object, e As EventArgs) Handles Btn_StaffsInsert.Click

        Frm_UpdateDeleteStaffs_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Grd_Staffs_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_Staffs.CellClick

        Dim Current_row As Integer = e.RowIndex
        current_code = Grd_Staffs.CurrentRow.Cells(0).Value

        Try
            Pict_Staffs.BackgroundImage = Image.FromFile("StaffsPictures/" & current_code & ".jpg")
        Catch ex As Exception
            Pict_Staffs.BackgroundImage = Image.FromFile("StaffsPictures/NoPict.png")
        End Try


    End Sub


End Class