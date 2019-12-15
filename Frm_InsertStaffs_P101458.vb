Public Class Frm_InsertStaffs_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer
    Dim DefaultPicture As String = Application.StartupPath & "\StaffsPictures\NoPict.png"

    Private Sub Frm_InsertStaffs_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Cmb_StaffPosition.DataSource = run_sql_query("select distinct Fld_StaffPosition from Tbl_Staffs_P101458")
        Cmb_StaffPosition.DisplayMember = "Fld_StaffPosition"

        Cmb_StaffDept.DataSource = run_sql_query("select distinct Fld_Department from Tbl_Staffs_P101458")
        Cmb_StaffDept.DisplayMember = "Fld_Department"

        Refresh_Grid()
        Txt_StaffId.Text = GenerateStaffId()


        Pict_InsertStaffs.BackgroundImage = Image.FromFile(DefaultPicture)
        Txt_InsertPict.Text = DefaultPicture


    End Sub

    Private Sub Refresh_Grid()

        Grd_InsertStaffs.DataSource = run_sql_query("select * from Tbl_Staffs_P101458")

        Grd_InsertStaffs.Columns(0).HeaderText = "Staff Id"
        Grd_InsertStaffs.Columns(1).HeaderText = "Staff Name"
        Grd_InsertStaffs.Columns(2).HeaderText = "Staff Position "
        Grd_InsertStaffs.Columns(3).HeaderText = "Staff Qualification"
        Grd_InsertStaffs.Columns(4).HeaderText = "Staff Department"

        Txt_StaffId.Text = GenerateStaffId()

    End Sub

    Private Function GenerateStaffId() As String

        Dim LastStaffMatric As String = run_sql_query("select Max(Fld_StaffId) as MaxStaffId from Tbl_Staffs_P101458").Rows(0).Item("MaxStaffId")

        'MsgBox(LastStaffMatric)

        Dim NewStaffId As String = "TTS" & Format((Mid(LastStaffMatric, 4) + 1), "000")

        Return NewStaffId


    End Function


    Private Sub Pnl_InsertStaffs_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_InsertStaffs.MouseDown


        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pict_Close_Click(sender As Object, e As EventArgs) Handles Pict_Close.Click

        Me.Close()
        MsgBox("Wish you have a nice day ahead",, "GoodBye")

    End Sub

    Private Sub Pict_MainMenu_Click(sender As Object, e As EventArgs) Handles Pict_MainMenu.Click

        Frm_MainMenu.Show()
        Me.Close()


    End Sub

    Private Sub Pict_InsertStaffBack_Click(sender As Object, e As EventArgs) Handles Pict_InsertStaffBack.Click

        Dim ReturnPreviousForm = MsgBox("Return to Staff Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Staffs_P101458.Show()
            Me.Close()

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_Insert.Click

        Dim mysql As String = "Insert into Tbl_Staffs_P101458 values ('" & Txt_StaffId.Text & "','" & Txt_StaffName.Text & "','" & Cmb_StaffDept.Text & "' , '" & Cmb_StaffPosition.Text & " ','" & Txt_StaffQualification.Text & "')"



        Dim mywriter As New OleDb.OleDbCommand(mysql, myconnection2)

        Try
            mywriter.Connection.Open()
            mywriter.ExecuteNonQuery()
            mywriter.Connection.Close()

            My.Computer.FileSystem.CopyFile(Txt_InsertPict.Text, "StaffsPictures\" & Txt_StaffId.Text & ".jpg")

            MsgBox("INSERT successful!",, "SUCCESS")

            Refresh_Grid()
            Txt_StaffId.Text = GenerateStaffId()
            Txt_StaffName.Text = ""
            Txt_StaffQualification.Text = ""


            Pict_InsertStaffs.BackgroundImage = Image.FromFile(DefaultPicture)
            Txt_InsertPict.Text = DefaultPicture


        Catch ex As Exception



            Beep()
            MsgBox("Oops! The is an mistake in the INSERT command: " & ex.Message)


            mywriter.Connection.Close()

        End Try


    End Sub

    Private Sub Btn_InsertPict_Click(sender As Object, e As EventArgs) Handles Btn_InsertPict.Click

        Dim mydesktop As String = My.Computer.FileSystem.SpecialDirectories.Desktop

        OpenFileDialog1.InitialDirectory = mydesktop
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "JPG files (*.jpg)|*.jpg"
        OpenFileDialog1.ShowDialog()

        Pict_InsertStaffs.BackgroundImage = Image.FromFile(OpenFileDialog1.FileName)
        Txt_InsertPict.Text = OpenFileDialog1.FileName


    End Sub

    Private Sub Pnl_InsertStaffs_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_InsertStaffs.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class