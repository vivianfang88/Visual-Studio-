Public Class Frm_InsertProducts_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer
    Dim DefaultPicture As String = Application.StartupPath & "\Pictures\NoImageAvailable.jpg"



    Private Sub Pict_Close_Click(sender As Object, e As EventArgs) Handles Pict_Close.Click

        Me.Close()
        MsgBox("Wish you have a pleasant day =D",, "ByeBye")


    End Sub

    Private Sub Pict_Home_Click(sender As Object, e As EventArgs) Handles Pict_Home.Click


        Frm_MainMenu.Show()
        Me.Close()


    End Sub

    Private Sub Pnl_InsertProduct_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_InsertProduct.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Frm_InsertProducts_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Cmb_Category.DataSource = run_sql_query("select distinct Fld_ClockCategory from Tbl_Products_P101458")
        Cmb_Category.DisplayMember = "Fld_ClockCategory"

        Refresh_Grid()
        Txt_ClockId.Text = NewCode()
        Pict_InsertProducts.BackgroundImage = Image.FromFile(DefaultPicture)
        Txt_InsertPict.Text = DefaultPicture


    End Sub

    Private Sub Refresh_Grid()

        Grd_InsertProduct.DataSource = run_sql_query("select * from Tbl_Products_P101458")

        Grd_InsertProduct.Columns(0).HeaderText = "Clock Id"
        Grd_InsertProduct.Columns(1).HeaderText = "Clock Name"
        Grd_InsertProduct.Columns(2).HeaderText = "Price"
        Grd_InsertProduct.Columns(3).HeaderText = "Clock Category"
        Grd_InsertProduct.Columns(4).HeaderText = "Clock Dimension"
        Grd_InsertProduct.Columns(5).HeaderText = "Stock Quantity"
        Grd_InsertProduct.Columns(6).HeaderText = "Clock Material"
        Grd_InsertProduct.Columns(7).HeaderText = "Customer Description"

        Txt_ClockId.Text = NewCode()

    End Sub

    Private Function NewCode() As String

        Dim LastId As String = run_sql_query("select max(Fld_ClockId) as MaxId from Tbl_Products_P101458").Rows(0).Item("MaxId")

        'MsgBox(LastId)

        Dim New_InsertId As String = "TT" & Mid(LastId, 3) + 1

        'MsgBox(New_InsertId)

        Return New_InsertId

    End Function

    Private Sub Btn_Insert_Click(sender As Object, e As EventArgs) Handles Btn_Insert.Click

        'Dim mysql As String = "insert into tbl_students_a123456 values ('matric','name','dept')"
        Dim mysql As String = "Insert into Tbl_Products_P101458 values ('" & Txt_ClockId.Text & "','" & Txt_ClockName.Text & "','" & Cmb_Category.Text & "' , '" & Txt_Description.Text & " ','" & Txt_Dimension.Text & "'," & Txt_Price.Text & ", '" & Txt_Material.Text & " '," & Txt_Quantity.Text & ")"



        Dim mywriter As New OleDb.OleDbCommand(mysql, myconnection2)

        Try
            mywriter.Connection.Open()
            mywriter.ExecuteNonQuery()
            mywriter.Connection.Close()

            My.Computer.FileSystem.CopyFile(Txt_InsertPict.Text, "Pictures\" & Txt_ClockId.Text & ".jpg")

            MsgBox("INSERT successful!",, "SUCCESS")

            Refresh_Grid()
            Txt_ClockName.Text = ""
            Txt_Description.Text = ""
            Txt_Dimension.Text = ""
            Txt_Material.Text = ""
            Txt_Price.Text = ""
            Txt_Quantity.Text = ""

            Pict_InsertProducts.BackgroundImage = Image.FromFile(DefaultPicture)
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

        Pict_InsertProducts.BackgroundImage = Image.FromFile(OpenFileDialog1.FileName)
        Txt_InsertPict.Text = OpenFileDialog1.FileName

    End Sub


    Private Sub Pict_BackInsertProduct_Click(sender As Object, e As EventArgs) Handles Pict_BackInsertProduct.Click

        Dim ReturnPreviousForm = MsgBox("Return to Product Form?", MsgBoxStyle.YesNo, "Question")
        If ReturnPreviousForm = MsgBoxResult.Yes Then

            Frm_Products_P101458.Show()
            Me.Close()
        End If

    End Sub

    Private Sub Pnl_InsertProduct_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_InsertProduct.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class