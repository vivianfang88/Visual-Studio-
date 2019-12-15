Public Class Frm_Products_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer



    Private Sub Frm_Products_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Lst_ClockId.DataSource = run_sql_query("select * from Tbl_Products_P101458")

        Lst_ClockId.DisplayMember = "Fld_ClockId"

        Refresh_text(Lst_ClockId.Text)

    End Sub

    Private Sub Refresh_text(ClockId As String)


        Dim mysql As String = "select * from Tbl_Products_P101458 where Fld_ClockID='" & ClockId & "'"

        Dim mydata As New DataTable

        Dim myreader As New OleDb.OleDbDataAdapter(mysql, myconnection)

        myreader.Fill(mydata)

        Txt_ClockId.Text = mydata.Rows(0).Item("Fld_ClockID")
        Txt_ClockName.Text = mydata.Rows(0).Item("Fld_ClockName")
        Txt_ClockPrice.Text = mydata.Rows(0).Item("Fld_Price")
        Txt_ClockCategory.Text = mydata.Rows(0).Item("Fld_ClockCategory")
        Txt_ClockDimension.Text = mydata.Rows(0).Item("Fld_Dimension")
        Txt_ClockQuantity.Text = mydata.Rows(0).Item("Fld_StockQuantity")
        Txt_ClockMaterial.Text = mydata.Rows(0).Item("Fld_Material")
        Txt_ClockDescription.Text = mydata.Rows(0).Item("Fld_Description")




        Try
            Pict_Products.BackgroundImage = Image.FromFile("pictures/" & ClockId & ".jpg")
        Catch ex As Exception
            Pict_Products.BackgroundImage = Image.FromFile("pictures/NoImageAvailable.jpg")
        End Try

    End Sub

    Private Sub Lst_ClockId_MouseClick(sender As Object, e As MouseEventArgs) Handles Lst_ClockId.MouseClick

        Refresh_text(Lst_ClockId.Text)

    End Sub





    Private Sub Pnl_WindowProducts_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_WindowProducts.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pict_Close_Click(sender As Object, e As EventArgs) Handles Pict_Close.Click

        Me.Close()
        Application.Exit()
        MsgBox("Thank you & wish you have a nice day ahead!",, "ByeBye")


    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Frm_MainMenu.Show()
        Me.Close()


    End Sub

    Private Sub Btn_InsertProducts_Click(sender As Object, e As EventArgs) Handles Btn_InsertProducts.Click

        Frm_InsertProducts_P101458.Show()
        Me.Close()


    End Sub

    Private Sub Btn_UpdateProducts_Click(sender As Object, e As EventArgs) Handles Btn_UpdateProducts.Click

        Frm_UpdateDeleteProducts_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Pnl_WindowProducts_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_WindowProducts.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class