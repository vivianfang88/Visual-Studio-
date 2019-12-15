Public Class Frm_ItemPurchasedRecord_P101458
    Dim ClickCurrent_Code As String
    Dim Current_Code As String
    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles Pict_ItemsPRHome.Click


        Frm_MainMenu.Show()
        Me.Close()

    End Sub

    Private Sub Pict_ItemsPRClose_Click(sender As Object, e As EventArgs) Handles Pict_ItemsPRClose.Click


        MsgBox("Wish you have a pleasant day",, "ByeBye")
        Me.Close()


    End Sub


    Private Sub Get_Current_Codes()


        If Grd_ItemsPR.SelectedRows.Count > 0 Then
            Dim Current_row As Integer = Grd_ItemsPR.CurrentRow.Index
            Current_Code = Grd_ItemsPR(0, Current_row).Value
            Try
                Pict_ItemsPR.BackgroundImage = Image.FromFile("Pictures/" & Current_Code & ".jpg")
            Catch ex As Exception
                Pict_ItemsPR.BackgroundImage = Image.FromFile("Pictures/NoImageAvailable.jpg")
            End Try


        End If




    End Sub

    Private Sub Grd_ItemsPR_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_ItemsPR.CellClick


        Dim Current_row As Integer = e.RowIndex
        ClickCurrent_Code = Grd_ItemsPR.CurrentRow.Cells(0).Value


        Try
            Pict_ItemsPR.BackgroundImage = Image.FromFile("Pictures/" & ClickCurrent_Code & ".jpg")
        Catch ex As Exception
            Pict_ItemsPR.BackgroundImage = Image.FromFile("Pictures/NoImageAvailable.jpg")
        End Try





    End Sub

    Private Sub Refresh_Grid()


        Grd_ItemsPR.DataSource = run_sql_query("select * from Tbl_ItemsPurchasedRecord_P101458 where Fld_ClockId like '%" & Txt_SearchItemsPR.Text & "%' order by Fld_OrderId")


        If Grd_ItemsPR.Rows.Count > 0 Then

            Grd_ItemsPR.Columns(0).HeaderText = "Clock Id"
            Grd_ItemsPR.Columns(1).HeaderText = "Order Id"
            Grd_ItemsPR.Columns(2).HeaderText = "Customer Id"
            Grd_ItemsPR.Columns(3).HeaderText = "Quantity"
            Grd_ItemsPR.Columns(4).HeaderText = "Unit Price"



        End If




    End Sub

    Private Sub Txt_SearchItemsPR_TextChanged(sender As Object, e As EventArgs) Handles Txt_SearchItemsPR.TextChanged


        Refresh_Grid()
        Get_Current_Codes()

    End Sub



    Private Sub Frm_ItemPurchasedRecord_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Refresh_Grid()
        Get_Current_Codes()


    End Sub

    Private Sub Pnl_ItemsPR_MouseDown(sender As Object, e As MouseEventArgs) Handles Pnl_ItemsPR.MouseDown


        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub



    Private Sub Pnl_ItemsPR_MouseMove(sender As Object, e As MouseEventArgs) Handles Pnl_ItemsPR.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If

    End Sub
End Class