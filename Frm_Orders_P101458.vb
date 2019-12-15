Public Class Frm_Orders_P101458

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer


    Private Sub Panel1_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel1.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y


    End Sub

    Private Sub Frm_Orders_P101458_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Cmb_ProcessBy.DataSource = run_sql_query("select * from Tbl_Staffs_P101458")
        Cmb_ProcessBy.DisplayMember = "Fld_StaffId"


        Refresh_Grid()



    End Sub

    Private Sub Refresh_Grid()
        Dim TotalPrice As String = "=" & Txt_TotalPrice.Text

        If Txt_TotalPrice.Text = "" Then
            TotalPrice = ">= 0"
        End If


        Grd_Orders.DataSource = run_sql_query("select * from Tbl_Orders_P101458 where Fld_OrderId Like '%" & Txt_OrderId.Text & "%' and Fld_CustId Like '%" & Txt_CustId.Text & "%' and Fld_ProcessBy Like '%" & Cmb_ProcessBy.SelectedText & "%' and Fld_TotalPrice " & TotalPrice & " order by Fld_OrderId")



        Grd_Orders.Columns(0).HeaderText = "Order Id"
        Grd_Orders.Columns(1).HeaderText = "Customer Id"
        Grd_Orders.Columns(2).HeaderText = "Process By "
        Grd_Orders.Columns(3).HeaderText = "Total Price"
        Grd_Orders.Columns(4).HeaderText = "Purchased Date"




    End Sub




    Private Sub Pict_OrderClose_Click(sender As Object, e As EventArgs) Handles Pict_OrderClose.Click

        Me.Close()
        MsgBox("Wish you have a nice day ahead!")

    End Sub

    Private Sub Pict_OrderHome_Click(sender As Object, e As EventArgs) Handles Pict_OrderHome.Click


        Frm_MainMenu.Show()
        Me.Close()

    End Sub


    Private Sub Txt_CustId_TextChanged(sender As Object, e As EventArgs) Handles Txt_CustId.TextChanged, Txt_TotalPrice.TextChanged, Txt_OrderId.TextChanged

        Refresh_Grid()
    End Sub

    Private Sub Btn_MoreInfosPP_Click(sender As Object, e As EventArgs) Handles Btn_MoreInfosPP.Click

        Frm_ItemPurchasedRecord_P101458.Show()
        Me.Close()

    End Sub



    Private Sub Grd_Orders_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grd_Orders.CellClick

        Grd_PurchasedProductView.Visible = True
        Grd_PurchasedProductView.DataSource = run_sql_query("Select Tbl_Products_P101458.Fld_ClockId , Tbl_Products_P101458.Fld_ClockName From Tbl_ItemsPurchasedRecord_P101458 INNER JOIN Tbl_Products_P101458 On Tbl_Products_P101458.Fld_ClockId = Tbl_ItemsPurchasedRecord_P101458.Fld_ClockId WHERE Tbl_ItemsPurchasedRecord_P101458.Fld_OrderId='" & Grd_Orders.Rows(e.RowIndex).Cells(0).Value.ToString & "'")

        Grd_PurchasedProductView.Columns(0).HeaderText = "Clock Id"
        Grd_PurchasedProductView.Columns(1).HeaderText = "Clock Name"



    End Sub



    Private Sub Panel1_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel1.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint

        End If



    End Sub



End Class