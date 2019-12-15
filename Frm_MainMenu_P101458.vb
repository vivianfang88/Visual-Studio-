Public Class Frm_MainMenu

    Dim newpoint As New System.Drawing.Point
    Dim X, Y As Integer


    Private Sub Pnl_Window_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_Staffs.MouseDown

        X = Control.MousePosition.X - Me.Location.X
        Y = Control.MousePosition.Y - Me.Location.Y

    End Sub

    Private Sub Pnl_Window_MouseMove(sender As Object, e As MouseEventArgs) Handles Btn_Staffs.MouseMove

        If e.Button = Windows.Forms.MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= (X)
            newpoint.Y -= (Y)
            Me.Location = newpoint


        End If

    End Sub

    Private Sub Btn_Products_Click(sender As Object, e As EventArgs) Handles Btn_MMProducts.Click


        Frm_Products_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Btn_MMStaffs.Click


        Frm_Staffs_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Btn_MMCustomer_Click(sender As Object, e As EventArgs) Handles Btn_MMCustomer.Click


        Frm_Customers_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Btn_MMOrder_Click(sender As Object, e As EventArgs) Handles Btn_MMOrder.Click


        Frm_Orders_P101458.Show()
        Me.Close()

    End Sub

    Private Sub Btn_MMItemsPR_Click(sender As Object, e As EventArgs) Handles Btn_MMItemsPR.Click


        Frm_ItemPurchasedRecord_P101458.Show()
        Me.Close()

    End Sub



    Private Sub Pict_Close_MouseClick(sender As Object, e As MouseEventArgs) Handles Pict_Close.MouseClick

        Me.Close()
        MsgBox("Wish you have a nice day ahead!",, "ByeBye")


    End Sub
End Class
