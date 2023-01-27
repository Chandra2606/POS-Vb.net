Imports System.Data.OleDb

Public Class Form_Barang_Masuk

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtsup.TextChanged

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If cbjenisbayar.Text = "KREDIT" Then
            Call koneksi()
            cmd = New OleDbCommand("insert into tbl_jual(faktur_jual,tgl_jual,jam,grand_total,dibayar,kembalian,kasir)values('" & v0nofaktur.Text & _
                                   "','" & v0tanggal.Text & _
                                   "','" & v0jam.Text & _
                                   "','" & v0grandtotal.Text & _
                                   "','" & v0dibayar.Text & _
                                   "','" & v0kembalian.Text & _
                                   "','" & v0kasir.Text & _
                                   "')", conn)


            cmd.ExecuteNonQuery()

        End If
    End Sub
End Class