Imports System.Data.OleDb

Public Class FormCariBarang
    Sub tampil_barang()
        cmd = New OleDbCommand("select * from tbl_barang", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))
        Loop
    End Sub
    Private Sub FormCariBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampil_barang()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Form_penjualan.v0kdbarang.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        Form_penjualan.v0namabarang.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
        Form_penjualan.v0jenis.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
        Form_penjualan.v0satuan.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
        Form_penjualan.v0harga.Text = DataGridView1.Rows(e.RowIndex).Cells(5).Value
        Me.Close()
        Form_penjualan.v0qty.Focus()



    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        cmd = New OleDbCommand("select * from tbl_barang where nama_barang like'%" & TextBox1.Text & "%'", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))
        Loop
    End Sub
End Class