Imports System.Data.OleDb
Imports System.Drawing.Printing

Public Class Form_penjualan
    Dim WithEvents PD As New PrintDocument
    Dim PPD As New PrintPreviewDialog
    Dim t_qty As Long
    Dim panjang As Integer

    Sub no_faktur()
        cmd = New OleDbCommand("select * from tbl_jual where faktur_jual in(select max(faktur_jual)from tbl_jual)order by faktur_jual DESC", conn)
        rd = cmd.ExecuteReader
        rd.Read()
        If Not rd.HasRows Then
            v0nofaktur.Text = Format(Now, "yyMMdd") + "0001"
        Else
            If Microsoft.VisualBasic.Left(rd.GetString(0), 6) <> Format(Now, "yyMMdd") Then
                v0nofaktur.Text = Format(Now, "yyMMdd") + "0001"
            Else
                v0nofaktur.Text = rd.Item("faktur_jual") + 1

            End If
        End If
    End Sub



    Sub grand_total()
        Dim jumlah As Decimal = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            jumlah = jumlah + DataGridView1.Rows(i).Cells(6).Value
            v0grandtotal.Text = jumlah
        Next
        If v0grandtotal.Text = "" Then
            v0grandtotal.Text = "0"

        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form_penjualan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call no_faktur()
        v0kdbarang.Focus()
        v0kembalian.Text = "0"
        Call grand_total()
        v0kasir.Text = Form_Menu.lbl_nama.Text





    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        v0tanggal.Text = Format(Now, "dd/MM/yyyy")
        v0jam.Text = Format(Now, "HH:mm:ss")
    End Sub

    Private Sub v0namabarang_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles v0namabarang.TextChanged

    End Sub

    Private Sub v0kdbarang_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles v0kdbarang.KeyPress
        If e.KeyChar = Chr(13) Then
            cmd = New OleDbCommand("select * from tbl_barang where kode_barang='" & v0kdbarang.Text & "'", conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = True Then
                v0namabarang.Text = rd.Item("nama_barang")
                v0jenis.Text = rd.Item("jenis_barang")
                v0satuan.Text = rd.Item("satuan_barang")
                v0namabarang.Text = rd.Item("nama_barang")
                v0harga.Text = rd.Item("harga_jual")
                v0qty.Focus()

            Else
                v0namabarang.Text = ""
                v0jenis.Text = ""
                v0satuan.Text = ""
                v0harga.Text = ""
                v0kdbarang.Focus()
                MessageBox.Show("BARANG INI TIDAK DITEMUKAN / TIDAK TERDAFTAR DI DATABASE")
            End If
        End If
    End Sub
    Sub bersih_barang()
        v0kdbarang.Text = ""
        v0namabarang.Text = ""
        v0jenis.Text = ""
        v0satuan.Text = ""
        v0harga.Text = ""
        v0qty.Text = ""
        v0totalharga.Text = ""
        v0kdbarang.Focus()
    End Sub

    Private Sub v0qty_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles v0qty.TextChanged
        Try
            v0totalharga.Text = Val(v0harga.Text) * Val(v0qty.Text)


        Catch ex As Exception
            v0totalharga.Text = ""
        End Try
    End Sub

    Private Sub v0qty_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles v0qty.KeyPress
        If e.KeyChar = Chr(13) Then
            DataGridView1.Rows.Add(v0kdbarang.Text, v0namabarang.Text, v0jenis.Text, v0satuan.Text, v0harga.Text, v0qty.Text, v0totalharga.Text)
            Call bersih_barang()
            Call grand_total()
        End If
    End Sub

    Private Sub v0dibayar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles v0dibayar.TextChanged
        Try
            v0kembalian.Text = Val(v0dibayar.Text) - Val(v0grandtotal.Text)

        Catch ex As Exception
            v0kembalian.Text = "0"

        End Try


    End Sub

    Private Sub v0nofaktur_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles v0nofaktur.TextChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ubahpanjang()
        PPD.Document = PD
        PPD.ShowDialog()

        If v0grandtotal.Text = "0" Then
            MessageBox.Show("Belum Ada Barang Yang Di input")
        ElseIf Val(v0dibayar.Text) < Val(v0grandtotal.Text) Then
            MessageBox.Show("Pembayaran Kurang")
        Else
            '===simpan data ke tabel barang====
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
            '=====simpan rincian barang dari datagridview ke tbl_rinci_jual========
            For baris As Integer = 0 To DataGridView1.Rows.Count - 2
                cmd = New OleDbCommand("insert into tbl_rinci_jual values('" & v0nofaktur.Text & _
                                       "','" & DataGridView1.Rows(baris).Cells(0).Value & _
                                       "','" & DataGridView1.Rows(baris).Cells(5).Value & _
                                       "','" & DataGridView1.Rows(baris).Cells(6).Value & _
                                       "')", conn)
                cmd.ExecuteNonQuery()

                '====pengurangan stok====
                cmd = New OleDbCommand("select * from tbl_barang where kode_barang='" & DataGridView1.Rows(baris).Cells(0).Value & "'", conn)
                rd = cmd.ExecuteReader
                rd.Read()
                If rd.HasRows = True Then
                    cmd = New OleDbCommand("update tbl_barang set stok ='" & rd.Item("stok") - Val(DataGridView1.Rows(baris).Cells(5).Value) & _
                                           "'where kode_barang='" & DataGridView1.Rows(baris).Cells(0).Value & "'", conn)
                    cmd.ExecuteNonQuery()

                End If

            Next
            'If MessageBox.Show("Cetak Faktur..?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

            'AxCrystalReport1.SelectionFormula = "totext({tbl_jual.faktur_jual})= '" & v0nofaktur.Text & "'"
            'AxCrystalReport1.ReportFileName = "cetakfaktur.rpt"
            'AxCrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
            'AxCrystalReport1.Action = 1


            'Else



            '=====membersihkan data transaksi=====
            MessageBox.Show("Transaksi penjulan berhasil di simpan")
            DataGridView1.Rows.Clear()
            v0grandtotal.Text = "0"
            v0dibayar.Text = ""
            v0kembalian.Text = "0"
            Call no_faktur()
            Call bersih_barang()

        End If
        'End If



    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        If e.KeyChar = Chr(27) Then
            Dim baris As Integer
            baris = DataGridView1.CurrentCell.RowIndex
            Try
                DataGridView1.Rows.RemoveAt(baris)
                Call grand_total()
            Catch ex As Exception

            End Try
        End If



    End Sub

    Private Sub v0kdbarang_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles v0kdbarang.TextChanged
        cmd = New OleDbCommand("select * from tbl_barang where kode_barang like'%" & v0kdbarang.Text & "%'", conn)
        rd = cmd.ExecuteReader



    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Form_Barang.ShowDialog()
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.Rows.Clear()
        v0grandtotal.Text = "0"
        v0dibayar.Text = ""
        v0kembalian.Text = "0"
        Call no_faktur()
        Call bersih_barang()
    End Sub

    Private Sub v0dibayar_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles v0dibayar.KeyPress
        If e.KeyChar = Chr(13) Then
            Button1.Focus()

        End If
    End Sub

    Private Sub Button4_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FormCariBarang.ShowDialog()

    End Sub

    Private Sub PD_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PD.BeginPrint
        Dim PageSetup As New PageSettings
        PageSetup.PaperSize = New PaperSize("custom", 300, panjang)
        PD.DefaultPageSettings = PageSetup

    End Sub

    Private Sub PD_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PD.PrintPage
        Dim f10 As New Font("Times New Roman", 10, FontStyle.Regular)
        Dim f10b As New Font("Times New Roman", 10, FontStyle.Bold)
        Dim f14 As New Font("Times New Roman", 14, FontStyle.Bold)

        Dim leftmargin As Integer = PD.DefaultPageSettings.Margins.Left
        Dim centermargin As Integer = PD.DefaultPageSettings.PaperSize.Width / 2
        Dim rightmargin As Integer = PD.DefaultPageSettings.Margins.Right

        Dim kanan As New StringFormat
        Dim tengah As New StringFormat
        kanan.Alignment = StringAlignment.Far
        tengah.Alignment = StringAlignment.Center

        Dim garis As String
        garis = "----------------------------------------------------------------"

        e.Graphics.DrawString("TOKO CHANDRA BERKAH", f14, Brushes.Black, centermargin, 5, tengah)
        e.Graphics.DrawString("JL. Padang Painan Km.51 Koto xi tarusan", f10, Brushes.Black, centermargin, 25, tengah)

        e.Graphics.DrawString("No Faktur", f10, Brushes.Black, 0, 60)
        e.Graphics.DrawString(":", f10, Brushes.Black, 65, 60)
        e.Graphics.DrawString("" & v0nofaktur.Text & "", f10, Brushes.Black, 75, 60)

        e.Graphics.DrawString("" & v0tanggal.Text & " " & v0jam.Text & "", f10, Brushes.Black, 0, 75)

        e.Graphics.DrawString(garis, f10, Brushes.Black, 0, 90)
        e.Graphics.DrawString("Nama Barang", f10, Brushes.Black, 0, 100)
        e.Graphics.DrawString("Qty", f10, Brushes.Black, 135, 100)
        e.Graphics.DrawString("Total Harga", f10, Brushes.Black, rightmargin + 122, 100)
        e.Graphics.DrawString(garis, f10, Brushes.Black, 0, 108)
        DataGridView1.AllowUserToAddRows = False

        Dim tinggi As Integer
        Dim i As Long
        For baris As Integer = 0 To DataGridView1.RowCount - 1
            tinggi += 15
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(1).Value.ToString, f10, Brushes.Black, 0, 105 + tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(5).Value.ToString, f10, Brushes.Black, 137, 105 + tinggi)

            i = DataGridView1.Rows(baris).Cells(6).Value
            DataGridView1.Rows(baris).Cells(6).Value = Format(i, "##,##0")
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(6).Value.ToString, f10, Brushes.Black, rightmargin + 135, 105 + tinggi)
        Next
        tinggi = 110 + tinggi
        hitung()
        e.Graphics.DrawString(garis, f10, Brushes.Black, 0, 5 + tinggi)

        e.Graphics.DrawString("Total Bayar: " & v0grandtotal.Text & "", f10b, Brushes.Black, rightmargin + 55, 15 + tinggi)
        e.Graphics.DrawString("Jumlah :", f10b, Brushes.Black, 0, 15 + tinggi)
        e.Graphics.DrawString(t_qty, f10b, Brushes.Black, 65, 15 + tinggi)

        e.Graphics.DrawString("~Terima Kasih Telah Belanja~", f10, Brushes.Black, centermargin, 65 + tinggi, tengah)
        e.Graphics.DrawString("~Di Toko Kami~", f10, Brushes.Black, centermargin, 80 + tinggi, tengah)
    End Sub

    Private Sub v0tanggal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles v0tanggal.TextChanged

    End Sub
    Sub hitung()
        Dim hitungqty As Long = 0
        For baris As Long = 0 To DataGridView1.RowCount - 1
            hitungqty = hitungqty + DataGridView1.Rows(baris).Cells(5).Value

        Next
        t_qty = hitungqty

    End Sub
    Sub ubahpanjang()
        Dim rowcount As Integer
        panjang = 0

        rowcount = DataGridView1.Rows.Count
        panjang = rowcount * 15
        panjang = panjang + 220

    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub AxCrystalReport1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxCrystalReport1.Enter

    End Sub

End Class