Imports System.Data.OleDb
Public Class Form_user
    Sub tampil_user()

        cmd = New OleDbCommand("select * from tbl_user", conn)
        rd = cmd.ExecuteReader
        DataGridView1.Rows.Clear()
        Do While rd.Read = True
            DataGridView1.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4))
        Loop
    End Sub
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Form_user_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call tampil_user()
        Call bersih()



    End Sub

    Sub bersih()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Text = ""
        TextBox1.Focus()



    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MessageBox.Show("Semua Data Wajib Di isi!")
        Else
            'melakukan pengecekan data ke database apakah kode user sudah ada
            cmd = New OleDbCommand("select * from tbl_user where kode_user='" & TextBox1.Text & "'", conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = False Then
                'melakukan penyimpanan data ke database
                Call koneksi()

                cmd = New OleDbCommand("insert into tbl_user values('" & TextBox1.Text & _
                                     "','" & TextBox2.Text & _
                                     "','" & TextBox3.Text & _
                                     "','" & TextBox4.Text & _
                                     "','" & ComboBox1.Text & _
                                     "')", conn)


                cmd.ExecuteNonQuery()
                MessageBox.Show("Data berhasil Di di tambahkan !!")
                Call bersih()
                Call tampil_user()
            Else
                '------ melakukan edit data berdasarkan kode user yang dipanggil
                Call koneksi()
                cmd = New OleDbCommand("update tbl_user set nama_user = '" & TextBox2.Text & _
                                       "',username ='" & TextBox3.Text & _
                                       "',pwd ='" & TextBox4.Text & _
                                       "',lvl ='" & ComboBox1.Text & _
                                        "'where kode_user='" & TextBox1.Text & "'", conn)

                cmd.ExecuteNonQuery()
                MessageBox.Show("Data Berhasil Di di ubah!!!")
                Call bersih()
                Call tampil_user()
            End If

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call bersih()

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        '-----memanggil data berdasarkan kode
        cmd = New OleDbCommand("select * from tbl_user where kode_user='" & TextBox1.Text & "'", conn)
        rd = cmd.ExecuteReader
        rd.Read()
        If rd.HasRows = True Then
            TextBox2.Text = rd(1)
            TextBox3.Text = rd(2)
            TextBox4.Text = rd(3)
            ComboBox1.Text = rd(4)

            
        Else
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            ComboBox1.Text = ""


        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MessageBox.Show("Apakah Data Akan Dihapus....??", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            cmd = New OleDbCommand("delete from tbl_user where kode_user = '" & TextBox1.Text & "'", conn)
            cmd.ExecuteNonQuery()
            MessageBox.Show("Data berhasil Didelet !!")
            Call bersih()
            Call tampil_user()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class