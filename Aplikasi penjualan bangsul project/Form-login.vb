Imports System.Data.OleDb
Public Class Form_login

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("username atau password nya harus di isi")
        Else
            Call koneksi()
            cmd = New OleDbCommand("select * from tbl_user where username ='" & TextBox1.Text & "' and pwd='" & TextBox2.Text & "'", conn)
            rd = cmd.ExecuteReader
            rd.Read()
            If rd.HasRows = True Then
                Form_Menu.lbl_level.Text = rd.Item("lvl")
                Form_Menu.lbl_nama.Text = rd.Item("nama_user")
                Me.Hide()
                Form_Menu.Show()

            Else
                MessageBox.Show("password atau username salah!!!")
            End If

        End If
    End Sub

    Private Sub Form_login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Focus()
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Chr(13) Then
            Button1.Focus()

        End If
    End Sub
End Class