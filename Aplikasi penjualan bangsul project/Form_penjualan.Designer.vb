<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_penjualan
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_penjualan))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.v0grandtotal = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.AxCrystalReport1 = New AxCrystal.AxCrystalReport()
        Me.v0kasir = New System.Windows.Forms.TextBox()
        Me.v0tanggal = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.v0jam = New System.Windows.Forms.TextBox()
        Me.v0nofaktur = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.v0kembalian = New System.Windows.Forms.TextBox()
        Me.v0dibayar = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.v0kdbarang = New System.Windows.Forms.TextBox()
        Me.v0namabarang = New System.Windows.Forms.TextBox()
        Me.v0jenis = New System.Windows.Forms.TextBox()
        Me.v0satuan = New System.Windows.Forms.TextBox()
        Me.v0harga = New System.Windows.Forms.TextBox()
        Me.v0qty = New System.Windows.Forms.TextBox()
        Me.v0totalharga = New System.Windows.Forms.TextBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button4 = New System.Windows.Forms.Button()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.AxCrystalReport1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1378, 210)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.GroupBox2.Controls.Add(Me.v0grandtotal)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(588, 25)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(787, 182)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Grand Total"
        '
        'v0grandtotal
        '
        Me.v0grandtotal.BackColor = System.Drawing.Color.Black
        Me.v0grandtotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0grandtotal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.v0grandtotal.Font = New System.Drawing.Font("DS-Digital", 120.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0grandtotal.ForeColor = System.Drawing.Color.Lime
        Me.v0grandtotal.Location = New System.Drawing.Point(3, 25)
        Me.v0grandtotal.Multiline = True
        Me.v0grandtotal.Name = "v0grandtotal"
        Me.v0grandtotal.ReadOnly = True
        Me.v0grandtotal.Size = New System.Drawing.Size(781, 154)
        Me.v0grandtotal.TabIndex = 19
        Me.v0grandtotal.Text = "2.000.00"
        Me.v0grandtotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox3
        '
        Me.GroupBox3.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.GroupBox3.Controls.Add(Me.AxCrystalReport1)
        Me.GroupBox3.Controls.Add(Me.v0kasir)
        Me.GroupBox3.Controls.Add(Me.v0tanggal)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.v0jam)
        Me.GroupBox3.Controls.Add(Me.v0nofaktur)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.White
        Me.GroupBox3.Location = New System.Drawing.Point(3, 25)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(585, 182)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Transaksi"
        '
        'AxCrystalReport1
        '
        Me.AxCrystalReport1.Enabled = True
        Me.AxCrystalReport1.Location = New System.Drawing.Point(166, 25)
        Me.AxCrystalReport1.Name = "AxCrystalReport1"
        Me.AxCrystalReport1.OcxState = CType(resources.GetObject("AxCrystalReport1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxCrystalReport1.Size = New System.Drawing.Size(28, 28)
        Me.AxCrystalReport1.TabIndex = 17
        '
        'v0kasir
        '
        Me.v0kasir.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0kasir.Location = New System.Drawing.Point(297, 121)
        Me.v0kasir.Name = "v0kasir"
        Me.v0kasir.ReadOnly = True
        Me.v0kasir.Size = New System.Drawing.Size(260, 29)
        Me.v0kasir.TabIndex = 14
        Me.v0kasir.Text = "Chandra"
        '
        'v0tanggal
        '
        Me.v0tanggal.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0tanggal.Location = New System.Drawing.Point(297, 62)
        Me.v0tanggal.Name = "v0tanggal"
        Me.v0tanggal.ReadOnly = True
        Me.v0tanggal.Size = New System.Drawing.Size(260, 29)
        Me.v0tanggal.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(293, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 24)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Kasir"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(293, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 24)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Tanggal"
        '
        'v0jam
        '
        Me.v0jam.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0jam.Location = New System.Drawing.Point(13, 121)
        Me.v0jam.Name = "v0jam"
        Me.v0jam.ReadOnly = True
        Me.v0jam.Size = New System.Drawing.Size(260, 29)
        Me.v0jam.TabIndex = 14
        '
        'v0nofaktur
        '
        Me.v0nofaktur.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0nofaktur.Location = New System.Drawing.Point(13, 62)
        Me.v0nofaktur.Name = "v0nofaktur"
        Me.v0nofaktur.ReadOnly = True
        Me.v0nofaktur.Size = New System.Drawing.Size(260, 29)
        Me.v0nofaktur.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 93)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 24)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Jam"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 24)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "No Faktur"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.GroupBox4.Controls.Add(Me.Button3)
        Me.GroupBox4.Controls.Add(Me.Button2)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.v0kembalian)
        Me.GroupBox4.Controls.Add(Me.v0dibayar)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.Color.White
        Me.GroupBox4.Location = New System.Drawing.Point(0, 637)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1378, 122)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Pembayaran"
        '
        'Button3
        '
        Me.Button3.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.White
        Me.Button3.Location = New System.Drawing.Point(1079, 50)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(126, 49)
        Me.Button3.TabIndex = 30
        Me.Button3.Text = "Close"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Location = New System.Drawing.Point(943, 50)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(126, 49)
        Me.Button2.TabIndex = 29
        Me.Button2.Text = "Data Baru"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Warna
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(808, 51)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(126, 49)
        Me.Button1.TabIndex = 28
        Me.Button1.Text = "Simpan"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'v0kembalian
        '
        Me.v0kembalian.BackColor = System.Drawing.Color.Black
        Me.v0kembalian.Font = New System.Drawing.Font("DS-Digital", 60.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0kembalian.ForeColor = System.Drawing.Color.Yellow
        Me.v0kembalian.Location = New System.Drawing.Point(403, 35)
        Me.v0kembalian.Name = "v0kembalian"
        Me.v0kembalian.ReadOnly = True
        Me.v0kembalian.Size = New System.Drawing.Size(388, 87)
        Me.v0kembalian.TabIndex = 18
        Me.v0kembalian.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'v0dibayar
        '
        Me.v0dibayar.BackColor = System.Drawing.Color.Black
        Me.v0dibayar.Font = New System.Drawing.Font("DS-Digital", 60.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0dibayar.ForeColor = System.Drawing.Color.Lime
        Me.v0dibayar.Location = New System.Drawing.Point(137, 35)
        Me.v0dibayar.Name = "v0dibayar"
        Me.v0dibayar.Size = New System.Drawing.Size(260, 87)
        Me.v0dibayar.TabIndex = 2
        Me.v0dibayar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(399, 14)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(108, 24)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Kembalian"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(133, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 24)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Dibayar"
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column7})
        Me.DataGridView1.Location = New System.Drawing.Point(0, 248)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1383, 383)
        Me.DataGridView1.TabIndex = 3
        '
        'Column1
        '
        Me.Column1.HeaderText = "Kode Barang"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 200
        '
        'Column2
        '
        Me.Column2.HeaderText = "Nama Barang"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 350
        '
        'Column3
        '
        Me.Column3.HeaderText = "Jenis Barang"
        Me.Column3.Name = "Column3"
        Me.Column3.Width = 200
        '
        'Column4
        '
        Me.Column4.HeaderText = "Satuan"
        Me.Column4.Name = "Column4"
        '
        'Column5
        '
        Me.Column5.HeaderText = "Harga"
        Me.Column5.Name = "Column5"
        Me.Column5.Width = 200
        '
        'Column6
        '
        Me.Column6.HeaderText = "QTY"
        Me.Column6.Name = "Column6"
        '
        'Column7
        '
        Me.Column7.HeaderText = "Total Harga"
        Me.Column7.Name = "Column7"
        Me.Column7.Width = 200
        '
        'v0kdbarang
        '
        Me.v0kdbarang.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0kdbarang.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0kdbarang.Location = New System.Drawing.Point(43, 213)
        Me.v0kdbarang.Name = "v0kdbarang"
        Me.v0kdbarang.Size = New System.Drawing.Size(202, 29)
        Me.v0kdbarang.TabIndex = 1
        '
        'v0namabarang
        '
        Me.v0namabarang.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0namabarang.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0namabarang.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0namabarang.Location = New System.Drawing.Point(242, 213)
        Me.v0namabarang.Name = "v0namabarang"
        Me.v0namabarang.ReadOnly = True
        Me.v0namabarang.Size = New System.Drawing.Size(349, 29)
        Me.v0namabarang.TabIndex = 99
        '
        'v0jenis
        '
        Me.v0jenis.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0jenis.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0jenis.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0jenis.Location = New System.Drawing.Point(592, 213)
        Me.v0jenis.Name = "v0jenis"
        Me.v0jenis.ReadOnly = True
        Me.v0jenis.Size = New System.Drawing.Size(199, 29)
        Me.v0jenis.TabIndex = 22
        '
        'v0satuan
        '
        Me.v0satuan.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0satuan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0satuan.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0satuan.Location = New System.Drawing.Point(792, 213)
        Me.v0satuan.Name = "v0satuan"
        Me.v0satuan.ReadOnly = True
        Me.v0satuan.Size = New System.Drawing.Size(101, 29)
        Me.v0satuan.TabIndex = 20
        '
        'v0harga
        '
        Me.v0harga.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0harga.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0harga.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0harga.Location = New System.Drawing.Point(894, 213)
        Me.v0harga.Name = "v0harga"
        Me.v0harga.ReadOnly = True
        Me.v0harga.Size = New System.Drawing.Size(198, 29)
        Me.v0harga.TabIndex = 21
        Me.v0harga.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'v0qty
        '
        Me.v0qty.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0qty.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0qty.Location = New System.Drawing.Point(1093, 213)
        Me.v0qty.Name = "v0qty"
        Me.v0qty.Size = New System.Drawing.Size(99, 29)
        Me.v0qty.TabIndex = 19
        Me.v0qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'v0totalharga
        '
        Me.v0totalharga.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.v0totalharga.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.v0totalharga.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.v0totalharga.Location = New System.Drawing.Point(1193, 213)
        Me.v0totalharga.Name = "v0totalharga"
        Me.v0totalharga.ReadOnly = True
        Me.v0totalharga.Size = New System.Drawing.Size(190, 29)
        Me.v0totalharga.TabIndex = 23
        Me.v0totalharga.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Button4
        '
        Me.Button4.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Button4.Location = New System.Drawing.Point(-3, 213)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(40, 29)
        Me.Button4.TabIndex = 17
        Me.Button4.Text = "Cari"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'PrintDocument1
        '
        '
        'Form_penjualan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackgroundImage = Global.Aplikasi_penjualan_bangsul_project.My.Resources.Resources.Background
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1378, 759)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.v0totalharga)
        Me.Controls.Add(Me.v0qty)
        Me.Controls.Add(Me.v0harga)
        Me.Controls.Add(Me.v0satuan)
        Me.Controls.Add(Me.v0jenis)
        Me.Controls.Add(Me.v0namabarang)
        Me.Controls.Add(Me.v0kdbarang)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form_penjualan"
        Me.ShowInTaskbar = False
        Me.Text = "Form_penjualan"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.AxCrystalReport1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents v0kasir As System.Windows.Forms.TextBox
    Friend WithEvents v0tanggal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents v0jam As System.Windows.Forms.TextBox
    Friend WithEvents v0nofaktur As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents v0kembalian As System.Windows.Forms.TextBox
    Friend WithEvents v0dibayar As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents v0grandtotal As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents v0kdbarang As System.Windows.Forms.TextBox
    Friend WithEvents v0namabarang As System.Windows.Forms.TextBox
    Friend WithEvents v0jenis As System.Windows.Forms.TextBox
    Friend WithEvents v0satuan As System.Windows.Forms.TextBox
    Friend WithEvents v0harga As System.Windows.Forms.TextBox
    Friend WithEvents v0qty As System.Windows.Forms.TextBox
    Friend WithEvents v0totalharga As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents AxCrystalReport1 As AxCrystal.AxCrystalReport
End Class
