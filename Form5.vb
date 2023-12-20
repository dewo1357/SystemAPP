Imports System.Data.OleDb
Public Class Form5

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        Dim Background As String = "C:\Users\Acer\OneDrive\Documents\Visual Studio 2010\Projects\ConsoleApplication1\bin\Debug\WindowsApplication3\WindowsApplication3\background.JPG"


        'Background
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        If IO.File.Exists(Background) Then
            Me.BackgroundImage = Image.FromFile(Background)
        Else
            MsgBox("File Tidak Ditemukan")
        End If

        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.Text = "PENJUALAN"
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        ' Add any initialization after the InitializeComponent() call.
        Button()
        kolom()
        TextBox()
        Label()
        group()


    End Sub

    '============GroupBox============='
    Private groupbox As New GroupBox()
    Private groupbox1 As New GroupBox()
    Private groupbox2 As New GroupBox()
    Private groupbox3 As New GroupBox()
    Private groupbox4 As New GroupBox()
    Private groupbox5 As New GroupBox()

    Sub group()

        '1
        groupbox1.Location = New Point(5, 30)
        groupbox1.Size = New Size(500, 150)
        groupbox1.BackColor = Color.DarkGray
        groupbox1.Text = "INFO TRANSAKSI"
        groupbox1.Font = New Font("Helvetiva", 20, FontStyle.Bold)
        groupbox1.ForeColor = Color.White
        Me.Controls.Add(groupbox1)

        '2
        groupbox2.Location = New Point(510, 30)
        groupbox2.Size = New Size(850, 150)
        groupbox2.BackColor = Color.DarkGray
        groupbox2.Font = New Font("Helvetiva", 20, FontStyle.Bold)
        groupbox2.ForeColor = Color.White
        groupbox2.Text = "Total Pembayaran"
        Me.Controls.Add(groupbox2)

        '3
        groupbox.Size = New Size(1500, 180)
        groupbox.Location = New Point(0, 0)
        groupbox.BackColor = Color.DarkGray
        groupbox.Dock = DockStyle.None
        groupbox.Font = New Font("Helvetiva", 20, FontStyle.Bold)
        groupbox.ForeColor = Color.White
        Me.Controls.Add(groupbox)

        groupbox3.Size = New Size(1000, 183)
        groupbox3.Location = New Point(1200, 800)
        groupbox3.BackColor = Color.DarkGray
        groupbox3.Dock = DockStyle.Bottom
        groupbox3.Text = "PEMBAYARAN"
        groupbox3.Font = New Font("Helvetiva", 28, FontStyle.Bold)
        groupbox3.ForeColor = Color.White
        Me.Controls.Add(groupbox3)

        groupbox4.Size = New Size(730, 180)
        groupbox4.Location = New Point(0, 180)
        groupbox4.BackColor = Color.DarkGray
        groupbox4.Text = "INPUT BARANG"
        groupbox4.Font = New Font("Helvetiva", 23, FontStyle.Bold)
        groupbox4.ForeColor = Color.White
        Me.Controls.Add(groupbox4)

        groupbox5.Size = New Size(700, 180)
        groupbox5.Location = New Point(724, 180)
        groupbox5.BackColor = Color.DarkGray
        groupbox5.Text = "INFORMASI"
        groupbox5.Font = New Font("Helvetiva", 23, FontStyle.Bold)
        groupbox5.ForeColor = Color.White
        Me.Controls.Add(groupbox5)
    End Sub

    '============LABEL============='
    Private Label1 As New Label()
    Private Label2 As New Label()
    Private Label3 As New Label()
    Private Label4 As New Label()
    Private Label5 As New Label()
    Private Label6 As New Label()
    Private Label7 As New Label()
    Private Label8 As New Label()
    Private Label9 As New Label()
    Private Label10 As New Label()
    Private Label11 As New Label()
    Private Label12 As New Label()
    Private Label13 As New Label()
    Private Label14 As New Label()


    Sub Label()
        Label1.Text = "No Transaksi"
        Label1.Size = New Size(150, 50)
        Label1.BackColor = Color.DarkGray
        Label1.Location = New Point(20, 70)
        Label1.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label1.ForeColor = Color.White
        Me.Controls.Add(Label1)

        Label2.Text = "Waktu"
        Label2.Size = New Size(150, 50)
        Label2.BackColor = Color.DarkGray
        Label2.Location = New Point(20, 125)
        Label2.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label2.ForeColor = Color.White
        Me.Controls.Add(Label2)

        Label3.Text = "Tanggal"
        Label3.Size = New Size(150, 50)
        Label3.BackColor = Color.DarkGray
        Label3.Location = New Point(300, 70)
        Label3.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label3.ForeColor = Color.White
        Me.Controls.Add(Label3)

        Label4.Text = "Kasir"
        Label4.Size = New Size(150, 50)
        Label4.BackColor = Color.DarkGray
        Label4.Location = New Point(302, 125)
        Label4.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label4.ForeColor = Color.White
        Me.Controls.Add(Label4)

        Label5.Text = "DIBAYAR"
        Label5.Size = New Size(150, 50)
        Label5.BackColor = Color.DarkGray
        Label5.Location = New Point(15, 630)
        Label5.Font = New Font("Helvetica", 15, FontStyle.Bold)
        Label5.ForeColor = Color.White
        Me.Controls.Add(Label5)

        Label6.Text = "KEMBALIAN"
        Label6.Size = New Size(150, 50)
        Label6.BackColor = Color.DarkGray
        Label6.Location = New Point(397, 630)
        Label6.Font = New Font("Helvetica", 15, FontStyle.Bold)
        Label6.ForeColor = Color.White
        Me.Controls.Add(Label6)

        Label7.Text = "Kode Barang"
        Label7.Size = New Size(150, 50)
        Label7.BackColor = Color.DarkGray
        Label7.Location = New Point(5, 230)
        Label7.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label7.ForeColor = Color.White
        Me.Controls.Add(Label7)

        Label8.Text = "Nama Barang"
        Label8.Size = New Size(150, 50)
        Label8.BackColor = Color.DarkGray
        Label8.Location = New Point(5, 295)
        Label8.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label8.ForeColor = Color.White
        Me.Controls.Add(Label8)

        Label9.Text = "Jenis Barang"
        Label9.Size = New Size(150, 50)
        Label9.BackColor = Color.DarkGray
        Label9.Location = New Point(250, 230)
        Label9.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label9.ForeColor = Color.White
        Me.Controls.Add(Label9)

        Label10.Text = "Satuan Barang"
        Label10.Size = New Size(150, 50)
        Label10.BackColor = Color.DarkGray
        Label10.Location = New Point(250, 295)
        Label10.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label10.ForeColor = Color.White
        Me.Controls.Add(Label10)

        Label11.Text = "Harga"
        Label11.Size = New Size(150, 50)
        Label11.BackColor = Color.DarkGray
        Label11.Location = New Point(490, 230)
        Label11.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label11.ForeColor = Color.White
        Me.Controls.Add(Label11)

        Label12.Text = "Quantity"
        Label12.Size = New Size(150, 50)
        Label12.BackColor = Color.DarkGray
        Label12.Location = New Point(490, 295)
        Label12.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label12.ForeColor = Color.White
        Me.Controls.Add(Label12)

        Label13.Text = "Total Harga Produk"
        Label13.Size = New Size(300, 50)
        Label13.BackColor = Color.DarkGray
        Label13.Location = New Point(866, 230)
        Label13.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label13.ForeColor = Color.White
        Me.Controls.Add(Label13)

        Label14.Text = "Urutan "
        Label14.Size = New Size(300, 50)
        Label14.BackColor = Color.DarkGray
        Label14.Location = New Point(739, 230)
        Label14.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Label14.ForeColor = Color.White
        Me.Controls.Add(Label14)
    End Sub

    '============TextBox============='
    Public TextBox1 As New TextBox()
    Public TextBox2 As New TextBox()
    Public TextBox3 As New TextBox()
    Public TextBox4 As New ComboBox()
    Public TextBox5 As New TextBox()
    Private TextBox6 As New TextBox()
    Public TextBox7 As New TextBox()
    Public TextBox8 As New TextBox()
    Private TextBox9 As New TextBox()
    Private TextBox10 As New ComboBox()
    Private TextBox11 As New ComboBox()
    Private TextBox12 As New TextBox()
    Private TextBox13 As New TextBox()
    Public TextBox14 As New TextBox()
    Private TextBox15 As New TextBox()

    Public Sub hapusdata()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox5.Text = 0
        TextBox6.Text = 0
        TextBox7.Text = 0
        TextBox8.Clear()
        TextBox10.SelectedIndex = 1
        TextBox11.SelectedIndex = 1
        TextBox12.Text = 0
        TextBox13.Text = 0
        TextBox14.Text = 0
        TextBox15.Text = 0
        Call kolom()
    End Sub

    Sub TextBox()
        TextBox1.Location = New Point(25, 95)
        TextBox1.Size = New Size(190, 70)
        TextBox1.BorderStyle = BorderStyle.None
        TextBox1.Font = New Font("Sogoe UI", 11, FontStyle.Bold)
        Me.Controls.Add(TextBox1)

        TextBox2.Location = New Point(25, 150)
        TextBox2.Size = New Size(190, 70)
        TextBox2.BorderStyle = BorderStyle.None
        TextBox2.Font = New Font("Sogoe UI", 11, FontStyle.Bold)
        Me.Controls.Add(TextBox2)

        TextBox3.Location = New Point(303, 95)
        TextBox3.Size = New Size(190, 70)
        TextBox3.BorderStyle = BorderStyle.None
        TextBox3.Font = New Font("Sogoe UI", 11, FontStyle.Bold)
        Me.Controls.Add(TextBox3)

        TextBox4.Location = New Point(303, 150)
        TextBox4.Size = New Size(190, 70)
        TextBox4.Font = New Font("Sogoe UI", 7, FontStyle.Bold)
        TextBox4.Items.Add("Sadewo Widyanto")
        TextBox4.Items.Add("Audia Prastiwi")
        TextBox4.Items.Add("Muhammad Hafiz Syahputra")
        TextBox4.SelectedIndex = 0
        Me.Controls.Add(TextBox4)

        TextBox5.Location = New Point(400, 670)
        TextBox5.Size = New Size(380, 70)
        TextBox5.BorderStyle = BorderStyle.None
        TextBox5.Font = New Font("DS-Digital", 47, FontStyle.Bold)
        TextBox5.TextAlign = HorizontalAlignment.Right
        TextBox5.Text = 0
        Me.Controls.Add(TextBox5)

        TextBox6.Location = New Point(20, 670)
        TextBox6.Size = New Size(370, 250)
        TextBox6.BorderStyle = BorderStyle.None
        TextBox6.Font = New Font("DS-Digital", 47, FontStyle.Bold)
        TextBox6.TextAlign = HorizontalAlignment.Right
        TextBox6.Text = 0
        Me.Controls.Add(TextBox6)

        TextBox7.Location = New Point(520, 65)
        TextBox7.Size = New Size(830, 100)
        TextBox7.BorderStyle = BorderStyle.None
        TextBox7.Font = New Font("DS-Digital", 70, FontStyle.Bold)
        TextBox7.Multiline = True
        TextBox7.TextAlign = HorizontalAlignment.Right
        TextBox7.Text = 0
        Me.Controls.Add(TextBox7)

        TextBox8.Location = New Point(10, 255)
        TextBox8.Size = New Size(220, 190)
        TextBox8.BorderStyle = BorderStyle.None
        TextBox8.Font = New Font("Sogoe UI", 18, FontStyle.Bold)
        Me.Controls.Add(TextBox8)
        AddHandler TextBox8.TextChanged, AddressOf kode_id


        TextBox9.Location = New Point(10, 320)
        TextBox9.Size = New Size(220, 200)
        TextBox9.BorderStyle = BorderStyle.None
        TextBox9.Font = New Font("Sogoe UI", 18, FontStyle.Bold)
        TextBox9.Text = ""
        Me.Controls.Add(TextBox9)

        TextBox10.Location = New Point(250, 255)
        TextBox10.Size = New Size(220, 200)
        TextBox10.Font = New Font("Sogoe UI", 13, FontStyle.Bold)
        TextBox10.Items.Add("Makanan")
        TextBox10.Items.Add("Alat Elektronik")
        TextBox10.Items.Add("Minuman")
        TextBox10.Items.Add("Furniture")
        TextBox10.SelectedIndex = 0
        Me.Controls.Add(TextBox10)

        TextBox11.Location = New Point(250, 320)
        TextBox11.Size = New Size(220, 200)
        TextBox11.Items.Add("Unit")
        TextBox11.Items.Add("Pcs")
        TextBox11.Items.Add("Bungkus")
        TextBox11.SelectedIndex = 0
        TextBox11.Font = New Font("Sogoe UI", 13, FontStyle.Bold)
        Me.Controls.Add(TextBox11)


        TextBox12.Location = New Point(490, 255)
        TextBox12.Size = New Size(220, 180)
        TextBox12.BorderStyle = BorderStyle.None
        TextBox12.Font = New Font("DS-Digital", 20, FontStyle.Bold)
        TextBox12.TextAlign = HorizontalAlignment.Right
        TextBox12.Text = 0
        Me.Controls.Add(TextBox12)


        TextBox13.Location = New Point(490, 320)
        TextBox13.Size = New Size(220, 180)
        TextBox13.BorderStyle = BorderStyle.None
        TextBox13.Font = New Font("DS-Digital", 20, FontStyle.Bold)
        TextBox13.TextAlign = HorizontalAlignment.Right
        Me.Controls.Add(TextBox13)

        TextBox14.Location = New Point(745, 255)
        TextBox14.Size = New Size(100, 93)
        TextBox14.BorderStyle = BorderStyle.None
        TextBox14.Font = New Font("DS-Digital", 75, FontStyle.Bold)
        TextBox14.Multiline = True
        TextBox14.TextAlign = HorizontalAlignment.Center
        TextBox14.Text = 0
        Me.Controls.Add(TextBox14)


        TextBox15.Location = New Point(870, 255)
        TextBox15.Size = New Size(480, 93)
        TextBox15.BorderStyle = BorderStyle.None
        TextBox15.Font = New Font("DS-Digital", 75, FontStyle.Bold)
        TextBox15.Multiline = True
        TextBox15.TextAlign = HorizontalAlignment.Right
        TextBox15.Text = 0
        Me.Controls.Add(TextBox15)

    End Sub

    '============Button============='
    Private Tombol1 As New Button()
    Private Tombol2 As New Button()
    Private Tombol3 As New Button()
    Private Tombol4 As New Button()
    Private Tombol5 As New Button()

    Sub Button()

        Tombol1.Text = "PROSES"
        Tombol1.Size = New Size(180, 70)
        Tombol1.Location = New Point(800, 670)
        Tombol1.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol1.BackColor = Color.White


        Tombol1.FlatStyle = FlatStyle.Flat
        Tombol1.FlatAppearance.BorderSize = 0
        Tombol1.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol1.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol1.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Raduis As New Drawing2D.GraphicsPath


        ' Memulai figure untuk path
        Raduis.StartFigure()

        ' Menambahkan elips pada sudut kiri atas
        Raduis.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        Raduis.AddLine(10, 0, Tombol1.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        Raduis.AddArc(New Rectangle(Tombol1.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        Raduis.AddLine(Tombol1.Width, 20, Tombol1.Width, Tombol1.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        Raduis.AddArc(New Rectangle(Tombol1.Width - 25, Tombol1.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        Raduis.AddLine(Tombol1.Width - 10, Tombol1.Width, 20, Tombol1.Height)
        ' Menambahkan elips pada sudut kiri bawah
        Raduis.AddArc(New Rectangle(0, Tombol1.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Raduis.CloseFigure()

        Tombol1.Region = New Region(Raduis)
        Me.Controls.Add(Tombol1)
        AddHandler Tombol1.Click, AddressOf Form5_MouseClick1
        AddHandler Tombol1.Click, AddressOf Form5_MouseClick5


        Tombol2.Text = "KELUAR"
        Tombol2.Size = New Size(175, 70)
        Tombol2.Location = New Point(1180, 670)
        Tombol2.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol2.BackColor = Color.White


        Tombol2.FlatStyle = FlatStyle.Flat
        Tombol2.FlatAppearance.BorderSize = 0
        Tombol2.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol2.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol2.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius As New Drawing2D.GraphicsPath

        Radius.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        Radius.AddLine(10, 0, Tombol2.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        Radius.AddArc(New Rectangle(Tombol2.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        Radius.AddLine(Tombol2.Width, 20, Tombol2.Width, Tombol2.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        Radius.AddArc(New Rectangle(Tombol2.Width - 25, Tombol2.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        Radius.AddLine(Tombol2.Width - 10, Tombol2.Width, 20, Tombol2.Height)
        ' Menambahkan elips pada sudut kiri bawah
        Radius.AddArc(New Rectangle(0, Tombol2.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Radius.CloseFigure()


        Tombol2.Region = New Region(Radius)
        Me.Controls.Add(Tombol2)
        AddHandler Tombol2.Click, AddressOf Form5_MouseClick
        ' Memulai figure untuk path

        Me.Controls.Add(Tombol2)
        AddHandler Tombol2.Click, AddressOf Form5_MouseClick

        Tombol3.Text = "HAPUS"
        Tombol3.Size = New Size(175, 70)
        Tombol3.Location = New Point(993, 670)
        Tombol3.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol3.BackColor = Color.White


        Tombol3.FlatStyle = FlatStyle.Flat
        Tombol3.FlatAppearance.BorderSize = 0
        Tombol3.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol3.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol3.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius1 As New Drawing2D.GraphicsPath

        Radius1.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        Radius1.AddLine(10, 0, Tombol3.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        Radius1.AddArc(New Rectangle(Tombol3.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        Radius1.AddLine(Tombol3.Width, 20, Tombol3.Width, Tombol3.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        Radius1.AddArc(New Rectangle(Tombol3.Width - 25, Tombol3.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        Radius1.AddLine(Tombol3.Width - 10, Tombol3.Width, 20, Tombol3.Height)
        ' Menambahkan elips pada sudut kiri bawah
        Radius1.AddArc(New Rectangle(0, Tombol3.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Radius1.CloseFigure()


        Tombol3.Region = New Region(Radius1)
        Me.Controls.Add(Tombol3)
        AddHandler Tombol3.Click, AddressOf Form5_MouseClick3
        ' Memulai figure untuk path


        Tombol4.Text = "BAYAR"
        Tombol4.Size = New Size(555, 45)
        Tombol4.Location = New Point(800, 613)
        Tombol4.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol4.BackColor = Color.White


        Tombol4.FlatStyle = FlatStyle.Flat
        Tombol4.FlatAppearance.BorderSize = 0
        Tombol4.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol4.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol4.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius12 As New Drawing2D.GraphicsPath

        Radius12.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        Radius12.AddLine(10, 0, Tombol4.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        Radius12.AddArc(New Rectangle(Tombol4.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        Radius12.AddLine(Tombol4.Width, 20, Tombol4.Width, Tombol4.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        Radius12.AddArc(New Rectangle(Tombol4.Width - 25, Tombol4.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        Radius12.AddLine(Tombol4.Width - 10, Tombol4.Width, 20, Tombol4.Height)
        ' Menambahkan elips pada sudut kiri bawah
        Radius12.AddArc(New Rectangle(0, Tombol4.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Radius12.CloseFigure()


        Tombol4.Region = New Region(Radius12)
        Me.Controls.Add(Tombol4)
        AddHandler Tombol4.Click, AddressOf Form5_MouseClick4
    End Sub

    Private gg As New DataGridView()
    Private Kolom1 As New DataGridViewButtonColumn()
    Private Kolom2 As New DataGridViewButtonColumn()
    Private Kolom3 As New DataGridViewButtonColumn()
    Private Kolom4 As New DataGridViewButtonColumn()
    Private Kolom5 As New DataGridViewButtonColumn()
    Private Kolom6 As New DataGridViewButtonColumn()
    Private Kolom7 As New DataGridViewButtonColumn()



    Sub kolom()
        Call Koneksi()

        gg.Location = New Point(0, 355)
        gg.Size = New Size(1400, 230)
        Me.Controls.Add(gg)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        If gg.Columns.Count = 0 Then

            Kolom1.HeaderText = "Kode Transaksi"
            gg.Columns.Add(Kolom1)

            Kolom2.HeaderText = "Nama Barang"
            gg.Columns.Add(Kolom2)

            Kolom3.HeaderText = "Jenis Barang"
            gg.Columns.Add(Kolom3)

            Kolom4.HeaderText = "Satuan Barang"
            gg.Columns.Add(Kolom4)

            Kolom5.HeaderText = "Harga"
            gg.Columns.Add(Kolom5)

            Kolom6.HeaderText = "Quantity"
            gg.Columns.Add(Kolom6)

            Kolom7.HeaderText = "Total Harga"
            gg.Columns.Add(Kolom7)
        Else
            gg.Rows.Clear()
        End If
        'Width Column

        gg.Columns(0).Width = 80
        gg.Columns(1).Width = 250
        gg.Columns(2).Width = 100
        gg.Columns(3).Width = 250
        gg.Columns(4).Width = 200
        gg.Columns(5).Width = 200
        gg.Columns(6).Width = 250
        gg.RowTemplate.Height = 30
        gg.ColumnHeadersHeight = 50

        cmd = New OleDbCommand("SELECT c.kode_id, t.nama_barang , t.jenis_barang, t.satuan,t.harga_jual ,c.qty,c.total_harga FROM sumtotal AS c INNER JOIN table_barang AS t ON t.nama_barang = c.nama_barang ;", conn)
        rd = cmd.ExecuteReader()
        gg.Rows.Clear()
        Do While rd.Read() = True
            gg.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))
        Loop


    End Sub

    Private Sub Form5_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Me.Close()
    End Sub

    Public Urutan As Double

    Private Sub Form5_MouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call Koneksi()
        Dim Tanggal As DateTime = Date.Now
        Dim Waktu As DateTime = Date.Now


        ' Tampilkan total_harga pada TextBox yang sesuai

        Try
            If TextBox8.Text = String.Empty Or TextBox9.Text = String.Empty Or TextBox10.Text = String.Empty Or TextBox11.Text = String.Empty Or TextBox12.Text = String.Empty Or TextBox13.Text = String.Empty Then
                MessageBox.Show("Data Tidak Lengkap")
                Call kolom()
            Else
                ' Ambil nilai total_harga yang baru saja dimasukkan
                TextBox2.Text = Waktu.ToShortTimeString()
                Dim Kode2 As String = Microsoft.VisualBasic.Right(TextBox2.Text, 2)
                TextBox3.Text = Tanggal.ToShortDateString()
                Dim Kode3 As String = Microsoft.VisualBasic.Left(TextBox3.Text, 2)
                Urutan = Convert.ToDouble(TextBox14.Text) + 1
                TextBox14.Text = Urutan
                TextBox1.Text = Microsoft.VisualBasic.Right(TextBox8.Text, 3) + Convert.ToString(Urutan) + TextBox14.Text + Microsoft.VisualBasic.Left(TextBox3.Text, 2)
                Dim totalHargaBaru As Double = Convert.ToDouble(TextBox12.Text) * Convert.ToDouble(TextBox13.Text)
                TextBox15.Text = totalHargaBaru
                Dim Total_bayar As Integer = Convert.ToDouble(TextBox7.Text) + totalHargaBaru
                TextBox7.Text = Total_bayar
                If TextBox8.Text = 0 Or TextBox1.Text = String.Empty Or TextBox13.Text = String.Empty Or TextBox7.Text = String.Empty Then
                    MsgBox("Data Harus Di isi")
                Else
                    cmd = New OleDbCommand("SELECT * FROM transaksi WHERE kode_id = '" & TextBox8.Text & "'", conn)
                    rd = cmd.ExecuteReader
                    If rd.HasRows = False Then
                        Try
                            Using insertcmd = New OleDbCommand("Insert Into transaksi (kode_id,No_Transaksi,quantity,total_harga) VALUES (@kode_id,@No_transaksi,@quantity,@total_harga)", conn)
                                insertcmd.Parameters.AddWithValue("@kode_id", TextBox8.Text)
                                insertcmd.Parameters.AddWithValue("@No_Transaksi,", TextBox1.Text)
                                insertcmd.Parameters.AddWithValue("@quantity", TextBox13.Text)
                                insertcmd.Parameters.AddWithValue("@total_harga", TextBox15.Text)
                                insertcmd.ExecuteNonQuery()
                            End Using
                        Catch ex As Exception
                            MsgBox("Kesalahan" & ex.Message)
                        End Try
                    End If
                    Using Insertcmd = New OleDbCommand("Insert Into table_penjualan (kode_id,No_Transaksi,Tanggal,waktu,kasir) Values (@kode_id,@No_Transaksi,@Tanggal,@waktu,@kasir)", conn)
                        Insertcmd.Parameters.AddWithValue("@kode_id", TextBox8.Text)
                        Insertcmd.Parameters.AddWithValue("@No_Transaksi", TextBox1.Text)
                        Insertcmd.Parameters.AddWithValue("@Tanggal", TextBox2.Text)
                        Insertcmd.Parameters.AddWithValue("@waktu", TextBox3.Text)
                        Insertcmd.Parameters.AddWithValue("@kasir", TextBox4.Text)
                        Insertcmd.ExecuteNonQuery()
                    End Using

                    Using Insertcmd = New OleDbCommand("Insert Into sumtotal (kode_id,nama_barang,qty,harga,total_harga) Values (@kode_id,@nama_barang,@qty,@harga,@total_harga)", conn)
                        Insertcmd.Parameters.AddWithValue("@kode_id", TextBox1.Text)
                        Insertcmd.Parameters.AddWithValue("@nama_barang", TextBox9.Text)
                        Insertcmd.Parameters.AddWithValue("@qty", TextBox13.Text)
                        Insertcmd.Parameters.AddWithValue("@harga", TextBox12.Text)
                        Insertcmd.Parameters.AddWithValue("@total_harga", TextBox15.Text)
                        Insertcmd.ExecuteNonQuery()
                    End Using
                    MsgBox("Data Transaksi Berhasil Di Input")
                    kolom()
                End If


            End If
        Catch ex As Exception
            MsgBox("KESALAHAN : " & ex.Message)
        End Try

    End Sub

    Sub updatebarang()

        Dim kembalian As Double = Convert.ToDouble(TextBox7.Text) - Convert.ToDouble(TextBox6.Text)

        If TextBox1.Text = String.Empty Or TextBox2.Text = String.Empty Or TextBox3.Text = String.Empty Or TextBox4.Text = String.Empty Then
            MsgBox("Identitas Transaksi Tidak Diketahui")
        Else
            cmd = New OleDbCommand("SELECT * FROM transaksi WHERE kode_id = '" & TextBox8.Text & "'", conn)
            rd = cmd.ExecuteReader
            If rd.HasRows = False Then
                Try
                    Using insertcmd = New OleDbCommand("Insert Into transaksi (kode_id,No_Transaksi,quantity,total_harga) VALUES (@kode_id,@No_transaksi,@quantity,@total_harga)", conn)
                        insertcmd.Parameters.AddWithValue("@kode_id", TextBox8.Text)
                        insertcmd.Parameters.AddWithValue("@No_Transaksi,", TextBox1.Text)
                        insertcmd.Parameters.AddWithValue("@quantity", TextBox13.Text)
                        insertcmd.Parameters.AddWithValue("@total_harga", TextBox15.Text)
                        insertcmd.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    MsgBox("Kesalahan" & ex.Message)
                End Try
            End If


            cmd = New OleDbCommand("Update table_penjualan set dibayar = '" & TextBox6.Text & _
                                   "',kembalian = '" & Convert.ToDouble(TextBox6.Text) - Convert.ToDouble(TextBox7.Text) & _
                                   "' where No_Transaksi = '" & TextBox1.Text & "' ", conn)
            cmd.ExecuteNonQuery()
            If TextBox6.Text < Convert.ToDouble(TextBox7.Text) Then
                MsgBox("Maaf, Saldo Anda Tidak Mencukupi.")
                Call hapusdata()
                Call delete_total()
                Call kolom()
            Else

                If MsgBox("Apakah Yakin Dengan Pesanan Anda??", MsgBoxStyle.YesNoCancel) = Windows.Forms.DialogResult.Yes Then
                    MsgBox("PESANAN TELAH SELESAI DI PROSES")
                    ' Tampilkan total_harga pada TextBox yang sesuai
                    TextBox5.Text = Convert.ToString(kembalian)
                Else
                    MsgBox("PESANAN DI BATALKAN")
                End If

            End If
        End If

    End Sub

    Sub labelloop()
        For i As Integer = 0 To Convert.ToDouble(TextBox14.Text)
            Dim Label As New Label()
            Label.Text = "Lable Ke - " & i
            Label.Location = New Point(10, 20 + 10)
            Label.ForeColor = Color.White
            Me.Controls.Add(Label)
        Next
    End Sub



    Sub delete()
        cmd = New OleDbCommand("Delete from transaksi where kode_id = '" & TextBox8.Text & "'", conn)
        cmd.ExecuteNonQuery()
    End Sub

    Sub delete_transaksi()
        cmd = New OleDbCommand("Delete from table_penjualan where No_Transaksi = '" & TextBox1.Text & "'", conn)
        cmd.ExecuteNonQuery()
    End Sub

    Sub delete_total()
        cmd = New OleDbCommand("Delete from sumtotal", conn)
        cmd.ExecuteNonQuery()
    End Sub

    Private Sub Form5_MouseClick4(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call updatebarang()
    End Sub

    Private Sub Form5_MouseClick3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call hapusdata()
        Call delete_total()
    End Sub

    Private Sub Form5_MouseClick5(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If TextBox1.Text = String.Empty Or TextBox2.Text = String.Empty Or TextBox3.Text = String.Empty Or TextBox4.Text = String.Empty Then
            MsgBox("Identitas Transaksi Tidak Diketahui")
        Else
            cmd = New OleDbCommand("select * from table_penjualan where No_Transaksi = '" & TextBox1.Text & "'", conn)
            rd = cmd.ExecuteReader()
            If rd.HasRows = False Then

            End If
        End If

    End Sub

    Private Sub Form5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged

        Using cmd = New OleDbCommand("select * from table_barang where kode_id = @kode_id", conn)
            cmd.Parameters.AddWithValue("@kode_id", TextBox8.Text)
            rd = cmd.ExecuteReader()
            gg.Rows.Clear()
            Do While rd.Read() = True
                gg.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))
            Loop


        End Using

    End Sub

    Private Sub kode_id(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged
        Using cmd = New OleDbCommand("SELECT kode_id, nama_barang, jenis_barang, satuan, harga_jual from table_barang where kode_id = @kode_id", conn)
            cmd.Parameters.AddWithValue("@kode_id", TextBox8.Text)
            rd = cmd.ExecuteReader()
            rd.Read()
            If rd.HasRows = True Then
                TextBox9.Text = rd(1)
                TextBox10.Text = rd(2)
                TextBox11.Text = rd(3)
                TextBox12.Text = rd(4)
            Else
                TextBox9.Text = " "
                TextBox10.Text = " "
                TextBox11.Text = " "
                TextBox12.Text = 0
                TextBox13.Text = 0

            End If
        End Using
    End Sub
    Private Sub Form5_MouseClick10(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call Form7.ShowDialog()
    End Sub


End Class