Imports System.Data.OleDb

Public Class Form2

    Public TextBox1 As New TextBox()
    Private judul As New Label()
    Private Tombol1 As New Button()
    Private Tombol2 As New Button()
    Private Tombol3 As New Button()
    Public Sub Setting()
        'BackGround Settings
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.Text = "Data Barang"
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Dim Background As String = "C:\Users\Acer\OneDrive\Documents\Visual Studio 2010\Projects\ConsoleApplication1\bin\Debug\WindowsApplication3\WindowsApplication3\background.JPG"


        'Background
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        If IO.File.Exists(Background) Then
            Me.BackgroundImage = Image.FromFile(Background)
        Else
            MsgBox("File Tidak Ditemukan")
        End If


        judul.Text = "DATA BARANG"
        judul.Size = New Size(1000, 100)
        judul.ForeColor = Color.OrangeRed
        judul.Location = New Point(10, 10)
        judul.Font = New Font("Helvetica", 70, FontStyle.Bold)
        judul.BackColor = Color.Transparent
        Me.Controls.Add(judul)

        'Tombol TAMBAH

        Tombol1.Text = "Tambah Barang"
        Tombol1.Size = New Size(300, 40)
        Tombol1.Location = New Point(25, 170)
        Tombol1.BackColor = Color.White
        Tombol1.Font = New Font("Sogoe UI", 15, FontStyle.Bold)

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
        AddHandler Tombol1.Click, AddressOf Form2_MouseClick1

        'Tombol KELUAR

        Tombol2.Text = "Keluar"
        Tombol2.Size = New Size(300, 40)
        Tombol2.Location = New Point(350, 170)
        Tombol2.BackColor = Color.White
        Tombol2.Font = New Font("Sogoe UI", 15, FontStyle.Bold)

        Tombol2.FlatStyle = FlatStyle.Flat
        Tombol2.FlatAppearance.BorderSize = 0
        Tombol2.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol2.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol2.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius As New Drawing2D.GraphicsPath
        AddHandler Tombol2.Click, AddressOf Form2_MouseClick

        ' Memulai figure untuk path
        Raduis.StartFigure()

        ' Menambahkan elips pada sudut kiri atas
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

        Tombol2.Region = New Region(Raduis)
        Me.Controls.Add(Tombol2)


        Tombol3.Text = "ARUS BARANG"
        Tombol3.Size = New Size(300, 40)
        Tombol3.Location = New Point(1050, 700)
        Tombol3.BackColor = Color.White
        Tombol3.Font = New Font("Sogoe UI", 15, FontStyle.Bold)

        Tombol3.FlatStyle = FlatStyle.Flat
        Tombol3.FlatAppearance.BorderSize = 0
        Tombol3.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol3.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol3.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius1 As New Drawing2D.GraphicsPath
        AddHandler Tombol3.Click, AddressOf Form2_MouseClick2

        ' Memulai figure untuk path
        Raduis.StartFigure()

        ' Menambahkan elips pada sudut kiri atas
        Radius.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        Radius.AddLine(10, 0, Tombol3.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        Radius.AddArc(New Rectangle(Tombol3.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        Radius.AddLine(Tombol3.Width, 20, Tombol3.Width, Tombol3.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        Radius.AddArc(New Rectangle(Tombol3.Width - 25, Tombol3.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        Radius.AddLine(Tombol3.Width - 10, Tombol3.Width, 20, Tombol3.Height)
        ' Menambahkan elips pada sudut kiri bawah
        Radius.AddArc(New Rectangle(0, Tombol3.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Radius.CloseFigure()

        Tombol3.Region = New Region(Raduis)
        Me.Controls.Add(Tombol3)
        'TextBox

        TextBox1.Location = New Point(750, 180)
        TextBox1.Size = New Size(600, 100)
        TextBox1.Font = New Font("Helvetica", 12)
        TextBox1.BorderStyle = BorderStyle.None
        TextBox1.BackColor = Color.White
        Me.Controls.Add(TextBox1)
        AddHandler TextBox1.TextChanged, AddressOf Form2_TextChanged

        'Label Cari
        Dim Label2 As New Label()
        Label2.Text = "Cari"
        Label2.Font = New Font("Helvetica", 18, FontStyle.Bold)
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Transparent
        Label2.Location = New Point(670, 175)
        Label2.Size = New Size(580, 70)
        Me.Controls.Add(Label2)

    End Sub



    Sub columns()



        'Data Grid



    End Sub

    Public Data As New DataGridView()

    Public Sub Datagridview()

        Data.Location = New Point(10, 250)
        Data.Size = New Size(1345, 432)
        Me.Controls.Add(Data)

        If Data.Columns.Count = 0 Then
            Dim Kolom1 As New DataGridViewButtonColumn()
            Kolom1.HeaderText = "Kode ID"
            Data.Columns.Add(Kolom1)

            Dim Kolom2 As New DataGridViewButtonColumn()
            Kolom2.HeaderText = "Nama"
            Data.Columns.Add(Kolom2)


            Dim Kolom3 As New DataGridViewButtonColumn()
            Kolom3.HeaderText = "Jenis"
            Data.Columns.Add(Kolom3)

            Dim Kolom7 As New DataGridViewButtonColumn()
            Kolom7.HeaderText = "satuan"
            Data.Columns.Add(Kolom7)


            Dim Kolom4 As New DataGridViewButtonColumn()
            Kolom4.HeaderText = "Harga Modal"
            Data.Columns.Add(Kolom4)

            Dim Kolom5 As New DataGridViewButtonColumn()
            Kolom5.HeaderText = "Harga Jual"
            Data.Columns.Add(Kolom5)


            Dim Kolom8 As New DataGridViewButtonColumn()
            Kolom8.HeaderText = "Stok"
            Data.Columns.Add(Kolom8)



        End If

        'Size Width
        Data.Columns(0).Width = 100
        Data.Columns(1).Width = 265
        Data.Columns(2).Width = 265
        Data.Columns(3).Width = 150
        Data.Columns(4).Width = 180
        Data.Columns(5).Width = 183
        Data.Columns(6).Width = 155
        Data.RowTemplate.Height = 30
        Data.ColumnHeadersHeight = 50

        Data.BackgroundColor = Color.LightGray
        Call Koneksi()
        cmd = New OleDbCommand("select * from table_barang ", conn)
        rd = cmd.ExecuteReader()
        Data.Rows.Clear()
        Do While rd.Read() = True
            Data.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6))
        Loop
    End Sub

    Sub CariData()

    End Sub


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Me.Close()
    End Sub

    Private Sub Form2_MouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Form3.ShowDialog()
    End Sub

    Private Sub Form2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Datagridview()
        cmd = New OleDbCommand("select * from table_barang where nama_barang like'%" & TextBox1.Text & "%'", conn)
        rd = cmd.ExecuteReader()
        Data.Rows.Clear()
        Do While rd.Read() = True
            Data.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(5))
        Loop
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Setting()
        Datagridview()
        Call Form6.tampilan()

    End Sub

    Private Sub Form2_MouseClick2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Form6.ShowDialog()
    End Sub
End Class