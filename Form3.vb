Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class Form3


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        label()
        tombol()
        TextBox()

        Me.Size = New Size(827, 500)
        Me.Text = "Form Tambah Data"

        Me.CenterToScreen()


        Dim Judul As New Label()
        Judul.Text = "TAMBAH BARANG"
        Judul.Size = New Size(1000, 1000)
        Judul.Location = New Point(200, 10)
        Judul.Font = New Font("Helvetica", 35, FontStyle.Bold)
        Judul.BackColor = Color.Transparent
        Me.Controls.Add(Judul)

        Me.BackColor = Color.DarkGray
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow

        ' Add any initialization after the InitializeComponent() call.

        ' Import fungsi API Windows


    End Sub



    'Insialisasi TextBox
    Private TextBox1 As New TextBox
    Private TextBox2 As New TextBox
    Private ComboBox1 As New ComboBox
    Private TextBox4 As New ComboBox
    Private TextBox5 As New TextBox
    Private TextBox6 As New TextBox
    Private TextBox7 As New TextBox
    'Fungsi Menghapus Isi Pada TextBox
    Sub Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.Text = " "
        TextBox4.Text = " "
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
    End Sub

    'Settings TextBox
    Sub TextBox()

        TextBox1.Size = New Size(580, 50)
        TextBox1.Location = New Point(200, 90)
        TextBox1.Font = New Font("Sogoe UI", 12)
        TextBox1.BorderStyle = BorderStyle.None

        Me.Controls.Add(TextBox1)
        AddHandler TextBox1.TextChanged, AddressOf Form3_TextChanged

        TextBox2.Size = New Size(580, 50)
        TextBox2.Location = New Point(200, 140)
        TextBox2.Font = New Font("Sogoe UI", 12)
        TextBox2.BorderStyle = BorderStyle.None

        Me.Controls.Add(TextBox2)

        ComboBox1.Size = New Size(580, 50)
        ComboBox1.Location = New Point(200, 190)
        ComboBox1.Font = New Font("Sogoe UI", 10)
        ComboBox1.Items.Add("Makanan")
        ComboBox1.Items.Add("Alat Elektronik")
        ComboBox1.Items.Add("Minuman")
        ComboBox1.Items.Add("Furniture")
        Me.Controls.Add(ComboBox1)

        TextBox4.Size = New Size(580, 50)
        TextBox4.Location = New Point(200, 240)
        TextBox4.Font = New Font("Sogoe UI", 10)
        TextBox4.Items.Add("Unit")
        TextBox4.Items.Add("Pcs")
        TextBox4.Items.Add("Bungkus")
        Me.Controls.Add(TextBox4)

        TextBox5.Size = New Size(580, 50)
        TextBox5.Location = New Point(200, 290)
        TextBox5.Font = New Font("Sogoe UI", 12)
        TextBox5.BorderStyle = BorderStyle.None

        Me.Controls.Add(TextBox5)

        TextBox6.Size = New Size(580, 50)
        TextBox6.Location = New Point(200, 340)
        TextBox6.Font = New Font("Sogoe UI", 12)
        TextBox6.BorderStyle = BorderStyle.None
        Me.Controls.Add(TextBox6)

        TextBox7.Size = New Size(580, 50)
        TextBox7.Location = New Point(200, 390)
        TextBox7.Font = New Font("Sogoe UI", 12)
        TextBox7.BorderStyle = BorderStyle.None
        Me.Controls.Add(TextBox7)
    End Sub

    'Inisialisasi Tombol
    Private Tombol1 As New Button()
    Private Tombol2 As New Button()
    Private Tombol3 As New Button()
    Private Tombol4 As New Button()

    Sub tombol()

        Tombol3.Text = "BATAL"
        Tombol3.Size = New Size(130, 40)
        Tombol3.Location = New Point(500, 430)
        Tombol3.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Tombol3.BackColor = Color.White
        Tombol3.ForeColor = Color.Black


        Tombol3.FlatStyle = FlatStyle.Flat
        Tombol3.FlatAppearance.BorderSize = 0
        Tombol3.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol3.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol3.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption


        Dim ellips As New Drawing2D.GraphicsPath

        ' Memulai figure untuk path
        ellips.StartFigure()

        ' Menambahkan elips pada sudut kiri atas
        ellips.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        ellips.AddLine(10, 0, Tombol3.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        ellips.AddArc(New Rectangle(Tombol3.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        ellips.AddLine(Tombol3.Width, 20, Tombol3.Width, Tombol3.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        ellips.AddArc(New Rectangle(Tombol3.Width - 25, Tombol3.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        ellips.AddLine(Tombol3.Width - 10, Tombol3.Width, 20, Tombol3.Height)
        ' Menambahkan elips pada sudut kiri bawah
        ellips.AddArc(New Rectangle(0, Tombol3.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        ellips.CloseFigure()

        Tombol3.Region = New Region(ellips)
        AddHandler Tombol3.Click, AddressOf Form3_MouseClick2
        Me.Controls.Add(Tombol3)



        Tombol1.Text = "SIMPAN"
        Tombol1.Size = New Size(130, 40)
        Tombol1.Location = New Point(200, 430)
        Tombol1.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Tombol1.BackColor = Color.White
        Tombol1.ForeColor = Color.Black


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
        AddHandler Tombol1.Click, AddressOf Form3_tambah_data



        Tombol2.Text = "HAPUS"
        Tombol2.Size = New Size(130, 40)
        Tombol2.Location = New Point(350, 430)
        Tombol2.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Tombol2.BackColor = Color.White
        Tombol2.ForeColor = Color.Black


        Tombol2.FlatStyle = FlatStyle.Flat
        Tombol2.FlatAppearance.BorderSize = 0
        Tombol2.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol2.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol2.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim Radius As New Drawing2D.GraphicsPath
        AddHandler Tombol2.Click, AddressOf Form3_MouseClick3

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



        Tombol4.Text = "KELUAR"
        Tombol4.Size = New Size(130, 40)
        Tombol4.Location = New Point(650, 430)
        Tombol4.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Tombol4.BackColor = Color.White
        Tombol4.ForeColor = Color.Black



        Tombol4.FlatStyle = FlatStyle.Flat
        Tombol4.FlatAppearance.BorderSize = 0
        Tombol4.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol4.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol4.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption

        Dim ellips2 As New Drawing2D.GraphicsPath

        AddHandler Tombol4.Click, AddressOf Form3_MouseClick

        ' Memulai figure untuk path
        ellips2.StartFigure()

        ' Menambahkan elips pada sudut kiri atas
        ellips2.AddArc(New Rectangle(0, 0, 20, 20), 180, 90)
        ' Menambahkan garis menuju sudut kanan atas
        ellips2.AddLine(10, 0, Tombol4.Width - 20, 0)
        ' Menambahkan elips pada sudut kanan atas
        ellips2.AddArc(New Rectangle(Tombol4.Width - 20, 0, 20, 20), -90, 90)
        ' Menambahkan garis menuju sudut kanan bawah
        ellips2.AddLine(Tombol4.Width, 20, Tombol4.Width, Tombol4.Height - 10)
        ' Menambahkan elips pada sudut kanan bawah
        ellips2.AddArc(New Rectangle(Tombol4.Width - 25, Tombol4.Height - 25, 25, 25), 0, 90)
        ' Menambahkan garis menuju sudut kiri bawah
        ellips2.AddLine(Tombol4.Width - 10, Tombol4.Width, 20, Tombol4.Height)
        ' Menambahkan elips pada sudut kiri bawah
        ellips2.AddArc(New Rectangle(0, Tombol4.Height - 20, 20, 20), 90, 90)

        ' Menutup figure dan memulai figure baru
        Radius.CloseFigure()

        Tombol4.Region = New Region(ellips2)
        Me.Controls.Add(Tombol4)
    End Sub

    Public Sub tambah_data()

    End Sub

    Sub update_data()

    End Sub

    Sub label()
        Dim Label1 As New Label()
        Label1.Text = "Kode ID"
        Label1.Location = New Point(20, 90)
        Label1.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Me.Controls.Add(Label1)

        Dim Label2 As New Label()
        Label2.Text = "Nama"
        Label2.Location = New Point(20, 140)
        Label2.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Me.Controls.Add(Label2)

        Dim Label3 As New Label()
        Label3.Text = "Jenis"
        Label3.Location = New Point(20, 190)
        Label3.Font = New Font("Sogoe UI", 14, FontStyle.Bold)
        Me.Controls.Add(Label3)

        Dim Label4 As New Label()
        Label4.Text = "Harga Modal"
        Label4.Size = New Size(180, 50)
        Label4.Location = New Point(20, 290)
        Label4.Font = New Font("Sogoe UI", 14, FontStyle.Bold)
        Me.Controls.Add(Label4)

        Dim Label5 As New Label()
        Label5.Text = "Harga Jual"
        Label5.Location = New Point(20, 340)
        Label5.Size = New Size(180, 50)
        Label5.Font = New Font("Sogoe UI", 14, FontStyle.Bold)
        Me.Controls.Add(Label5)

        Dim Label6 As New Label()
        Label6.Text = "Stok"
        Label6.Location = New Point(20, 390)
        Label6.Size = New Size(180, 50)
        Label6.Font = New Font("Sogoe UI", 14, FontStyle.Bold)
        Me.Controls.Add(Label6)

        Dim Label7 As New Label()
        Label7.Text = "Satuan"
        Label7.Location = New Point(20, 240)
        Label7.Size = New Size(180, 50)
        Label7.Font = New Font("Sogoe UI", 14, FontStyle.Bold)
        Me.Controls.Add(Label7)
    End Sub



    Private Sub Form3_MouseClick2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Clear()
    End Sub

    'Private Sub Form3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    ' If e.KeyChar = ChrW(Keys.Enter) Then
    ' TextBox2.Focus()
    '   End If
    ' End Sub

    ' Private Sub Form3_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '  If e.KeyChar = ChrW(Keys.Enter) Then
    '    TextBox4.Focus()
    '  End If
    ' End Sub

    ' Private Sub Form3_KeyPress2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '   If e.KeyChar = ChrW(Keys.Enter) Then
    '  TextBox5.Focus()
    ' End If
    ' End Sub

    'Private Sub Form3_KeyPress3(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    ' If e.KeyChar = ChrW(Keys.Enter) Then
    '   TextBox6.Focus()
    ' End If
    ' End Sub



    Private Sub Form3_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Me.Close()
    End Sub

    Private Sub Form3_tambah_data(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call Koneksi()
        Try
            If TextBox1.Text = String.Empty Or TextBox2.Text = String.Empty Or ComboBox1.Text = String.Empty Or TextBox4.Text = String.Empty Or TextBox5.Text = String.Empty Or TextBox6.Text = String.Empty Or TextBox7.Text = String.Empty Then
                MessageBox.Show("Data Tidak Lengkap")
            Else
                cmd = New OleDbCommand("SELECT * FROM table_barang WHERE kode_id = '" & TextBox1.Text & "'", conn)
                rd = cmd.ExecuteReader
                If rd.HasRows = False Then
                    Using insertcmd = New OleDbCommand("INSERT INTO table_barang(kode_id,nama_barang,jenis_barang,satuan,harga_modal,harga_jual,stok)VALUES(@kode_id,@nama_barang,@jenis_barang,@satuan,@harga_modal,@harga_jual,@stok)", conn)
                        insertcmd.Parameters.AddWithValue("@kode_id", TextBox1.Text)
                        insertcmd.Parameters.AddWithValue("@nama_arang", TextBox2.Text)
                        insertcmd.Parameters.AddWithValue("@jenis_barang", ComboBox1.Text)
                        insertcmd.Parameters.AddWithValue("@satuan", TextBox4.Text)
                        insertcmd.Parameters.AddWithValue("@harga_modal", TextBox5.Text)
                        insertcmd.Parameters.AddWithValue("@harga_jual", TextBox6.Text)
                        insertcmd.Parameters.AddWithValue("@stok", TextBox7.Text)
                        insertcmd.ExecuteNonQuery()
                    End Using
                    MsgBox("Data Berhasil Ditambah!")
                    Call Form2.Datagridview()
                Else
                    cmd = New OleDbCommand("update table_barang set nama_barang = '" & TextBox2.Text &
                                          "',jenis_barang = '" & ComboBox1.Text &
                                          "',satuan = '" & TextBox4.Text &
                                          "',harga_modal ='" & TextBox5.Text &
                                          "',harga_jual ='" & TextBox6.Text &
                                          "',stok = '" & TextBox7.Text &
                                          "'where kode_id = '" & TextBox1.Text & "'", conn)
                    cmd.ExecuteNonQuery()
                    MsgBox("Data Berhasil Di Perbarui")
                    Call Form2.Datagridview()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Terjadi kesalahan OleDb: " & ex.Message)
        End Try

    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeComponent()
        Call Koneksi()
        Me.Text = "TAMBAH BARANG"

    End Sub

    Private Sub Form3_MouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call Clear()
    End Sub

    Private Sub Form3_MouseClick3(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        If MsgBox("Apakah Anda Yakin Menghapus Data Ini?", MsgBoxStyle.YesNoCancel) = Windows.Forms.DialogResult.Yes Then
            cmd = New OleDbCommand("delete from table_barang where kode_id = '" & TextBox1.Text & "' ", conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Hapus")
            Call Clear()

        End If
        Call Form2.Datagridview()
    End Sub

    Private Sub Form3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        cmd = New OleDbCommand("select * from table_barang where kode_id ='" & TextBox1.Text & "'", conn)
        rd = cmd.ExecuteReader()
        rd.Read()
        If rd.HasRows = True Then
            TextBox2.Text = rd(1)
            ComboBox1.Text = rd(2)
            TextBox4.Text = rd(3)
            TextBox5.Text = rd(4)
            TextBox6.Text = rd(5)
            TextBox7.Text = rd(6)
        Else
            TextBox2.Text = " "
            ComboBox1.Text = " "
            TextBox4.Text = " "
            TextBox5.Text = " "
            TextBox6.Text = " "
        End If
    End Sub

    Sub panggilan()
        cmd = New OleDbCommand("select * from table_barang where kode_id ='" & TextBox1.Text & "'", conn)
        rd = cmd.ExecuteReader()
        rd.Read()
        If rd.HasRows = True Then
            TextBox2.Text = rd(1)
            ComboBox1.Text = rd(2)
            TextBox4.Text = rd(4)
            TextBox5.Text = rd(5)
            TextBox6.Text = rd(6)
            TextBox7.Text = rd(3)
        Else
            TextBox2.Text = " "
            ComboBox1.Text = " "
            TextBox4.Text = " "
            TextBox5.Text = " "
            TextBox6.Text = " "
        End If
    End Sub
End Class