Imports System.Data.OleDb
Public Class Form6

    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Size = New Size(1350, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.WindowState = FormWindowState.Maximized
        Me.Text = "Arus Barang"


        'Background
        Dim Background As String = "C:\Users\Acer\OneDrive\Documents\Visual Studio 2010\Projects\ConsoleApplication1\bin\Debug\WindowsApplication3\WindowsApplication3\background.JPG"
        If IO.File.Exists(Background) Then
            Me.BackgroundImage = Image.FromFile(Background)
        Else
            MsgBox("File Tidak Ditemukan")
        End If


        Call tampilan()

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        label()
        tampilan()


    End Sub

    Private judul As New Label
    Private TextBox1 As New TextBox
    Sub label()
        judul.Text = "HASIL STOK AKHIR"
        judul.Size = New Size(1000, 100)
        judul.ForeColor = Color.OrangeRed
        judul.Location = New Point(230, 20)
        judul.Font = New Font("Helvetica", 70, FontStyle.Bold)
        judul.BackColor = Color.Transparent
        Me.Controls.Add(judul)

        TextBox1.Location = New Point(750, 185)
        TextBox1.Size = New Size(600, 100)
        TextBox1.Font = New Font("Helvetica", 12)
        TextBox1.BorderStyle = BorderStyle.None
        TextBox1.BackColor = Color.White
        Me.Controls.Add(TextBox1)
        AddHandler TextBox1.TextChanged, AddressOf Form6_TextChanged

        'Label Cari
        Dim Label2 As New Label()
        Label2.Text = "Cari"
        Label2.Font = New Font("Helvetica", 18, FontStyle.Bold)
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Transparent
        Label2.Location = New Point(670, 180)
        Label2.Size = New Size(580, 70)
        Me.Controls.Add(Label2)

    End Sub

    Public Data As New DataGridView()

    Public Sub tampilan()

        Data.Location = New Point(10, 250)
        Data.Size = New Size(1348, 400)
        Data.BackColor = Color.LightGray
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
            Kolom8.HeaderText = "Stok Awal"
            Data.Columns.Add(Kolom8)

            Dim Kolom9 As New DataGridViewButtonColumn()
            Kolom9.HeaderText = "Stok Keluar"
            Data.Columns.Add(Kolom9)

            Dim Kolom10 As New DataGridViewButtonColumn()
            Kolom10.HeaderText = "Sisa Stok"
            Data.Columns.Add(Kolom10)


        End If

        'Size Width
        Data.Columns(0).Width = 100
        Data.Columns(1).Width = 230
        Data.Columns(2).Width = 233
        Data.Columns(3).Width = 140
        Data.Columns(4).Width = 122
        Data.Columns(5).Width = 140
        Data.Columns(6).Width = 140
        Data.RowTemplate.Height = 30
        Data.ColumnHeadersHeight = 50


        Data.BackgroundColor = Color.LightGray
        Call Koneksi()
        cmd = New OleDbCommand("select e.kode_id,e.nama_barang,e.jenis_barang,e.satuan,e.harga_modal,e.harga_jual,e.stok,sum(t.quantity) as total_quantity ,(e.stok - sum(t.quantity)) as sisa_stok from table_barang as e INNER JOIN transaksi as t ON t.kode_id = e.kode_id Group By e.kode_id,e.nama_barang,e.jenis_barang,e.satuan,e.harga_modal,e.harga_jual,e.stok ", conn)
        rd = cmd.ExecuteReader()
        Data.Rows.Clear()
        Do While rd.Read() = True
            Data.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6), rd(7), rd(8))
        Loop
    End Sub

    Private Sub Form6_TextChanged(sender As Object, e As EventArgs)
        Call tampilan()
        cmd = New OleDbCommand("select e.kode_id,e.nama_barang,e.jenis_barang,e.satuan,e.harga_modal,e.harga_jual,e.stok,sum(t.quantity) as total_quantity ,(e.stok - sum(t.quantity)) as sisa_stok from table_barang as e INNER JOIN transaksi as t ON t.kode_id = e.kode_id Where e.nama_barang Like '%" & TextBox1.Text & "%' Group By e.kode_id,e.nama_barang,e.jenis_barang,e.satuan,e.harga_modal,e.harga_jual,e.stok;", conn)
        rd = cmd.ExecuteReader()
        Data.Rows.Clear()
        Do While rd.Read() = True
            Data.Rows.Add(rd(0), rd(1), rd(2), rd(3), rd(4), rd(5), rd(6), rd(7), rd(8))
        Loop
    End Sub
End Class