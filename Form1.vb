Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeComponent()

        Tombol()
        Koneksi()
        gambar()
    End Sub

    Public Sub Tombol()

        'Label


        'Label
        Dim Label2 As New Label()
        Label2.Text = "Project By :" & vbNewLine & "Sadewo Widyanto"
        Label2.Size = New Size(500, 300)
        Label2.Location = New Point(1100, 600)
        Label2.ForeColor = Color.White
        Label2.Font = New Font("Sogoe UI", 18, FontStyle.Bold)
        Label2.BackColor = Color.Transparent
        Me.Controls.Add(Label2)

        'Tombol
        Dim Tombol1 As New Button()


        Tombol1.Text = "BARANG"
        Tombol1.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol1.ForeColor = Color.White
        Tombol1.BackColor = Color.Transparent
        Tombol1.Location = New Point(360, 150)
        Tombol1.Size = New Size(300, 50)

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

        Dim Tombol2 As New Button()

        Tombol2.Text = "PENJUALAN"
        Tombol2.Font = New Font("Sogoe UI", 15, FontStyle.Bold)
        Tombol2.BackColor = Color.Transparent
        Tombol2.ForeColor = Color.White
        Tombol2.Location = New Point(700, 150)
        Tombol2.Size = New Size(300, 50)
        AddHandler Tombol2.Click, AddressOf Form1_MouseClick1

        Tombol2.FlatStyle = FlatStyle.Flat
        Tombol2.FlatAppearance.BorderSize = 0
        Tombol2.FlatAppearance.BorderColor = SystemColors.ControlDarkDark
        Tombol2.FlatAppearance.MouseOverBackColor = SystemColors.GradientActiveCaption
        Tombol2.FlatAppearance.MouseDownBackColor = SystemColors.GradientInactiveCaption
        Me.Controls.Add(Tombol2)

        Dim Radius As New Drawing2D.GraphicsPath

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

        Dim Background As String = "C:\Users\Acer\OneDrive\Documents\Visual Studio 2010\Projects\ConsoleApplication1\bin\Debug\WindowsApplication3\WindowsApplication3\background.JPG"


        'Background
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.SizableToolWindow
        If IO.File.Exists(Background) Then
            Me.BackgroundImage = Image.FromFile(Background)
        Else
            MsgBox("File Tidak Ditemukan")
        End If

        AddHandler Tombol1.Click, AddressOf Form1_MouseClick
        Me.Text = "CRUD SYSTEM"


    End Sub

    Sub gambar()
        Dim PictureBox As New PictureBox
        Dim Imagepath As String = "C:\\Users\Acer\Downloads\crud1.png"
        Dim GAMBAR As Image = Image.FromFile(Imagepath)
        PictureBox.Image = GAMBAR
        PictureBox.Size = New Size(1200, 500)
        PictureBox.BackColor = Color.Transparent
        PictureBox.Location = New Point(170, 250)
        Me.Controls.Add(PictureBox)
    End Sub

    Private Sub Form1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Form2.ShowDialog()
    End Sub

    Private Sub Form1_MouseClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Call Form5.delete_total()
        Form5.ShowDialog()
    End Sub

    Private Sub Form1_MouseClick2(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Form7.ShowDialog()
    End Sub
End Class
