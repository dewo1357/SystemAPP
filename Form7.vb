Imports System.Data.OleDb
Public Class Form7

    Private Sub Form7_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Koneksi()
        Me.Size = New Size(650, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ForeColor = Color.Black
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.BackColor = Color.DarkGray
        Me.Text = "INFORMATION"
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        call judul()
        Call tampil_data()
        Call groupbox()



    End Sub


    Sub tampil_data()
        Call Koneksi()
        cmd = New OleDbCommand("Select * from sumtotal", conn)
        rd = cmd.ExecuteReader()

        Dim xpos As Integer = 30
        Dim ypos As Integer = 230



        While rd.Read()
            Dim Label1 As New Label()
            Label1.Size = New Size(570, 20)
            Label1.Text = rd("kode_id").ToString() & "                                      " & rd("nama_barang").ToString() & "                " & rd("qty").ToString() & "                   " & rd("harga").ToString() & "                   " & rd("total_harga").ToString()
            Label1.ForeColor = Color.White
            Label1.Location = New Point(xpos, ypos)
            Label1.Font = New Font("Helvetica", 12)
            Label1.BackColor = Color.DarkGray
            ypos += 25
            Me.Controls.Add(Label1)

        End While


    End Sub

    Private GroupBox1 As New GroupBox()
    Private GroupBox2 As New GroupBox()
    Private GroupBox3 As New GroupBox()
    Private GroupBox4 As New GroupBox()

    Sub groupbox()
        GroupBox1.Size = New Size(605, 140)
        GroupBox1.Location = New Point(20, 0)
        GroupBox1.BackColor = Color.DarkGray
        GroupBox1.Dock = DockStyle.None
        GroupBox1.Font = New Font("Helvetiva", 20, FontStyle.Bold)
        GroupBox1.ForeColor = Color.White
        GroupBox1.Text = "INFORMATION STORE"
        Me.Controls.Add(GroupBox1)

        GroupBox2.Size = New Size(620, 430)
        GroupBox2.Location = New Point(10, 145)
        GroupBox2.BackColor = Color.DarkGray
        GroupBox2.Dock = DockStyle.None
        GroupBox2.Font = New Font("Helvetiva", 20, FontStyle.Bold)
        GroupBox2.ForeColor = Color.White
        GroupBox2.Text = "List Products"
        Me.Controls.Add(GroupBox2)

        GroupBox3.Size = New Size(620, 80)
        GroupBox3.Location = New Point(10, 575)
        GroupBox3.BackColor = Color.DarkGray
        GroupBox3.Dock = DockStyle.None
        GroupBox3.Font = New Font("Helvetiva", 14, FontStyle.Bold)
        GroupBox3.ForeColor = Color.White
        GroupBox3.Text = "Total"
        Me.Controls.Add(GroupBox3)
    End Sub

    'Judul

    Sub judul

        Dim Label00 As New Label() ' Invoice
        Label00.Text = "|" & Form5.TextBox1.Text & " | " & Form5.TextBox2.Text & " | " & Form5.TextBox3.Text & "|"
        Label00.Size = New Size(500, 50)
        Label00.Font = New Font("Helvetica", 20, FontStyle.Bold)
        Label00.Location = New Point(25, 40)
        Label00.ForeColor = Color.White
        Label00.BackColor = Color.DarkGray
        Me.Controls.Add(Label00)

        Dim Label003 As New Label() ' Kasir
        Label003.Text = "Kasir         :" & "  " & Form5.TextBox4.Text
        Label003.Size = New Size(400, 30)
        Label003.Font = New Font("Sogoe UI", 15)
        Label003.Location = New Point(30, 90)
        Label003.ForeColor = Color.White
        Label003.BackColor = Color.DarkGray
        Me.Controls.Add(Label003)

        Dim Label12 As New Label()
        Label12.Size = New Size(600, 30)
        Label12.Text = "No" & "                       " & "Nama" & "          " & "QTY" & "         " & "Harga" & "           " & "Total"
        Label12.ForeColor = Color.White
        Label12.Location = New Point(25, 190)
        Label12.Font = New Font("Sogoe UI", 13, Fontstyle.Bold)
        Label12.BackColor = Color.DarkGray
        Me.Controls.Add(Label12)

        Dim Label004 As New Label() ' Total Pembayaran
        Label004.Text = Form5.TextBox7.Text
        Label004.Size = New Size(300, 50)
        Label004.Font = New Font("DS-Digital", 40)
        Label004.Location = New Point(75, 600)
        Label004.ForeColor = Color.White
        Label004.BackColor = Color.DarkGray
        Me.Controls.Add(Label004)


    End Sub

    Private Sub Form7_TextChanged(sender As Object, e As EventArgs) Handles Me.TextChanged

    End Sub
End Class