Public Class Form4

    Public textbox As New List(Of TextBox)

    Public Sub New()
        InitializeComponent()
        'BackGround
        Me.Size = New Size(827, 500)
        Me.Text = "Form Tambah Data"
        Me.BackColor = Color.CadetBlue
        Me.CenterToScreen()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow

        For i As Integer = 1 To 7
            Dim hoho As New TextBox()

            hoho.Size = New Size(580, 50)
            hoho.Location = New Point(200, (i * 50))
            hoho.Font = New Font("Sogoe UI", 12)
            hoho.BorderStyle = BorderStyle.FixedSingle

            Me.Controls.Add(hoho)
        Next

        Dim Tombol3 As New Button()
        Tombol3.Text = "HAPUS"
        Tombol3.Size = New Size(130, 40)
        Tombol3.Location = New Point(500, 380)
        Tombol3.Font = New Font("Sogoe UI", 12, FontStyle.Bold)
        Tombol3.BackColor = Color.LightCyan
        Me.Controls.Add(Tombol3)

    End Sub

    Private Sub Form4_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For Each hoho As TextBox In textbox
            cleartext(hoho)
        Next
    End Sub

    Sub cleartext(ByVal textbox As TextBox)
        textbox.Clear()
    End Sub


    Public Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Shared Sub main()
        Dim Myform As New Form4
        Application.Run(Myform)
    End Sub


End Class