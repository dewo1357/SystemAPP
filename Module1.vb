Imports System.Data.OleDb

Module Module1

    Public cmd As OleDbCommand
    Public conn As OleDbConnection
    Public rd As OleDbDataReader

    Sub Koneksi()
        conn = New OleDbConnection("provider=microsoft.jet.oledb.4.0; data source= C:\\Users\Acer\OneDrive\Documents\Visual Studio 2010\Projects\ConsoleApplication1\bin\Debug\WindowsApplication3\WindowsApplication3\bin\Debug\hehehehe.mdb")
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If

        'MsgBox("Data Terlihat")
    End Sub

End Module
