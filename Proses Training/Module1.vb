Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types

Public Module Module1
    Public con As OracleConnection
    Public cmd As OracleCommand
    Public dr As OracleDataReader
    Public da As OracleDataAdapter
    Public ds As DataSet
    Public str As String
    Public pesan As String

    Sub koneksi()
        str = "Data source =(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 192.168.88.38)(PORT = 1521))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = bhihrdb)));User id = admin;Password = magang"
        con = New OracleConnection(str)
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
    End Sub
End Module
