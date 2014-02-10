Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Public Class Form1
    Dim row As Int16
    Dim counter As Integer
    Dim datable As New DataTable
    Dim id_hrd As String

    'Untuk Load id dari tabel training ke combobox
    Sub Load_IdTran()
        cmd = New OracleCommand("select ID_TRAN from T_TRAINING", con)
        dr = cmd.ExecuteReader
        While dr.Read
            ComboBox1.Items.Add(dr("ID_TRAN"))
        End While
    End Sub

    'Load Data peserta ke datagridview                                                                                                                                                                                                                                                                                               
    Sub Load_DataPeserta()
        id_hrd = ComboBox1.SelectedItem.ToString
        con.Close()

        da = New OracleDataAdapter("select B.npk, A.nama, D.N_JABATAN AS JABATAN, E.nama_departemen AS DEPARTEMEN" &
                                   " from karyawan A join t_training_mem B on A.NPK=B.NPK" &
                                   " join t_training C on B.ID_TRAN=C.ID_TRAN" &
                                   " join jab D on A.jab=D.ID_JABATAN" &
                                   " join dept E on A.Dept=E.ID_DEPARTEMEN where C.ID_TRAN='" + id_hrd + "'", con)
        ds = New DataSet
        da.Fill(ds, "karyawan")
        DataGridView1.DataSource = ds.Tables("karyawan")
        DataGridView1.ReadOnly = True
    End Sub

    'Load data transaksi training ke dataset untuk ditampilkan ke crystalreport berdasarkan Id
    Sub Load_DataUmum()
        id_hrd = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta," &
                                   " a.jenis_t,a.jenis_p,a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m," &
                                   " to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s, a.sumber,a.ket, f.nama as pembuat," &
                                   " j.n_jabatan as jab_buat,g.nama as pemeriksa, k.n_jabatan as jab_periksa," &
                                   " h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui," &
                                   " m.n_jabatan as jab_mengetahui, b.npk,c.nama,d.n_jabatan,e.nama_departemen" &
                                   " from t_training a left join t_training_mem b on a.id_tran=b.id_tran" &
                                   " left join karyawan c on b.npk=c.npk left join jab d on c.jab=d.id_jabatan" &
                                   " left join dept e on c.dept=e.id_departemen left join karyawan f on f.npk=a.lev_1" &
                                   " left join karyawan g on g.npk=a.lev_2 left join karyawan h on h.npk=a.lev_3" &
                                   " left join karyawan i on i.npk=a.lev_4 left join jab j on j.id_jabatan=f.jab" &
                                   " left join jab k on k.id_jabatan=g.jab left join jab l on l.id_jabatan=h.jab" &
                                   " left join jab m on m.id_jabatan=i.jab where a.id_tran='" + id_hrd + "'", con)
        da.Fill(ds, "t_training")
        da.Fill(datable)
        If datable.Rows.Count = 0 Then
            Label2.Text = "-"
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            RichTextBox1.Clear()
            TextBox10.Clear()
            TextBox11.Clear()
            TextBox12.Clear()
            Label21.Text = "-"
            Label22.Text = "-"
            Label23.Text = "-"
            Label24.Text = "-"
        Else
            Label2.Text = datable.Rows(0)("status").ToString()
            TextBox1.Text = datable.Rows(0)("tanggal").ToString()
            TextBox2.Text = datable.Rows(0)("id_minta").ToString()
            TextBox3.Text = datable.Rows(0)("jenis_t").ToString()
            TextBox4.Text = datable.Rows(0)("jenis_p").ToString()
            TextBox5.Text = datable.Rows(0)("sumber").ToString()
            RichTextBox1.Text = datable.Rows(0)("ket").ToString()
            TextBox10.Text = datable.Rows(0)("tempat").ToString()
            TextBox11.Text = datable.Rows(0)("tanggal_m").ToString()
            TextBox12.Text = datable.Rows(0)("tanggal_s").ToString()
            Label21.Text = datable.Rows(0)("pembuat").ToString()
            Label22.Text = datable.Rows(0)("pemeriksa").ToString()
            Label23.Text = datable.Rows(0)("menyetujui").ToString()
            Label24.Text = datable.Rows(0)("mengetahui").ToString()
        End If

        datable.Clear()
    End Sub

    'Load semua data peserta training ke dataset untuk dicetak ke crystalreport
    Sub CetakSemua()
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta, a.jenis_t," &
                                   "a.jenis_p,a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m,to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s," &
                                   " a.sumber,a.ket, f.nama as pembuat,j.n_jabatan as jab_buat,g.nama as pemeriksa, k.n_jabatan as jab_periksa," &
                                   "h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui, m.n_jabatan as jab_mengetahui," &
                                   " b.npk,c.nama,d.n_jabatan,e.nama_departemen from t_training a left join t_training_mem b on a.id_tran=b.id_tran" &
                                   " left join karyawan c on b.npk=c.npk left join jab d on c.jab=d.id_jabatan left join dept e on c.dept=e.id_departemen" &
                                   " left join karyawan f on f.npk=a.lev_1 left join karyawan g on g.npk=a.lev_2 left join karyawan h on h.npk=a.lev_3" &
                                   " left join karyawan i on i.npk=a.lev_4 left join jab j on j.id_jabatan=f.jab left join jab k on k.id_jabatan=g.jab" &
                                   " left join jab l on l.id_jabatan=h.jab left join jab m on m.id_jabatan=i.jab", con)
        ds = New DataSet
        da.Fill(ds, "t_training")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call koneksi()
        Call Load_DataPeserta()
        Call Load_DataUmum()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        id_hrd = ComboBox1.SelectedItem.ToString
        If id_hrd = "All" Then
            If (MsgBox("Lanjutkan  Semua Data?", vbYesNoCancel) = MsgBoxResult.Yes) Then
                Call CetakSemua()
                Form2.Show()
            End If
        Else
            Form2.Show()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call koneksi()
        Form3.Show()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call Load_IdTran()
    End Sub

End Class
