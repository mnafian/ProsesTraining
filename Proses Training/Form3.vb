Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Public Class Form3
    'Pencarian data peserta training
    Dim jenis_p As String
    Dim nama As String

    'Load Id transaksi ke combobox.
    Sub LoadJenisTran()
        cmd = New OracleCommand("select distinct JENIS_P from T_TRAINING", con)
        dr = cmd.ExecuteReader
        While dr.Read
            ComboBox1.Items.Add(dr("JENIS_P"))
        End While
    End Sub

    'Untuk load data training berdasarkan jenis training yang diikuti ke datagridview.
    Sub LoadDataByJenis()
        jenis_p = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta,a.jenis_t," &
                                   "a.jenis_p,b.npk,c.nama,d.n_jabatan,e.nama_departemen,a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m," &
                                   "to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s,a.sumber,a.ket from t_training a" &
                                   " join t_training_mem b on a.id_tran=b.id_tran join karyawan c on b.npk=c.npk" &
                                   " join jab d on c.jab=d.id_jabatan join dept e on c.dept=e.id_departemen where a.jenis_p='" + jenis_p + "'", con)
        ds = New DataSet
        da.Fill(ds, "t_training")
        DataGridView1.DataSource = ds.Tables("t_training")
        DataGridView1.ReadOnly = True
    End Sub

    'Untuk cetak semua data training berdasarkan jenis training yang diikuti
    Sub CetakDataByJenis()
        jenis_p = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta, a.jenis_t,a.jenis_p," &
                                   " a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m,to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s," &
                                   " a.sumber,a.ket, f.nama as pembuat,j.n_jabatan as jab_buat,g.nama as pemeriksa, k.n_jabatan as jab_periksa," &
                                   " h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui, m.n_jabatan as jab_mengetahui," &
                                   " b.npk,c.nama,d.n_jabatan,e.nama_departemen from t_training a " &
                                   " left join t_training_mem b on a.id_tran=b.id_tran left join karyawan c on b.npk=c.npk " &
                                   " left join jab d on c.jab=d.id_jabatan left join dept e on c.dept=e.id_departemen " &
                                   " left join karyawan f on f.npk=a.lev_1 left join karyawan g on g.npk=a.lev_2 " &
                                   " left join karyawan h on h.npk=a.lev_3 left join karyawan i on i.npk=a.lev_4 " &
                                   " left join jab j on j.id_jabatan=f.jab left join jab k on k.id_jabatan=g.jab " &
                                   " left join jab l on l.id_jabatan=h.jab left join jab m on m.id_jabatan=i.jab where a.jenis_p='" + jenis_p + "'", con)
        ds = New DataSet
        da.Fill(ds, "t_training")
    End Sub

    'Untuk pencarian data training dan ditampilkan ke datagridview
    Sub LoadDataByTxt()
        jenis_p = ComboBox1.SelectedItem.ToString
        nama = TextBox1.Text.ToString.ToUpper
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta,a.jenis_t,a.jenis_p,b.npk,c.nama,d.n_jabatan,e.nama_departemen,a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m,to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s,a.sumber,a.ket,f.nama as pembuat,g.nama as pemeriksa,h.nama as menyetujui,i.nama as mengetahui from t_training a join t_training_mem b on a.id_tran=b.id_tran join karyawan c on b.npk=c.npk join jab d on c.jab=d.id_jabatan join dept e on c.dept=e.id_departemen join karyawan f on f.npk=a.lev_1 join karyawan g on g.npk=a.lev_2 join karyawan h on h.npk=a.lev_3 join karyawan i on i.npk=a.lev_4 where a.jenis_p='" + jenis_p + "' and c.nama like '%" + nama + "%'", con)
        ds = New DataSet
        da.Fill(ds, "t_training")
        DataGridView1.DataSource = ds.Tables("t_training")
        DataGridView1.ReadOnly = True
    End Sub

    'Untuk cetak hasil pencarian data training dan diisi ke dataset untuk dicetak ke crystalreport.
    Sub CetakDataByTxt()
        jenis_p = ComboBox1.SelectedItem.ToString
        nama = TextBox1.Text.ToString.ToUpper
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.tgl_cr, 'fmdd MON yyyy')as tanggal,a.id_minta, a.jenis_t,a.jenis_p,a.tempat,to_char(a.tgl_1, 'fmdd MON yyyy')as tanggal_m,to_char(a.tgl_2, 'fmdd MON yyyy')as tanggal_s, a.sumber,a.ket, f.nama as pembuat,j.n_jabatan as jab_buat,g.nama as pemeriksa, k.n_jabatan as jab_periksa,h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui, m.n_jabatan as jab_mengetahui, b.npk,c.nama,d.n_jabatan,e.nama_departemen from t_training a left join t_training_mem b on a.id_tran=b.id_tran left join karyawan c on b.npk=c.npk left join jab d on c.jab=d.id_jabatan left join dept e on c.dept=e.id_departemen left join karyawan f on f.npk=a.lev_1 left join karyawan g on g.npk=a.lev_2 left join karyawan h on h.npk=a.lev_3 left join karyawan i on i.npk=a.lev_4 left join jab j on j.id_jabatan=f.jab left join jab k on k.id_jabatan=g.jab left join jab l on l.id_jabatan=h.jab left join jab m on m.id_jabatan=i.jab where a.jenis_p='" + jenis_p + "' and c.nama like '%" + nama + "%'", con)
        ds = New DataSet
        da.Fill(ds, "t_training")
        DataGridView1.DataSource = ds.Tables("t_training")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call LoadJenisTran()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call koneksi()
        Call LoadDataByJenis()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call koneksi()
        jenis_p = ComboBox1.SelectedItem.ToString
        If jenis_p = "Jenis Training" Then
            MsgBox("Pilih Jenis Training Terlebih Dahulu")
        Else
            Call LoadDataByTxt()
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        'Percabangan untuk antisipasi user belum memilih jenis training
        jenis_p = ComboBox1.SelectedItem.ToString
        If jenis_p = "Jenis Training" Then
            MsgBox("Pilih Jenis Training Terlebih Dahulu")
        ElseIf TextBox1.Text = "" Then
            Call CetakDataByJenis()
            Form4.Show()
        ElseIf TextBox1.Text <> "" Then
            Call CetakDataByTxt()
            Form4.Show()
        End If
    End Sub
End Class