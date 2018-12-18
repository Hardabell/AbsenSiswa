Imports System.Data.SqlClient
Public Class Form1
    Dim Conn As SqlConnection
    Dim Da As SqlDataAdapter
    Dim Ds As DataSet
    Dim Cmd As SqlCommand
    Dim RD As SqlDataReader
    Dim LokasiDB As String
    Sub Koneksi()
        LokasiDB = "data source=DESKTOP-VS68HBG;initial catalog=SIMABSENSI;integrated security=true;UID=sa;PWD=lostsaga12"
        Conn = New SqlConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Sub KondisiAwal()
        Koneksi()
        Da = New SqlDataAdapter("Select * from tbSiswa", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "tbSiswa")
        DataGridView1.DataSource = (Ds.Tables("tbSiswa"))
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("PRIA")
        ComboBox1.Items.Add("WANITA")
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Dim rn As New Random

        Dim n = rn.Next(1, 9999999)
        TextBox1.Text = n

        TextBox1.Enabled = False

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data belum lengkap")
        Else
            Call Koneksi()
            Dim simpan As String = "insert into tbSiswa (NIS,NamaSiswa,JenisKelamin,AlamatSiswa,NamaWali,NomorWali) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "')"
            Cmd = New SqlCommand(simpan, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Berhasil Input Data Siswa")
            Call KondisiAwal()
        End If
    End Sub
    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("Select * From TBL_SISWA  where NISN='" & TextBox1.Text & "'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("NISN Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("NamaSiswa")
                TextBox3.Text = RD.Item("Kelas")
                TextBox4.Text = RD.Item("TeleponOrangTua")
                ComboBox1.Text = RD.Item("JenisKelamin")
                TextBox2.Focus()
            End If
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Koneksi()
        Dim edit As String = "update TBL_SISWA set NamaSiswa='" & TextBox2.Text & "',Kelas='" & TextBox3.Text & "',TeleponOrangTua='" & TextBox4.Text & "',JenisKelamin='" & ComboBox1.Text & "' where NISN='" & TextBox1.Text & "'"
        Cmd = New SqlCommand(edit, Conn)
        CMD.ExecuteNonQuery()
        MsgBox("Data Siswa Berhasil di Update")
        Call KondisiAwal()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang akan di hapus dengan Masukan NISN dan ENTER")
        Else
            If MessageBox.Show("Anda yakin ingin menghapus siswa tersebut?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As SqlCommand
                Dim hapus As String = "delete From TBL_SISWA  where NISN='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, Conn)
                CMD.ExecuteNonQuery()
                MsgBox("Data Siswa Berhasil diHapus")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
