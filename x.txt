Imports System.Data.OleDb
Public Class NguyenChiBao
    Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & Application.StartupPath & "\QLBS.accdb;"
    Dim Cn As New OleDbConnection(connStr)
    Dim daKhoa As OleDbDataAdapter
    Dim daBacSi As OleDbDataAdapter

    Dim tKhoa As New DataTable
    Dim BindingKhoa As New BindingSource
    Dim tBacSi As New DataTable
    Dim BindingBacSi As New BindingSource

    Private Sub NguyenChiBao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Cn.Open()

        BindingKhoa.DataSource = tKhoa
        DataGridView1.DataSource = BindingKhoa
        BindingBacSi.DataSource = tBacSi
        DataGridView2.DataSource = BindingBacSi

        BindingNavigator1.BindingSource = BindingKhoa
        BindingNavigator2.BindingSource = BindingBacSi
    End Sub

    Private Sub btnXemDuLieu2Grid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnXemDuLieu2Grid.Click
        tKhoa.Clear()
        tBacSi.Clear()

        Dim sKhoa As String = "SELECT * FROM Khoa"
        daKhoa = New OleDbDataAdapter(sKhoa, Cn)
        daKhoa.Fill(tKhoa)
        DataGridView1.DataSource = BindingKhoa

        Dim sBacSi As String = "SELECT * FROM BacSi"
        daBacSi = New OleDbDataAdapter(sBacSi, Cn)
        daBacSi.Fill(tBacSi)
        DataGridView2.DataSource = BindingBacSi
    End Sub

    Private Sub btnXoaDuLieuTrenGrid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnXoaDuLieuTrenGrid.Click
        DataGridView1.DataSource = Nothing
        DataGridView2.DataSource = Nothing
    End Sub

    Private Sub btnXemDSBSThem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnXemDSBSThem.Click
        DataGridView3.DataSource = tBacSi.GetChanges(DataRowState.Added)
    End Sub

    Private Sub btnBoQua_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBoQua.Click
        DataGridView3.DataSource = Nothing
        tBacSi.RejectChanges()
    End Sub

    Private Sub btnLuuDSBSThem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLuuDSBSThem.Click
        Dim insertCommand As New OleDbCommand("INSERT INTO BacSi (MaBS, Ho, Ten, Nu, NgaySinh, MaKhoa) VALUES (@MaBS, @Ho, @Ten, @Nu, @NgaySinh, @MaKhoa)", Cn)
        insertCommand.Parameters.Add("@MaBS", OleDbType.VarChar, 10, "MaBS")
        insertCommand.Parameters.Add("@Ho", OleDbType.VarChar, 25, "Ho")
        insertCommand.Parameters.Add("@Ten", OleDbType.VarChar, 7, "Ten")
        insertCommand.Parameters.Add("@Nu", OleDbType.Boolean, 1, "Nu")
        insertCommand.Parameters.Add("@NgaySinh", OleDbType.Date, 8, "NgaySinh")
        insertCommand.Parameters.Add("@MaKhoa", OleDbType.VarChar, 10, "MaKhoa")
        daBacSi.InsertCommand = insertCommand

        'Lấy các thay đổi từ DataTable (danh sách các hàng đã thêm)
        Dim changes As DataTable = tBacSi.GetChanges(DataRowState.Added)

        'Xác nhận và cập nhật các thay đổi vào CSDL
        daBacSi.Update(changes)

        'Đánh dấu các thay đổi đã được lưu
        tBacSi.AcceptChanges()

        DataGridView3.DataSource = Nothing
    End Sub

    Private Sub btnKetThuc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKetThuc.Click
        Cn.Close()
        Me.Close()
    End Sub
End Class