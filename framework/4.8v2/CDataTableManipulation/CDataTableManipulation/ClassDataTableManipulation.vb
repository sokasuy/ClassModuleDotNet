Public Class CDataTableManipulation
    Private myCShowMessage As New CShowMessage.CShowMessage
    Public Sub BulkUpdateSingleColumn(ByRef _dt As DataTable, _columnName As String, _newValue As Object, Optional _rowFilter As String = "")
        Try
            ' Validasi kolom
            If Not _dt.Columns.Contains(_columnName) Then
                Throw New ArgumentException($"Kolom '{_columnName}' tidak ditemukan di DataTable.")
            End If

            Dim rows() As DataRow

            If String.IsNullOrWhiteSpace(_rowFilter) Then
                rows = _dt.Select() ' Semua baris
            Else
                rows = _dt.Select(_rowFilter)
            End If
            For Each row In rows
                row(_columnName) = _newValue
            Next

            'For Each row As DataRow In dt.Select(rowFilter)
            '    row(columnName) = newValue
            'Next
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "BulkUpdateSingleColumn Error")
        End Try
    End Sub

    Public Sub BulkUpdateMultipleColumns(_dt As DataTable, _columnValues As Dictionary(Of String, Object), Optional _rowFilter As String = "")
        Try
            ' Validasi kolom
            For Each colName In _columnValues.Keys
                If Not _dt.Columns.Contains(colName) Then
                    Throw New ArgumentException($"Kolom '{colName}' tidak ditemukan di DataTable.")
                End If
            Next

            ' Ambil baris yang sesuai
            Dim rows As DataRow() = If(String.IsNullOrWhiteSpace(_rowFilter), _dt.Select(), _dt.Select(_rowFilter))

            ' Update kolom di setiap baris
            For Each row In rows
                For Each kvp In _columnValues
                    row(kvp.Key) = kvp.Value
                Next
            Next
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "BulkUpdateMultipleColumns Error")
        End Try
    End Sub

    ''' <summary>
    ''' Mengubah tipe kolom di DataTable
    ''' </summary>
    ''' <param name="_dt">DataTable yang ingin diubah</param>
    ''' <param name="_columnName">Nama kolom yang akan diubah</param>
    ''' <param name="_newType">Tipe baru (misalnya GetType(String))</param>
    Public Sub ConvertColumnType(ByVal _dt As DataTable, ByVal _columnName As String, ByVal _newType As Type)
        Try
            If Not _dt.Columns.Contains(_columnName) Then
                Throw New ArgumentException($"Kolom '{_columnName}' tidak ditemukan di DataTable.")
            End If

            ' Nama kolom baru sementara
            Dim tempColName As String = _columnName & "_temp"

            ' Tambahkan kolom baru dengan tipe data baru
            _dt.Columns.Add(tempColName, _newType)

            ' Copy data lama ke kolom baru dengan konversi
            For Each row As DataRow In _dt.Rows
                If row.IsNull(_columnName) Then
                    row(tempColName) = DBNull.Value
                Else
                    Try
                        row(tempColName) = Convert.ChangeType(row(_columnName), _newType)
                    Catch ex As Exception
                        ' Jika gagal konversi, fallback ke string
                        row(tempColName) = row(_columnName).ToString()
                    End Try
                End If
            Next

            ' Hapus kolom lama
            _dt.Columns.Remove(_columnName)

            ' Rename kolom baru ke nama asli
            _dt.Columns(tempColName).ColumnName = _columnName
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConvertColumnType Error")
        End Try
    End Sub

    ''' <summary>
    ''' Cloning datatable dengan Filter
    ''' </summary>
    ''' <param name="_source">DataTable yang ingin di clone</param>
    ''' <param name="_filter">Filter kolom</param>
    ''' <param name="_columns">Filter jika mau ambil kolom tertentu saja yang di clone</param>
    Public Function CloneWithFilter(_source As DataTable, _filter As String, Optional _columns() As String = Nothing) As DataTable
        Try
            Dim view As New DataView(_source)
            view.RowFilter = _filter

            If _columns IsNot Nothing AndAlso _columns.Length > 0 Then
                Return view.ToTable(True, _columns) ' True = duplikat dihapus
            Else
                Return view.ToTable()
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CloneWithFilter Error")
            Return Nothing
        End Try
    End Function

    Public Sub RenameColumn(ByRef _dt As DataTable, _oldName As String, _newName As String)
        Try
            ' Pastikan kolom lama ada
            If Not _dt.Columns.Contains(_oldName) Then
                Throw New ArgumentException("Kolom dengan nama '" & _oldName & "' tidak ditemukan.")
            End If

            ' Pastikan kolom baru belum ada
            If _dt.Columns.Contains(_newName) Then
                Throw New ArgumentException("Kolom dengan nama '" & _newName & "' sudah ada.")
            End If

            ' Ganti nama kolom
            _dt.Columns(_oldName).ColumnName = _newName

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RenameColumn Error")
        End Try
    End Sub
End Class
