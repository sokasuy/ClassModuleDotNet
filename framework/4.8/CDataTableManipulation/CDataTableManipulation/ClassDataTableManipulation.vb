Public Class CDataTableManipulation
    Private myCShowMessage As New CShowMessage.CShowMessage
    Public Sub BulkUpdateSingleColumn(ByRef dt As DataTable, columnName As String, newValue As Object, Optional rowFilter As String = "")
        Try
            ' Validasi kolom
            If Not dt.Columns.Contains(columnName) Then
                Throw New ArgumentException($"Kolom '{columnName}' tidak ditemukan di DataTable.")
            End If

            Dim rows() As DataRow

            If String.IsNullOrWhiteSpace(rowFilter) Then
                rows = dt.Select() ' Semua baris
            Else
                rows = dt.Select(rowFilter)
            End If
            For Each row In rows
                row(columnName) = newValue
            Next

            'For Each row As DataRow In dt.Select(rowFilter)
            '    row(columnName) = newValue
            'Next
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "BulkUpdateSingleColumn Error")
        End Try
    End Sub

    Public Sub BulkUpdateMultipleColumns(dt As DataTable, columnValues As Dictionary(Of String, Object), Optional rowFilter As String = "")
        Try
            ' Validasi kolom
            For Each colName In columnValues.Keys
                If Not dt.Columns.Contains(colName) Then
                    Throw New ArgumentException($"Kolom '{colName}' tidak ditemukan di DataTable.")
                End If
            Next

            ' Ambil baris yang sesuai
            Dim rows As DataRow() = If(String.IsNullOrWhiteSpace(rowFilter), dt.Select(), dt.Select(rowFilter))

            ' Update kolom di setiap baris
            For Each row In rows
                For Each kvp In columnValues
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
    ''' <param name="dt">DataTable yang ingin diubah</param>
    ''' <param name="columnName">Nama kolom yang akan diubah</param>
    ''' <param name="newType">Tipe baru (misalnya GetType(String))</param>
    Public Sub ConvertColumnType(ByVal dt As DataTable, ByVal columnName As String, ByVal newType As Type)
        Try
            If Not dt.Columns.Contains(columnName) Then
                Throw New ArgumentException($"Kolom '{columnName}' tidak ditemukan di DataTable.")
            End If

            ' Nama kolom baru sementara
            Dim tempColName As String = columnName & "_temp"

            ' Tambahkan kolom baru dengan tipe data baru
            dt.Columns.Add(tempColName, newType)

            ' Copy data lama ke kolom baru dengan konversi
            For Each row As DataRow In dt.Rows
                If row.IsNull(columnName) Then
                    row(tempColName) = DBNull.Value
                Else
                    Try
                        row(tempColName) = Convert.ChangeType(row(columnName), newType)
                    Catch ex As Exception
                        ' Jika gagal konversi, fallback ke string
                        row(tempColName) = row(columnName).ToString()
                    End Try
                End If
            Next

            ' Hapus kolom lama
            dt.Columns.Remove(columnName)

            ' Rename kolom baru ke nama asli
            dt.Columns(tempColName).ColumnName = columnName
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConvertColumnType Error")
        End Try
    End Sub
End Class
