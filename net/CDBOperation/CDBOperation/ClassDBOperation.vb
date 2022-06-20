Imports System.Windows.Forms
'Imports CShowMessage.CShowMessage
'Imports CStringManipulation.CStringManipulation
'Imports CMiscFunction.CMiscFunction
'Imports CDBConnection.CDBConnection
'Imports System.Data.SqlClient

Public Class CDBOperation
    Private myCShowMessage As New CShowMessage.CShowMessage
    Private myCStringManipulation As New CStringManipulation.CStringManipulation
    Private myCMiscFunction As New CMiscFunction.CMiscFunction
    Private myCDBConnection As New CDBConnection.CDBConnection

    Public Sub CommandConstructor(_conn As Object, ByRef _comm As Object, _text As String, _queryType As String, Optional _checkPetik As Boolean = False, Optional _isTransaction As Boolean = False, Optional _transaction As Object = Nothing)
        'OK
        'untuk menyusun command / query yang akan digunakan OLEDB Connection
        Try
            If (TypeOf _conn Is OleDb.OleDbConnection) Then
                'Jika koneksi oledb (access, sql)
                _comm = New OleDb.OleDbCommand
            ElseIf (TypeOf _conn Is Npgsql.NpgsqlConnection) Then
                'Jika koneksi postgresql
                _comm = New Npgsql.NpgsqlCommand
                'defaultnya gak dikasih petik
                If _checkPetik Then
                    'Klw harus dikasih petik
                    _text = myCStringManipulation.SetStringForQueryPgsql(_queryType, _text)
                End If
            ElseIf (TypeOf _conn Is SqlConnection) Then
                _comm = New SqlCommand
            End If
            If (_isTransaction) Then
                _comm.Transaction = _transaction
            End If
            _comm.Connection = _conn
            _comm.CommandType = CommandType.Text
            _comm.CommandText = _text
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CommandConsructor Error")
        End Try
    End Sub

    Public Function ExecuteReaderFunction(_comm As Object) As Object
        'OK
        'untuk membaca data yang diquery dari database
        ExecuteReaderFunction = Nothing
        Try
            ExecuteReaderFunction = _comm.ExecuteReader
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExecuteReaderFunction Error")
        End Try
    End Function

    Public Sub ExecuteNonQueryFunction(_comm As Object, Optional _updateString As String = "")
        'OK
        'untuk menjalankan segala macam perintah manipulasi database (INSERT, UPDATE, DELETE,)
        'yang sebelumnya telah dibuat command text nya
        Try
            _comm.ExecuteNonQuery()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExecuteNonQueryFunction Error")
        End Try
    End Sub

    Public Function GetDataTableUsingAdapterOle(_conn As OleDb.OleDbConnection, _dataAdapter As OleDb.OleDbDataAdapter, _stSQL As String, _tablename As String) As DataTable
        'OK
        'untuk membuat satu tabel yang sumbernya diambil dari database
        'diperlukan untuk mengisi datagrid, flexgrid, trueDBCombo, dll yang sifatnya mengambil satu paket records, seperti klw ingin mengambil satu isi tabel tertentu
        '>>>NOTE: agar bisa mengupdate data tabel yang sourcenya dari multiple table, dataAdapter yang digunakan harus berbeda, kalau tidak akan tampil error seperti ini
        'An unhandled exception of type 'System.InvalidOperationException' occurred in system.data.dll
        'Additional information: Missing the DataColumn 'KOLOM_A' in the DataTable 'NAMA DATA TABEL' for the SourceColumn 'KOLOM_A'.
        Try
            Dim dtSet As New DataSet
            _dataAdapter = New OleDb.OleDbDataAdapter(_stSQL, _conn)
            Dim cb As New OleDb.OleDbCommandBuilder(_dataAdapter)
            dtSet.Tables.Add(New DataTable(_tablename))
            'dtSet = GetDataSet(Conn, OleSQL, tablename)
            _dataAdapter.FillSchema(dtSet, SchemaType.Mapped, _tablename)
            _dataAdapter.Fill(dtSet.Tables(_tablename))
            GetDataTableUsingAdapterOle = dtSet.Tables(_tablename)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetDataTableUsingAdapterOle Error")
            GetDataTableUsingAdapterOle = Nothing
        End Try
    End Function

    Public Function GetDataTableUsingAdapterNpgsql(_conn As Npgsql.NpgsqlConnection, _dataAdapter As Npgsql.NpgsqlDataAdapter, _stSQL As String, _tablename As String) As DataTable
        'OK
        'untuk membuat satu tabel yang sumbernya diambil dari database
        'diperlukan untuk mengisi datagrid, flexgrid, trueDBCombo, dll yang sifatnya mengambil satu paket records, seperti klw ingin mengambil satu isi tabel tertentu
        '>>>NOTE: agar bisa mengupdate data tabel yang sourcenya dari multiple table, dataAdapter yang digunakan harus berbeda, kalau tidak akan tampil error seperti ini
        'An unhandled exception of type 'System.InvalidOperationException' occurred in system.data.dll
        'Additional information: Missing the DataColumn 'KOLOM_A' in the DataTable 'NAMA DATA TABEL' for the SourceColumn 'KOLOM_A'.
        Try
            Dim dtSet As New DataSet
            _stSQL = myCStringManipulation.SetStringForComm(_conn, "select", _stSQL)
            _dataAdapter = New Npgsql.NpgsqlDataAdapter(_stSQL, _conn)
            Dim cb As New Npgsql.NpgsqlCommandBuilder(_dataAdapter)
            dtSet.Tables.Add(New DataTable(_tablename))
            'dtSet = GetDataSet(Conn, OleSQL, tablename)
            _dataAdapter.FillSchema(dtSet, SchemaType.Mapped, _tablename)
            _dataAdapter.Fill(dtSet.Tables(_tablename))
            GetDataTableUsingAdapterNpgsql = dtSet.Tables(_tablename)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetDataTableUsingAdapterNpgsql Error")
            GetDataTableUsingAdapterNpgsql = Nothing
        End Try
    End Function

    Public Function GetDataTableUsingReader(_conn As Object, _comm As Object, _reader As Object, _stSQL As String, _tablename As String, Optional _checkPetik As Boolean = False) As DataTable
        'OK
        'untuk membuat satu tabel yang sumbernya diambil dari database
        'diperlukan untuk mengisi datagrid, flexgrid, trueDBCombo, dll yang sifatnya mengambil satu paket records, seperti klw ingin mengambil satu isi tabel tertentu
        '>>>NOTE: agar bisa mengupdate data tabel yang sourcenya dari multiple table, dataAdapter yang digunakan harus berbeda, kalau tidak akan tampil error seperti ini
        'An unhandled exception of type 'System.InvalidOperationException' occurred in system.data.dll
        'Additional information: Missing the DataColumn 'KOLOM_A' in the DataTable 'NAMA DATA TABEL' for the SourceColumn 'KOLOM_A'.
        Try
            'menggunakan reader
            Dim dtSet As New DataSet
            'tString = myCStringManipulation.SetStringForComm(conn, "select", oleSQL)
            Call CommandConstructor(_conn, _comm, _stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)
            dtSet.Tables.Add(New DataTable(_tablename))
            dtSet.Load(_reader, LoadOption.PreserveChanges, _tablename)
            GetDataTableUsingReader = dtSet.Tables(_tablename)
            _reader.close()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetDataTableUsingReader Error")
            GetDataTableUsingReader = Nothing
        End Try
    End Function

    Public Sub FillDataSet(_conn As OleDb.OleDbConnection, _dataAdapter As OleDb.OleDbDataAdapter, ByRef _dtSet As DataSet, _stSQL As String, _tablename As String)
        'OK
        'untuk membuat semacam data set yang dapat berisi lebih dari satu tabel (array of table)
        '>>>NOTE: agar bisa mengupdate data tabel yang sourcenya dari multiple table, _dataAdapter yang digunakan harus berbeda, kalau tidak akan tampil error seperti ini
        'An unhandled exception of type 'System.InvalidOperationException' occurred in system.data.dll
        'Additional information: Missing the DataColumn 'KOLOM_A' in the DataTable 'NAMA DATA TABEL' for the SourceColumn 'KOLOM_A'.
        Try
            _stSQL = myCStringManipulation.SetStringForComm(_conn, "select", _stSQL)
            _dataAdapter = New OleDb.OleDbDataAdapter(_stSQL, _conn)
            Dim cb As New OleDb.OleDbCommandBuilder(_dataAdapter)
            _dtSet.Tables.Add(New DataTable(_tablename))
            _dataAdapter.FillSchema(_dtSet, SchemaType.Mapped, _tablename)
            _dataAdapter.Fill(_dtSet.Tables(_tablename))
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "FillDataSet Error")
        End Try
    End Sub

    Public Function GetDataIndividual(_conn As Object, _comm As Object, _reader As Object, _stSQL As String, Optional _checkPetik As Boolean = False) As String
        'OK
        'untuk mengambil satu data dengan menggunakan ole _reader
        Try
            _stSQL = Trim(_stSQL)
            'tString = myCStringManipulation.SetStringForComm(_conn, "select", stSql)
            Call CommandConstructor(_conn, _comm, _stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            GetDataIndividual = Nothing

            While _reader.Read
                If Not (IsDBNull(_reader(0))) Then
                    GetDataIndividual = _reader(0)
                End If
            End While
            _reader.close()
        Catch ex As Exception
            GetDataIndividual = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetDataIndividual Error")
        End Try
    End Function

    Public Sub ConstructorInsertData(_conn As Object, _comm As Object, _reader As Object, _myDataTableSource As DataTable, _myTableDestination As String, Optional _mySchema As String = "", Optional _myKriteria As String = Nothing)
        Try
            Dim mJumlahKolom As Integer
            Dim mNamaKolom(mJumlahKolom) As String
            Dim idx As Integer
            Dim newValues As String
            Dim newFields As String

            mJumlahKolom = _myDataTableSource.Columns.Count
            ReDim mNamaKolom(mJumlahKolom)
            For i As Integer = 0 To mJumlahKolom - 1
                mNamaKolom(i) = _myDataTableSource.Columns.Item(i).ColumnName
            Next

            idx = 0
            While idx < _myDataTableSource.Rows.Count
                If Not IsNothing(_myKriteria) Then
                    If (_myDataTableSource.Rows(idx).Item(_myKriteria) = True) Then
                        'IS NEW = TRUE YANG DI INSERT
                        newValues = Nothing
                        newFields = Nothing
                        For i As Integer = 0 To mJumlahKolom - 1
                            If Not IsNothing(newValues) Then
                                newValues &= "," & myCStringManipulation.SetInsertValues(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                newFields &= "," & myCStringManipulation.SetInsertFields(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).ColumnName)
                            Else
                                newValues = myCStringManipulation.SetInsertValues(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                newFields = myCStringManipulation.SetInsertFields(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).ColumnName)
                            End If
                        Next
                        If (_mySchema.Length > 0) Then
                            Call InsertData(_conn, _comm, _mySchema & "." & _myTableDestination, newValues, newFields)
                        Else
                            Call InsertData(_conn, _comm, _myTableDestination, newValues, newFields)
                        End If
                    End If
                Else
                    newValues = Nothing
                    newFields = Nothing
                    For i As Integer = 0 To mJumlahKolom - 1
                        If Not IsNothing(newValues) Then
                            newValues &= "," & myCStringManipulation.SetInsertValues(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                            newFields &= "," & myCStringManipulation.SetInsertFields(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).ColumnName)
                        Else
                            newValues = myCStringManipulation.SetInsertValues(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                            newFields = myCStringManipulation.SetInsertFields(_myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).ColumnName)
                        End If
                    Next
                    If (_mySchema.Length > 0) Then
                        Call InsertData(_conn, _comm, _mySchema & "." & _myTableDestination, newValues, newFields)
                    Else
                        Call InsertData(_conn, _comm, _myTableDestination, newValues, newFields)
                    End If
                End If

                idx += 1
                If (idx Mod 1000 = 0) Then
                    GC.Collect()
                End If
            End While
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConstructorInsertData Error")
        End Try
    End Sub

    Public Sub ConstructorUpdateData(_conn As Object, _comm As Object, _reader As Object, _myQueryDestination As String, _myDataTableSource As DataTable, _myTableDestination As String, _myUniqueIdentifier As String, Optional _mySchema As String = "", Optional _myKriteria As String = Nothing)
        Try
            'MASIH ERROR HARUS DIPERBAIKI DULU
            '18 MARET 2019
            Dim idx As Integer
            Dim updateString As String
            Dim whereString As String
            Dim mJumlahKolom As Integer
            Dim myDataTableDestination As New DataTable

            mJumlahKolom = _myDataTableSource.Columns.Count
            While idx < _myDataTableSource.Rows.Count
                If Not IsNothing(_myKriteria) Then
                    If (_myDataTableSource.Rows(idx).Item(_myKriteria) = False) Then
                        'ISNEW = FALSE YANG DI UPDATE
                        'UNTUK UPDATE BB SUBMATERIAL YANG LAMA
                        updateString = Nothing
                        myDataTableDestination = GetDataTableUsingReader(_conn, _comm, _reader, _myQueryDestination, "T_" & _myTableDestination)
                        For i As Integer = 0 To mJumlahKolom - 1
                            If _myDataTableSource.Columns(i).ColumnName <> _myKriteria Then
                                'Hanya kalau nama kolomnya tidak sama dengan nama kriterianya
                                'contoh kriterianya isnew, jadi kolom isnew tidak perlu di cek
                                If Not IsDBNull(_myDataTableSource.Rows(idx).Item(i)) Then
                                    Select Case _myDataTableSource.Columns(i).DataType.Name
                                        Case "Byte", "Decimal", "Double", "Int16", "Int32", "Int64", "SByte", "Single", "UInt16", "UInt32", "UInt64"
                                            If IsDBNull(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) Then
                                                myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName) = 0
                                            End If
                                            If (Double.Parse(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) <> Double.Parse(IIf(IsDBNull(_myDataTableSource.Rows(idx).Item(i)), 0, _myDataTableSource.Rows(idx).Item(i)))) Then
                                                If Not IsNothing(updateString) Then
                                                    updateString &= "," & myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                                Else
                                                    updateString = myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                                End If
                                            End If
                                        Case Else
                                            If IsDBNull(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) Then
                                                myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName) = ""
                                            End If
                                            If (Trim(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) <> IIf(IsDBNull(_myDataTableSource.Rows(idx).Item(i)), "", Trim(_myDataTableSource.Rows(idx).Item(i)))) Then
                                                If Not IsNothing(updateString) Then
                                                    updateString &= "," & myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                                Else
                                                    updateString = myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                                End If
                                            End If
                                    End Select
                                End If
                            End If
                        Next
                        If Not IsNothing(updateString) Then
                            whereString = _myUniqueIdentifier & "=" & myDataTableDestination.Rows(0).Item(_myUniqueIdentifier)
                            If (_mySchema.Length > 0) Then
                                Call UpdateData(_conn, _comm, _mySchema & "." & _myTableDestination, updateString, whereString)
                            Else
                                Call UpdateData(_conn, _comm, _myTableDestination, updateString, whereString)
                            End If
                        End If
                    End If
                Else
                    updateString = Nothing
                    myDataTableDestination = GetDataTableUsingReader(_conn, _comm, _reader, _myQueryDestination, "T_" & _myTableDestination)
                    For i As Integer = 0 To mJumlahKolom - 1
                        If Not IsDBNull(_myDataTableSource.Rows(idx).Item(i)) Then
                            Select Case _myDataTableSource.Columns(i).DataType.Name
                                Case "Byte", "Decimal", "Double", "Int16", "Int32", "Int64", "SByte", "Single", "UInt16", "UInt32", "UInt64"
                                    If IsDBNull(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) Then
                                        myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName) = 0
                                    End If
                                    If (Double.Parse(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) <> Double.Parse(IIf(IsDBNull(_myDataTableSource.Rows(idx).Item(i)), 0, _myDataTableSource.Rows(idx).Item(i)))) Then
                                        If Not IsNothing(updateString) Then
                                            updateString &= "," & myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                        Else
                                            updateString = myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                        End If
                                    End If
                                Case Else
                                    If IsDBNull(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) Then
                                        myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName) = ""
                                    End If
                                    If (Trim(myDataTableDestination.Rows(0).Item(_myDataTableSource.Columns(i).ColumnName)) <> IIf(IsDBNull(_myDataTableSource.Rows(idx).Item(i)), "", Trim(_myDataTableSource.Rows(idx).Item(i)))) Then
                                        If Not IsNothing(updateString) Then
                                            updateString &= "," & myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                        Else
                                            updateString = myCStringManipulation.SetUpdateString(_myDataTableSource.Columns(i).ColumnName, _myDataTableSource.Rows(idx).Item(i), _myDataTableSource.Columns(i).DataType.Name)
                                        End If
                                    End If
                            End Select
                        End If
                    Next
                    If Not IsNothing(updateString) Then
                        whereString = _myUniqueIdentifier & "=" & myDataTableDestination.Rows(0).Item(_myUniqueIdentifier)
                        If (_mySchema.Length > 0) Then
                            Call UpdateData(_conn, _comm, _mySchema & "." & _myTableDestination, updateString, whereString)
                        Else
                            Call UpdateData(_conn, _comm, _myTableDestination, updateString, whereString)
                        End If
                    End If
                End If
                idx += 1
                If (idx Mod 1000 = 0) Then
                    GC.Collect()
                End If
            End While
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConstructorUpdateData Error")
        End Try
    End Sub

    Public Function TransactionData(_conn As Object, _comm As Object, _transaction As Object, _strQuery As String, Optional _checkPetik As Boolean = False) As Boolean
        Try
            'Dim tr_ As SqlTransaction = _conn.BeginTransaction(IsolationLevel.ReadCommitted)
            _transaction = _conn.BeginTransaction(IsolationLevel.ReadCommitted)
            Call CommandConstructor(_conn, _comm, _strQuery, "insert", _checkPetik, True, _transaction)
            Call ExecuteNonQueryFunction(_comm)

            Try
                _transaction.Commit()
                TransactionData = True
            Catch ex As SqlException
                _transaction.Rollback()
                TransactionData = False
                Call myCShowMessage.ShowErrMsg(ex.Message)
            Finally
                _transaction.Dispose()
            End Try
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "TransactSQLInsert Error")
            Return False
        End Try
    End Function

    Public Sub InsertData(_conn As Object, _comm As Object, _tablename As String, _valuetexts As String, Optional _fieldnames As String = "", Optional _checkPetik As Boolean = False)
        'OK
        'untuk memasukkan banyak data sekaligus ke dalam database
        'digunakan apabila memasukkan data baru yang primary key nya auto number
        Try
            If _fieldnames.Length = 0 Then
                Call CommandConstructor(_conn, _comm, "INSERT INTO " & _tablename & " VALUES (" & _valuetexts & ");", "insert", _checkPetik)
            Else
                Call CommandConstructor(_conn, _comm, "INSERT INTO " & _tablename & " (" & _fieldnames & ") VALUES (" & _valuetexts & ");", "insert", _checkPetik)
            End If
            'InputBox("", "", "INSERT INTO " & _tablename & " (" & _fieldnames & ") VALUES (" & _valuetexts & ");")
            Call ExecuteNonQueryFunction(_comm)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "InsertData Error")
        End Try
    End Sub

    Public Sub UpdateData(_conn As Object, _comm As Object, _tablename As String, _setString As String, Optional _whereString As String = "", Optional _checkPetik As Boolean = False)
        'OK
        'untuk mengupdate banyak data sekaligus pada database
        'sebaiknya klw bisa satu2, lebih baik pakai satu2, karena update bersamaan lebih berat bebannya daripada satu per satu
        Try
            If _whereString.Length = 0 Then
                Call CommandConstructor(_conn, _comm, "UPDATE " & _tablename & " SET " & _setString & ";", "update", _checkPetik)
            Else
                Call CommandConstructor(_conn, _comm, "UPDATE " & _tablename & " SET " & _setString & " WHERE " & _whereString & ";", "update", _checkPetik)
            End If
            'InputBox("", "", "UPDATE " & tablename & " SET " & setString & " WHERE " & whereString & ";")
            Call ExecuteNonQueryFunction(_comm)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "UpdateData Error")
        End Try
    End Sub

    Public Function ExecuteCmd(_conn As Object, _comm As Object, _stSQL As String, Optional _mTipe As String = "", Optional _timeDelaySeconds As Byte = 0, Optional ByRef _callerFsName As String = "", Optional _checkPetik As Boolean = False) As Boolean
        'OK
        'cek apa bisa menunggu s/d perintah selesai diexecute sebelum terus ke baris berikutnya
        '04Jun10 dipindah ke module ModFrmGnrl, krn penting
        'butuh ref: MS ADO 2.5 lib (msado25.tlb) utk ADODB.Connection
        'contoh :
        'Call ExecuteCmd(Conn, stSQL, , "PrepareChemUsgPctPerSO 1")
        'ctt:
        'fs ini didesign utk eksekusi perintah tunggal atau berantai, yg jika terjadi error, perintah-2 di bawahnya tetap boleh dieksekusi
        'jika butuh yg jika terjadi error, eksekusi berikutnya harus dihentikan, lakukan dgn cara normal : Conn.Execute stSQL
        'eksekusi del record tidak akan bisa jalan bila tabel tempat record tsb menjadi datasource utk MS Data Control, yg menjadi data source utk TDbGrid (tdbg7.ocx), krn record terkunci. Tapi bisa didel bila memanggil fs DelAllMSDtControlRecord.

        Try
            If (Trim(_stSQL) = "") Then
                Call myCShowMessage.ShowWarning("SQL Belum Didefinisikan...", "ExecuteCmd Cancelled")
                ExecuteCmd = False
                Exit Function
            End If
            Call CommandConstructor(_conn, _comm, _stSQL, _mTipe, _checkPetik)
            Call ExecuteNonQueryFunction(_comm)
            '20Apr10...
            If (_timeDelaySeconds > 0) Then
                Call myCMiscFunction.Wait(_timeDelaySeconds)
            End If
            ExecuteCmd = True
        Catch ex As Exception
            ExecuteCmd = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExecuteCmd Error")
        End Try
    End Function

    Function DelDbRecords(_conn As Object, _comm As Object, _stTableName As String, Optional _stCrit As String = "", Optional _dbType As String = "pgsql", Optional _checkPetik As Boolean = False) As Boolean
        'OK
        'bisa dipakai utk menggantikan fs DelTmpDbData, tapi hati-hati dng koneksi yg dipakai >>> biasanya pakai AdoCon3 atau AdoCon2
        Dim stSQLDel As String
        Try
            _stTableName = Trim(_stTableName)
            If (_stTableName = "") Then
                Call myCShowMessage.ShowWarning("Table Name Belum Diset..." & vbCrLf & vbCrLf & "Penghapusan Data Table Dibatalkan...", "DelDbRecords Error")
                DelDbRecords = False
                Exit Function
            End If
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If (Left(_stTableName, 1) <> "[") And Not (_stTableName.Contains(".")) Then
                    _stTableName = "[" & _stTableName & "]"
                End If
            End If

            If (_stCrit <> "") Then
                If Not (UCase(Left(Trim(_stCrit), 5)) = "WHERE") Then
                    _stCrit = " WHERE " & _stCrit
                End If
            End If
            If (_dbType = "access") Then
                stSQLDel = "Delete T.* From " & _stTableName & " as T " & _stCrit & ";"
            ElseIf (_dbType = "sql" Or _dbType = "pgsql") Then
                stSQLDel = "Delete From " & _stTableName & " " & _stCrit & ";"
            Else
                stSQLDel = Nothing
            End If
            If Not IsNothing(stSQLDel) Then
                Call CommandConstructor(_conn, _comm, stSQLDel, "delete", _checkPetik)
                Call ExecuteNonQueryFunction(_comm)
                DelDbRecords = True
            Else
                Call myCShowMessage.ShowWarning("Command query untuk delete keliru!!" & ControlChars.NewLine & "Silahkan hubungi EDP")
                DelDbRecords = False
            End If
        Catch ex As Exception
            DelDbRecords = False
            Call myCShowMessage.ShowErrMsg("Pesan Error:" & ex.Message, "DelDbRecords Error")
        End Try
    End Function

    Public Function UpdateDb(_conn As Object, _comm As Object, _stTableName As String, _stFieldName As String, _stFieldValue As String, _bFieldType As Byte, _stCrit As String, Optional _setUpdatingDate As Boolean = False, Optional _cancelIfWouldCreateDuplicateKey As Boolean = False, Optional _checkPetik As Boolean = False) As Boolean
        'SEHARUSNYA SUDAH OK, AKAN DI TES LEBIH LANJUT UNTUK LEBIH MEMASTIKAN
        'Hanya bisa buat update 1 records dalam 1 waktu, kalau mw update banyak langsung, pakai update data
        'hanya saja pada update data belum ada pengecekan duplicate key dll...
        'jadi diasumsikan user sudah tahu apa yang harus dilakukan
        '28Oct2011 sebaiknya diberi parameter utk menentukan db apa yg dipakai, apakah access, MSSQL, atau lainnya
        'paling baik utk dipakai update data di db, dgn entry via grid, atau entry via control yg lain dalam bentuk array. >>> lih. PrjCrdMgmt utk contoh.
        'contoh :
        'stSQLF = "where coBriefname='" & txtBriefName.Text & "'"
        'Call UpdateDb(AdoCon, "[Master EntitiesAddress]", "EntCountry", "Indonesia", 1, stSQLF)
        'Call UpdateDb(AdoCon, DbPathMstPIM & "[Master EntitiesProfile]", dbVal, Trim(txtGeneral(Index).Text), 1, stSQLF)
        'Call UpdateDb(AdoCon, DbPathTr & "[PIMTr PO]", dbval, Trim(cboDlvDt(Index).BoundText), tpval, stSQLF)

        'metoda lain yg bisa dipakai utk grid / array
        'buka RS, set nilai di grid / array dgn loop for-next
        'definisikan SQL sesuai urutan tampilan yg dipakai di grid / urutan array, dgn filter keyfield tsb (utamanya utk grid, 1 kolom hrs merupakan key dan tdk bisa diubah nilainya, dan diambil nilainya saat afterrowcolchange)
        'pada event validate atau afteredit, set sbb: AdoRS.Update slFld(Col), .TextMatrix(Row, Col) >>> updating dilakukan utk setiap editing pada cell atau anggota array.
        'cara lain :
        'lebih aman, tapi sedikit lebih lama, pakai UpdateInsertDb
        'butuh ref : Microsoft ActiveX Data Objects 2.6 Library (msado15.dll)
        Try
            Dim stSQLUpdate As String
            Dim CheckResult As Boolean
            Dim NewFieldValue As String
            Dim Xist As Short

            'di trim semua awal2, biar hilang semua spasi di awal dan akhir
            'biar aman
            _stTableName = Trim(_stTableName)
            _stFieldName = Trim(_stFieldName)
            _stFieldValue = Trim(_stFieldValue)

            'blok berikut utk pengecekan nama field dan tabel, utamanya utk Access
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Not (Left(Trim(_stFieldName), 1) = "[") Then
                    _stFieldName = "[" & _stFieldName & "]"
                End If
                Xist = InStr(1, UCase(_stTableName), " AS ")
                If (Xist = 0) Then
                    Xist = InStr(1, UCase(_stTableName), ".DBO.") 'utk mengatasi masalah nama tabel utk sql
                    If (Xist = 0) Then
                        If (Left(Trim(_stTableName), 1) <> "[") And Not (_stTableName.Contains(".")) Then 'bermasalah kalau ada deklarasi alias >>> 14Sep11 ditambahkan pengecekan alias dulu
                            _stTableName = "[" & _stTableName & "]"
                        End If
                    End If
                    '            stSQLUpdate = "Update " & _stTableName & " as T "    '14Sep2011
                Else
                    '            stSQLUpdate = "Update " & _stTableName & " "         '14Sep2011
                End If
            End If

            If Not (UCase(Left(Trim(_stCrit), 5)) = "WHERE") Then
                _stCrit = " WHERE " & _stCrit 'kalau kriterianya ngaco, nggak ketahuan
            End If

            NewFieldValue = myCMiscFunction.FormatFieldValForDbUpdate(_stFieldValue, _bFieldType)

            If (_cancelIfWouldCreateDuplicateKey) Then
                If (NewFieldValue = "Null") Then 'krn dipastikan bhw field ini adalah keyfield, maka nggak boleh null
                    UpdateDb = False
                    Call myCShowMessage.ShowWarning("Key Field Value Tidak Boleh 'Null'..." & vbCrLf & vbCrLf & "Update Data Dibatalkan...", "UpdateDB Cancelled")
                    Exit Function
                Else
                    CheckResult = IsExistRecords(_conn, _comm, _stFieldName, _stTableName, _stFieldName & "=" & NewFieldValue)
                    If (CheckResult) Then
                        UpdateDb = False
                        Call myCShowMessage.ShowWarning("Sudah Ada Key Field Dengan Nilai '" & NewFieldValue & "'..." & vbCrLf & vbCrLf & "Update Data Dibatalkan Karena Tidak Boleh Ada Duplicate Value Untuk Key Field...", "UpdateDB Cancelled")
                        Exit Function
                    End If
                End If
            End If

            If Not (_setUpdatingDate) Then
                stSQLUpdate = "Update " & _stTableName & " " & "set " & _stFieldName & "= " & NewFieldValue & " " & _stCrit & ";"
            Else
                'utk Access
                stSQLUpdate = "Update " & _stTableName & " " & "set " & _stFieldName & "= " & NewFieldValue & ", UpdatingDate=Now() " & _stCrit & ";"
                'ConvDateForMsSQL(Now() - 2)       '>>> utk SQL pakai cara ini
                stSQLUpdate = "Update " & _stTableName & " " & "set " & _stFieldName & "= " & NewFieldValue & ", UpdatingDate=" & Now.ToOADate() - 2 & " " & _stCrit & ";"
            End If
            Call CommandConstructor(_conn, _comm, stSQLUpdate, "update", _checkPetik)
            Call ExecuteNonQueryFunction(_comm)
            UpdateDb = True
        Catch ex As Exception
            UpdateDb = False
            Call myCShowMessage.ShowErrMsg("Pesan Error:" & ex.Message, "UpdateDb Error")
        End Try
    End Function

    Public Function UpdateInsertDb(_conn As Object, _comm As Object, _reader As Object, _tableName As String, _fieldName As String, _stFieldValue As String, _bFieldType As Byte, _stCrit As String, Optional _cancelIfWouldCreateDuplicateKey As Boolean = False, Optional _checkPetik As Boolean = False) As Boolean
        'belum di tes
        'fs ini akan mengupdate bila record sdh ada, atau menyisipkan record baru bila belum ada sebelumnya
        'dipakai bila duplikasi data pd field tsb oleh krn update ini tidak bermasalah >>> normalnya dipakai utk tabel yg hanya mempunyai 1 key field (yg bukan autonumber)
        '>>> tidak bisa dipakai untuk search dan insert dgn multiple criteria
        'contoh :
        'Call UpdateInsertDb(AdoCon, DbPathMstPIM & "[Master EntitiesProfile]", dbVal, Trim(txtGeneral(Index).Text), 1, stSQLF)
        Try
            Dim stSQLUpdate As String
            Dim NewFieldValue As String
            Dim CheckResult As Boolean

            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Not (Left(Trim(_fieldName), 1) = "[") Then
                    _fieldName = "[" & _fieldName & "]"
                End If
            End If

            If Not (UCase(Left(Trim(_stCrit), 5)) = "WHERE") Then
                _stCrit = " WHERE " & _stCrit
            End If

            NewFieldValue = myCMiscFunction.FormatFieldValForDbUpdate(_stFieldValue, _bFieldType)

            If (_cancelIfWouldCreateDuplicateKey = True) Then
                If (_stFieldValue = "Null") Then 'krn dipastikan bhw field ini adalah keyfield, maka nggak boleh null
                    UpdateInsertDb = False
                    Call myCShowMessage.ShowWarning("Key Field Value Tidak Boleh 'Null'..." & vbCrLf & vbCrLf & "Update / Insert Data Dibatalkan...", "UpdateInsertDB Cancelled")
                    Call myCDBConnection.CloseConn(_conn, -1)
                    Exit Function
                Else
                    'IsExistRecords ada di module 'ModDomainAggregateFs'
                    CheckResult = IsExistRecords(_conn, _comm, _fieldName, _tableName, _fieldName & "=" & NewFieldValue)
                    If (CheckResult = True) Then
                        UpdateInsertDb = False
                        Call myCShowMessage.ShowWarning("Sudah Ada Key Field Dengan Nilai '" & NewFieldValue & "'..." & vbCrLf & vbCrLf & "Update / Insert Data Dibatalkan Karena Tidak Boleh Ada Duplicate Value Untuk Key Field...", "UpdateInsertDB Cancelled")
                        Call myCDBConnection.CloseConn(_conn, -1)
                        Exit Function
                    End If
                End If
            End If

            stSQLUpdate = "Select T.* From " & _tableName & " as T " & _stCrit & ";"
            Call CommandConstructor(_conn, _comm, stSQLUpdate, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            Dim update As Boolean

            While _reader.Read
                If (_reader(_fieldName) <> NewFieldValue) Then
                    'jika tidak sama >>> update
                    stSQLUpdate = "Update " & _tableName & " " & "set " & _fieldName & "= " & NewFieldValue & " " & _stCrit & ";"
                    Call CommandConstructor(_conn, _comm, stSQLUpdate, "update", _checkPetik)
                    Call ExecuteNonQueryFunction(_comm)
                    update = True
                End If
            End While
            _reader.Close()
            'kalau bukan update berarti itu record baru jadi di insert
            If (update = False) Then
                'If Not (FieldValue = "Null") Then      '>>> di sini, jika field bukan key field, nggak masalah kalau null; utk key field, sdh dicek di atas
                '22Dec09 ini hanya bisa dipakai utk insert sederhana, dimana kriteria hanya 1 field, dan bila tidak ada, yg diinsert adalah 1 field tsb.
                stSQLUpdate = "insert into " & _tableName & " (" & _fieldName & ") " & "values (" & NewFieldValue & ");"
                Call CommandConstructor(_conn, _comm, stSQLUpdate, "insert", _checkPetik)
                Call ExecuteNonQueryFunction(_comm)
                'End If
            End If
            UpdateInsertDb = True
        Catch ex As Exception
            UpdateInsertDb = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "UpdateInsertDb Error")
        End Try
    End Function

    Public Function InsertRecIfNotExist(_conn As Object, _comm As Object, _reader As Object, _tableName As String, _fieldName As String, _fieldValue As String, _bFieldType As Byte, _keyFieldName As String, _keyFieldValue As String, Optional _getRecID As Boolean = False, Optional _recIDFieldName As String = "", Optional ByRef _recID As Object = Nothing, Optional _checkPetik As Boolean = False) As Boolean
        'OK, cuman ada beberapa baris code yang disesuaikan
        'kebanyakan di comment, karena malah menyebabkan error
        'fs ini akan menyisipkan record baru bila belum ada sebelumnya.
        'terutama dipakai utk data yg mempunyai 2 keyfield, dimana 1 keyfield dipakai scr internal utk komputer (normalnya autonumber), dan keyfield lain adalah yg dikenal oleh user (eksternal, yi: search code >>> normalnya string).
        'e.g. utk tabel entities, dimana ada FinID utk user di accounting, dan BriefName utk user lainnya, dan ada EntId (utk komputer, dan utk join antara kedua tabel tsb)
        'bisa dipakai utk di main table dan utk di sub table, utk tabel-2 dng kyfield internal dan eksternal
        'catatan :
        'KeyFieldValue harus dalam format yg sesuai. e.g. utk string, nilai yg dipassing harus : 'KeyFieldValue'; sedang utk numerik : KeyFieldValue.
        'nama dan value KeyField dipakai utk pembentukan filter, utk mencek apakah data tsb sdh ada / tidak.
        'FieldValue meski string, nggak perlu diberi single quote di kiri dan kanannya saat memanggil fs ini, krn penambahan quote sdh diurus di fs ini
        'contoh :
        'Call InsertRecIfNotExist(AdoCon, DbPathMstPIM & "[Master Entities]", "CoBriefName", txtBriefName.Text, 1, "EntFinID", " '" & txtFinanceID.Text & "' ")
        'penting :
        'ctt: validation rule di Access db (Like "???.???" spt di field EntFinID) hrs dihapus, krn akan mengakibatkan error di pemanggilan fs InsertRecIfNotExist.
        Try
            Dim stSQLInsert As String
            Dim stCrit As String
            Dim Xist As Short
            '    Dim RecID As Variant

            'Trim semua untuk menghilangkan spasi, agar aman waktu insert
            _tableName = Trim(_tableName)
            _keyFieldName = Trim(_keyFieldName)
            _fieldName = Trim(_fieldName)
            _keyFieldValue = Trim(_keyFieldValue)
            _fieldValue = Trim(_fieldValue)

            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                Xist = InStr(1, UCase(_tableName), " AS ")
                If (Xist = 0) Then
                    Xist = InStr(1, UCase(_tableName), ".DBO.") 'utk mengatasi masalah nama tabel utk sql
                    If (Xist = 0) Then
                        If (Left(Trim(_tableName), 1) <> "[") And Not (_tableName.Contains(".")) Then 'bermasalah kalau ada deklarasi alias >>> 14Sep11 ditambahkan pengecekan alias dulu
                            _tableName = "[" & _tableName & "]"
                        End If
                    End If
                    '            stSQLUpdate = "Update " & stTableName & " as T "    '14Sep2011
                Else
                    '            stSQLUpdate = "Update " & stTableName & " "         '14Sep2011
                End If
            End If


            '==========================================================================================
            'kalau ini di-uncomment, maka tidak bisa digunakan untuk insert data banyak sekaligus
            'karena apabila ada banyak fieldname maka akan jadi begini [field_A,field_B,field_C], dan akan error
            'jadi apabila mau insert data, diasumsikan semua format data yang dikirim sudah benar
            'jadi sebelum masuk sini, nama field yang terdiri dari 2 kata, CO: NAMA ITEM
            'sebelum masuk sini harus sudah begini [NAMA ITEM] agar tidak error
            '==========================================================================================

            'If Not (Left(Trim(FieldName), 1) = "[") Then
            '    FieldName = "[" & FieldName & "]"
            'End If

            'If Not (Left(KeyFieldName, 1) = "[") Then
            '    KeyFieldName = "[" & KeyFieldName & "]"
            'End If

            If ((Trim(_keyFieldName) <> "") And (Trim(_keyFieldValue) <> "")) Then
                If (_getRecID = False) Then
                    stCrit = " WHERE T." & _keyFieldName & "=" & _keyFieldValue & ";"
                Else
                    'RecIDFieldName adalah autonumber >>> KRITERIA SENGAJA DINOLKAN, utk mengatasi masalah jika utk 1 keperluan, e.g. purchase, ada 2 orang yg bisa dihubungi dr Co yg sama.
                    stCrit = " WHERE T." & _keyFieldName & "=" & _keyFieldValue & " and T." & _recIDFieldName & "=0;"
                End If
            Else
                Call myCShowMessage.ShowWarning("Kesalahan Definisi Kriteria...", "InsertRecIfNotExist Cancelled")
                stCrit = ""
                InsertRecIfNotExist = False
            End If

            'kalau di uncomment, tidak akan bisa update masal
            'FieldValue = FormatFieldValForDbUpdate(FieldValue, bFieldType)

            stSQLInsert = "Select T.* From " & _tableName & " as T " & stCrit
            Call CommandConstructor(_conn, _comm, stSQLInsert, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            Dim Exist As Boolean
            While _reader.Read
                'sdh ada record, nggak usah insert data
                InsertRecIfNotExist = False
                Exist = True
            End While
            _reader.Close()
            If (Exist = False) Then
                If Not (_fieldValue = "Null") Then '>>> dipakai krn keyfield nggak boleh null
                    stSQLInsert = "insert into " & _tableName & " (" & _keyFieldName & ", " & _fieldName & ") " & "values (" & _keyFieldValue & ", " & _fieldValue & ")"
                    Call CommandConstructor(_conn, _comm, stSQLInsert, "insert", _checkPetik)
                    Call ExecuteNonQueryFunction(_comm)
                    InsertRecIfNotExist = True
                    If (_getRecID = True) Then
                        'AdoRS.Requery
                        'AdoRS.MoveLast      'pakai dmax dgn kriteria spt nilai insert, krn autonumber >>> meski ada resiko utk multiuser
                        'RecID = AdoRS.Fields(RecIDFieldName).Value
                        '>>> DMax ada di ModDomainAggregateFs : Public Function DMax(pConn As Object, pFieldName As String, pTableName As String, Optional pFilter As String = "") As Variant
                        _recID = DMax(_conn, _comm, _recIDFieldName, _tableName, _keyFieldName & "=" & _keyFieldValue & " AND " & _fieldName & "=" & _fieldValue) '12apr11 sebelumnya pakai AdoCon >>> cek ulang fs-2 yg memanggil fs ini
                    End If
                Else
                    InsertRecIfNotExist = False
                End If
            Else
                InsertRecIfNotExist = False
            End If
            'reclaim ram
            stSQLInsert = Nothing
            stCrit = Nothing
        Catch ex As Exception
            InsertRecIfNotExist = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "InsertRecIfNotExist Error")
        End Try
    End Function

    Public Function DCount(_conn As Object, _comm As Object, _reader As Object, ByRef _pFldName As String, ByRef _pTblName As String, Optional ByRef _pFilter As String = "", Optional _checkPetik As Boolean = False) As Integer
        'belum di tes
        '------------------------------------------------------------
        ' Purpose: mendapatkan jumlah record dari suatu tabel
        ' Parameters:
        ' Date: Januari,13 05' Time: 14:13
        '------------------------------------------------------------
        'contoh:
        'RecCount = DCount(Adocon, "keyfield", "table_name")
        'Ctt:
        '- jika field yg menjadi acuan tidak ada isinya (null), maka tidak akan dihitung
        'return value akan = 0 jika tidak ada record / tidak ada record yg sesuai kriteria
        Dim stSQL As String
        Try
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Left(Trim(_pFldName), 1) <> "[" Then
                    _pFldName = "[" & Trim(_pFldName) & "]"
                End If
                If Left(Trim(_pTblName), 1) <> "[" And Not (_pTblName.Contains(".")) Then
                    _pTblName = "[" & Trim(_pTblName) & "]"
                End If
            End If

            stSQL = "select count(" & _pFldName & ") as jml from " & _pTblName
            If Trim(_pFilter) <> "" Then stSQL = stSQL & " where " & _pFilter & ";"

            'return value akan = 0 jika tidak ada record / tidak ada record yg sesuai kriteria
            Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            DCount = 0

            While _reader.Read
                'untuk query2 SUM, MAX, MIN, COUNT, dan mgkn yang lainnya yang belum ketemu, harus dilakukan pengecekan isNull pada _reader
                'karena di cek _reader ada record namun null
                If Not (IsDBNull(_reader("jml"))) Then
                    'DCount = oleReader("jml")
                    'karena menghitung jumlahnya, bukan mengambil nilainya
                    DCount += 1
                End If
            End While
            _reader.Close()
        Catch ex As Exception
            DCount = 0
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DCount Error")
        End Try
    End Function

    Public Function DCountDistinct(_conn As Object, _comm As Object, _reader As Object, _pFldName As String, _pTblName As String, Optional _pFilter As String = "", Optional _checkPetik As Boolean = False) As Integer
        'belum di tes

        '------------------------------------------------------------
        ' Purpose: mendapatkan jumlah record dari suatu tabel, dengan pengelompokan dahulu pada field _pFldName tsb >>> memperoleh DISTINCT VALUE dulu, baru dihitung
        ' Parameters:
        ' Date: November,03 10' Time: 13:52
        '------------------------------------------------------------
        'contoh:
        'RecCount = DCountDistinct(Adocon, "keyfield", "table_name")
        'Ctt:
        '- jika field yg menjadi acuan tidak ada isinya (null), maka tidak akan dihitung

        Dim stSQL As String
        Try
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Left(Trim(_pFldName), 1) <> "[" Then
                    _pFldName = "[" & Trim(_pFldName) & "]"
                End If
                If Left(Trim(_pTblName), 1) <> "[" And Not (_pTblName.Contains(".")) Then
                    _pTblName = "[" & Trim(_pTblName) & "]"
                End If
            End If

            stSQL = "SELECT DISTINCT " & _pFldName & " as F from " & _pTblName
            If Trim(_pFilter) <> "" Then stSQL = stSQL & " where " & _pFilter & ";"
            'return value akan = 0 jika tidak ada record / tidak ada record yg sesuai kriteria

            Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            DCountDistinct = 0

            While _reader.Read
                'Seharusnya select distinct tidak perlu melakukan pengecekan
                'kalau hasil query tidak ada maka tidak akan masuk fungsi while ini
                'If Not (IsDBNull(oleReader("F"))) Then
                'DCountDistinct = oleReader("F")
                'karena menghitung jumlahnya, bukan mengambil nilainya
                DCountDistinct += 1
                'End If
            End While
            _reader.Close()
        Catch ex As Exception
            DCountDistinct = 0
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DCountDistinct Error")
        End Try
    End Function

    Public Function IsExistRecords(_conn As Object, _comm As Object, _reader As Object, _pFieldName As String, _pTableName As String, Optional _pCriteria As String = "", Optional _checkPetik As Boolean = False) As Boolean
        'OK
        'dipakai utk mencek apakah record ada dan field tsb ada isinya
        'ada fs dgn nama yg hampir sama, tapi utk DAO, di module 'ModFrmGnrl'
        'contoh pemakaian :
        'Dim CheckResult As Boolean     ', atau Dim RecordIsExist as Boolean
        'CheckResult = IsExistRecords(AdoCon, "EntID", "Master EntitiesProfile", "EntBriefName='" & txtBriefName.Text & "' ")
        'If (CheckResult = True) Then
        'Cancel = True
        'End If
        Dim RecordsIsExist As Object
        Dim Xist As Short
        Try
            'blok berikut utk pengecekan nama field dan tabel, utamanya utk Access
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                'Jika koneksi oledb (access, sql)
                If Not (Left(_pFieldName, 1) = "[") Then
                    _pFieldName = "[" & _pFieldName & "]"
                End If
                Xist = InStr(1, UCase(_pTableName), " AS ") 'utk mengatasi masalah alias
                If (Xist = 0) Then
                    Xist = InStr(1, UCase(_pTableName), ".DBO.") 'utk mengatasi masalah nama tabel utk sql
                    If (Xist = 0) Then
                        If (Left(_pTableName, 1) <> "[") And Not (_pTableName.Contains(".")) Then 'bermasalah kalau ada deklarasi alias >>> 14Sep11 ditambahkan pengecekan alias dulu
                            _pTableName = "[" & _pTableName & "]"
                        End If
                    End If
                End If
            End If

            RecordsIsExist = DLookup(_conn, _comm, _reader, _pFieldName, _pTableName, _pCriteria, _checkPetik)
            If IsDBNull(RecordsIsExist) Then
                IsExistRecords = False
            Else
                'jika record ada meski field yg diminta datanya null, empty string atau empty
                IsExistRecords = True
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "IsExistRecords Error")
            IsExistRecords = False
        End Try
    End Function

    Public Function DLookup(_conn As Object, _comm As Object, _reader As Object, _pFieldName As String, _pTableName As String, Optional _pCriteria As String = "", Optional _checkPetik As Boolean = False) As Object
        'OK
        '------------------------------------------------------------
        ' Purpose: mendapatkan nilai suatu field dari sebuah tabel, atau gabungan dari beberapa field (>>> pakai syntax SQL dan diawali serta diakhiri kurung)
        ' Parameters:
        ' Date: Januari,13 05' Time: 14:13
        '------------------------------------------------------------
        'Contoh:
        'STOCK_MIN = NZ(DLookup(AdoCon, "[MINIMUM]", "Item", "(Item.KODE_ITEM)='" & Me.XKODE_ITEM & "' "), 0)
        'TxtOrdBy(1).Text = NZ(DLookup(AdoCon, "T.EntAddress", "[Master EntitiesAddress] AS T", "(((T.EntAddrC)=0) AND ((T.EntBriefName)='" & .DCbSO(0).Text & "'));"), "")
        'jika utk mencek keberadaan record, bisa pakai fs IsExistRecord di module ModFrmGnrl (utk DAO >>> pindah lagi ke ModDbGenFsDAO), atau fs IsExistRecords di module ini (utk ADO)
        'jika ingin mengetahui apakah suatu tabel tidak ada isinya, atau tabel ada isinya tapi record yg dicari tidak ada, pakai fs DLookup, dan cek nilai DRsStatus

        'pengecekan return value : jika record yg dicari tidak ada : If IsNull(STOCK_MIN) Then MsgBox "No Record"
        'sebetulnya fs ini menggantikan fs sederhana sbb yg tanpa error checking :
        'SQLText = "Select SOOrigin From " & Db3Name & "[SDCMTr SOIndex] Where SONr=" & AdoRS.Fields(0).Value
        'vSOOrigin = AdoCon.Execute(SQLText)!SOOrigin
        'ctt:
        'utk menggantikan fs seperti berikut:
        '        SQLText = "SELECT InvTrType FROM " & db4Name & "[LDTr StockTrIndex] WHERE TrNr=" & sNr
        '        AdoRS.Open SQLText, AdoCon, adOpenStatic, adLockReadOnly
        '        If Not AdoRS.EOF Then
        '            AdoRS.MoveFirst
        '            If AdoRS.Fields(0).Value <> "" Then vInvType = AdoRS.Fields(0).Value
        '        End If
        '        AdoRS.Close
        Dim stSQL As String
        Dim DRsStatus As String

        Try
            Dim ReturnValue As Object
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                stSQL = "select top 1 " & _pFieldName & " as var from " & _pTableName
                If Trim(_pCriteria) <> "" Then
                    stSQL = stSQL & " where " & _pCriteria & ";"
                End If
            ElseIf (TypeOf _conn Is Npgsql.NpgsqlConnection) Then
                stSQL = "select " & _pFieldName & " as var from " & _pTableName
                If Trim(_pCriteria) <> "" Then
                    stSQL = stSQL & " where " & _pCriteria & " limit 1;"
                Else
                    stSQL = stSQL & " limit 1;"
                End If
            Else
                stSQL = Nothing
            End If

            If Not IsNothing(stSQL) Then
                DRsStatus = "" '>>> variable penanda jika rs.EOF, agar bisa mengetahui apakah record memang tidak ada, atau nilai yg dikembalikan memang null
                '>>> dibutuhkan jika hasil pencarian nilai pada 1 field dipakai utk melakukan update pada field lain di record tsb

                Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)
                _reader = ExecuteReaderFunction(_comm)

                'inisialisasi awal jadi kalau tidak masuk while akan berisi ini
                'sama saja dengan RS.EOF
                DLookup = DBNull.Value
                DRsStatus = "EOF"

                While _reader.Read
                    DRsStatus = "OK"
                    ReturnValue = _reader("var")
                    If (IsDBNull(ReturnValue) Or IsNothing(ReturnValue) Or (Trim(ReturnValue) = "")) Then
                        DLookup = "" 'jika field memang null, tapi record ada, return dr fs ini adalah empty string
                    Else
                        DLookup = ReturnValue
                    End If
                End While
                _reader.Close()
            Else
                DLookup = False
                Call myCShowMessage.ShowWarning("String query tidak teridentifikasi atau kosong!!")
            End If
        Catch ex As Exception
            DLookup = DBNull.Value
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DLookUp Error")
        End Try
    End Function

    Public Function DMin(_conn As Object, _comm As Object, _reader As Object, _pFieldName As String, _pTableName As String, Optional _pFilter As String = "", Optional _checkPetik As Boolean = False) As Object
        'belum di tes
        'pengecekan return value : jika record yg dicari tidak ada : If IsNull(STOCK_MIN) Then MsgBox "No Record"
        'ada fs yg sama di module 'MdlGeneral', tapi dgn urutan parameter yg berbeda >>> 09Jul09 nggak dipakai lagi
        'ctt:
        '_pTableName langsung diset dgn alias T
        Dim stSQL As String
        Try
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Not ((Left(_pFieldName, 1) = "[") And (Right(_pFieldName, 1) = "]")) Then _pFieldName = "[" & _pFieldName & "]"
            End If

            stSQL = "SELECT Min(T." & _pFieldName & ") AS MinValue " & "FROM " & _pTableName & " as T "
            If Trim(_pFilter) <> "" Then stSQL = stSQL & " WHERE (" & _pFilter & ");"
            Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            'sblmnya diinisialisasi Null dulu agar apabila record tidak ada DMin akan Null
            DMin = DBNull.Value

            While _reader.Read
                'untuk query2 SUM, MAX, MIN, COUNT, harus dilakukan pengecekan isNull pada _reader
                'karena di cek _reader ada record namun null
                If Not (IsDBNull(_reader("MinValue"))) Then
                    DMin = myCStringManipulation.NZ(_reader("MinValue"))
                End If
            End While
            _reader.Close()
        Catch ex As Exception
            DMin = DBNull.Value
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DMin Error")
        End Try
    End Function

    Public Function DMax(_conn As Object, _comm As Object, _reader As Object, _pFieldName As String, _pTableName As String, Optional _pFilter As String = "", Optional _checkPetik As Boolean = False) As Object
        'OK
        'pengecekan return value : jika record yg dicari tidak ada, e.g :
        'If IsNull(STOCK_MIN) Then MsgBox "No Record"
        'ada fs yg sama di module 'MdlGeneral', tapi dgn urutan parameter yg berbeda >>> 09Jul09 nggak dipakai lagi
        'contoh:
        'NewDNSINr = NZ(DMax(AdoCon, "DNSINr", Db4Name & "[SDCMTr DNSIIndex]")) + 1
        Dim stSQL As String
        Try
            If (TypeOf _conn Is OleDb.OleDbConnection Or TypeOf _conn Is SqlConnection) Then
                If Not ((Left(_pFieldName, 1) = "[") And (Right(_pFieldName, 1) = "]")) Then _pFieldName = "[" & _pFieldName & "]"
            End If

            stSQL = "SELECT Max(T." & _pFieldName & ") AS MaxValue " & "FROM " & _pTableName & " as T "
            If Trim(_pFilter) <> "" Then stSQL = stSQL & " WHERE (" & _pFilter & ");"

            Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)
            _reader = ExecuteReaderFunction(_comm)

            'kalau RS.EOF, karena tidak masuk while
            DMax = DBNull.Value

            While _reader.Read
                'untuk query2 SUM, MAX, MIN, COUNT, harus dilakukan pengecekan isNull pada _reader
                'karena di cek _reader ada record namun null
                If Not (IsDBNull(_reader("MaxValue"))) Then
                    DMax = myCStringManipulation.NZ(_reader("MaxValue"))
                End If
            End While
            _reader.Close()
        Catch ex As Exception
            DMax = DBNull.Value
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DMax Error")
        End Try
    End Function

    Public Function DSum(_conn As Object, _comm As Object, _reader As Object, _pFldName As String, _pTblName As String, Optional _pFilter As String = "", Optional _checkPetik As Boolean = False) As Object
        'OK
        'untuk memperoleh jumlah isi dari sebuah field
        Dim stSQL As String
        Try
            stSQL = "SELECT Sum(" & _pFldName & ") AS jml FROM " & _pTblName & " as T "
            If Trim(_pFilter) <> "" Then stSQL = stSQL & " WHERE (" & _pFilter & ");"

            'Call OpenConn(pConn)

            'buat pengecekan macam2
            'stSQL = "select MAX(InvOB) as maks from [Resume Goods Transactions BNSN] as T where(T.FinYear = 1990)"

            Call CommandConstructor(_conn, _comm, stSQL, "select", _checkPetik)

            _reader = ExecuteReaderFunction(_comm)

            'sementara saja buat pengecekan
            'Dim a As Boolean
            'a = oleReader.Read()
            'MsgBox(a)

            DSum = Nothing

            While _reader.Read
                'buat pengecekan juga, hanya sementara
                'MsgBox(IsDBNull(oleReader("maks")))

                'untuk query2 SUM, MAX, MIN, COUNT, harus dilakukan pengecekan isNull pada _reader
                'karena di cek _reader ada record namun null
                If Not (IsDBNull(_reader("jml"))) Then
                    DSum = myCStringManipulation.NZ(_reader("jml"))
                End If
            End While
            _reader.Close()
        Catch ex As Exception
            DSum = DBNull.Value
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DSum Error")
        End Try
    End Function

    Public Function GetSpecificRecord(_conn As Object, _comm As Object, _reader As Object, _pFldName As String, _pTblName As String, Optional _sortingMethod As String = "desc", Optional _pFilter As String = Nothing, Optional _dbType As String = "pgsql", Optional _checkPetik As Boolean = False) As Object
        'OK
        'untuk memperoleh jumlah isi dari sebuah field
        Dim stSQL As String
        Try
            If Not IsNothing(_pFilter) Then
                If (_dbType = "pgsql") Then
                    stSQL = "SELECT " & _pFldName & " FROM " & _pTblName & " WHERE " & _pFilter & " ORDER BY " & _pFldName & " " & _sortingMethod & " LIMIT 1;"
                Else
                    stSQL = "SELECT TOP 1 " & _pFldName & " FROM " & _pTblName & " WHERE " & _pFilter & " ORDER BY " & _pFldName & " " & _sortingMethod & ";"
                End If
            Else
                If (_dbType = "pgsql") Then
                    stSQL = "SELECT " & _pFldName & " FROM " & _pTblName & " ORDER BY " & _pFldName & " " & _sortingMethod & " LIMIT 1;"
                Else
                    stSQL = "SELECT TOP 1 " & _pFldName & " FROM " & _pTblName & " ORDER BY " & _pFldName & " " & _sortingMethod & ";"
                End If
            End If
            GetSpecificRecord = GetDataIndividual(_conn, _comm, _reader, stSQL)
        Catch ex As Exception
            GetSpecificRecord = DBNull.Value
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetSpecificRecord Error")
        End Try
    End Function

    Public Function GetFormulationRecord(_conn As Object, _comm As Object, _reader As Object, _pFldName As String, _pTblName As String, Optional _formulationMethod As String = "Sum", Optional _pFilter As String = Nothing, Optional _dbType As String = "pgsql", Optional _checkPetik As Boolean = False) As Object
        'OK
        'untuk memperoleh jumlah isi dari sebuah field
        Dim stSQL As String
        Try
            If Not IsNothing(_pFilter) Then
                'If (_dbType = "pgsql") Then
                stSQL = "SELECT " & _formulationMethod & "(" & _pFldName & ") as i FROM " & _pTblName & " WHERE " & _pFilter & ";"
                'Else
                '    stSQL = "SELECT " & _formulationMethod & "(" & _pFldName & ") as i FROM " & _pTblName & " WHERE " & _pFilter & ";"
                'End If
            Else
                'If (_dbType = "pgsql") Then
                stSQL = "SELECT " & _formulationMethod & "(" & _pFldName & ") as i FROM " & _pTblName & ";"
                'Else
                '    stSQL = "SELECT " & _formulationMethod & "(" & _pFldName & ") as i FROM " & _pTblName & ";"
                'End If
            End If
            GetFormulationRecord = GetDataIndividual(_conn, _comm, _reader, stSQL)
        Catch ex As Exception
            GetFormulationRecord = DBNull.Value
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetFormulationRecord Error")
        End Try
    End Function

    Public Sub SetCbo_(_conn As Object, _comm As Object, _reader As Object, _stSQL As String, ByRef _myDataTableCbo As DataTable, ByRef _myBindingSourceCbo As BindingSource, ByRef _cbo As ComboBox, _table As String, _valueMember As String, _displayMember As String, ByRef _isCboPrepared As Boolean, Optional _gantiKriteria As Boolean = False)
        Try
            _isCboPrepared = False
            If IsNothing(_myDataTableCbo.DataSet) Or _gantiKriteria Then
                _myDataTableCbo = GetDataTableUsingReader(_conn, _comm, _reader, _stSQL, _table)
            End If
            _myBindingSourceCbo.DataSource = _myDataTableCbo
            With _cbo
                .DataSource = _myBindingSourceCbo
                .ValueMember = _valueMember
                .DisplayMember = _displayMember
                .SelectedIndex = -1
            End With
            _isCboPrepared = True
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetCbo_ Error")
        End Try
    End Sub

    Public Function GetRegNumber(_conn As Object, _comm As Object, _reader As Object, _regCodePrefix As String, _tableName As String, _regCode As String, _sortingType As String, _docType As String, Optional _dbType As String = "sql") As String
        Try
            Dim mPeriodeDoc As String
            Dim myNoUrutDoc As Integer
            Dim myNoDoc As String
            Dim myNumberDocString As String

            mPeriodeDoc = IIf(Now.Month.ToString.Length = 1, "0" & Now.Month, Now.Month) & Now.Year.ToString.Substring(2, 2) & "." & IIf(Now.Day.ToString.Length > 1, Now.Day, "0" & Now.Day)

            myNoDoc = GetSpecificRecord(_conn, _comm, _reader, _regCode, _tableName, _sortingType, "(" & _regCode & " Like '" & _regCodePrefix & "-" & mPeriodeDoc & "%')", _dbType)

            If Not IsNothing(myNoDoc) Then
                myNoUrutDoc = myNoDoc.Substring(myNoDoc.Length - 4, 4)
                myNoUrutDoc += 1
                If (myNoUrutDoc > 9999) Then
                    Return Nothing
                    Call myCShowMessage.ShowInfo("Selamat usaha anda berjaya.. " & ControlChars.NewLine & "Cheers... ^_^")
                    Call myCShowMessage.ShowInfo("Silahkan hubungi developer untuk menaikkan batas " & _docType & " yang dapat dibuat dalam 1 hari!!")
                    Exit Function
                Else
                    myNumberDocString = IIf(myNoUrutDoc.ToString.Length = 1, "000" & myNoUrutDoc, IIf(myNoUrutDoc.ToString.Length = 2, "00" & myNoUrutDoc, IIf(myNoUrutDoc.ToString.Length = 3, "0" & myNoUrutDoc, myNoUrutDoc)))
                End If
                myNoDoc = _regCodePrefix & "-" & mPeriodeDoc & "." & myNumberDocString
            Else
                myNoDoc = _regCodePrefix & "-" & mPeriodeDoc & "." & "0001"
            End If
            GetRegNumber = myNoDoc
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetRegNumber Error")
            GetRegNumber = Nothing
        End Try
    End Function

    Public Sub SaveUpdatedAt(_conn As Object, _comm As Object, _reader As Object, ByRef _newValues As String, _dbaseTbl As String, Optional _dbType As String = "access")
        'OK
        Try
            'untuk menghindari kesamaan waktu pada saat penyimpanan data, jadi di cek dulu apakah sudah ada data dengan waktu ini
            'hal ini dapat terjadi karena komputer selalu sinkronisasi data dengan internet
            Dim isExist As Boolean
            'Dim isExist2 As Boolean
            Dim cek As Boolean = False
            Dim idx As Integer = 0

            isExist = True
            idx = 0
            While isExist
                If (_dbType = "access") Then
                    isExist = IsExistRecords(_conn, _comm, _reader, "updated_at", _dbaseTbl, "updated_at=#" & Now.AddSeconds(idx) & "#")
                ElseIf (_dbType = "sql" Or _dbType = "pgsql") Then
                    isExist = IsExistRecords(_conn, _comm, _reader, "updated_at", _dbaseTbl, "updated_at='" & Now.AddSeconds(idx) & "'")
                End If
                If Not isExist Then
                    _newValues &= ",'" & Now.AddSeconds(idx) & "'"
                Else
                    idx += 1
                End If
            End While
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SaveUpdatedAt Error")
        End Try
    End Sub

    Public Sub EditUpdatedAt(_conn As Object, _comm As Object, _reader As Object, ByRef _updateString As String, _dbaseTbl As String, Optional _dbType As String = "access")
        'OK
        Try
            'untuk menghindari kesamaan waktu pada saat penyimpanan data
            Dim isExist As Boolean
            Dim cek As Boolean = False
            Dim idx As Integer = 0

            isExist = True
            idx = 0
            While isExist
                If (_dbType = "access") Then
                    isExist = IsExistRecords(_conn, _comm, _reader, "updated_at", _dbaseTbl, "updated_at=#" & Now.AddSeconds(idx) & "#")
                ElseIf (_dbType = "sql" Or _dbType = "pgsql") Then
                    isExist = IsExistRecords(_conn, _comm, _reader, "updated_at", _dbaseTbl, "updated_at='" & Now.AddSeconds(idx) & "'")
                End If
                If Not isExist Then
                    _updateString &= ",updated_at='" & Now.AddSeconds(idx) & "'"
                Else
                    idx += 1
                End If
            End While
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "EditUpdatedAt Error")
        End Try
    End Sub

    Public Function SetAutoKode(_conn As Object, _comm As Object, _reader As Object, _tblName As String, _prefixKode As String, Optional _dbType As String = "access") As String
        Try
            'DIGANTI DENGAN FUNGSI SetDynamicAutoKode
            Dim mCountAddZero As Integer
            Dim mTempKode As String
            Dim myAutoKode As String
            Dim stSQL As String

            If (_dbType = "pgsql") Then
                stSQL = "SELECT (" & _tblName & ".kode) " &
                        "FROM(" & _tblName & ") " &
                        "ORDER BY " & _tblName & ".id DESC " &
                        "LIMIT 1;"
            Else
                stSQL = "SELECT TOP 1 (" & _tblName & ".kode) " &
                        "FROM(" & _tblName & ") " &
                        "ORDER BY " & _tblName & ".id DESC;"
            End If

            myAutoKode = GetDataIndividual(_conn, _comm, _reader, stSQL)

            If Not IsNothing(myAutoKode) Then
                'jika kategori sudah ada isinya, maka ambil yang kode paling akhir, terus ditambah 1
                'diambil yang kode angkanya saja, huruf K di depan dibuang dulu
                myAutoKode = myAutoKode.Substring(1, 6)
                myAutoKode += 1
                If (myAutoKode.Count < 6) Then
                    'jika kurang dari 4, contoh seharusnya 0005, kalau masih 5 saja berarti kan kurang dari 4, harus ditambahin angka 0 sebanyak 3 biji di depannya
                    mCountAddZero = 6 - myAutoKode.Count
                    mTempKode = Nothing
                    For a As Integer = 0 To mCountAddZero - 1
                        mTempKode &= "0"
                    Next a
                    myAutoKode = _prefixKode & mTempKode & myAutoKode
                End If
            Else
                'klw data kategori masih kosong, maka auto kode diisi lgsg
                myAutoKode = _prefixKode & "000001"
            End If
            Return myAutoKode
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAutoKode Error")
            Return Nothing
        End Try
    End Function

    Public Function SetDynamicAutoKode(_conn As Object, _comm As Object, _reader As Object, _tblName As String, _kodeField As String, _idField As String, _prefixKode As String, _digitLength As Integer, _incDate As Boolean, Optional _dateValue As Date = Nothing, Optional _dbType As String = "access", Optional _midKode As String = Nothing, Optional _incDay As Boolean = True) As String
        Try
            Dim mCountAddZero As Integer
            Dim zeroNumber As String
            Dim myAutoKode As String
            Dim stSQL As String

            If Not IsNothing(_midKode) Then
                _prefixKode &= _midKode
            End If

            If (_incDate) Then
                'Jika include tanggal di kodenya
                Dim mPeriodeDoc As String
                If (_incDay) Then
                    'DIGIT TANGGAL IKUT
                    'URUTANNYA ddmmyy
                    'mPeriodeDoc = _dateValue.Year & IIf(_dateValue.Month.ToString.Length = 1, "0" & _dateValue.Month, _dateValue.Month) & IIf(_dateValue.Day < 10, "0" & _dateValue.Day, _dateValue.Day)
                    mPeriodeDoc = IIf(_dateValue.Day < 10, "0" & _dateValue.Day, _dateValue.Day) & IIf(_dateValue.Month.ToString.Length = 1, "0" & _dateValue.Month, _dateValue.Month) & _dateValue.Year.ToString.Substring(2, 2)
                Else
                    'DIGIT TANGGAL TIDAK IKUT
                    'URUTANNYA mmyy
                    mPeriodeDoc = IIf(_dateValue.Month.ToString.Length = 1, "0" & _dateValue.Month, _dateValue.Month) & _dateValue.Year.ToString.Substring(2, 2)
                End If
                _prefixKode &= "" & mPeriodeDoc
            End If

            If (_dbType = "pgsql") Then
                stSQL = "SELECT " & _kodeField & " " &
                        "FROM " & _tblName & " " &
                        "WHERE " & _kodeField & " like '" & _prefixKode & "%'" &
                        "ORDER BY " & _idField & " DESC " &
                        "LIMIT 1;"
            Else
                stSQL = "SELECT TOP 1 " & _kodeField & " " &
                        "FROM " & _tblName & " " &
                        "WHERE " & _kodeField & " like '" & _prefixKode & "%' " &
                        "ORDER BY " & _idField & " DESC;"
            End If
            myAutoKode = GetDataIndividual(_conn, _comm, _reader, stSQL)

            If (_digitLength > 0) Then
                If Not IsNothing(myAutoKode) Then
                    'jika kategori sudah ada isinya, maka ambil yang kode paling akhir, terus ditambah 1
                    'diambil yang kode angkanya saja, huruf K di depan dibuang dulu
                    myAutoKode = myAutoKode.Substring(myAutoKode.Length - _digitLength, _digitLength)
                    myAutoKode += 1
                    If (myAutoKode.Count <= _digitLength) Then
                        'jika kurang dari 4, contoh seharusnya 0005, kalau masih 5 saja berarti kan kurang dari 4, harus ditambahin angka 0 sebanyak 3 biji di depannya
                        mCountAddZero = _digitLength - myAutoKode.Count
                        zeroNumber = Nothing
                        For a As Integer = 0 To mCountAddZero - 1
                            zeroNumber &= "0"
                        Next a
                        myAutoKode = _prefixKode & zeroNumber & myAutoKode
                    Else
                        Call myCShowMessage.ShowWarning("Counter yang diperbolehkan melebihi batas sebanyak " & _digitLength & " digit" & ControlChars.NewLine & "Silahkan hubungi Developer")
                    End If
                Else
                    'klw data kategori masih kosong, maka auto kode diisi lgsg
                    zeroNumber = Nothing
                    For i As Integer = 0 To _digitLength - 2
                        zeroNumber &= "0"
                    Next
                    myAutoKode = _prefixKode & zeroNumber & "1"
                End If
            Else
                myAutoKode = _prefixKode
            End If
            Return myAutoKode
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetDynamicAutoKode Error")
            Return Nothing
        End Try
    End Function

    Public Function SetDynamicKDR(_conn As Object, _comm As Object, _reader As Object, _tblName As String, _kodeField As String, _idField As String, _prefixDate As Date, _identityValue As String, _digitLength As Integer, Optional _dbType As String = "access") As String
        Try
            Dim mCountAddZero As Integer
            Dim zeroNumber As String
            Dim myAutoKode As String
            Dim stSQL As String
            Dim myPrefixKode As String

            Dim mPeriodeDoc As String
            mPeriodeDoc = Format(_prefixDate, "ddMMMyyyy")
            myPrefixKode = mPeriodeDoc & _identityValue

            If (_dbType = "pgsql") Then
                stSQL = "SELECT (" & _tblName & "." & _kodeField & ") " &
                        "FROM (" & _tblName & ") " &
                        "WHERE " & _tblName & "." & _kodeField & " like '" & myPrefixKode & "%'" &
                        "ORDER BY " & _tblName & "." & _idField & " DESC " &
                        "LIMIT 1;"
            Else
                stSQL = "SELECT TOP 1 " & _kodeField & " " &
                        "FROM " & _tblName & " " &
                        "WHERE " & _kodeField & " like '" & myPrefixKode & "%' " &
                        "ORDER BY " & _idField & " DESC;"
            End If
            myAutoKode = GetDataIndividual(_conn, _comm, _reader, stSQL)

            If (_digitLength > 0) Then
                If Not IsNothing(myAutoKode) Then
                    'jika kategori sudah ada isinya, maka ambil yang kode paling akhir, terus ditambah 1
                    'diambil yang kode angkanya saja, huruf K di depan dibuang dulu
                    myAutoKode = myAutoKode.Substring(myAutoKode.Length - _digitLength, _digitLength)
                    myAutoKode += 1
                    If (myAutoKode.Count <= _digitLength) Then
                        'jika kurang dari 4, contoh seharusnya 0005, kalau masih 5 saja berarti kan kurang dari 4, harus ditambahin angka 0 sebanyak 3 biji di depannya
                        mCountAddZero = _digitLength - myAutoKode.Count
                        zeroNumber = Nothing
                        For a As Integer = 0 To mCountAddZero - 1
                            zeroNumber &= "0"
                        Next a
                        myAutoKode = myPrefixKode & zeroNumber & myAutoKode
                    Else
                        Call myCShowMessage.ShowWarning("Counter yang diperbolehkan melebihi batas sebanyak " & _digitLength & " digit" & ControlChars.NewLine & "Silahkan hubungi Developer")
                    End If
                Else
                    'klw data kategori masih kosong, maka auto kode diisi lgsg
                    zeroNumber = Nothing
                    For i As Integer = 0 To _digitLength - 2
                        zeroNumber &= "0"
                    Next
                    myAutoKode = myPrefixKode & zeroNumber & "1"
                End If
            Else
                'Jika digit length nya = 0
                myAutoKode = myPrefixKode
            End If
            Return myAutoKode
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetDynamicKDR Error")
            Return Nothing
        End Try
    End Function
End Class
