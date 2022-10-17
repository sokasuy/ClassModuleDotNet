Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class CDBConnection
    Private myCShowMessage As New CShowMessage.CShowMessage
    Private myCFileIO As New CFileIO.CFileIO
    Private DbPath As String
    Private ConnType As String
    Private prvdrTypeAccess As String
    Private prvdrTypeExcel As String
    'Private usrName As String
    'Private usrPass As String
    Private serverName As String
    Private prvdrTypeSql As String
    Private serverPort As String
    Private prvdrTypePGSql As String
    Private GDatabase As String
    Private sysDB As String

    Public Sub SetDBParameter(ByRef _conn As Object, _dbType As String, _serverName As String, _dbPath As String, _prvdrType As String, _filePathFileName As String, _serverPort As String, Optional _userName As String = "", Optional _password As String = "", Optional _sysDB As String = "")
        Try
            If (_dbType = "ACCESS") Then
                DbPath = myCFileIO.ReadIniFile(_dbType, _dbPath, _filePathFileName)
                prvdrTypeAccess = myCFileIO.ReadIniFile(_dbType, _prvdrType, _filePathFileName)
                'usrName = myCFileIO.ReadIniFile(_dbType, _userName, _filePathFileName)
                'usrPass = myCFileIO.ReadIniFile(_dbType, _password, _filePathFileName)
                Call SetDbPath(_conn, _dbType, serverName, DbPath, prvdrTypeAccess, _userName, _password, sysDB)
            ElseIf (_dbType = "EXCEL") Then
                DbPath = _dbPath
                prvdrTypeExcel = myCFileIO.ReadIniFile(_dbType, _prvdrType, _filePathFileName)
                'usrName = myCFileIO.ReadIniFile(_dbType, _userName, _filePathFileName)
                'usrPass = myCFileIO.ReadIniFile(_dbType, _password, _filePathFileName)
                Call SetDbPath(_conn, _dbType, serverName, DbPath, prvdrTypeExcel, _userName, _password)
            ElseIf (_dbType = "SQL") Then
                serverName = myCFileIO.ReadIniFile(_dbType, _serverName, _filePathFileName)
                prvdrTypeSql = myCFileIO.ReadIniFile(_dbType, _prvdrType, _filePathFileName)
                DbPath = myCFileIO.ReadIniFile(_dbType, _dbPath, _filePathFileName)
                'usrName = myCFileIO.ReadIniFile(_dbType, _userName, _filePathFileName)
                'usrPass = myCFileIO.ReadIniFile(_dbType, _password, _filePathFileName)
                Call SetDbPath(_conn, _dbType, serverName, DbPath, prvdrTypeSql, _userName, _password)
            ElseIf (_dbType = "SQLSrv") Then
                '3 November 2021 Baru pake SqlConnection
                ConnType = myCFileIO.ReadIniFile(_dbType, "CONN_TYPE", _filePathFileName)
                serverName = myCFileIO.ReadIniFile(_dbType, _serverName, _filePathFileName)
                serverPort = myCFileIO.ReadIniFile(_dbType, _serverPort, _filePathFileName)
                prvdrTypeSql = Nothing
                DbPath = myCFileIO.ReadIniFile(_dbType, _dbPath, _filePathFileName)
                'usrName = myCFileIO.ReadIniFile(_dbType, _userName, _filePathFileName)
                'usrPass = myCFileIO.ReadIniFile(_dbType, _password, _filePathFileName)
                Call SetDbPath(_conn, _dbType, serverName, DbPath, prvdrTypeSql, _userName, _password,, serverPort,, ConnType)
            ElseIf (_dbType = "PGSQL") Then
                serverName = myCFileIO.ReadIniFile(_dbType, _serverName, _filePathFileName)
                serverPort = myCFileIO.ReadIniFile(_dbType, _serverPort, _filePathFileName)
                prvdrTypePGSql = myCFileIO.ReadIniFile(_dbType, _prvdrType, _filePathFileName)
                DbPath = myCFileIO.ReadIniFile(_dbType, _dbPath, _filePathFileName)
                Call SetDbPath(_conn, _dbType, serverName, DbPath, prvdrTypePGSql, _userName, _password,, serverPort)
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetDBParameter Error")
        End Try
    End Sub

    Private Function SetDbPath(ByRef _conn As Object, _dbType As String, _serverName As String, _databaseName As String, _prvdrType As String, Optional _usrID As String = "", Optional _dbPsw As String = "", Optional _sysDb As String = "", Optional _serverPort As Integer = 0, Optional _excelVersion As String = "xls", Optional _connType As String = "") As Boolean
        'OK
        'dipakai utk menset nilai Gdatabase
        Try
            If (_dbType = "ACCESS") Then
                If (_databaseName.Chars(0) = "\") Then
                    'kalau database ada di jaringan
                    'misal \\server_name\folder\MyDB.mdb
                    GDatabase = _databaseName
                Else
                    If Not (_databaseName.Chars(2) = "\") Then
                        'ini kalau pada file setting.ini hanya berisi nama mdb nya
                        GDatabase = Application.StartupPath & "\" & _databaseName
                    Else
                        'kalau ada path maka akan masuk else sini
                        GDatabase = _databaseName
                    End If
                End If
                _conn = New OleDb.OleDbConnection
                Call SetAndOpenConnDB(_dbType, _conn, _prvdrType, "", GDatabase, _sysDb, 0, _usrID, _dbPsw)
            ElseIf (_dbType = "SQL") Then
                _conn = New OleDb.OleDbConnection
                Call SetAndOpenConnDB(_dbType, _conn, _prvdrType, serverName, DbPath, "", 0, _usrID, _dbPsw)
            ElseIf (_dbType = "SQLSrv") Then
                _conn = New SqlConnection
                Call SetAndOpenConnDB(_dbType, _conn, _prvdrType, serverName, DbPath, "", serverPort, _usrID, _dbPsw,, _connType)
            ElseIf (_dbType = "PGSQL") Then
                _conn = New Npgsql.NpgsqlConnection
                Call SetAndOpenConnDB(_dbType, _conn, _prvdrType, serverName, DbPath, "", serverPort, _usrID, _dbPsw)
            ElseIf (_dbType = "EXCEL") Then
                _conn = New OleDb.OleDbConnection
                Call SetAndOpenConnDB(_dbType, _conn, _prvdrType, serverName, DbPath, "", serverPort, _usrID, _dbPsw, _excelVersion)
            End If
            SetDbPath = True
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetDbPath Error")
            SetDbPath = False
        End Try
    End Function

    Public Sub SetAndOpenConnForExcel(ByRef _conn As Object, _dataSource As String, _excelVersion As String, _prvdrType As String)
        'OK
        Try
            Dim stProvider As String
            'khusus untuk excel, koneksi di new di sini, karena biasanya koneksi ke excel dipakai hanya untuk kondisi tertentu saja
            'bukan untuk database permanen
            _conn = New OleDb.OleDbConnection
            Call CloseConn(_conn)

            Select Case _prvdrType
                Case "3.51", "4.0"
                    stProvider = "Provider=Microsoft.Jet.OLEDB." & _prvdrType & ";" 'utk Access 97  : Microsoft.Jet.OLEDB.3.51 = MSJET35.DLL 24-Apr-1998, di \system32\  DAN utk Access 2000 - 2003
                Case Else
                    stProvider = "Provider=Microsoft.ACE.OLEDB." & _prvdrType & ";" 'utk Access 2007 dst
            End Select

            If (_excelVersion.ToLower = "xls") Then
                'khusus untuk file excel 97-2003
                _conn.ConnectionString = stProvider & "Data source=" & _dataSource & ";Extended Properties=Excel 8.0;"
            ElseIf (_excelVersion.ToLower = "xlsx") Then
                'kalau buat file excel 2007
                _conn.ConnectionString = stProvider & "Data Source=" & _dataSource & ";Extended Properties=Excel 12.0 Xml;"
            End If
            Call OpenConn(_conn)

        Catch ex As Exception
            Call CloseConn(_conn)
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAndOpenConnForExcel Error")
        End Try
    End Sub

    Public Function ExcelConStr(_prvdrType As String, _dataSource As String, _excelVersion As String) As String
        Try
            'Dim connSql As String
            Dim stProvider As String

            Select Case _prvdrType
                Case "3.51", "4.0"
                    stProvider = "Provider=Microsoft.Jet.OLEDB." & _prvdrType & ";" 'utk Access 97  : Microsoft.Jet.OLEDB.3.51 = MSJET35.DLL 24-Apr-1998, di \system32\
                Case "12.0"
                    stProvider = "Provider=Microsoft.ACE.OLEDB." & _prvdrType & ";" 'utk Access 2007 dst
                    'stProvider = "Provider=Microsoft.ACE.OLEDB.12.0;" 'utk Access 2007 dst
            End Select

            If (_excelVersion.ToLower = "xls") Then
                'khusus untuk file excel 97-2003
                ExcelConStr = stProvider & "data source=" & _dataSource & ";Extended Properties=Excel 8.0;"
            ElseIf (_excelVersion.ToLower = "xlsx") Then
                'kalau buat file excel 2007
                ExcelConStr = stProvider & "Data Source=" & _dataSource & ";Extended Properties=Excel 12.0 Xml;"
            End If

            '"HDR=Yes;" indicates that the first row contains columnnames, not data. "HDR=No;" indicates the opposite.
            '"IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. Note that this option might affect excel sheet write access negative.

            'ExcelConStr = connSql
        Catch ex As Exception
            ExcelConStr = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExcelConStr Error")
        End Try
    End Function

    Public Function SetAndOpenConnDB(_mType As String, ByRef _conn As Object, _prvdrType As String, _serverName As String, _dbName As String, Optional _sysDb As String = "", Optional _serverPort As Integer = 0, Optional _usrID As String = "", Optional _dbPsw As String = "", Optional _excelVersion As String = "xls", Optional _connType As String = "") As Boolean
        Try
            Call CloseConn(_conn)

            With _conn
                If (_mType = "SQL") Then
                    .ConnectionString = SetSqlConStr(_prvdrType, _serverName, _dbName, _usrID, _dbPsw)
                ElseIf (_mType = "SQLSrv") Then
                    .ConnectionString = SetSqlConStrNew(_connType, _serverName, _serverPort, _dbName, _usrID, _dbPsw)
                ElseIf (_mType = "ACCESS") Then
                    .ConnectionString = SetAccessConStr(_prvdrType, _dbName, _usrID, _dbPsw, _sysDb)
                ElseIf (_mType = "PGSQL") Then
                    .ConnectionString = SetPgSqlConStr(_prvdrType, _serverName, _dbName, _serverPort, _usrID, _dbPsw)
                ElseIf (_mType = "EXCEL") Then
                    .ConnectionString = ExcelConStr(_prvdrType, _dbName, _excelVersion)
                End If
            End With

            If (OpenConn(_conn) = False) Then
                'error bisa dikarenakan lokasi SysDb yg salah; tipe db, dbuser/psw yg keliru
                Call myCShowMessage.ShowWarning("Koneksi ke Server / Database Gagal Dibentuk..." & vbCrLf & vbCrLf & "Silakan Beritahukan Masalah Ini ke Network Administrator...")
                SetAndOpenConnDB = False
            Else
                SetAndOpenConnDB = True
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAndOpenConnDB Error")
            SetAndOpenConnDB = False
        End Try
    End Function

    Public Function SetPgSqlConStr(_prvdrType As String, _serverName As String, _dbName As String, _serverPort As Integer, Optional _usrID As String = "", Optional _dbPsw As String = "") As String
        '>>> dipakai jika Db tdk diset usernya
        Try
            Dim stProvider As String = ""
            Dim conSql As String

            'stProvider = "Provider=" & PrvdrType & ";"
            If (Trim(_dbPsw) = "") Then
                'tanpa Password
                'ConSqlStr = stProvider & ";Persist Security Info=False;Initial Catalog=HSI;Data Source=AKTSVR"
                'ConSql = "Server=<" & ServerName & ">;Database=<" & DbName & ">;User=<" & UsrID & ">;Port=<" & ServerPort & ">;"
                'ConSql = "Data Source=" & ServerName & ";Port=" & ServerPort & ";Initial Catalog=" & DbName & ";"
                conSql = "Server=" & _serverName & ";Port=" & _serverPort & ";Database=" & _dbName & ";CommandTimeout=200;"
            Else
                'ConSqlStr = stProvider & ";Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=HSI;Data Source=HSIDEVSVR"
                'ConSql = "Server=<" & ServerName & ">;Database=<" & DbName & ">;User=<" & UsrID & ">;Password=<" & DbPsw & ">;Port=<" & ServerPort & ">;"
                'ConSql = "Data Source=" & ServerName & ";Port=" & ServerPort & ";Initial Catalog=" & DbName & ";User ID=" & UsrID & ";password=" & DbPsw & ";"
                conSql = "Server=" & _serverName & ";Port=" & _serverPort & ";Database=" & _dbName & ";User Id=" & _usrID & ";Password=" & _dbPsw & ";CommandTimeout=200;"
            End If
            'SetPgSqlConStr = stProvider & ConSql
            SetPgSqlConStr = conSql
        Catch ex As Exception
            SetPgSqlConStr = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetPgSqlConStr Error")
        End Try
    End Function

    Public Function SetAndOpenConnForMSSQL(ByRef _conn As OleDb.OleDbConnection, _comm As OleDb.OleDbCommand, _prvdrType As String, _sqlServerName As String, _dbName As String, Optional _usrID As String = "", Optional _dbPsw As String = "") As Boolean
        'ok
        'MEI 2016 GAK DIPAKAI
        'contoh :
        'Call SetAndOpenConnForMSSQL(ConSql, "8.0", "HSIDEVSVR", "HSI", Dbuser, Dbpass)
        'Call SetAndOpenConnForMSSQL(AdoCon, "8.0", "HSISERVER", "HSI", "sw", "sw")
        'ctt:
        'login time out di SQL manager harus diset 30 detik, diubah dari defaultnya 4 detik
        Try
            'tutup dulu koneksinya
            Call CloseConn(_conn)

            With _conn
                .ConnectionString = SetSqlConStr(_prvdrType, _sqlServerName, _dbName, _usrID, _dbPsw)
            End With
            'With Comm
            '    .CommandTimeout = 120
            'End With
            If (OpenConn(_conn) = False) Then
                'error bisa dikarenakan lokasi SysDb yg salah; tipe db, dbuser/psw yg keliru
                Call myCShowMessage.ShowWarning("Koneksi ke Server / Database Gagal Dibentuk..." & vbCrLf & vbCrLf & "Silakan Beritahukan Masalah Ini ke Network Administrator...")
                SetAndOpenConnForMSSQL = False
            Else
                SetAndOpenConnForMSSQL = True
            End If
            'SetAndOpenConnForMSSQL = OpenConn(Conn)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAndOpenConnForMSSQL Error")
            SetAndOpenConnForMSSQL = Nothing
        End Try
    End Function

    Public Function SetSqlConStr(_prvdrType As String, _sqlServerName As String, _dbName As String, Optional _usrID As String = "", Optional _dbPsw As String = "") As String
        'ok
        'contoh :
        'stConSQL = SetSqlConStr("8.0", "HSIDEVSVR", "HSI", Dbuser, Dbpass)     '>>> dipakai jika Db diset usernya
        'stConSQL = SetSqlConStr("8.0", "HSIDEVSVR", "HSI")                                  '>>> dipakai jika Db tdk diset usernya
        Try
            Dim stProvider As String = ""
            Dim ConSql As String

            'Select Case PrvdrType
            '    '        Case ""
            '    '            stProvider = "Provider=;"        'utk
            '    Case "8.0"
            '        stProvider = "Provider=SQLOLEDB;"
            '    Case "8.1"
            '        stProvider = "Provider=SQLOLEDB.1;"           'sql vs 8.0
            'End Select
            stProvider = "Provider=" & _prvdrType & ";"
            If (Trim(_dbPsw) = "") Then
                'tanpa Password
                'ConSqlStr = stProvider & ";Persist Security Info=False;Initial Catalog=HSI;Data Source=AKTSVR"
                ConSql = "Persist Security Info=False;Initial Catalog=" & _dbName & ";Data Source=" & _sqlServerName
            Else
                'ConSqlStr = stProvider & ";Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=HSI;Data Source=HSIDEVSVR"
                ConSql = "User ID=" & _usrID & ";Password=" & _dbPsw & ";Persist Security Info=True;Initial Catalog=" & _dbName & ";Data Source=" & _sqlServerName & ";"
            End If
            SetSqlConStr = stProvider & ConSql
        Catch ex As Exception
            SetSqlConStr = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetSqlConStr Error")
        End Try
    End Function

    Public Function SetSqlConStrNew(_connType As String, _myServerAddress As String, _serverPort As String, _myDataBase As String, Optional _myUsername As String = "", Optional _myPassword As String = "") As String
        Try
            'SetSqlConStrNew = "Data Source=172.16.0.1\hsiserver;Initial Catalog=" & _myDataBase & ";User ID=" & _myUsername & ";Password=" & _myPassword & ";Persist Security Info=True;"
            If (_connType = "SQL_INSTANCE") Then
                SetSqlConStrNew = "Server=" & _myServerAddress & ";Database=" & _myDataBase & ";User Id=" & _myUsername & ";Password=" & _myPassword & ";"
            ElseIf (_connType = "IP_ADDRESS") Then
                SetSqlConStrNew = "Data Source=" & _myServerAddress & "," & _serverPort & ";Network Library=DBMSSOCN;Initial Catalog=" & _myDataBase & ";User ID=" & _myUsername & ";Password=" & _myPassword & ";"
            End If
        Catch ex As Exception
            SetSqlConStrNew = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetSqlConStrNew Error")
        End Try
    End Function

    Public Function SetAndOpenOleConForAccess(ByRef _conn As OleDb.OleDbConnection, _prvdrType As String, _dtSourceName As String, Optional _usrID As String = "", Optional _dbPsw As String = "", Optional _sysDb As String = "", Optional _cmdTimeOut As Short = 240) As Boolean
        'OK
        'MEI 2016 GAK DIPAKAI
        'untuk membuka koneksi ole untuk db access
        'Contoh: SetAndOpenOleConForAccess(OleCon, "4.0", Gdatabase, "", "", "SYSTEM.mdw")
        Try
            Call CloseConn(_conn, -1) 'tutup dulu koneksi sebelum pengesetan
            With _conn
                'mengambil koneksi string dari fungsi SetAccessConStr
                .ConnectionString = SetAccessConStr(_prvdrType, _dtSourceName, _usrID, _dbPsw, _sysDb)
            End With
            If (OpenConn(_conn) = False) Then
                'error bisa dikarenakan lokasi SysDb yg salah; tipe db, dbuser/psw yg keliru
                Call myCShowMessage.ShowWarning("Koneksi ke Server / Database Gagal Dibentuk..." & vbCrLf & vbCrLf & "Silakan Beritahukan Masalah Ini ke Network Administrator...")
                SetAndOpenOleConForAccess = False
            Else
                SetAndOpenOleConForAccess = True
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAndOpenOleConForAccess Error")
            SetAndOpenOleConForAccess = False
        End Try
    End Function

    Public Function SetAccessConStr(_prvdrType As String, _dtSourceName As String, Optional _usrID As String = "", Optional _dbPsw As String = "", Optional _sysDb As String = "") As String
        'OK
        'untuk penyusunan koneksi string untuk access
        'Contoh: SetAccessConStr("4.0", Gdatabase, "", "", "SYSTEM.mdw")
        Try
            Dim stProvider As String = ""
            Dim conStrProvider As String
            Dim conStrSys As String
            Dim conStr As String

            Select Case _prvdrType
                Case "3.51", "4.0"
                    stProvider = "Provider=Microsoft.Jet.OLEDB." & _prvdrType & ";" 'utk Access 97  : Microsoft.Jet.OLEDB.3.51 = MSJET35.DLL 24-Apr-1998, di \system32\
                Case Else
                    stProvider = "Provider=Microsoft.ACE.OLEDB." & _prvdrType & ";" 'utk Access 2007 dst
            End Select

            If (Trim(_sysDb) = "") Then
                'tanpa sysDB
                conStrProvider = stProvider & "Data Source="
                conStrSys = ";Persist Security Info=False"
            Else
                'dengan sysDB
                conStrProvider = stProvider & "Password=" & _dbPsw & ";User ID=" & _usrID & ";Data Source="
                conStrSys = ";Persist Security Info=True;Jet OLEDB:System database=" & _sysDb
            End If
            conStr = conStrProvider & _dtSourceName & conStrSys

            SetAccessConStr = conStr
        Catch ex As Exception
            SetAccessConStr = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetAccessConStr Error")
        End Try
    End Function

    Public Sub PrepareConn(ByRef _conn As Object, ByRef _comm As Object, _connStr As String, Optional _connTimeOut As Integer = 120, Optional _cmdTimeOut As Integer = 120)
        'belum di tes
        'ctt: dgn cara ini, koneksi tidak perlu dideklarasikan scr umum (Public), tapi cukup ConnectionStringnya yg Public
        'koneksi dideklarasikan pada level form, dan diset saat form_load.
        'keuntungan metoda ini adalah utk proyek dgn banyak koneksi ke bermacam db, yi koneksi ke suatu db dibentuk dan dipertahankan seperlunya saja
        'dan setelah itu koneksinya dihapus (bukan cuma ditutup)
        'sedangkan utk prj dengan jumlah koneksi yg sedikit (ke 1 atau 2 db), lebih baik jika koneksi yg dideklarasikan sbg public dan diset saat MDI form diload,
        'dan dipassingkan ke form yg memerlukannya. Koneksi hanya perlu dibuka saat Form_Load, dan ditutup saat event UnLoad. Koneksi diset nothing hanya saat MDIform ditutup.
        'contoh :
        'deklarasi pada level form :
        'Dim Con As New ADODB.Connection
        'pada event Form_Load :
        'Call PrepareConn (Con, stConDbTmp)
        'Call OpenConn(Con)
        'pada event Form_UnLoad :
        'RemoveConn Con
        Try
            If (_conn.ConnectionString <> "") Then
                Call CloseConn(_conn, -1) 'agar tdk error bila sdh pernah diset sebelumnya
            End If
            With _conn
                .ConnectionString = _connStr
            End With
            With _comm
                .CommandTimeout = _cmdTimeOut
            End With
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "PrepareConn Error")
        End Try
    End Sub

    Function OpenConn(ByRef _conn As Object) As Boolean
        'OK
        'membuka koneksi ke database
        'Contoh: OpenConn(oleConn)
        Try
            'cek apakah status koneksi sedang dalam close atw tidak
            'klw closed maka koneksi dibuka (Open). Klw sudah dalam kondisi Open, dipaksa Open lagi, maka akan error.
            If _conn.State = ConnectionState.Closed Then _conn.Open()

            OpenConn = True
        Catch ex As Exception
            OpenConn = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "OpenConn Error")
        End Try
    End Function

    Sub CloseConn(ByRef _conn As Object, Optional _waitingTimeAllowanceInSeconds As Short = 0)
        'OK
        'untuk menutup koneksi ke database
        'Contoh: CloseConn(oleConn, -1)
        Try
            Dim DelayEnd As Double

            Select Case _waitingTimeAllowanceInSeconds
                Case 0
                    'tanpa waktu tunggu
                    If _conn.State <> ConnectionState.Closed Then _conn.Close() '>>> memaksa utk close. Kelihatannya berbahaya, sebaiknya tunggu s/d state = open, baru ditutup
                Case Is > 0
                    'toleransi s/d sekian detik
                    DelayEnd = DateAdd(DateInterval.Second, _waitingTimeAllowanceInSeconds, Now).ToOADate
                    While DateDiff(DateInterval.Second, Now, System.DateTime.FromOADate(DelayEnd)) > 0 'If date1 refers to a later point in time than date2, the DateDiff function returns a negative number.
                        'tunggu
                    End While
                    _conn.Close()
                Case Else
                    'cek apakah dalam kondisi open? kalau iya maka di close
                    If _conn.State = ConnectionState.Open Then _conn.Close()
            End Select
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CloseConn Error")
        End Try
    End Sub

    Sub RemoveConn(ByRef _conn As Object, ByRef _comm As Object)
        'OK
        'untuk me-remove koneksi ke db
        'contoh : RemoveConn(oleConn)
        Try
            CloseConn(_conn, -1)
            If Not (IsNothing(_comm)) Then
                _comm.Dispose()
                _comm = Nothing
            End If
            _conn.Dispose()
            _conn = Nothing
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RemoveConn Error")
        End Try
    End Sub
End Class
