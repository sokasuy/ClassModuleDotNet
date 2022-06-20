Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Security.Cryptography
Imports System.Windows.Forms

Public Class CStringManipulation
    Private enc As System.Text.UTF8Encoding
    Private encryptor As ICryptoTransform
    Private decryptor As ICryptoTransform
    Private KEY_128 As Byte() = {42, 1, 52, 67, 231, 13, 94, 101, 123, 6, 0, 12, 32, 91, 4, 111, 31, 70, 21, 141, 123, 142, 234, 82, 95, 129, 187, 162, 12, 55, 98, 23}
    Private IV_128 As Byte() = {234, 12, 52, 44, 214, 222, 200, 109, 2, 98, 45, 76, 88, 53, 23, 78}
    Private symmetricKey As RijndaelManaged = New RijndaelManaged()
    Private myCShowMessage As New CShowMessage.CShowMessage


    Public Function NZ(_value As Object, Optional _replacementValue As Object = 0) As Object
        'OK
        'If IsNull(value) Or IsEmpty(value) Or value = "" Then
        'harus dipecah jadi if dan else if karena kalau jadi 1 maka klw value berisi DBNull.Value
        'akan error kalau dipakai Trim(Value)
        Try
            If IsDBNull(_value) Then 'utk memastikan, fs pemakai yg passing Value, hrs menulis dalam bentuk trim(Value)
                NZ = _replacementValue 'diubah 10 Mar 06
            ElseIf IsNothing(_value) Then
                NZ = _replacementValue
            ElseIf Trim(_value) = "" Then
                NZ = _replacementValue 'diubah 10 Mar 06
            Else
                NZ = _value
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "NZ Error")
            NZ = Nothing
        End Try
    End Function

    Public Function ZN(_value As Object) As Object
        'belum di tes
        'harus dipecah jadi if dan else if karena kalau jadi 1 maka klw value berisi DBNull.Value
        'akan error kalau dipakai Trim(Value)
        Try
            If IsDBNull(_value) Then 'utk memastikan, fs pemakai yg passing Value, hrs menulis dalam bentuk trim(Value)
                ZN = System.DBNull.Value
            ElseIf (_value = 0) Or (Trim(_value) = "") Then
                ZN = System.DBNull.Value
            Else
                ZN = _value
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ZN Error")
            ZN = Nothing
        End Try
    End Function

    Public Function ConvertDate(_dataTgl As String) As Date
        'belum di tes
        Try
            ConvertDate = CDate(Mid(_dataTgl, 4, 3) & Mid(_dataTgl, 1, 3) & Mid(_dataTgl, 7, 4))
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConvertDate Error")
            ConvertDate = Nothing
        End Try
    End Function

    Public Function ConvDateForMsSQL(_stDate As String) As String
        'belum di tes
        'mengubah nilai tanggal dari format tanggal menjadi Long.
        'Input dari grid atau control, output dalam Long yg diformat string, utk save ke db SQL.
        'contoh :
        'ConvDateForMsSQL(.TextMatrix(Row, 12))
        'ctt:
        'utk nilai jam, disimpan dalam 10 digit desimal >>> setting di db pakai desimal, scala 10
        Try
            Dim lOutput As String

            If (IsDate(_stDate)) Then
                lOutput = CStr(CLng(CDate(_stDate).ToOADate - 2))
            Else
                lOutput = "Null"
            End If
            ConvDateForMsSQL = lOutput
            lOutput = Nothing
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConvDateForMsSQL Error")
            ConvDateForMsSQL = Nothing
        End Try
    End Function

    Public Function BitToBoolean(_bitValue As Object) As Boolean
        'belum di tes
        'dipakai utk konversi nilai CheckBox di form
        '07Apr10 fs ini nggak bisa dipakai utk checkbox VSFG, krn 1 berarti true dan 2 berarti false >>> pakai fs VSFGBitToBoolean

        'e.g. : Call SetPOField(Not (BitToBoolean(Me.xtunai)))      '>>> bila checkbox Tunai ditandai, field PO didisabled; dan sebaliknya
        Try
            If (Val(_bitValue) = 0) Then
                BitToBoolean = False
            Else
                BitToBoolean = True
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "BitToBoolean Error")
            BitToBoolean = False
        End Try
    End Function

    Public Function BooleanToBit(_boolValue As Object) As Short
        'belum di tes
        'dipakai utk konversi nilai utk mengisi CheckBox di form
        'tidak akan memicu event _Click
        '07Apr10 fs ini nggak bisa dipakai utk checkbox VSFG, krn 1 berarti true dan 2 berarti false >>> pakai fs VSFGBitToBoolean
        Try
            If (_boolValue = False) Then
                BooleanToBit = 0
            Else
                BooleanToBit = 1 'bisa digunakan utk Access
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "BooleanToBit Error")
            BooleanToBit = Nothing
        End Try
    End Function

    Public Function Enkripsi(_txt As String) As String
        'OK
        'untuk enkripsi password
        Try
            Dim i As Short
            Dim hrf As String

            For i = 1 To Len(_txt)
                hrf = hrf & Chr(Asc(Mid(_txt, i, 1)) + 40)
            Next i

            Enkripsi = hrf
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "Enkripsi error")
            Enkripsi = Nothing
        End Try
    End Function

    Public Function Dekripsi(_txt As String) As String
        'OK
        'untuk dekripsi password
        Try
            Dim i As Short
            Dim hrf As String

            For i = 1 To Len(_txt)
                hrf = hrf & Chr(Asc(Mid(_txt, i, 1)) - 40)
            Next i

            Dekripsi = hrf
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "Dekripsi error")
            Dekripsi = Nothing
        End Try
    End Function

    Public Function GetRijndaelEncrypt(_txt As String) As String
        Try
            Dim sPlainText As String = _txt
            If Not String.IsNullOrEmpty(sPlainText) Then
                symmetricKey.Mode = CipherMode.CBC

                enc = New System.Text.UTF8Encoding
                encryptor = symmetricKey.CreateEncryptor(KEY_128, IV_128)

                Dim memoryStream As MemoryStream = New MemoryStream()
                Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, Me.encryptor, CryptoStreamMode.Write)
                cryptoStream.Write(Me.enc.GetBytes(sPlainText), 0, sPlainText.Length)
                cryptoStream.FlushFinalBlock()
                GetRijndaelEncrypt = Convert.ToBase64String(memoryStream.ToArray())
                memoryStream.Close()
                cryptoStream.Close()
            Else
                GetRijndaelEncrypt = Nothing
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetRijndaelEncrypt error")
            GetRijndaelEncrypt = Nothing
        End Try
    End Function

    Public Function GetRijndaelDecrypt(_txt As String) As String
        Try
            symmetricKey.Mode = CipherMode.CBC

            enc = New System.Text.UTF8Encoding
            decryptor = symmetricKey.CreateDecryptor(KEY_128, IV_128)

            Dim cypherTextBytes As Byte() = Convert.FromBase64String(_txt)
            Dim memoryStream As MemoryStream = New MemoryStream(cypherTextBytes)
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, Me.decryptor, CryptoStreamMode.Read)
            Dim plainTextBytes(cypherTextBytes.Length) As Byte
            Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
            memoryStream.Close()
            cryptoStream.Close()
            GetRijndaelDecrypt = Me.enc.GetString(plainTextBytes, 0, decryptedByteCount)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetRijndaelDecrypt error")
            GetRijndaelDecrypt = Nothing
        End Try
    End Function

    Public Function GetMD5Hash(_strToHash As String) As String
        'OK
        Try
            Dim md5Obj As New Security.Cryptography.MD5CryptoServiceProvider
            Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(_strToHash)
            bytesToHash = md5Obj.ComputeHash(bytesToHash)
            Dim strResult As String = ""
            For Each b As Byte In bytesToHash
                strResult += b.ToString("x2")
            Next
            Return strResult
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetMD5Hash Error")
            Return Nothing
        End Try
    End Function

    Public Function GetSHA1Hash(_strToHash As String) As String
        'OK
        Try
            Dim sha1Obj As New Security.Cryptography.SHA1CryptoServiceProvider
            Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(_strToHash)
            bytesToHash = sha1Obj.ComputeHash(bytesToHash)
            Dim strResult As String = ""
            For Each b As Byte In bytesToHash
                strResult += b.ToString("x2")
            Next
            Return strResult
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetSHA1Hash Error")
            Return Nothing
        End Try
    End Function

    Public Function EmailAddressCheck(_emailAddress As String) As Boolean
        'OK
        Try
            Dim pattern As String = "^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"
            Dim emailAddressMatch As Match = Regex.Match(_emailAddress, pattern)
            If emailAddressMatch.Success Then
                EmailAddressCheck = True
            Else
                EmailAddressCheck = False
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "EmailAddressCheck Error")
            EmailAddressCheck = False
        End Try
    End Function

    Public Function CleanInputInteger(_str As String) As String
        'OK
        Try
            'cek validasi angka pada segala macam input text, kalau bukan angka dimasukkan maka akan otomatis dihilangkan
            Return System.Text.RegularExpressions.Regex.Replace(_str, "[a-zA-Z\b\s-.,]", "")
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CleanInputInteger Error")
            Return Nothing
        End Try
    End Function

    Public Function CleanInputDouble(_str As String) As String
        'OK
        Try
            'sama dengan CleanInputInteger hanya saja karakter . tidak di eliminasi klw ada
            'Return System.Text.RegularExpressions.Regex.Replace(str, "[a-zA-Z\b\s-]", "")
            Return System.Text.RegularExpressions.Regex.Replace(_str, "[a-zA-Z\b\s]", "")
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CleanInputDouble Error")
            Return Nothing
        End Try
    End Function

    Public Function CleanInputAlphabet(_str As String) As String
        'OK
        Try
            'cek validasi huruf pada segala macam input text, selain huruf akan otomatis dihilangkan
            Return System.Text.RegularExpressions.Regex.Replace(_str, "[0-9\b\s-]", "")
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CleanInputAlphabet Error")
            Return Nothing
        End Try
    End Function

    Public Function CleanInputTime(_str As String) As String
        'OK
        Try
            'cek validasi time

            'Regular Expressions: Time hh: mm validation 24 hours format:  

            '([0-1]\d|2[0-3])([0-5]\d)

            'If you Then need seconds too: 

            '([0-1]\d|2[0-3])([0-5]\d)(([0-5]\d))?

            'Edit: To prevent trailing AM/PM, partial match can't be allowed or you have to put the expression between ^ and $.

            '^([0-1]\d|2[0-3])([0-5]\d)(:([0-5]\d))?$

            'Return System.Text.RegularExpressions.Regex.Replace(_str, "^([0-1]\d|2[0-3]):([0-5]\d)(:([0-5]\d))?$", "")
            If System.Text.RegularExpressions.Regex.Match(_str, "^([0-1]\d|2[0-3]):([0-5]\d)(:([0-5]\d))?$").Success Or System.Text.RegularExpressions.Regex.Match(_str, "^([0-1]\d|2[0-3]):([0-5]\d)?$").Success Then
                Return _str
            Else
                'Return System.Text.RegularExpressions.Regex.Match(_str, "^([0-1]\d|2[0-3]):([0-5]\d)(:([0-5]\d))?$").Value
                Return Nothing
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CleanInputTime Error")
            Return Nothing
        End Try
    End Function

    Public Function FromExcelSerialDate(_serialDate As Integer) As DateTime
        'OK
        Try
            If _serialDate > 59 Then _serialDate -= 1 ''// Excel/Lotus 2/29/1900 bug
            Return New DateTime(1899, 12, 31).AddDays(_serialDate)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "FromExcelSerialDate Error")
            Return Nothing
        End Try
    End Function

    Public Function SafeSqlLiteral(_strValue As String, Optional _intLevel As Integer = 1) As String
        'OK
        ' _intLevel represent how thorough the value will be checked for dangerous code
        ' _intLevel (1) - Do just the basic. This level will already counter most of the SQL injection attacks
        ' _intLevel (2) - &nbsp; (non breaking space) will be added to most words used in SQL queries to prevent unauthorized access to the database. Safe to be printed back into HTML code. Don't use for usernames or passwords

        Try
            If Not IsDBNull(_strValue) Then
                If _intLevel > 0 Then
                    _strValue = Trim(_strValue)
                    _strValue = Replace(_strValue, "'", "''") ' Most important one! This line alone can prevent most injection attacks
                    '_strValue = Replace(_strValue, """", "''")
                    '_strValue = Replace(_strValue, "--", "")
                    '_strValue = Replace(_strValue, "[", "[[]")
                    '_strValue = Replace(_strValue, "%", "[%]")
                End If

                If _intLevel > 1 Then
                    Dim myArray As Array
                    myArray = Split("xp_ ;update ;insert ;select ;drop ;alter ;create ;rename ;delete ;replace ", ";")
                    Dim i, i2, intLenghtLeft As Integer
                    For i = LBound(myArray) To UBound(myArray)
                        Dim rx As New Regex(myArray(i), RegexOptions.Compiled Or RegexOptions.IgnoreCase)
                        Dim matches As MatchCollection = rx.Matches(_strValue)
                        i2 = 0
                        For Each match As Match In matches
                            Dim groups As GroupCollection = match.Groups
                            intLenghtLeft = groups.Item(0).Index + Len(myArray(i)) + i2
                            _strValue = Left(_strValue, intLenghtLeft - 1) & "&nbsp;" & Right(_strValue, Len(_strValue) - intLenghtLeft)
                            i2 += 5
                        Next
                    Next
                End If

                '_strValue = Replace(_strValue, ";", ";&nbsp;")
                '_strValue = Replace(_strValue, "_", "[_]")

                Return _strValue
            Else
                Return _strValue
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SafeSqlLiteral Error")
            Return Nothing
        End Try
    End Function

    Public Sub ValidateTextBox(ByRef _textBox As TextBox, _tbName As String, Optional _formatNumber As String = "#,##0.00")
        'OK
        'BISA KOMA
        Try
            Dim temp As String
            Dim formatString As String = _formatNumber
            Dim tamp As Double
            temp = Trim(_textBox.Text)

            If (temp <> CleanInputDouble(Trim(_textBox.Text))) Then
                Call myCShowMessage.ShowWarning(_tbName & " harus berupa angka!", "Perhatian")
                _textBox.Text = CleanInputDouble(Trim(_textBox.Text))
                If Not (_textBox.Text.Length = 0) Then
                    tamp = _textBox.Text
                    'dikalikan 10 dulu agar angka di belakang koma bisa banyak
                    'misal tamp = 3.1456789 klw dikali 10 maka jadi 31.456789
                    'lalu kena fungsi math.ceiling jadi 31.5 (math.ceiling pasti dibulatkan ke atas)
                    'lalu terakhir dibagi 10 lagi maka jadi 3.15
                    tamp = tamp * 10
                    tamp = Math.Round(tamp * 10.0) / 10
                    tamp = tamp / 10
                    'tamp = -Int(-(tamp))
                    _textBox.Text = tamp.ToString(formatString)
                Else
                    _textBox.Text = "0.00"
                End If
            Else
                If Not (CleanInputDouble(Trim(_textBox.Text)).Length = 0) Then
                    tamp = CleanInputDouble(Trim(_textBox.Text))
                    tamp = tamp * 10
                    tamp = Math.Round(tamp * 10.0) / 10
                    tamp = tamp / 10
                    'tamp = -Int(-(tamp))
                    _textBox.Text = tamp.ToString(formatString)
                Else
                    _textBox.Text = "0.00"
                End If
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ValidateTextBox Error")
        End Try
    End Sub

    Public Sub ValidateTextBoxNumber(ByRef _textBox As TextBox, _tbName As String)
        'OK
        'TANPA KOMA
        Try
            Dim temp As String
            Dim formatString As String = "#,##"
            Dim tamp As Double
            temp = Trim(_textBox.Text)

            If (temp <> CleanInputInteger(Trim(_textBox.Text))) Then
                Call myCShowMessage.ShowWarning(_tbName & " harus berupa angka!", "Perhatian")
                _textBox.Text = CleanInputInteger(Trim(_textBox.Text))
                If Not (_textBox.Text.Length = 0) Then
                    tamp = _textBox.Text
                    'fungsinya sama dengan math.ceiling cuman klw yg ini hasilnya bulat tanpa koma
                    tamp = -Int(-(tamp))
                    _textBox.Text = tamp.ToString(formatString)
                Else
                    _textBox.Text = "0"
                End If
            Else
                If Not (CleanInputInteger(Trim(_textBox.Text)).Length = 0) Then
                    tamp = CleanInputInteger(Trim(_textBox.Text))
                    If (tamp <> 0) Then
                        'fungsinya sama dengan math.ceiling cuman klw yg ini hasilnya bulat tanpa koma
                        tamp = -Int(-(tamp))
                        _textBox.Text = tamp.ToString(formatString)
                    End If
                Else
                    _textBox.Text = "0"
                End If
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ValidateTextBoxNumber Error")
        End Try
    End Sub

    Public Sub ValidateLabel(ByRef _label As Label, _lblName As String)
        'OK
        Try
            Dim temp As String
            Dim formatString As String = "#,##0.00"
            Dim tamp As Double
            temp = Trim(_label.Text)

            If (temp <> CleanInputDouble(Trim(_label.Text))) Then
                Call myCShowMessage.ShowWarning(_lblName & " harus berupa angka!", "Perhatian")
                _label.Text = CleanInputDouble(Trim(_label.Text))
                If Not (_label.Text.Length = 0) Then
                    tamp = _label.Text
                    tamp = tamp * 10
                    tamp = Math.Ceiling(tamp * 10.0) / 10
                    'tamp = -Int(-(tamp))
                    tamp = tamp / 10
                    _label.Text = tamp.ToString(formatString)
                Else
                    _label.Text = "0.00"
                End If
            Else
                If Not (CleanInputDouble(Trim(_label.Text)).Length = 0) Then
                    tamp = CleanInputDouble(Trim(_label.Text))
                    tamp = tamp * 10
                    'tamp = -Int(-(tamp))
                    tamp = Math.Ceiling(tamp * 10.0) / 10
                    tamp = tamp / 10
                    _label.Text = tamp.ToString(formatString)
                Else
                    _label.Text = "0.00"
                End If
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ValidateLabel Error")
        End Try
    End Sub

    Public Function FindFirstNumericIndexInString(_string As String) As Integer
        Try
            'Dim mString As String = "PVC 12,7 X 21,5 CM OUTRE NEW GOLDEN"
            'Hasilnya index = 4
            Dim index As Integer = -1
            For i As Integer = 0 To _string.Length - 1
                If Char.IsDigit(_string(i)) Then
                    index = i
                    Exit For
                End If
            Next
            Return index
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "FindFirstNumericIndexInString Error")
            Return Nothing
        End Try
    End Function

    Public Function ConvertBulanToString(_kodeBulan As Integer) As String
        Try
            Select Case _kodeBulan
                Case 1
                    Return "Januari"
                Case 2
                    Return "Februari"
                Case 3
                    Return "Maret"
                Case 4
                    Return "April"
                Case 5
                    Return "Mei"
                Case 6
                    Return "Juni"
                Case 7
                    Return "Juli"
                Case 8
                    Return "Agustus"
                Case 9
                    Return "September"
                Case 10
                    Return "Oktober"
                Case 11
                    Return "November"
                Case 12
                    Return "Desember"
                Case Else
                    Return "Bulan Tidak terdefinisi"
            End Select
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ConvertBulanToString Error")
            Return "Bulan Tidak terdefinisi"
        End Try
    End Function

    Public Function CountChars(_value As String, _char As Char) As Integer
        Try
            Return _value.Count(Function(c) c = _char)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CountChars Error")
            Return Nothing
        End Try
    End Function

    Public Function SetStringForComm(_conn As Object, _tipe As String, _stSql As String) As String
        Try
            If (TypeOf _conn Is Npgsql.NpgsqlConnection) Then
                'Jika koneksi postgresql
                SetStringForComm = SetStringForQueryPgsql(_tipe, _stSql)
            Else
                'Klw bukan postgresql kirimkan balik tanpa diapa2kan
                Return _stSql
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetStringForComm Error")
            Return Nothing
        End Try
    End Function

    Public Function SetStringForQueryPgsql(_tipe As String, _value As String) As String
        Try
            'Dim mTempStr As String
            'Dim findIdx As Integer
            Dim mTempStringAwal As String
            'Dim mTempStringAkhir As String
            Dim mArrField() As String
            Dim mNewField As String
            Dim mIdx As Integer
            Dim mCrit() As String
            If Not IsNothing(_value) Then
                If (_tipe = "insert") Then
                    'INSERT INTO public."MsUser"('"Password", "Username")
                    'VALUES(?, ?);
                    _value = _value.Replace("insert ", "INSERT ")
                    _value = _value.Replace(" into ", " INTO ")
                    _value = _value.Replace(" values ", " VALUES ")
                    'mTempStringAwal = _value.Substring(_value.IndexOf("("), (_value.IndexOf(")") - _value.IndexOf("(") + 1))
                    'mTempStringAkhir = mTempStringAwal.Replace(" ", "")
                    '_value = _value.Replace(mTempStringAwal, mTempStringAkhir)

                    'Ambil nama tabel
                    mTempStringAwal = _value.Substring(_value.IndexOf("INTO") + 5, (_value.IndexOf("(") - _value.IndexOf("INTO")) - 6)
                    If Not (mTempStringAwal.Contains(".")) Then
                        mNewField = "public." & mTempStringAwal
                        _value = _value.Replace(mTempStringAwal, mNewField)
                    End If

                    'Diambil kriteria dari query yang sudah pasti
                    mCrit = New String() {" ", ",", "INSERT", "INTO", "VALUES", "(", ")"}

                    'Di split berdasar query tersebut, yang sudah ada di daftar dihilangkan, hasil yang kosong juga gak diikutkan
                    mArrField = _value.Split(mCrit, StringSplitOptions.RemoveEmptyEntries)
                    mIdx = 0

                    'SEMUA PENGATURAN " ditaruh di sini
                    While mIdx < mArrField.Length
                        mNewField = Nothing
                        If (mArrField(mIdx).Contains(".")) Then
                            'Jika ada . nya, berarti nama tabel, ada schema nya
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(".") + 1) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf(".") + 1, mArrField(mIdx).Length - (mArrField(mIdx).IndexOf(".") + 1)) & """"
                            End If
                        End If
                        If (mArrField(mIdx).Contains("=")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<>")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<>")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<=")) Then
                            'Klw ada <=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains(">=")) Then
                            'Klw ada >=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(">=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If IsNothing(mNewField) Then
                            If Not mNewField.Contains("""") Then
                                If Not (mArrField(mIdx) = "=") Then
                                    'Agar tanda = yang di inner join gak ikut dikasih ""
                                    mNewField = """" & mArrField(mIdx) & """"
                                End If
                            End If
                        End If
                        _value = _value.Replace(mArrField(mIdx), mNewField)

                        mIdx += 1
                    End While

                    'If (_value.Contains(".")) Then
                    '    'Jika sudah ada . nya
                    '    'gak perlu diapa2kan, karena berarti sudah ada schema name nya
                    'Else
                    '    'Kalau blm ada nama schemanya, maka ditambahkan public
                    '    'Semua yang gak ada schemanya dianggap sbg public
                    '    'STRING INSERT INTO sepanjang 12 char
                    '    _value = _value.Insert(12, "public.")
                    'End If
                    ''cari titik
                    'findIdx = _value.IndexOf(".")
                    'If (_value.Chars(findIdx + 1) <> """") Then
                    '    'Klw setelah . tidak ada ", maka ditambahkan "
                    '    _value = _value.Insert(findIdx + 1, """")
                    'End If
                    ''cari (
                    'findIdx = _value.IndexOf("(")
                    'If (_value.Chars(findIdx - 1) <> """") Then
                    '    'Klw setelah . tidak ada ", maka ditambahkan "
                    '    _value = _value.Insert(findIdx - 1, """")
                    'End If
                    ''cari ,
                    'If (_value.Contains(",")) Then
                    '    'ambil semua dulu field nya yang akan diisiin, ilangin petik2nya
                    '    mTempStringAwal = _value.Substring(_value.IndexOf("(") + 1, (_value.IndexOf(")") - _value.IndexOf("(")) - 1)
                    '    mArrField = mTempStringAwal.Split(",")
                    '    'Jika ada ,
                    '    mIdx = 0
                    '    While mIdx < mArrField.Length
                    '        'Di depan sama di belakang string nya dikasih petik "
                    '        mNewField = """" & mArrField(mIdx) & """"
                    '        'Lalu di replace kan sama nilai yang lama, di string aslinya
                    '        _value = _value.Replace(mArrField(mIdx), mNewField)
                    '        mIdx += 1
                    '    End While
                    'End If
                    SetStringForQueryPgsql = _value
                ElseIf (_tipe = "select") Then
                    'SELECT "Password", "Username"
                    'FROM public."MsUser";
                    _value = _value.Replace("select ", "SELECT ")
                    _value = _value.Replace(" from ", " FROM ")
                    _value = _value.Replace(" where ", " WHERE ")
                    _value = _value.Replace(" having ", " HAVING ")
                    _value = _value.Replace(" group by ", " GROUP BY ")
                    _value = _value.Replace(" inner join ", " INNER JOIN ")
                    _value = _value.Replace(" left join ", " LEFT JOIN ")
                    _value = _value.Replace(" right join ", " RIGHT JOIN ")
                    _value = _value.Replace(" outer join ", " OUTER JOIN ")
                    _value = _value.Replace(" union ", " UNION ")
                    _value = _value.Replace(" like ", " LIKE ")
                    _value = _value.Replace(" not ", " NOT ")
                    _value = _value.Replace(" and ", " AND ")
                    _value = _value.Replace(" or ", " OR ")
                    _value = _value.Replace(" between ", " BETWEEN ")
                    _value = _value.Replace(" as ", " AS ")
                    _value = _value.Replace(" on ", " ON ")
                    _value = _value.Replace(" top ", " TOP ")
                    InputBox("", "", _value)

                    'Ambil nama tabel
                    If (_value.Contains("WHERE")) Then
                        'Klw ada where nya, where diambil pertama kali, karena urutan query where itu paling dulu setelah from
                        mTempStringAwal = _value.Substring(_value.IndexOf("FROM") + 5, (_value.IndexOf("WHERE") - _value.IndexOf("FROM")) - 6)
                    ElseIf (_value.Contains("GROUP BY")) Then
                        'Klw gak ada where, maka selanjutnya cek ada group by atau nggak,
                        'having gak perlu di cek, karena kalau ada having, pasti ada group by, jadi tetap group by dulu yang di cek
                        mTempStringAwal = _value.Substring(_value.IndexOf("FROM") + 5, (_value.IndexOf("GROUP BY") - _value.IndexOf("FROM")) - 6)
                    Else
                        'Klw gak ada semua itu, maka tinggal cek ;
                        mTempStringAwal = _value.Substring(_value.IndexOf("FROM") + 5, (_value.IndexOf(";") - _value.IndexOf("FROM")) - 6)
                    End If

                    'Di split dulu berdasar spasi
                    mArrField = mTempStringAwal.Split(" ")
                    mIdx = 0
                    'Cek ada join atau union nggak, kalau ada join, berarti tabel nya lbh dari 1
                    If (mTempStringAwal.Contains("JOIN") Or mTempStringAwal.Contains("UNION")) Then
                        While mIdx < mArrField.Length
                            If Not (mArrField(mIdx).Contains("JOIN")) And Not (mArrField(mIdx).Contains("UNION")) And Not (mArrField(mIdx).Contains("INNER")) And Not (mArrField(mIdx).Contains("OUTER")) And Not (mArrField(mIdx).Contains("LEFT")) And Not (mArrField(mIdx).Contains("RIGHT")) Then
                                'Yang bukan join atau union, ditambahin public.
                                If Not (mArrField(mIdx).Contains(".")) Then
                                    mNewField = "public." & mArrField(mIdx)
                                    _value = _value.Replace(mArrField(mIdx), mNewField)
                                End If
                            End If
                        End While
                    Else
                        'Klw gak ada join atau union
                        While mIdx < mArrField.Length
                            If Not (mArrField(mIdx).Contains(".")) Then
                                mNewField = "public." & mArrField(mIdx)
                                _value = _value.Replace(mArrField(mIdx), mNewField)
                            End If
                            mIdx += 1
                        End While
                    End If

                    'Diambil kriteria dari query yang sudah pasti
                    mCrit = New String() {" ", ",", "SELECT", "TOP", "FROM", "WHERE", "GROUP BY", "INNER JOIN", "OUTER JOIN", "LEFT JOIN", "RIGHT JOIN", "UNION", "HAVING", "LIKE", "NOT", "END", "BETWEEN", "AS", "ON", "AND", "OR"}

                    'Di split berdasar query tersebut, yang sudah ada di daftar dihilangkan, hasil yang kosong juga gak diikutkan
                    mArrField = _value.Split(mCrit, StringSplitOptions.RemoveEmptyEntries)
                    mIdx = 0

                    'SEMUA PENGATURAN " ditaruh di sini
                    While mIdx < mArrField.Length
                        mNewField = Nothing
                        If (mArrField(mIdx).Contains(".")) Then
                            'Jika ada . nya, berarti nama tabel, ada schema nya
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(".") + 1) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf(".") + 1, mArrField(mIdx).Length - (mArrField(mIdx).IndexOf(".") + 1)) & """"
                            End If
                        End If
                        If (mArrField(mIdx).Contains("=")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<>")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<>")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<=")) Then
                            'Klw ada <=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains(">=")) Then
                            'Klw ada >=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(">=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If IsNothing(mNewField) Then
                            If Not mNewField.Contains("""") Then
                                If Not (mArrField(mIdx) = "=") Then
                                    'Agar tanda = yang di inner join gak ikut dikasih ""
                                    mNewField = """" & mArrField(mIdx) & """"
                                End If
                            End If
                        End If
                        _value = _value.Replace(mArrField(mIdx), mNewField)

                        mIdx += 1
                    End While
                    SetStringForQueryPgsql = _value
                ElseIf (_tipe = "delete") Then
                    'DELETE FROM public."MsUser"
                    'WHERE <condition>;
                    _value = _value.Replace("delete ", "DELETE ")
                    _value = _value.Replace(" from ", " FROM ")
                    _value = _value.Replace(" where ", " WHERE ")
                    _value = _value.Replace(" inner join ", " INNER JOIN ")
                    _value = _value.Replace(" left join ", " LEFT JOIN ")
                    _value = _value.Replace(" right join ", " RIGHT JOIN ")
                    _value = _value.Replace(" outer join ", " OUTER JOIN ")
                    _value = _value.Replace(" like ", " LIKE ")
                    _value = _value.Replace(" not ", " NOT ")
                    _value = _value.Replace(" and ", " AND ")
                    _value = _value.Replace(" or ", " OR ")
                    _value = _value.Replace(" between ", " BETWEEN ")
                    _value = _value.Replace(" as ", " AS ")
                    _value = _value.Replace(" on ", " ON ")
                    _value = _value.Replace(" top ", " TOP ")

                    'AMBIL NAMA TABEL
                    If (_value.Contains("WHERE")) Then
                        'Klw ada where nya, where diambil pertama kali, karena urutan query where itu paling dulu setelah from
                        mTempStringAwal = _value.Substring(_value.IndexOf("FROM") + 5, (_value.IndexOf("WHERE") - _value.IndexOf("FROM")) - 6)
                    Else
                        'Klw gak ada semua itu, maka tinggal cek ;
                        mTempStringAwal = _value.Substring(_value.IndexOf("FROM") + 5, (_value.IndexOf(";") - _value.IndexOf("FROM")) - 6)
                    End If
                    'Di cek sudah ada skemanya blm, klw blm ada ditambahin
                    If Not (mTempStringAwal.Contains(".")) Then
                        mNewField = "public." & mTempStringAwal
                        _value = _value.Replace(mTempStringAwal, mNewField)
                    End If

                    'Diambil kriteria dari query yang sudah pasti
                    mCrit = New String() {" ", ",", "DELETE", "TOP", "FROM", "WHERE", "INNER JOIN", "OUTER JOIN", "LEFT JOIN", "RIGHT JOIN", "LIKE", "NOT", "END", "BETWEEN", "AS", "ON", "AND", "OR"}

                    'Di split berdasar query tersebut, yang sudah ada di daftar dihilangkan, hasil yang kosong juga gak diikutkan
                    mArrField = _value.Split(mCrit, StringSplitOptions.RemoveEmptyEntries)
                    mIdx = 0
                    'SEMUA PENGATURAN " ditaruh di sini
                    While mIdx < mArrField.Length
                        mNewField = Nothing
                        If (mArrField(mIdx).Contains(".")) Then
                            'Jika ada . nya, berarti nama tabel, ada schema nya
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(".") + 1) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf(".") + 1, mArrField(mIdx).Length - (mArrField(mIdx).IndexOf(".") + 1)) & """"
                            End If
                        End If
                        If (mArrField(mIdx).Contains("=")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<>")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<>")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<=")) Then
                            'Klw ada <=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains(">=")) Then
                            'Klw ada >=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(">=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If IsNothing(mNewField) Then
                            If Not mNewField.Contains("""") Then
                                If Not (mArrField(mIdx) = "=") Then
                                    'Agar tanda = yang di inner join gak ikut dikasih ""
                                    mNewField = """" & mArrField(mIdx) & """"
                                End If
                            End If
                        End If
                        _value = _value.Replace(mArrField(mIdx), mNewField)

                        mIdx += 1
                    End While
                    SetStringForQueryPgsql = _value
                ElseIf (_tipe = "update") Then
                    'UPDATE public."MsUser"
                    'SET "Password"=?, "Username"=?
                    'WHERE <condition>;

                    _value = _value.Replace("update ", "UPDATE ")
                    _value = _value.Replace(" set ", " SET ")
                    _value = _value.Replace(" where ", " WHERE ")
                    _value = _value.Replace(" inner join ", " INNER JOIN ")
                    _value = _value.Replace(" left join ", " LEFT JOIN ")
                    _value = _value.Replace(" right join ", " RIGHT JOIN ")
                    _value = _value.Replace(" outer join ", " OUTER JOIN ")
                    _value = _value.Replace(" like ", " LIKE ")
                    _value = _value.Replace(" not ", " NOT ")
                    _value = _value.Replace(" and ", " AND ")
                    _value = _value.Replace(" or ", " OR ")
                    _value = _value.Replace(" between ", " BETWEEN ")
                    _value = _value.Replace(" as ", " AS ")
                    _value = _value.Replace(" on ", " ON ")

                    'AMBIL NAMA TABEL
                    mTempStringAwal = _value.Substring(_value.IndexOf("UPDATE") + 7, (_value.IndexOf("SET") - _value.IndexOf("UPDATE")) - 8)
                    'Di cek sudah ada skemanya blm, klw blm ada ditambahin
                    If Not (mTempStringAwal.Contains(".")) Then
                        mNewField = "public." & mTempStringAwal
                        _value = _value.Replace(mTempStringAwal, mNewField)
                    End If

                    'Diambil kriteria dari query yang sudah pasti
                    mCrit = New String() {" ", ",", "UPDATE", "TOP", "SET", "WHERE", "INNER JOIN", "OUTER JOIN", "LEFT JOIN", "RIGHT JOIN", "LIKE", "NOT", "END", "BETWEEN", "AS", "ON", "AND", "OR"}

                    'Di split berdasar query tersebut, yang sudah ada di daftar dihilangkan, hasil yang kosong juga gak diikutkan
                    mArrField = _value.Split(mCrit, StringSplitOptions.RemoveEmptyEntries)
                    mIdx = 0
                    'SEMUA PENGATURAN " ditaruh di sini
                    While mIdx < mArrField.Length
                        mNewField = Nothing
                        If (mArrField(mIdx).Contains(".")) Then
                            'Jika ada . nya, berarti nama tabel, ada schema nya
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(".") + 1) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf(".") + 1, mArrField(mIdx).Length - (mArrField(mIdx).IndexOf(".") + 1)) & """"
                            End If
                        End If
                        If (mArrField(mIdx).Contains("=")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<>")) Then
                            'Klw ada = atau <>
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<>")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains("<=")) Then
                            'Klw ada <=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf("<=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If (mArrField(mIdx).Contains(">=")) Then
                            'Klw ada >=
                            If Not mArrField(mIdx).Contains("""") Then
                                mNewField = """" & mArrField(mIdx).Substring(0, mArrField(mIdx).IndexOf(">=")) & """" & mArrField(mIdx).Substring(mArrField(mIdx).IndexOf("="), mArrField(mIdx).Length - (mArrField(mIdx).IndexOf("=") + 1))
                            End If
                        End If
                        If IsNothing(mNewField) Then
                            If Not mNewField.Contains("""") Then
                                If Not (mArrField(mIdx) = "=") Then
                                    'Agar tanda = yang di inner join gak ikut dikasih ""
                                    mNewField = """" & mArrField(mIdx) & """"
                                End If
                            End If
                        End If
                        _value = _value.Replace(mArrField(mIdx), mNewField)

                        mIdx += 1
                    End While
                    SetStringForQueryPgsql = _value
                Else
                    Return Nothing
                End If
            Else
                SetStringForQueryPgsql = Nothing
            End If

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetStringForQuery Error")
            Return Nothing
        End Try
    End Function

    Public Function SetInsertValues(_addValues As Object, _dataType As String) As String
        Try
            If Not IsDBNull(_addValues) Then
                If Not IsNothing(_addValues) Then
                    Select Case _dataType
                        Case "Byte", "Decimal", "Double", "Int16", "Int32", "Int64", "SByte", "Single", "UInt16", "UInt32", "UInt64"
                            SetInsertValues = _addValues
                        Case "String", "Char", "Guid"
                            If (_addValues <> "") Then
                                SetInsertValues = "'" & SafeSqlLiteral(_addValues, 1) & "'"
                            Else
                                SetInsertValues = "Null"
                            End If
                        Case "TimeSpan"
                            If Not IsNothing(_addValues.ToString) Then
                                SetInsertValues = "'" & _addValues.ToString & "'"
                            Else
                                SetInsertValues = "Null"
                            End If
                        Case "Boolean"
                            If Not IsNothing(_addValues) Then
                                SetInsertValues = "'" & _addValues & "'"
                            Else
                                SetInsertValues = "Null"
                            End If
                        Case "DateTime"
                            If (_addValues.ToString.Contains(":")) Then
                                SetInsertValues = "'" & _addValues & "'"
                            Else
                                SetInsertValues = "'" & Format(_addValues, "dd-MMM-yyyy") & "'"
                            End If
                        Case Else
                            If (_addValues <> "") Then
                                SetInsertValues = "'" & SafeSqlLiteral(_addValues, 1) & "'"
                            Else
                                SetInsertValues = "Null"
                            End If
                    End Select
                Else
                    SetInsertValues = "Null"
                End If
            Else
                SetInsertValues = "Null"
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetInsertValues Error")
            Return Nothing
        End Try
    End Function

    Public Function SetInsertFields(_addValues As Object, _addFields As Object) As String
        Try
            SetInsertFields = _addFields
            'If Not IsDBNull(addValues) Then
            '    SetInsertFields = addFields
            'Else
            '    Return Nothing
            'End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetInsertFields Error")
            Return Nothing
        End Try
    End Function

    Public Function SetUpdateString(_updateFields As String, _updateValues As Object, _dataType As String) As String
        Try
            If Not IsDBNull(_updateValues) Then
                Select Case _dataType
                    Case "Boolean", "Byte", "Decimal", "Double", "Int16", "Int32", "Int64", "SByte", "Single", "UInt16", "UInt32", "UInt64"
                        SetUpdateString = _updateFields & "=" & _updateValues
                    Case "String", "Char", "Guid", "TimeSpan"
                        SetUpdateString = _updateFields & "=" & "'" & SafeSqlLiteral(_updateValues, 1) & "'"
                    Case "DateTime"
                        SetUpdateString = _updateFields & "=" & "'" & Format(_updateValues, "dd-MMM-yyyy") & "'"
                    Case Else
                        SetUpdateString = _updateFields & "=" & "'" & SafeSqlLiteral(_updateValues, 1) & "'"
                End Select
            Else
                SetUpdateString = "Null"
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetUpdateString Error")
            Return Nothing
        End Try
    End Function

    Public Sub ReplaceStringInDataTable(ByRef _dataTable As DataTable, _columnName As String, _oldString As String, _newString As String)
        Try
            Dim idx As Integer = 0
            While idx < _dataTable.Rows.Count
                If Not IsDBNull(_dataTable.Rows(idx).Item(_columnName)) Then
                    _dataTable.Rows(idx).Item(_columnName) = _dataTable.Rows(idx).Item(_columnName).ToString.Replace(_oldString, _newString)
                End If

                idx += 1
            End While
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReplaceStringInDataTable Error")
        End Try
    End Sub

    Public Function Terbilang(_angka As String) As String
        Try
            Dim nilai As String
            Dim tanda As String
            Dim desimal As String
            Dim intAngka As Long
            Dim satuan As Long
            Dim puluhan As Long
            Dim ratusan As Long
            Dim nilaiHuruf As String
            Dim multi As Integer
            Dim riil As String
            Dim pecahan As Long
            Dim i As Integer

            Dim digit(10) As String
            Dim order(15) As String

            digit(0) = "Zero"
            digit(1) = "One"
            digit(2) = "Two"
            digit(3) = "Three"
            digit(4) = "Four"
            digit(5) = "Five"
            digit(6) = "Six"
            digit(7) = "Seven"
            digit(8) = "Eight"
            digit(9) = "Nine"
            digit(10) = "Ten"
            digit(11) = "Eleven"
            digit(12) = "Twelv"
            digit(13) = "Thirteen"
            digit(14) = "Fourteen"
            digit(15) = "Fiveteen"
            digit(16) = "Sixteen"
            digit(17) = "Seventeen"
            digit(18) = "Eighteen"
            digit(19) = "Nineteen"
            digit(20) = "Twenty"

            order(0) = ""
            order(1) = "ty"
            order(2) = "Hundred"
            order(3) = "Tousand"
            order(6) = "Million"
            order(9) = "Billion"
            order(12) = "Triliun"
            order(15) = "Kuadriliun"

            _angka = IIf(_angka = "", _angka = "0", _angka)

            nilai = ""
            tanda = IIf(Left$(_angka, 1) = "-", "minus ", "")
            _angka = Replace(_angka, "-", "")
            pecahan = InStr(_angka, ".")
            'pecahan = IIf(pecahan = 0, 1, pecahan)
            If pecahan = 0 Then
                riil = _angka
                desimal = ""
            ElseIf pecahan = 1 Then
                riil = Left$(_angka, 0)
                desimal = Right$(_angka, Len(_angka) - pecahan)
            Else
                riil = Left$(_angka, pecahan - 1)
                desimal = Right$(_angka, Len(_angka) - pecahan)
            End If

            While (riil <> "")

                intAngka = CInt(Right$(riil, 3))

                'ambil satuan,puluhan,ratus
                satuan = intAngka Mod 10
                puluhan = (intAngka Mod 100 - satuan) / 10
                ratusan = (intAngka - puluhan * 10 - satuan) / 100

                'konversi ratusan
                nilai = digit(ratusan) & " " & order(2) & " "

                'konversi puluhan dan satuan
                If puluhan = 0 Then
                    nilai = IIf(nilai <> "", nilai, "") & digit(satuan)
                ElseIf puluhan = 1 Then
                    If satuan = 0 Then
                        nilai = IIf(nilai <> "", nilai, "") & "Se" & order(1)
                    ElseIf satuan = 1 Then
                        nilai = IIf(nilai <> "", nilai, "") & "Sebelas"
                    Else
                        nilai = IIf(nilai <> "", nilai, "") & digit(satuan) & "belas"
                    End If
                Else
                    nilai = IIf(nilai <> "", nilai, "") & digit(puluhan) & " puluh "
                End If

                'gabungkan dengan hasil sebelumnya
                nilaiHuruf = IIf(nilai <> "", nilai & (IIf(nilai = "se", "", " ")) & order(multi) & " ", "")

                multi = multi + 3
                'riil = Replace(riil, intAngka, "", Len(CStr(riil)) - multi)
                Dim substract&
                substract = IIf(Len(CStr(riil)) < 3, Len(CStr(riil)), 3)

                riil = Left(CStr(riil), Len(CStr(riil)) - substract)

            End While

            nilai = ""

            For i = 1 To Len(desimal)
                nilai = nilai & IIf(nilai <> "", " ", "") & digit(Mid(desimal, i, 1))
            Next


            Terbilang = tanda & Trim(nilaiHuruf) & (IIf(nilai <> "", " Point " & nilai, ""))
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "Terbilang Error")
            Terbilang = Nothing
        End Try
    End Function

    ' Convert a number between 0 and 999 into words.
    Private Function GroupToWords(_num As Integer) As String
        Try
            Static one_to_nineteen() As String = {"zero", "one",
            "two", "three", "four", "five", "six", "seven",
            "eight", "nine", "ten", "eleven", "twelve",
            "thirteen", "fourteen", "fifteen", "sixteen",
            "seventeen", "eightteen", "nineteen"}
            Static multiples_of_ten() As String = {"twenty",
                "thirty", "forty", "fifty", "sixty", "seventy",
                "eighty", "ninety"}

            ' If the number is 0, return an empty string.
            If _num = 0 Then Return ""

            ' Handle the hundreds digit.
            Dim digit As Integer
            Dim result As String = ""
            If _num > 99 Then
                digit = _num \ 100
                _num = _num Mod 100
                result = one_to_nineteen(digit) & " hundred"
            End If

            ' If _num = 0, we have hundreds only.
            If _num = 0 Then Return result.Trim()

            ' See if the rest is less than 20.
            If _num < 20 Then
                ' Look up the correct name.
                result &= " " & one_to_nineteen(_num)
            Else
                ' Handle the tens digit.
                digit = _num \ 10
                _num = _num Mod 10
                result &= " " & multiples_of_ten(digit - 2)

                ' Handle the final digit.
                If _num > 0 Then
                    result &= " " & one_to_nineteen(_num)
                End If
            End If

            Return result.Trim()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GroupToWords Error")
            Return Nothing
        End Try
    End Function

    ' Return a word representation of the whole number value.
    Private Function NumberToString(_num As Double) As String
        Try
            ' Remove any fractional part.
            _num = Int(_num)

            ' If the number is 0, return zero.
            If _num = 0 Then Return "zero"

            Static groups() As String = {"", "thousand", "million",
                "billion", "trillion", "quadrillion", "?", "??",
                "???", "????"}
            Dim result As String = ""

            ' Process the groups, smallest first.
            Dim quotient As Double
            Dim remainder As Integer
            Dim group_num As Integer = 0
            Do While _num > 0
                ' Get the next group of three digits.
                quotient = Int(_num / 1000)
                remainder = CInt(_num - quotient * 1000)
                _num = quotient

                ' Convert the group into words.
                result = GroupToWords(remainder) &
                    " " & groups(group_num) & ", " &
                    result

                ' Get ready for the next group.
                group_num += 1
            Loop

            ' Remove the trailing ", ".
            If result.EndsWith(", ") Then
                result = result.Substring(0, result.Length - 2)
            End If

            Return result.Trim()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "NumberToString Error")
            Return Nothing
        End Try
    End Function

    ' Return a word representation of a dollar amount.
    Public Function DollarsToString(_money As Double, Optional _after_dollars As String = " DOLLARS AND ", Optional _after_cents As String = " CENTS") As String
        Try
            ' Get the dollar and cents parts.
            Dim dollars As Double = Int(_money)
            Dim cents As Double = CInt(100 * (_money - dollars))

            ' Convert the parts into words.
            Dim dollars_str As String = NumberToString(dollars)
            Dim cents_str As String = NumberToString(cents)

            Return dollars_str & _after_dollars & cents_str & _after_cents
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DollarsToString Error")
            Return Nothing
        End Try
    End Function

    Public Function TranslateWeekdayName(_weekdayName As String) As String
        Try
            Select Case _weekdayName
                Case "Sunday"
                    Return "Minggu"
                Case "Monday"
                    Return "Senin"
                Case "Tuesday"
                    Return "Selasa"
                Case "Wednesday"
                    Return "Rabu"
                Case "Thursday"
                    Return "Kamis"
                Case "Friday"
                    Return "Jumat"
                Case "Saturday"
                    Return "Sabtu"
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Return Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "TranslateWeekdayName Error")
        End Try
    End Function

    Public Function QueryBuilder(_queryType As String, _tablename As String, _values As String, Optional _fields As String = Nothing, Optional _where As String = Nothing) As String
        Try
            If (_queryType = "insert") Then
                If _fields.Length = 0 Then
                    QueryBuilder = "INSERT INTO " & _tablename & " VALUES (" & _values & ");"
                Else
                    QueryBuilder = "INSERT INTO " & _tablename & " (" & _fields & ") VALUES (" & _values & ");"
                End If
            ElseIf (_queryType = "update") Then
                If _where.Length = 0 Then
                    QueryBuilder = "UPDATE " & _tablename & " SET " & _values & ";"
                Else
                    QueryBuilder = "UPDATE " & _tablename & " SET " & _values & " WHERE " & _where & ";"
                End If
            ElseIf (_queryType = "delete") Then
                If _where.Length = 0 Then
                    QueryBuilder = "DELETE FROM " & _tablename & ";"
                Else
                    QueryBuilder = "DELETE FROM " & _tablename & " WHERE " & _where & ";"
                End If
            End If
        Catch ex As Exception
            QueryBuilder = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "QueryBuilder Error")
        End Try
    End Function

    Public Function CalculateDaysBetweenDates(_dateValueStart As Date, _dateValueEnd As Date) As UShort
        Try
            Dim ts As TimeSpan = _dateValueEnd.Subtract(_dateValueStart)

            If Convert.ToUInt16(ts.Days) >= 0 Then
                CalculateDaysBetweenDates = Convert.ToUInt16(ts.Days)
            Else
                Call myCShowMessage.ShowWarning("Invalid Input")
                CalculateDaysBetweenDates = 0
            End If
        Catch ex As Exception
            CalculateDaysBetweenDates = 0
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CalculateDaysBetweenDates Error")
        End Try
    End Function
End Class
