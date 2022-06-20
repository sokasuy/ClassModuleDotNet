Imports System.Windows.Forms

Public Class CMiscFunction
    Private myCShowMessage As New CShowMessage.CShowMessage

    Public Function FormatFieldValForDbUpdate(_stFieldValue As String, _bFieldType As Byte) As String
        'belum di tes
        Dim NewFieldVal As String
        Try
            NewFieldVal = Trim(_stFieldValue)
            Select Case _bFieldType
                Case 1
                    'string >>> mengasumsikan boleh empty string >>> utk SQL saja. utk Access normalnya diubah jadi Null
                    If NewFieldVal = "" Then
                        NewFieldVal = "Null"
                    Else
                        '>>> mestinya kiri-kanan dicek apa sdh ada petik tunggalnya
                        NewFieldVal = " '" & NewFieldVal & "' "
                    End If
                Case 2
                    'numeric
                    If NewFieldVal = "" Then
                        NewFieldVal = "Null"
                        '            Else
                        '                NewFieldVal = CStr(.TextMatrix(Row, Col))
                    End If
                Case 3
                    'date
                    If IsDate(NewFieldVal) = True Then
                        'utk MSSql
                        'NewFieldVal = ConvDateForMsSQL(.TextMatrix(Row, Col))
                        'NewFieldVal = ConvDateForMsSQL(NewFieldVal)       '>>> didisabled utk saat ini
                        'utk Access
                        NewFieldVal = "#" & NewFieldVal & "#"
                    Else
                        NewFieldVal = "Null"
                    End If
                Case 4
                    'boolean >>> bagaimana jika dalam nilai 0 atau 1 atau -1 ?
                    If NewFieldVal = "" Then
                        NewFieldVal = "False"
                    Else
                        If (IsNumeric(NewFieldVal)) Then
                            If (Val(NewFieldVal) = 0) Then
                                NewFieldVal = "False"
                            Else
                                NewFieldVal = "True"
                            End If
                            '                Else
                            '                    'diasumsikan user benar memasukkan true / false
                        End If
                    End If
            End Select
            FormatFieldValForDbUpdate = NewFieldVal
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "FormatFieldValForDbUpdate Error")
            FormatFieldValForDbUpdate = Nothing
        End Try
    End Function

    Public Sub Wait(_delay As Short, Optional _delayUnit As String = "s", Optional _dispHrGlass As Boolean = False)
        'belum di tes
        'utk memaksa cpu menunggu sebelum proses berikutnya dilakukan.
        'opsi lain adalah DelayUnit = "n" ( = menit )
        Try
            Dim DelayEnd As Double
            If (_dispHrGlass = True) Then Cursor.Current = Cursors.WaitCursor

            DelayEnd = DateAdd(_delayUnit, _delay, Now).ToOADate
            While DateDiff(DateInterval.Second, Now, DateTime.FromOADate(DelayEnd)) > 0 'If date1 refers to a later point in time than date2, the DateDiff function returns a negative number.
                'tunggu
            End While

            If (_dispHrGlass = True) Then Cursor.Current = Cursors.Default
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "Wait Error")
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Function HitungRangeBayar(_dateBayar As Date, _dateNota As Date) As Integer
        Try
            Dim mMonth As Integer
            Dim arrMyDateBayar(3) As String
            Dim arrMyDateNota(3) As String

            arrMyDateBayar = _dateBayar.ToShortDateString.Split("/")
            arrMyDateNota = _dateNota.ToShortDateString.Split("/")

            If (Integer.Parse(arrMyDateBayar(2)) = Integer.Parse(arrMyDateNota(2))) Then
                'Klw di tahun yang sama
                'bulan pembayaran cuman bisa sama dengan atau lebih besar dari bulan pembuatan nota
                If (Integer.Parse(arrMyDateBayar(0)) = Integer.Parse(arrMyDateNota(0))) Then
                    'klw di bulan yang sama
                    'berarti belum terhitung 1 bulan
                    mMonth = 0
                ElseIf (Integer.Parse(arrMyDateBayar(0)) > Integer.Parse(arrMyDateNota(0))) Then
                    'klw bulan pembayaran lebih dari bulan pembuatan nota
                    'bulannya dikurangi dulu
                    mMonth = Integer.Parse(arrMyDateBayar(0)) - Integer.Parse(arrMyDateNota(0))
                    'lalu di cek apakah tanggalnya sama atw gak
                    If (Integer.Parse(arrMyDateBayar(1)) = Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal bayar sama dengan tanggal pembuatan nota berarti sudah benar bulannya
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) > Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal bayar lebih dari tanggal pembuatan nota
                        mMonth += 1
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) < Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal bayar kurang dari tanggal pembuatan nota
                        'bulan tetap seperti klw tanggalnya sama
                        'mMonth += 1
                    End If
                End If
            ElseIf (Integer.Parse(arrMyDateBayar(2)) > Integer.Parse(arrMyDateNota(2))) Then
                'klw tahun pembayaran lebih dari tahun pembuatan nota
                If (Integer.Parse(arrMyDateBayar(0)) = Integer.Parse(arrMyDateNota(2))) Then
                    'klw di bulan yang sama
                    mMonth = (Integer.Parse(arrMyDateBayar(2)) - Integer.Parse(arrMyDateNota(2))) * 12
                    If (Integer.Parse(arrMyDateBayar(1)) = Integer.Parse(arrMyDateNota(1))) Then
                        'klw di tanggal yang sama
                        'berarti ya cuman banyaknya selisih tahun * 12
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) > Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran lebih dari tanggal nota, berarti 13 bulan
                        mMonth += 1
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) < Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran kurang dari tanggal nota, berarti ikut klw sama tanggalnya
                        'mMonth = 13
                    End If
                ElseIf (Integer.Parse(arrMyDateBayar(0)) > Integer.Parse(arrMyDateNota(0))) Then
                    'klw bulan pembayaran lebih dari bulan pembuatan nota
                    'berapa tahun * 12, lalu ditambah selisih bulan bayar dikurang nota
                    mMonth = ((Integer.Parse(arrMyDateBayar(2)) - Integer.Parse(arrMyDateNota(2))) * 12) + (Integer.Parse(arrMyDateBayar(0)) - Integer.Parse(arrMyDateNota(0)))
                    'lalu di cek apakah tanggalnya sama atw gak
                    If (Integer.Parse(arrMyDateBayar(1)) = Integer.Parse(arrMyDateNota(1))) Then
                        'klw di tanggal yang sama
                        'berarti ya cuman banyaknya selisih tahun * 12
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) > Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran lebih dari tanggal nota, berarti 13 bulan
                        mMonth += 1
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) < Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran kurang dari tanggal nota, berarti ikut klw sama tanggalnya
                        'mMonth = 13
                    End If
                ElseIf (Integer.Parse(arrMyDateBayar(0)) < Integer.Parse(arrMyDateNota(0))) Then
                    'klw bulan pembayaran kurang dari bulan pembuatan nota
                    'berapa tahun * 12, lalu ditambah selisih bulan bayar dikurang nota
                    'PS: bulan bayar lebih kecil dari bulan nota, jadi pasti hasilnya negatif
                    'misal bulan buat nota bulan 10, lalu bulan bayar bulan 5 tahun depannya
                    'maka akan jadi begini: 12 + (5-10)
                    mMonth = ((arrMyDateBayar(2) - Integer.Parse(arrMyDateNota(2))) * 12) + (arrMyDateBayar(0) - arrMyDateNota(0))
                    'lalu di cek apakah tanggalnya sama atw gak
                    If (Integer.Parse(arrMyDateBayar(1)) = Integer.Parse(arrMyDateNota(1))) Then
                        'klw di tanggal yang sama
                        'berarti ya cuman banyaknya selisih tahun * 12
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) > Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran lebih dari tanggal nota, berarti 13 bulan
                        mMonth += 1
                    ElseIf (Integer.Parse(arrMyDateBayar(1)) < Integer.Parse(arrMyDateNota(1))) Then
                        'klw tanggal pembayaran kurang dari tanggal nota, berarti ikut klw sama tanggalnya
                    End If
                End If
            End If
            HitungRangeBayar = mMonth
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "HitungRangeBayar Error")
            HitungRangeBayar = Nothing
        End Try
    End Function

    Public Function WeekNumberOfMonthAlternate(_date As Date) As Integer
        Try
            'KLW HANYA DIBAGI 7
            Dim weekNumber As Integer = Math.Floor(_date.Day / 7) + 1
            Return weekNumber
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "WeekNumberOfMonthAlternate Error")
        End Try
    End Function

    Public Function GetWeekNumberInAMonth(_date As Date) As Integer
        Try
            'MINGGU SEBENARNYA SESUAI DENGAN KALENDER
            Dim DayNo As Integer = 0
            Dim DateNo As Integer = 0
            Dim StartingDay As Integer = 0
            Dim PayPerEndDate As Date = _date
            Dim week As Integer = 1
            StartingDay = PayPerEndDate.AddDays(-(PayPerEndDate.Day - 1)).DayOfWeek
            DayNo = PayPerEndDate.DayOfWeek
            DateNo = PayPerEndDate.Day
            week = Fix(DateNo / 7)
            If DateNo Mod 7 > 0 Then
                week += 1
            End If
            If StartingDay > DayNo Then
                week += 1
            End If
            Return week
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetWeekNumberInAMonth Error")
            Return Nothing
        End Try
    End Function

    'Public Sub ResetForm(_root As Control)
    '    'TINGGAL DI COMMENT SAJA KARENA SUDAH DIPINDAH KE ClassFormManipulation
    '    '16 Desember 2021
    '    Try
    '        For Each ctrl As Control In _root.Controls
    '            ResetForm(ctrl)

    '            If TypeOf ctrl Is TextBox Then
    '                CType(ctrl, TextBox).Clear()
    '            ElseIf TypeOf ctrl Is ComboBox Then
    '                CType(ctrl, ComboBox).SelectedIndex = -1
    '            ElseIf TypeOf ctrl Is CheckBox Then
    '                CType(ctrl, CheckBox).Checked = False
    '            ElseIf TypeOf ctrl Is RichTextBox Then
    '                CType(ctrl, RichTextBox).Clear()
    '            ElseIf TypeOf ctrl Is DataGridView Then
    '                CType(ctrl, DataGridView).Rows.Clear()
    '            ElseIf TypeOf ctrl Is DateTimePicker Then
    '                CType(ctrl, DateTimePicker).ResetText()
    '            ElseIf TypeOf ctrl Is NumericUpDown Then
    '                CType(ctrl, NumericUpDown).ResetText()
    '            End If
    '        Next ctrl
    '    Catch ex As Exception
    '        Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ResetForm Error")
    '    End Try
    'End Sub

    'Public Sub SetReadOnly(_root As Control, _isReadOnly As Boolean)
    '    'TINGGAL DI COMMENT SAJA KARENA SUDAH DIPINDAH KE ClassFormManipulation
    '    '16 Desember 2021
    '    Try
    '        For Each ctrl As Control In _root.Controls
    '            If TypeOf ctrl Is TextBox Then
    '                CType(ctrl, TextBox).ReadOnly = _isReadOnly
    '            ElseIf TypeOf ctrl Is ComboBox Then
    '                CType(ctrl, ComboBox).Enabled = IIf(_isReadOnly, False, True)
    '            ElseIf TypeOf ctrl Is CheckBox Then
    '                CType(ctrl, CheckBox).Enabled = IIf(_isReadOnly, False, True)
    '            ElseIf TypeOf ctrl Is RichTextBox Then
    '                CType(ctrl, RichTextBox).ReadOnly = _isReadOnly
    '            ElseIf TypeOf ctrl Is DataGridView Then
    '                CType(ctrl, DataGridView).ReadOnly = _isReadOnly
    '            ElseIf TypeOf ctrl Is DateTimePicker Then
    '                CType(ctrl, DateTimePicker).Enabled = IIf(_isReadOnly, False, True)
    '            ElseIf TypeOf ctrl Is NumericUpDown Then
    '                CType(ctrl, NumericUpDown).ReadOnly = _isReadOnly
    '            End If
    '        Next ctrl
    '    Catch ex As Exception
    '        Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetReadOnly Error")
    '    End Try
    'End Sub

    'Public Sub SetCheckListBoxUserRights(ByRef _clbUserRight As CheckedListBox, _isSuperuser As Boolean, _formName As String, _dtTableUserRight As DataTable)
    '    'TINGGAL DI COMMENT SAJA KARENA SUDAH DIPINDAH KE ClassFormManipulation
    '    '16 Desember 2021
    '    Try
    '        If Not _isSuperuser Then
    '            'Hanya di cek kalau bukan superuser
    '            'Superuser bisa segalanya
    '            Dim foundRows As DataRow()
    '            foundRows = _dtTableUserRight.Select("formname='" & _formName & "'")
    '            For i As Integer = 0 To _clbUserRight.Items.Count - 1
    '                _clbUserRight.SetItemChecked(i, _dtTableUserRight.Rows(_dtTableUserRight.Rows.IndexOf(foundRows(0))).Item(_clbUserRight.Items(i)))
    '            Next
    '        Else
    '            For i As Integer = 0 To _clbUserRight.Items.Count - 1
    '                _clbUserRight.SetItemChecked(i, True)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetCheckListBoxUserRights Error")
    '    End Try
    'End Sub
End Class
