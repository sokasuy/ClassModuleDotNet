Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports ClosedXML.Excel
Imports System.IO
Imports System.Data
'Imports CShowMessage.CShowMessage

Public Class CFileIO
    Public Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (lpApplicationName As String, lpKeyName As String, lpDefault As String, lpReturnedString As String, nSize As Int32, lpFileName As String) As Int32
    Private myCShowMessage As New CShowMessage.CShowMessage

    Public Function ReadIniFile(_iniFileSection As String, _settingKeyWord As String, _file As String) As String
        'OK
        'untuk membaca File .INI yang biasanya digunakan untuk menyimpan berbagai settingan
        Try
            Dim ParamVal As String
            Dim LenParamVal As Long

            'Dim a As Integer
            ParamVal = Space$(1024)
            LenParamVal = GetPrivateProfileString(_iniFileSection, _settingKeyWord, "(error)", ParamVal, Len(ParamVal), _file)

            ReadIniFile = Left$(ParamVal, LenParamVal)
        Catch ex As Exception
            ReadIniFile = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReadIniFile Error")
        End Try
    End Function

    Public Function ReadIniAndCheckDefinedDirOrFileExistance(_iniFileSection As String, _settingKeyWord As String, _iniFilePathFileName As String, _definedFileDir As String) As Boolean
        'GAK DIPAKAI
        'Public Function ReadIniAndCheckDirOrFileExistance(INIFileSection As String, SettingKeyWord As String, INIFilePathFileName As String, ReadingResult As String) As Boolean
        'belum di tes
        'contoh:
        'If (ReadIniAndCheckDirOrFileExistance("EXTFILE", "DtSPath", FilePathFileName, StartUpDir) = False) Then
        'StartUpDir = App.Path                                   'jika belum diset di INI file, atau directory tsb tidak ada, pakai lokasi program
        'End If
        Try
            Dim stReadingResult As String
            'Dim stMsg As String

            stReadingResult = ReadIniFile(_iniFileSection, _settingKeyWord, _iniFilePathFileName)
            If (stReadingResult <> "") Then 'jika setting di INI file ada (setelah keyword= ada isinya)
                'cek keberadaan directory / file yg didefinisikan di INI file tsb
                If Dir(stReadingResult) <> "" Then
                    'pakai dir tsb. ok
                    _definedFileDir = stReadingResult
                    ReadIniAndCheckDefinedDirOrFileExistance = True
                Else
                    'jika file tsb tidak ada, beri msg utk keperluan koreksi INI file. ok
                    Call myCShowMessage.ShowErrMsg("Directory atau File : '" & stReadingResult & "' Tidak Ditemukan...", "ReadIniAndCheckDirOrFileExistance Error")
                    _definedFileDir = "" 'returned value harus empty, utk menghindari kesalahan
                    ReadIniAndCheckDefinedDirOrFileExistance = False
                End If
            Else
                'jika blm diset di INI file...Keyword sdh ditulis tapi sisi kanan '=' belum diisi...atau Section dan / atau Keyword belum didefinisikan. ok
                ReadIniAndCheckDefinedDirOrFileExistance = False
            End If
            stReadingResult = Nothing
        Catch ex As Exception
            _definedFileDir = ""
            ReadIniAndCheckDefinedDirOrFileExistance = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReadIniAndCheckDefinedDirOrFileExistance Error")
        End Try
    End Function

    Public Function ReadIniValue(_iniFileSection As String, _settingKeyWord As String, _iniPath As String) As String
        'OK  -->  GAK DIPAKAI
        'Fungsi sama dengan ReadIniFile hanya saja di sini file .ini diberlakukan seperti file biasa
        Try
            Dim NF As Short
            Dim Temp As String
            Dim LcaseTemp As String
            Dim ReadyToRead As Boolean

            'AssignVariables:
            NF = FreeFile()
            ReadIniValue = ""
            _iniFileSection = "[" & LCase(_iniFileSection) & "]"
            _settingKeyWord = LCase(_settingKeyWord)

            'EnsureFileExists:
            FileOpen(NF, _iniPath, Microsoft.VisualBasic.OpenMode.Binary)
            FileClose(NF)
            SetAttr(_iniPath, FileAttribute.Archive)

            'LoadFile:
            FileOpen(NF, _iniPath, Microsoft.VisualBasic.OpenMode.Input)
            While Not EOF(NF)
                Temp = LineInput(NF)
                LcaseTemp = LCase(Temp)
                If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
                If LcaseTemp = _iniFileSection Then ReadyToRead = True
                If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
                    If InStr(LcaseTemp, _settingKeyWord & "=") = 1 Then
                        ReadIniValue = Mid(Temp, 1 + Len(_settingKeyWord & "="))
                        FileClose(NF) : Exit Function
                    End If
                End If
            End While
            FileClose(NF)
        Catch ex As Exception
            ReadIniValue = Nothing
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReadIniValue Error")
        End Try
    End Function

    'Public Sub ExcelCleanUp(_oXL As Excel.Application, _oWB As Excel.Workbook, _oSheet As Excel.Worksheet)
    '    Try
    '        'untuk membersihkan sisa2 sampah dari pembuatan file excel, serta menutup proses dari excel
    '        'kalau fungsi ini tidak ada, maka setiap kali meng-generate excel, maka akan menambah 1 proses excel
    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()

    '        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oXL)
    '        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oSheet)
    '        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oWB)

    '        _oSheet = Nothing
    '        _oWB = Nothing
    '        _oXL = Nothing
    '    Catch ex As Exception
    '        Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExcelCleanUp Error")
    '    End Try
    'End Sub

    Public Sub ExcelCleanUp(_oXL As Excel.Application, _oWB As Excel.Workbook, _oSheet As Excel.Worksheet)
        Try
            ' Tutup workbook dan Excel dengan aman
            If _oWB IsNot Nothing Then
                Try
                    _oWB.Close(SaveChanges:=False)
                Catch
                End Try
            End If

            If _oXL IsNot Nothing Then
                Try
                    _oXL.Quit()
                Catch
                End Try
            End If

            ' Rilis COM object dari dalam ke luar
            If _oSheet IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oSheet)
            End If
            If _oWB IsNot Nothing Then
                '_oWB.Close(False)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oWB)
            End If
            If _oXL IsNot Nothing Then
                '_oXL.Quit()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oXL)
            End If

            ' Bersihkan referensi
            _oSheet = Nothing
            _oWB = Nothing
            _oXL = Nothing

            ' Paksa garbage collector untuk memastikan semua dilepas
            GC.Collect()
            GC.WaitForPendingFinalizers()
            'GC.Collect()
            'GC.WaitForPendingFinalizers()

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExcelCleanUp Error")
        End Try
    End Sub

    Function SheetExists(_sheetToFind As String, _file As String) As Boolean
        Dim oXL As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        'Dim sheet_found As Boolean
        Try
            'Dim SheetNameToCheck As String = "Sheet22"
            '~~> Opens an exisiting Workbook. Change path and filename as applicable
            oWB = oXL.Workbooks.Open(_file)
            '~~> Display Excel
            oXL.Visible = False
            '~~> Loop through the all the sheets in the workbook to find if name matches
            For Each oSheet In oWB.Sheets
                If oSheet.Name = _sheetToFind Then
                    SheetExists = True
                    Exit For
                Else
                    SheetExists = False
                End If
            Next
            oXL.Quit()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SheetExists Error")
            'harus ditaruh di finally, karena kalau ditaruh di try, kalau misal export excel gagal mungkin gara2 path yang dituju tidak ada
            'maka fungsi ini akan tetap dijalankan
            ExcelCleanUp(oXL, oWB, oSheet)
            SheetExists = False
        End Try
    End Function

    'Public Sub PopulateSheet(_dt As DataTable, _file As String, _sheetName As String, Optional _textColumn As Boolean = False, Optional _textColIndex As Byte = 0)
    '    Dim oXL As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
    '    Dim oWB As Excel.Workbook = Nothing
    '    Dim oSheet As Excel.Worksheet = Nothing
    '    Dim oRng As Excel.Range = Nothing

    '    Try
    '        Dim oRng As Excel.Range
    '        oXL.Visible = False

    '        oWB = oXL.Workbooks.Add
    '        oSheet = CType(oWB.ActiveSheet, Excel.Worksheet)

    '        Dim dc As DataColumn
    '        Dim dr As DataRow
    '        Dim colIndex As Integer = 0
    '        Dim rowIndex As Integer = 0
    '        For Each dc In _dt.Columns
    '            colIndex = colIndex + 1
    '            oXL.Cells(1, colIndex) = dc.ColumnName
    '            If (_textColumn And _textColIndex = colIndex) Then
    '                oXL.Columns(_textColIndex).NumberFormat = "@"
    '            End If
    '        Next
    '        For Each dr In _dt.Rows
    '            rowIndex = rowIndex + 1
    '            colIndex = 0
    '            For Each dc In _dt.Columns
    '                colIndex = colIndex + 1
    '                oXL.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
    '            Next
    '        Next

    '        oSheet.Name = _sheetName
    '        'oSheet.Cells.Select()
    '        'oSheet.Columns.AutoFit()
    '        'oSheet.Rows.AutoFit()
    '        Dim usedRange = oSheet.UsedRange
    '        usedRange.Columns.AutoFit()
    '        usedRange.Rows.AutoFit()

    '        'oXL.Visible = False
    '        'oXL.UserControl = True

    '        oWB.SaveAs(_file)
    '        oRng = Nothing
    '        oXL.Quit()

    '        'ExcelCleanUp(oXL, oWB, oSheet)
    '    Catch ex As Exception
    '        Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "PopulateSheet Error")
    '    Finally
    '        'harus ditaruh di finally, karena kalau ditaruh di try, kalau misal export excel gagal mungkin gara2 path yang dituju tidak ada
    '        'maka fungsi ini akan tetap dijalankan
    '        ExcelCleanUp(oXL, oWB, oSheet)
    '    End Try
    'End Sub

    Public Sub PopulateSheet(_dt As DataTable, _file As String, _sheetName As String, Optional _textColumn As Boolean = False, Optional _textColIndex As Byte = 0)

        Dim oXL As Excel.Application = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim oRng As Excel.Range = Nothing

        Try
            oXL = CType(CreateObject("Excel.Application"), Excel.Application)
            oXL.Visible = False

            oWB = oXL.Workbooks.Add()
            oSheet = CType(oWB.ActiveSheet, Excel.Worksheet)

            '--- Header
            Dim colIndex As Integer = 0
            For Each dc As DataColumn In _dt.Columns
                colIndex += 1
                oSheet.Cells(1, colIndex).Value = dc.ColumnName

                If _textColumn AndAlso _textColIndex = colIndex Then
                    CType(oSheet.Columns(colIndex), Excel.Range).NumberFormat = "@"
                End If
            Next

            '--- Data rows
            Dim rowIndex As Integer = 1
            For Each dr As DataRow In _dt.Rows
                rowIndex += 1
                colIndex = 0

                For Each dc As DataColumn In _dt.Columns
                    colIndex += 1
                    oSheet.Cells(rowIndex, colIndex).Value = dr(dc.ColumnName)
                Next
            Next

            oSheet.Name = _sheetName

            'AutoFit only used range
            oRng = oSheet.UsedRange
            oRng.Columns.AutoFit()
            oRng.Rows.AutoFit()

            'Save file
            oWB.SaveAs(_file)

        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "PopulateSheet Error")
        Finally
            ' Release COM object from inner → outer
            If oRng IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oRng)
            If oSheet IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheet)

            If oWB IsNot Nothing Then
                oWB.Close(False)
                Runtime.InteropServices.Marshal.FinalReleaseComObject(oWB)
            End If

            If oXL IsNot Nothing Then
                oXL.Quit()
                Runtime.InteropServices.Marshal.FinalReleaseComObject(oXL)
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub PopulateSheetFast(_dt As DataTable, _filePath As String, _sheetName As String, Optional _textColumn As Boolean = False, Optional _textColIndex As Integer = 0)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWB As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlRange As Excel.Range = Nothing

        Try
            'Buat Excel instance
            xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
            xlApp.Visible = False

            xlWB = xlApp.Workbooks.Add()
            xlSheet = CType(xlWB.ActiveSheet, Excel.Worksheet)
            xlSheet.Name = _sheetName

            Dim rows As Integer = _dt.Rows.Count
            Dim cols As Integer = _dt.Columns.Count

            '===========================================
            ' 1) Build Array Data (Header + Rows)
            '===========================================
            Dim arr(rows, cols - 1) As Object  'rows+1 untuk header

            'Header
            For c As Integer = 0 To cols - 1
                arr(0, c) = _dt.Columns(c).ColumnName
            Next

            'Rows
            For r As Integer = 0 To rows - 1
                For c As Integer = 0 To cols - 1
                    arr(r + 1, c) = _dt.Rows(r)(c)
                Next
            Next

            '===========================================
            ' 2) Tentukan Target Range
            '===========================================
            Dim lastCol As String = ColumnIndexToLetter(cols)
            Dim lastRow As Integer = rows + 1

            xlRange = xlSheet.Range("A1", lastCol & lastRow.ToString())

            '===========================================
            ' 3) Dump Array ke Excel (super cepat!)
            '===========================================
            xlRange.Value = arr

            '===========================================
            ' 4) Formatting
            '===========================================
            If _textColumn AndAlso _textColIndex > 0 Then
                CType(xlSheet.Columns(_textColIndex), Excel.Range).NumberFormat = "@"
            End If

            'Autofit hanya used range
            xlSheet.UsedRange.Columns.AutoFit()

            '===========================================
            ' 5) Save
            '===========================================
            xlWB.SaveAs(_filePath)

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "PopulateSheetFast")
        Finally
            'Release COM (inner → outer)
            If xlRange IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(xlRange)
            If xlSheet IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(xlSheet)

            If xlWB IsNot Nothing Then
                xlWB.Close(False)
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWB)
            End If

            If xlApp IsNot Nothing Then
                xlApp.Quit()
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp)
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    'Convert nomor kolom → ABCD...
    Private Function ColumnIndexToLetter(_index As Integer) As String
        Try
            Dim dividend As Integer = _index
            Dim columnName As String = ""
            Dim modulo As Integer

            While dividend > 0
                modulo = (dividend - 1) Mod 26
                columnName = Chr(65 + modulo) & columnName
                dividend = (dividend - modulo) \ 26
            End While

            Return columnName
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ColumnIndexToLetter")
            Return Nothing
        End Try
    End Function

    Public Sub ExportToExcelMultipleSheets(_tables As Dictionary(Of String, DataTable), _filePath As String)
        Dim oXL As Excel.Application = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing

        Try
            oXL = CType(CreateObject("Excel.Application"), Excel.Application)
            oXL.Visible = False
            oWB = oXL.Workbooks.Add()

            Dim sheetIndex As Integer = 1

            For Each kvp As KeyValuePair(Of String, DataTable) In _tables
                Dim sheetName As String = kvp.Key
                Dim dt As DataTable = kvp.Value

                ' Tambahkan sheet baru kalau bukan sheet pertama
                If sheetIndex = 1 Then
                    oSheet = CType(oWB.Sheets(1), Excel.Worksheet)
                Else
                    oSheet = CType(oWB.Sheets.Add(After:=oWB.Sheets(oWB.Sheets.Count)), Excel.Worksheet)
                End If

                ' Set nama sheet
                If sheetName.Length > 31 Then
                    sheetName = sheetName.Substring(0, 31) 'batas maksimal nama sheet Excel
                End If
                oSheet.Name = sheetName

                ' Tulis header
                Dim colIndex As Integer = 0
                For Each dc As DataColumn In dt.Columns
                    colIndex += 1
                    oSheet.Cells(1, colIndex) = dc.ColumnName
                Next

                ' Tulis data
                Dim rowIndex As Integer = 0
                For Each dr As DataRow In dt.Rows
                    rowIndex += 1
                    colIndex = 0
                    For Each dc As DataColumn In dt.Columns
                        colIndex += 1
                        oSheet.Cells(rowIndex + 1, colIndex) = dr(dc)
                    Next
                Next

                ' Rapiin tampilan
                oSheet.Columns.AutoFit()
                sheetIndex += 1
            Next

            ' Simpan file
            oWB.SaveAs(_filePath)
            oWB.Close(SaveChanges:=False)
            oXL.Quit()

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExportToExcelMultipleSheets Error")
        Finally
            ExcelCleanUp(oXL, oWB, oSheet)
        End Try
    End Sub

    Public Sub ExportToExcelMultipleSheetsFast(ByVal _filePath As String, ByVal _tables As Dictionary(Of String, DataTable),
        ByVal _columnsToFix As List(Of String)   ' ← daftar kolom yang perlu leading zero
    )
        Dim oXL As Excel.Application = Nothing
        Dim oWB As Excel.Workbook = Nothing

        Try
            oXL = New Excel.Application()
            oXL.Visible = False
            oWB = oXL.Workbooks.Add()

            Dim firstSheet As Boolean = True

            For Each kvp In _tables
                Dim sheetName = kvp.Key
                Dim dt = kvp.Value

                Dim oSheet As Excel.Worksheet

                If firstSheet Then
                    oSheet = CType(oWB.ActiveSheet, Excel.Worksheet)
                    firstSheet = False
                Else
                    oSheet = CType(oWB.Worksheets.Add(), Excel.Worksheet)
                End If

                oSheet.Name = sheetName

                ' =========================================
                ' 1) AUTO FIX LEADING ZERO
                ' =========================================
                If _columnsToFix IsNot Nothing AndAlso _columnsToFix.Count > 0 Then
                    For Each colFix In _columnsToFix
                        If dt.Columns.Contains(colFix) Then
                            For Each r As DataRow In dt.Rows
                                Dim raw = r(colFix).ToString().Trim()

                                ' Jika angka tapi butuh leading zero
                                If IsNumeric(raw) AndAlso raw.StartsWith("0") = False Then
                                    ' tambahkan leading zero di depan sesuai panjang aslinya
                                    ' Jika kamu punya aturan lain, bisa disesuaikan
                                    r(colFix) = "'" & raw
                                ElseIf Not IsNumeric(raw) AndAlso raw <> "" Then
                                    ' Kalau ada data campuran yang sebenarnya numeric
                                    If raw.All(AddressOf Char.IsDigit) Then
                                        r(colFix) = "'" & raw
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If

                ' =========================================
                ' 2) TULIS HEADER
                ' =========================================
                For col = 0 To dt.Columns.Count - 1
                    oSheet.Cells(1, col + 1) = dt.Columns(col).ColumnName
                Next

                ' =========================================
                ' 3) TULIS DATA
                ' =========================================
                Dim arr(dt.Rows.Count - 1, dt.Columns.Count - 1) As Object
                For r = 0 To dt.Rows.Count - 1
                    For c = 0 To dt.Columns.Count - 1
                        arr(r, c) = dt.Rows(r)(c)
                    Next
                Next

                Dim startCell As Excel.Range = oSheet.Range("A2")
                Dim endCell As Excel.Range = oSheet.Cells(dt.Rows.Count + 1, dt.Columns.Count)
                Dim writeRange As Excel.Range = oSheet.Range(startCell, endCell)

                writeRange.Value2 = arr

                ' Autofit
                oSheet.Columns.AutoFit()
            Next

            oWB.SaveAs(_filePath)
            oWB.Close()
            oXL.Quit()

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExportToExcelMultipleSheetsFast Error")
        Finally
            ' CLEANUP — Penting!!
            Try
                If oWB IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oWB)
                If oXL IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oXL)
            Catch
                Call myCShowMessage.ShowErrMsg("Pesan Error: Release Error", "ExportToExcelMultipleSheetsFast Error")
            End Try

            oWB = Nothing
            oXL = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub


    Public Sub ExportDataGridViewToExcel(_dataGridView As DataGridView, _file As String, _sheetName As String)
        Dim oXL As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Try
            oXL.Visible = False

            oWB = oXL.Workbooks.Add
            oSheet = CType(oWB.ActiveSheet, Excel.Worksheet)
            'oSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection

            'Export Header Names Start
            Dim columnsCount As Integer = _dataGridView.Columns.Count
            For Each column In _dataGridView.Columns
                oSheet.Cells(1, column.Index + 1).Value = column.Name
            Next
            'Export Header Name End


            'Export Each Row Start
            For i As Integer = 0 To _dataGridView.Rows.Count - 1
                Dim columnIndex As Integer = 0
                Do Until columnIndex = columnsCount
                    oSheet.Cells(i + 2, columnIndex + 1).Value = _dataGridView.Item(columnIndex, i).Value.ToString
                    columnIndex += 1
                Loop
            Next
            'Export Each Row End

            oSheet.Name = _sheetName
            oSheet.Cells.Select()
            oSheet.Columns.AutoFit()
            oSheet.Rows.AutoFit()

            oXL.Visible = False
            oXL.UserControl = True

            oWB.SaveAs(_file)
            oXL.Quit()

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExportDataGridViewToExcel Error")
        Finally
            'harus ditaruh di finally, karena kalau ditaruh di try, kalau misal export excel gagal mungkin gara2 path yang dituju tidak ada
            'maka fungsi ini akan tetap dijalankan
            ExcelCleanUp(oXL, oWB, oSheet)
        End Try
    End Sub

    Public Sub ExportDataGridViewToExcelFast(_dataGridView As DataGridView, _file As String, _sheetName As String)
        Try
            ' Pastikan file punya ekstensi .xlsx
            If Not _file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then
                _file &= ".xlsx"
            End If

            Using wb As New XLWorkbook()
                Dim ws = wb.Worksheets.Add(_sheetName)

                ' === Export Header ===
                For col As Integer = 0 To _dataGridView.Columns.Count - 1
                    ws.Cell(1, col + 1).Value = _dataGridView.Columns(col).HeaderText
                Next

                ' === Export Data ===
                For row As Integer = 0 To _dataGridView.Rows.Count - 1
                    For col As Integer = 0 To _dataGridView.Columns.Count - 1
                        Dim cellValue = _dataGridView.Rows(row).Cells(col).Value
                        If cellValue IsNot Nothing Then
                            ws.Cell(row + 2, col + 1).Value = cellValue
                        Else
                            ws.Cell(row + 2, col + 1).Value = ""
                        End If
                    Next
                Next

                ' Autofit kolom
                ws.Columns().AdjustToContents()

                wb.SaveAs(_file)
            End Using
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ExportDataGridViewToExcelFast Error")
        End Try
    End Sub

    Public Function ImportExcelToDataTable(_filePath As String, Optional _sheetName As String = Nothing, Optional _kolomFilter() As String = Nothing, Optional _skipBarisKosongTotal As Boolean = True) As DataTable
        Dim dt As New DataTable()
        Try
            Using workbook As New XLWorkbook(_filePath)
                Dim ws As IXLWorksheet

                ' Ambil sheet
                If String.IsNullOrEmpty(_sheetName) Then
                    ws = workbook.Worksheets.First()
                Else
                    ws = workbook.Worksheet(_sheetName)
                End If

                If ws Is Nothing Then
                    Throw New Exception("Worksheet tidak ditemukan.")
                End If

                ' Ambil range data
                Dim range = ws.RangeUsed()
                Dim firstRow = range.FirstRowUsed().RowNumber()
                Dim lastRow = range.LastRowUsed().RowNumber()
                Dim firstCol = range.FirstColumnUsed().ColumnNumber()
                Dim lastCol = range.LastColumnUsed().ColumnNumber()

                ' Tambahkan kolom dari header
                For col As Integer = firstCol To lastCol
                    Dim colName As String = ws.Cell(firstRow, col).GetString().Trim()
                    If String.IsNullOrEmpty(colName) Then
                        colName = "Column" & col
                    End If
                    dt.Columns.Add(colName)
                Next

                ' Tambahkan data mulai baris ke-2
                For row As Integer = firstRow + 1 To lastRow
                    Dim dr As DataRow = dt.NewRow()
                    For col As Integer = firstCol To lastCol
                        dr(col - firstCol) = ws.Cell(row, col).GetString().Trim()
                    Next

                    ' --- 1️ Cek baris kosong total (semua kolom kosong)
                    If _skipBarisKosongTotal Then
                        Dim semuaKosong = True
                        For Each value In dr.ItemArray
                            If Not String.IsNullOrEmpty(value.ToString()) Then
                                semuaKosong = False
                                Exit For
                            End If
                        Next
                        If semuaKosong Then Continue For
                    End If

                    ' --- 2️ Cek kolom filter (kalau ada)
                    If _kolomFilter IsNot Nothing AndAlso _kolomFilter.Length > 0 Then
                        Dim semuaFilterKosong As Boolean = True
                        For Each k In _kolomFilter
                            If dt.Columns.Contains(k) Then
                                If Not String.IsNullOrEmpty(dr(k).ToString()) Then
                                    semuaFilterKosong = False
                                    Exit For
                                End If
                            End If
                        Next
                        If semuaFilterKosong Then Continue For
                    End If

                    ' Tambahkan baris jika lolos semua filter
                    dt.Rows.Add(dr)

                    'If (row Mod 100 = 0) Then
                    '    GC.Collect()
                    'End If
                Next
            End Using

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ImportExcelToDataTable Error")
        End Try

        Return dt
    End Function


    Public Function CreateDirectory(_destinationPath As String) As Boolean
        Try
            System.IO.Directory.CreateDirectory(_destinationPath)
            CreateDirectory = True
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CreateDirectory Error")
            CreateDirectory = False
        End Try
    End Function

    Public Function CopyFile(_sourcePath As String, _destinationPath As String, Optional _overwrite As Boolean = False) As Boolean
        Try
            If (System.IO.File.Exists(_destinationPath)) Then
                'kalau file dengan nama yang sama ditemukan, maka tidak bisa copy dan tampilkan msgbox
                'MessageBox.Show("Peringatan sudah ada file dengan nama ini", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Dim tumpuk As Integer
                'tumpuk = MessageBox.Show("Menimpa file database yang sudah ada??", "Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                tumpuk = myCShowMessage.GetUserResponse("Menimpa file yang sudah ada??", "Overwrite")
                If (tumpuk = 6) Then
                    'KALAU IYA, maka akan ditumpuk / overwrite
                    System.IO.File.Copy(_sourcePath, _destinationPath, True)
                    'FileCopy(_sourcePath, _destinationPath)
                    CopyFile = True
                ElseIf (tumpuk = 7) Then
                    'KALAU TIDAK, maka tidak jadi di copy
                    Call myCShowMessage.ShowInfo("Operasi copy dibatalkan", "Proses copy batal")
                    CopyFile = False
                Else
                    CopyFile = False
                End If
            Else
                System.IO.File.Copy(_sourcePath, _destinationPath)
                'FileCopy(_sourcePath, _destinationPath)
                CopyFile = True
            End If
        Catch ex As Exception
            CopyFile = False
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CopyFile Error")
        End Try
    End Function

    Public Function DeleteFile(_sourcePath As String) As Boolean
        Try
            If (System.IO.File.Exists(_sourcePath)) Then
                System.IO.File.Delete(_sourcePath)
                'FileIO.FileSystem.DeleteFile(_sourcePath)
                DeleteFile = True
            Else
                Call myCShowMessage.ShowWarning("Delete gagal " & _sourcePath & " tidak ditemukan!!", "Delete Gagal")
                DeleteFile = False
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "DeleteFile Error")
            DeleteFile = False
        End Try
    End Function

    Public Function MoveFile(_sourcePath As String, _destinationPath As String) As Boolean
        Try
            If (System.IO.File.Exists(_sourcePath)) Then
                System.IO.File.Move(_sourcePath, _destinationPath)
                MoveFile = True
            Else
                Call myCShowMessage.ShowWarning("Move File gagal " & _sourcePath & " tidak ditemukan!!", "Move File Gagal")
                MoveFile = False
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "MoveFile Error")
            MoveFile = False
        End Try
    End Function

    Public Function RenameFile(_sourcePath As String, _destinationPath As String) As Boolean
        Try
            If (System.IO.File.Exists(_sourcePath)) Then
                Rename(_sourcePath, _destinationPath)
                RenameFile = True
            Else
                Call myCShowMessage.ShowWarning("Rename File gagal " & _sourcePath & " tidak ditemukan atau file di " & _destinationPath & " sudah ada!!", "Rename File Gagal")
                RenameFile = False
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RenameFile Error")
            RenameFile = False
        End Try
    End Function

    Public Function SetFilterForAttachment(Optional _docTypes As String = "all") As String
        Try
            If (_docTypes.ToLower = "images") Then
                'Berarti gambar saja
                SetFilterForAttachment = "Image file (*.jpeg,*.jpg,*.png,*.bmp,*.tiff)|*.jpeg;*.jpg;*.png;*.bmp;*.tiff"
            ElseIf (_docTypes.ToLower = "documents") Then
                'Berarti dokumen saja
                SetFilterForAttachment = "All Supported File|*.xls;*.xlsx;*.docx;*.doc;*.pptx;*.ppt;*.pdf;*.txt|Excel file (*.xls,*.xlsx)|*.xls;*.xlsx|Word file (*.doc,*.docx)|*.doc;*.docx|Power Point file (*.ppt,*.pptx)|*.ppt;*.pptx|PDF file (*.pdf)|*.pdf|Normal text file (*.txt)|*.txt"
            ElseIf (_docTypes.ToLower = "all") Then
                'Semuanya yang memungkinkan
                SetFilterForAttachment = "All Supported File|*.xls;*.xlsx;*.docx;*.doc;*.pptx;*.ppt;*.pdf;*.txt;*.jpeg;*.jpg;*.png;*.bmp;*.tiff|Excel file (*.xls,*.xlsx)|*.xls;*.xlsx|Word file (*.doc,*.docx)|*.doc;*.docx|Power Point file (*.ppt,*.pptx)|*.ppt;*.pptx|PDF file (*.pdf)|*.pdf|Normal text file (*.txt)|*.txt|Image file (*.jpeg,*.jpg,*.png,*.bmp,*.tiff)|*.jpeg;*.jpg;*.png;*.bmp;*.tiff"
            Else
                'untuk spesifik
                'Nanti saja
                SetFilterForAttachment = Nothing
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetFilterForAttachment Error")
            SetFilterForAttachment = Nothing
        End Try
    End Function

    Public Function SetProcessInfo(_processInfo As ProcessStartInfo, _fileExtension As String, _namaFile As String) As ProcessStartInfo
        Try
            'With _processInfo
            '    .Arguments = _namaFile
            '    .UseShellExecute = True
            '    '.Verb = "open"

            '    If (_fileExtension = ".txt") Then
            '        .FileName = "notepad"
            '    ElseIf (_fileExtension = ".pdf") Then
            '        'MENGGUNAKAN MS EDGE
            '        'processInfo.FileName = "msedge"
            '        'MENGGUNAKAN ADOBE ACROBAT
            '        .FileName = "acrobat"
            '    ElseIf (_fileExtension = ".xls" Or _fileExtension = ".xlsx") Then
            '        .FileName = "excel"
            '    ElseIf (_fileExtension = ".doc" Or _fileExtension = ".docx") Then
            '        .FileName = "winword"
            '    ElseIf (_fileExtension = ".jpeg" Or _fileExtension = ".jpg" Or _fileExtension = ".png") Then
            '        '.FileName = "winword"
            '    End If
            'End With
            'DIAPIT PETIK DUA (") AGAR NAMA FILE MAUPUN PATH NYA BISA ADA SPASINYA
            _processInfo = New ProcessStartInfo("""" & _namaFile & """")
            _processInfo.UseShellExecute = True
            SetProcessInfo = _processInfo
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetProcessInfo Error")
            SetProcessInfo = Nothing
        End Try
    End Function
End Class
