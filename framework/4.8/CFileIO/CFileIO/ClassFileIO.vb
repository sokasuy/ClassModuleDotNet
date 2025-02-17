Imports Microsoft.Office.Interop
Imports System.Windows.Forms
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

    Public Sub ExcelCleanUp(_oXL As Excel.Application, _oWB As Excel.Workbook, _oSheet As Excel.Worksheet)
        Try
            'untuk membersihkan sisa2 sampah dari pembuatan file excel, serta menutup proses dari excel
            'kalau fungsi ini tidak ada, maka setiap kali meng-generate excel, maka akan menambah 1 proses excel
            GC.Collect()
            GC.WaitForPendingFinalizers()

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oXL)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oSheet)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_oWB)

            _oSheet = Nothing
            _oWB = Nothing
            _oXL = Nothing
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

    Public Sub PopulateSheet(_dt As DataTable, _file As String, _sheetName As String, Optional _textColumn As Boolean = False, Optional _textColIndex As Byte = 0)
        Dim oXL As Excel.Application = CType(CreateObject("Excel.Application"), Excel.Application)
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        Try
            Dim oRng As Excel.Range
            oXL.Visible = False

            oWB = oXL.Workbooks.Add
            oSheet = CType(oWB.ActiveSheet, Excel.Worksheet)

            Dim dc As DataColumn
            Dim dr As DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0
            For Each dc In _dt.Columns
                colIndex = colIndex + 1
                oXL.Cells(1, colIndex) = dc.ColumnName
                If (_textColumn And _textColIndex = colIndex) Then
                    oXL.Columns(_textColIndex).NumberFormat = "@"
                End If
            Next
            For Each dr In _dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In _dt.Columns
                    colIndex = colIndex + 1
                    oXL.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next

            oSheet.Name = _sheetName
            oSheet.Cells.Select()
            oSheet.Columns.AutoFit()
            oSheet.Rows.AutoFit()

            oXL.Visible = False
            oXL.UserControl = True

            oWB.SaveAs(_file)
            oRng = Nothing
            oXL.Quit()

            'ExcelCleanUp(oXL, oWB, oSheet)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "PopulateSheet Error")
        Finally
            'harus ditaruh di finally, karena kalau ditaruh di try, kalau misal export excel gagal mungkin gara2 path yang dituju tidak ada
            'maka fungsi ini akan tetap dijalankan
            ExcelCleanUp(oXL, oWB, oSheet)
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
