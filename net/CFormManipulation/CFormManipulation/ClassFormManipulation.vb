Imports System.Windows.Forms
Public Class CFormManipulation
    Private myCShowMessage As New CShowMessage.CShowMessage

    Public Sub GoToForm(ByRef _myForm_parent As Form, ByRef _myForm_child As Form, Optional _isAnakMdi As Boolean = True, Optional _enabledParent As Boolean = False)
        'OK
        'digunakan untuk pindah dari form induk ke form anak
        Try
            _myForm_child.Owner = _myForm_parent
            If (_isAnakMdi) Then
                'kalau mdi child
                _myForm_child.MdiParent = _myForm_parent
            Else
                'kalau bukan mdi child
                _myForm_parent.Enabled = _enabledParent
            End If
            _myForm_child.Show()
            _myForm_child.BringToFront()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GoToForm Error")
        End Try
    End Sub

    Public Sub CloseMyForm(ByRef _myForm_parent As Form, ByRef _myForm_child As Form, Optional _hookFormClosed As Boolean = True, Optional _isAnakMdi As Boolean = True)
        'OK
        'digunakan untuk pindah dari form anak kembali ke form induk
        Try
            _myForm_parent = _myForm_child.Owner
            If Not (_isAnakMdi) Then
                _myForm_parent.Enabled = True
                _myForm_parent.BringToFront()
            End If
            If (_hookFormClosed = False) Then
                'kalau fungsi CloseMyForm dipanggil bukan pada saat event FormClosed
                'karena kalau pada saat event FormClosed, form yang bersangkutan sudah dalam kondisi closed
                'sehingga akan error kalau berusaha di close lagi
                _myForm_child.Close()
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CloseMyForm Error")
        End Try
    End Sub

    Public Sub ResetForm(_root As Control)
        Try
            For Each ctrl As Control In _root.Controls
                ResetForm(ctrl)

                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Clear()
                ElseIf TypeOf ctrl Is MaskedTextBox Then
                    CType(ctrl, MaskedTextBox).Clear()
                ElseIf TypeOf ctrl Is ComboBox Then
                    CType(ctrl, ComboBox).SelectedIndex = -1
                ElseIf TypeOf ctrl Is CheckBox Then
                    CType(ctrl, CheckBox).Checked = False
                ElseIf TypeOf ctrl Is RichTextBox Then
                    CType(ctrl, RichTextBox).Clear()
                ElseIf TypeOf ctrl Is DataGridView Then
                    CType(ctrl, DataGridView).Rows.Clear()
                ElseIf TypeOf ctrl Is DateTimePicker Then
                    CType(ctrl, DateTimePicker).ResetText()
                ElseIf TypeOf ctrl Is NumericUpDown Then
                    CType(ctrl, NumericUpDown).ResetText()
                End If
            Next ctrl
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ResetForm Error")
        End Try
    End Sub

    Public Sub SetReadOnly(_root As Control, _isReadOnly As Boolean)
        Try
            For Each ctrl As Control In _root.Controls
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).ReadOnly = _isReadOnly
                ElseIf TypeOf ctrl Is MaskedTextBox Then
                    CType(ctrl, MaskedTextBox).ReadOnly = _isReadOnly
                ElseIf TypeOf ctrl Is ComboBox Then
                    CType(ctrl, ComboBox).Enabled = IIf(_isReadOnly, False, True)
                ElseIf TypeOf ctrl Is CheckBox Then
                    CType(ctrl, CheckBox).Enabled = IIf(_isReadOnly, False, True)
                ElseIf TypeOf ctrl Is RichTextBox Then
                    CType(ctrl, RichTextBox).ReadOnly = _isReadOnly
                ElseIf TypeOf ctrl Is DataGridView Then
                    CType(ctrl, DataGridView).ReadOnly = _isReadOnly
                ElseIf TypeOf ctrl Is DateTimePicker Then
                    CType(ctrl, DateTimePicker).Enabled = IIf(_isReadOnly, False, True)
                ElseIf TypeOf ctrl Is NumericUpDown Then
                    CType(ctrl, NumericUpDown).ReadOnly = _isReadOnly
                End If
            Next ctrl
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetReadOnly Error")
        End Try
    End Sub

    Public Sub SetCheckListBoxUserRights(ByRef _clbUserRight As CheckedListBox, _isSuperuser As Boolean, _formName As String, _dtTableUserRight As DataTable)
        Try
            If Not _isSuperuser Then
                'Hanya di cek kalau bukan superuser
                'Superuser bisa segalanya
                Dim foundRows As DataRow()
                foundRows = _dtTableUserRight.Select("formname='" & _formName & "'")
                For i As Integer = 0 To _clbUserRight.Items.Count - 1
                    _clbUserRight.SetItemChecked(i, _dtTableUserRight.Rows(_dtTableUserRight.Rows.IndexOf(foundRows(0))).Item(_clbUserRight.Items(i)))
                Next
            Else
                For i As Integer = 0 To _clbUserRight.Items.Count - 1
                    _clbUserRight.SetItemChecked(i, True)
                Next
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetCheckListBoxUserRights Error")
        End Try
    End Sub

    Public Sub SetButtonSimpanAvailabilty(ByRef _button As Button, _clbUserRight As CheckedListBox, action As String)
        Try
            Dim idx As Integer

            'Hanya index 1 dan 2 saja yang bisa melakukan save
            Select Case action
                Case "load", "save"
                    idx = _clbUserRight.Items.IndexOf("Menambah")
                Case "edit"
                    idx = _clbUserRight.Items.IndexOf("Memperbaharui")
            End Select

            If (_clbUserRight.GetItemChecked(idx)) Then
                _button.Enabled = True
            Else
                _button.Enabled = False
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetButtonSimpanAvailabilty Error")
        End Try
    End Sub
End Class
