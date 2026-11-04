Imports System.Windows.Forms
Imports System.Diagnostics

Public Class CViewManager
    Private Class ComboBoxBindingInfo
        Public Property Combo As ComboBox
        Public Property DisplayMember As String
        Public Property ValueMember As String
    End Class

    Private myCShowMessage As New CShowMessage.CShowMessage
    Private dataViews As New Dictionary(Of String, DataView)
    'Private comboBoxes As New Dictionary(Of String, ComboBox)
    Private comboBoxes As New Dictionary(Of String, ComboBoxBindingInfo)

    ' Tambahkan DataView ke manager dengan nama key
    Public Sub AddView(_key As String, _dt As DataTable)
        Try
            If dataViews.ContainsKey(_key) Then
                dataViews(_key).Table.Clear()
                dataViews(_key).Table.Merge(_dt)
            Else
                dataViews(_key) = New DataView(_dt)
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "AddView Error")
        End Try
    End Sub

    ' Dapatkan DataView berdasarkan key
    Public Function GetView(_key As String) As DataView
        Try
            If dataViews.ContainsKey(_key) Then
                Return dataViews(_key)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetView Error")
            Return Nothing
        End Try
    End Function

    ' Set RowFilter ke DataView tertentu
    Public Sub Filter(_key As String, _rowFilter As String)
        Try
            If dataViews.ContainsKey(_key) Then
                dataViews(_key).RowFilter = _rowFilter
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "Filter Error")
        End Try
    End Sub

    ' Clear RowFilter
    Public Sub ClearFilter(_key As String)
        Try
            If dataViews.ContainsKey(_key) Then
                dataViews(_key).RowFilter = ""
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ClearFilter Error")
        End Try
    End Sub

    ' Daftarkan ComboBox untuk binding
    Public Sub RegisterComboBox(_key As String, _comboBox As ComboBox, _displayMember As String, _valueMember As String)
        Try
            If dataViews.ContainsKey(_key) Then
                Dim dt = dataViews(_key).Table

                If Not dt.Columns.Contains(_displayMember) Then
                    Throw New ArgumentException($"DisplayMember '{_displayMember}' tidak ditemukan di DataTable.")
                End If
                If Not dt.Columns.Contains(_valueMember) Then
                    Throw New ArgumentException($"ValueMember '{_valueMember}' tidak ditemukan di DataTable.")
                End If
                'comboBoxes(key) = comboBox
                'comboBox.DataSource = dataViews(key)
                'comboBox.DisplayMember = displayMember
                'comboBox.ValueMember = valueMember
                'comboBox.SelectedIndex = -1
                comboBoxes(_key) = New ComboBoxBindingInfo With {
                    .Combo = _comboBox,
                    .DisplayMember = _displayMember,
                    .ValueMember = _valueMember
                }
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RegisterComboBox Error")
        End Try
    End Sub

    ' Re-bind ComboBox setelah filter
    Public Sub RebindComboBox(_key As String, Optional _autoSelectFirst As Boolean = False)
        Try
            'If comboBoxes.ContainsKey(key) AndAlso dataViews.ContainsKey(key) Then
            '    Dim cmb = comboBoxes(key)
            '    cmb.DataSource = Nothing
            '    cmb.DataSource = dataViews(key)
            '    cmb.DisplayMember = displayMember
            '    cmb.ValueMember = valueMember
            '    cmb.SelectedIndex = -1
            'End If
            If comboBoxes.ContainsKey(_key) AndAlso dataViews.ContainsKey(_key) Then
                Dim info = comboBoxes(_key)
                With info.Combo
                    .DataSource = Nothing
                    .DataSource = dataViews(_key)
                    .DisplayMember = info.DisplayMember
                    .ValueMember = info.ValueMember
                    .SelectedIndex = If(_autoSelectFirst, 0, -1)
                End With
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RebindComboBox Error")
        End Try
    End Sub

    ' Ambil ComboBox berdasarkan key
    Public Function GetComboBox(_key As String) As ComboBox
        Try
            If comboBoxes.ContainsKey(_key) Then
                Return comboBoxes(_key).Combo
            Else
                Return Nothing
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetComboBox Error")
            Return Nothing
        End Try
    End Function

    ' Reset pilihan ComboBox
    Public Sub ClearSelection(_key As String)
        Try
            If comboBoxes.ContainsKey(_key) Then
                comboBoxes(_key).Combo.SelectedIndex = -1
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ClearSelection Error")
        End Try
    End Sub

    ' Set pilihan ComboBox berdasarkan value
    Public Sub SetSelectedValue(_key As String, _value As Object)
        Try
            If comboBoxes.ContainsKey(_key) Then
                comboBoxes(_key).Combo.SelectedValue = _value
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetSelectedValue Error")
        End Try
    End Sub

    ' Ambil nilai terpilih
    Public Function GetSelectedValue(_key As String) As Object
        Try
            If comboBoxes.ContainsKey(_key) Then
                Return comboBoxes(_key).Combo.SelectedValue
            Else
                Return Nothing
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetSelectedValue Error")
            Return Nothing
        End Try
    End Function

    ' Ambil teks tampilan terpilih
    Public Function GetSelectedText(_key As String) As String
        Try
            If comboBoxes.ContainsKey(_key) Then
                Return comboBoxes(_key).Combo.Text
            Else
                Return ""
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetSelectedText Error")
            Return ""
        End Try
    End Function

    Public Sub UnregisterComboBox(_key As String)
        Try
            If comboBoxes.ContainsKey(_key) Then
                comboBoxes(_key).Combo.DataSource = Nothing
                comboBoxes.Remove(_key)
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "UnregisterComboBox Error")
        End Try
    End Sub

    Public Sub RemoveView(_key As String)
        Try
            If dataViews.ContainsKey(_key) Then
                dataViews.Remove(_key)
            End If
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "RemoveView Error")
        End Try
    End Sub

    Public Sub ClearAllViews()
        Try
            dataViews.Clear()
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ClearAllViews Error")
        End Try
    End Sub

    Public Function HasView(_key As String) As Boolean
        Try
            Return dataViews.ContainsKey(_key)
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "HasView Error")
            Return Nothing
        End Try
    End Function

    Public Function HasComboBox(_key As String) As Boolean
        Try
            Return comboBoxes.ContainsKey(_key)
        Catch ex As Exception
            myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "HasComboBox Error")
            Return Nothing
        End Try
    End Function

    '<Conditional("DEBUG")>
    'Private Sub DebugLog(msg As String)
    '    Debug.WriteLine("[CViewManager] " & msg)
    'End Sub

    'Cara panggilnya
    'DebugLog($"Rebinding ComboBox {key}")

End Class