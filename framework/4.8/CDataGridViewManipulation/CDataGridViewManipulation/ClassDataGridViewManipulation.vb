Imports System.Windows.Forms
Imports System.Drawing

Public Class CDataGridViewManipulation
    Private myCShowMessage As New CShowMessage.CShowMessage

    Public Sub AutoNumberRowsForGridViewWithPaging(_dataGridView As DataGridView, _records As Integer)
        '>>>CARA 1 auto number tampil pada header row
        Try
            If _dataGridView IsNot Nothing Then
                Dim offSet As Integer = _records
                While (_records <= ((offSet + _dataGridView.Rows.Count) - 1))
                    _dataGridView.Rows(_records Mod 10).HeaderCell.Value = String.Format((_records + 1).ToString(), "0")
                    _records += 1
                End While
            End If
            _dataGridView.RowHeadersWidth = 80
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "AutoNumberRowsForGridViewWithPaging Error")
        End Try
    End Sub

    Public Sub AutoNumberRowsForGridViewWithoutPaging(_dataGridView As DataGridView, _records As Long)
        '>>>CARA 1 auto number tampil pada header row
        Try
            If _dataGridView IsNot Nothing Then
                Dim offSet As Long = 0
                'TANPA PAGING
                While (offSet < _records)
                    _dataGridView.Rows(offSet).HeaderCell.Value = String.Format((offSet + 1).ToString(), "0")
                    offSet += 1
                End While
            End If
            _dataGridView.RowHeadersWidth = 80
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "AutoNumberRowsForGridViewWithoutPaging Error")
        End Try
    End Sub

    Public Sub SetDGVColour(_dataGridView As DataGridView)
        '>>>KASIH WARNA KE KOLOM HEADER
        Try
            Dim cs As New System.Windows.Forms.DataGridViewCellStyle
            cs.BackColor = Color.LemonChiffon
            _dataGridView.AlternatingRowsDefaultCellStyle = cs
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetDGVColour Error")
        End Try
    End Sub

    Public Sub SetGridViewSortState(ByRef _dataGridView As DataGridView, sortMode As DataGridViewColumnSortMode)
        Try
            For Each cols As DataGridViewColumn In _dataGridView.Columns
                cols.SortMode = sortMode
            Next
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SetGridViewSortState Error")
        End Try
    End Sub
End Class

Public Class CalendarColumn
    Inherits DataGridViewColumn

    Public Sub New()
        MyBase.New(New CalendarCell())
    End Sub

    Public Overrides Property CellTemplate() As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(value As DataGridViewCell)

            ' Ensure that the cell used for the template is a CalendarCell.
            If (value IsNot Nothing) AndAlso
                Not value.GetType().IsAssignableFrom(GetType(CalendarCell)) _
                Then
                Throw New InvalidCastException("Must be a CalendarCell")
            End If
            MyBase.CellTemplate = value

        End Set
    End Property

End Class

Public Class CalendarCell
    Inherits DataGridViewTextBoxCell

    Public Sub New()
        ' Use the short date format.
        Me.Style.Format = "dd-MMM-yyyy"
    End Sub

    Public Overrides Sub InitializeEditingControl(rowIndex As Integer,
         initialFormattedValue As Object,
         dataGridViewCellStyle As DataGridViewCellStyle)

        ' Set the value of the editing control to the current cell value.
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue,
            dataGridViewCellStyle)

        Dim ctl As CalendarEditingControl =
            CType(DataGridView.EditingControl, CalendarEditingControl)

        ' Use the default row value when Value property is null.
        If (Me.Value Is Nothing) Then
            ctl.Value = CType(Me.DefaultNewRowValue, DateTime)
        Else
            ctl.Value = CType(Me.Value, DateTime)
        End If
    End Sub

    Public Overrides ReadOnly Property EditType() As Type
        Get
            ' Return the type of the editing control that CalendarCell uses.
            Return GetType(CalendarEditingControl)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType() As Type
        Get
            ' Return the type of the value that CalendarCell contains.
            Return GetType(DateTime)
        End Get
    End Property

    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
        Get
            ' Use the current date and time as the default value.
            Return DateTime.Now
        End Get
    End Property

End Class

Class CalendarEditingControl
    Inherits DateTimePicker
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer

    Public Sub New()
        Me.Format = DateTimePickerFormat.Short
    End Sub

    Public Property EditingControlFormattedValue() As Object _
        Implements IDataGridViewEditingControl.EditingControlFormattedValue

        Get
            Return Me.Value.ToShortDateString()
        End Get

        Set(value As Object)
            Try
                ' This will throw an exception of the string is 
                ' null, empty, or not in the format of a date.
                Me.Value = DateTime.Parse(CStr(value))
            Catch
                ' In the case of an exception, just use the default
                ' value so we're not left with a null value.
                Me.Value = DateTime.Now
            End Try
        End Set

    End Property

    Public Function GetEditingControlFormattedValue(context _
        As DataGridViewDataErrorContexts) As Object _
        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

        Return Me.Value.ToShortDateString()

    End Function

    Public Sub ApplyCellStyleToEditingControl(dataGridViewCellStyle As _
        DataGridViewCellStyle) _
        Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

        Me.Font = dataGridViewCellStyle.Font
        Me.CalendarForeColor = dataGridViewCellStyle.ForeColor
        Me.CalendarMonthBackground = dataGridViewCellStyle.BackColor

    End Sub

    Public Property EditingControlRowIndex() As Integer _
        Implements IDataGridViewEditingControl.EditingControlRowIndex

        Get
            Return rowIndexNum
        End Get
        Set(value As Integer)
            rowIndexNum = value
        End Set

    End Property

    Public Function EditingControlWantsInputKey(key As Keys,
         dataGridViewWantsInputKey As Boolean) As Boolean _
        Implements IDataGridViewEditingControl.EditingControlWantsInputKey

        ' Let the DateTimePicker handle the keys listed.
        Select Case key And Keys.KeyCode
            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right,
                Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp

                Return True

            Case Else
                Return Not dataGridViewWantsInputKey
        End Select

    End Function

    Public Sub PrepareEditingControlForEdit(selectAll As Boolean) _
        Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

        ' No preparation needs to be done.

    End Sub

    Public ReadOnly Property RepositionEditingControlOnValueChange() _
        As Boolean Implements _
        IDataGridViewEditingControl.RepositionEditingControlOnValueChange

        Get
            Return False
        End Get

    End Property

    Public Property EditingControlDataGridView() As DataGridView _
        Implements IDataGridViewEditingControl.EditingControlDataGridView

        Get
            Return dataGridViewControl
        End Get
        Set(value As DataGridView)
            dataGridViewControl = value
        End Set

    End Property

    Public Property EditingControlValueChanged() As Boolean _
        Implements IDataGridViewEditingControl.EditingControlValueChanged

        Get
            Return valueIsChanged
        End Get
        Set(value As Boolean)
            valueIsChanged = value
        End Set

    End Property

    Public ReadOnly Property EditingControlCursor() As Cursor _
        Implements IDataGridViewEditingControl.EditingPanelCursor

        Get
            Return MyBase.Cursor
        End Get

    End Property

    Protected Overrides Sub OnValueChanged(eventargs As EventArgs)

        ' Notify the DataGridView that the contents of the cell have changed.
        valueIsChanged = True
        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
        MyBase.OnValueChanged(eventargs)

    End Sub
End Class
