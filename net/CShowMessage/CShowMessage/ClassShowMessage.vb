Imports System.Windows.Forms

Public Class CShowMessage
    Public Sub ShowInfo(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(_stMsg, _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub ShowWarning(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            MessageBox.Show(_stMsg, _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Public Sub ShowErrMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            MessageBox.Show(_stMsg, _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Public Function GetUserResponse(ByRef _stMsg As String, Optional _msgCaption As String = "") As DialogResult
        'ok
        'Dim iResponse As Short
        If (_msgCaption = "") Then
            GetUserResponse = MessageBox.Show(_stMsg, IIf(_msgCaption = "", "Confirmation", _msgCaption & " Confirmation"), MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        Else
            GetUserResponse = MessageBox.Show(_stMsg, _msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        End If
        'GetUserResponse = iResponse
    End Function

    Public Function GetUserResponseWithCancel(ByRef _stMsg As String, Optional _msgCaption As String = "") As DialogResult
        'ok
        'Dim iResponse As Short
        If (_msgCaption = "") Then
            GetUserResponseWithCancel = MessageBox.Show(_stMsg, IIf(_msgCaption = "", "Confirmation", _msgCaption & " Confirmation"), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        Else
            GetUserResponseWithCancel = MessageBox.Show(_stMsg, _msgCaption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        End If
        'GetUserResponseWithCancel = iResponse
    End Function

    Public Sub ShowSavedMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " berhasil ditambahkan", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(_stMsg & " berhasil ditambahkan", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub ShowUpdatedMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " berhasil diperbaharui", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(_stMsg & " berhasil diperbaharui", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub ShowDeletedMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " berhasil dihapus", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show(_stMsg & " berhasil dihapus", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub ShowFailedSaveMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " gagal ditambahkan", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show(_stMsg & " gagal ditambahkan", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Public Sub ShowFailedUpdateMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " gagal diperbaharui", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show(_stMsg & " gagal diperbaharui", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Public Sub ShowFailedDeleteMsg(ByRef _stMsg As String, Optional _msgCaption As String = "")
        'ok
        If (_msgCaption = "") Then
            MessageBox.Show(_stMsg & " gagal dihapus", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show(_stMsg & " gagal dihapus", _msgCaption, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
End Class
