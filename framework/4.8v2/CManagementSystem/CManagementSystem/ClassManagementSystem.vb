Imports System.Management

Public Class CManagementSystem
    Private myCShowMessage As New CShowMessage.CShowMessage

    Public Function GetComputerName() As String
        Try
            Dim ComputerName As String
            ComputerName = Environment.MachineName
            Return ComputerName
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetComputerName Error")
            Return Nothing
        End Try
    End Function

    Public Function GetMacAddress() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty And CBool(mo.Properties("IPEnabled").Value) = True) Then
                    hardwareID = mo.Properties("MacAddress").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetMacAddress Error")
            Return Nothing
        End Try
    End Function

    Public Function GetCPU_ID() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_Processor")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty) Then
                    hardwareID = mo.Properties("ProcessorId").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetCPU_ID Error")
            Return Nothing
        End Try
    End Function

    Public Function GetBiosSerialNumber() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_Bios")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty) Then
                    hardwareID = mo.Properties("SerialNumber").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetBiosSerialNumber Error")
            Return Nothing
        End Try
    End Function

    Public Function GetComputerSystemProduct() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_ComputerSystemProduct")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty) Then
                    hardwareID = mo.Properties("UUID").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetComputerSystemProduct Error")
            Return Nothing
        End Try
    End Function

    Public Function GetComputerBaseboard() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_BaseBoard")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty) Then
                    hardwareID = mo.Properties("SerialNumber").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetComputerBaseboard Error")
            Return Nothing
        End Try
    End Function

    Public Function GetVolumeSerial() As String
        Try
            Dim hardwareID As String = String.Empty
            Dim mc As ManagementClass = New ManagementClass("Win32_LogicalDisk")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            For Each mo As ManagementObject In moc
                If (hardwareID = String.Empty) Then
                    hardwareID = mo.Properties("VolumeSerialNumber").Value.ToString()
                End If
            Next
            Return hardwareID
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetVolumeSerial Error")
            Return Nothing
        End Try
    End Function
End Class
