Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Public Class CNetworkOperation

    Public Const Client_Broadcast_port As Integer = 18740
    Public Const Client_Broadcast_Address_All As String = "255.255.255.255"
    Public sendingUdpClient As UdpClient
    'Public iep As IPEndPoint = New IPEndPoint(IPAddress.Any, 18740)

    Public receivingUdpClient As UdpClient
    Public RemoteIpEndPoint As New System.Net.IPEndPoint(System.Net.IPAddress.Any, Client_Broadcast_port)
    Public ThreadReceive As System.Threading.Thread
    Public BitDet As BitArray

    Public GLOIP As IPAddress
    Public GLOINTPORT As Integer
    Public bytCommand As Byte() = New Byte() {}
    Public udpClient As New UdpClient

    Private myCShowMessage As New CShowMessage.CShowMessage

    Public IP_ASAL_ As String
    Public BROADCAST_ As String

    Public Sub InitializeSender_All()
        Try
            sendingUdpClient = New UdpClient(Client_Broadcast_Address_All, Client_Broadcast_port) 'Use broadcastAddress for sending data locally (on LAN), otherwise you'll need the public (or global) IP address of the machine that you want to send your data to
            sendingUdpClient.EnableBroadcast = True
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "InitializeSender_All Error")
        End Try
    End Sub

    Public Sub InitializeReceive_All()
        Try
            receivingUdpClient = New System.Net.Sockets.UdpClient(Client_Broadcast_port)
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "InitializeSender_All Error")
        End Try
    End Sub

    Public Sub CloseUDPConnection()
        Try
            'ThreadReceive.Abort()
            'receivingUdpClient.Close()
            'sendingUdpClient.Close()
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "CloseUDPConnection Error")
        End Try
    End Sub

    Public Sub sendMessageAlternatif(_sendString As String)
        Try
            Dim pret As Integer

            GLOIP = IPAddress.Parse("172.16.0.0")
            GLOINTPORT = Client_Broadcast_port
            udpClient.Connect(GLOIP, GLOINTPORT)
            If Not _sendString = "" Then
                bytCommand = Encoding.ASCII.GetBytes(_sendString)
                pret = udpClient.Send(bytCommand, bytCommand.Length)
                'Console.WriteLine("No of bytes send " & pret)
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "sendMessageAlternatif Error")
        End Try
    End Sub

    Public Sub SendMessage(_sendString As String)
        Try
            Dim pret As Integer
            If Not _sendString = "" Then
                Dim data() As Byte = Encoding.ASCII.GetBytes(_sendString)
                pret = sendingUdpClient.Send(data, data.Length)
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "SendMessage Error")
        End Try
    End Sub

    Public Sub ReceiveMessage()
        Try
            Dim receiveBytes As [Byte]() = receivingUdpClient.Receive(RemoteIpEndPoint)

            IP_ASAL_ = RemoteIpEndPoint.Address.ToString
            BitDet = New BitArray(receiveBytes)
            Dim strReturnData As String = System.Text.Encoding.ASCII.GetString(receiveBytes)
            BROADCAST_ = strReturnData
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReceiveMessage Error")
        End Try
    End Sub

    Public Sub ReceiverStart()
        Try
            ThreadReceive = New System.Threading.Thread(AddressOf ReceiveMessage)
            If (ThreadReceive.ThreadState = Threading.ThreadState.Unstarted) Then
                ThreadReceive.Start()
            End If
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReceiverStart Error")
        End Try
    End Sub

    Public Sub ReceiverStop()
        Try

        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "ReceiverStart Error")
        End Try
    End Sub

    Public Sub GetBitDet()
        Try
            Dim tempStr2 As String
            Dim tempStr As String
            Dim i As Integer = 0
            Dim line As Integer = 0
            Dim mTextBox As String

            For j As Integer = 0 To BitDet.Length - 1
                If BitDet(j) = True Then
                    Console.Write("1 ")
                    tempStr2 = tempStr
                    tempStr = " 1" + tempStr2
                Else
                    Console.Write("0 ")
                    tempStr2 = tempStr
                    tempStr = " 0" + tempStr2
                End If
                i += 1
                If i = 8 And j <= (BitDet.Length - 1) Then
                    line += 1
                    i = 0
                    mTextBox = mTextBox + tempStr
                    tempStr = ""
                    tempStr2 = ""
                    mTextBox = mTextBox + vbCrLf
                    If j <> (BitDet.Length - 1) Then
                        mTextBox = mTextBox + line.ToString & ") "
                        Console.WriteLine()
                    End If
                End If
            Next
        Catch ex As Exception
            Call myCShowMessage.ShowErrMsg("Pesan Error: " & ex.Message, "GetBitDet Error")
        End Try
    End Sub
End Class
