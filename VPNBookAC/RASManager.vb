Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Net
Imports DotRas

Public Class RasManager
    Public Property VpnConnectionName As String
    Public Property Username As String
    Public Property Password As String
    Public Property Server As String
    Public Property IsConnected As Boolean = False
    Public Property ConnectedServerIp As IPAddress
    Public Property ConnectedClientIp As IPAddress
    Protected WithEvents PhoneBook As New RasPhoneBook
    Protected PhoneBookPath As String = RasPhoneBook.GetPhoneBookPath(RasPhoneBookType.AllUsers)
    Protected ConnectionHandle As RasHandle
    Public WithEvents Dialer As New RasDialer
    Public Structure VpnServers
        Shared Europe217 As String = "euro217.vpnbook.com"
        Shared Europe214 As String = "euro214.vpnbook.com"
        Shared UnitedStates1_WebOnly As String = "us1.vpnbook.com"           '(US VPN - optimized for fast web surfing; no p2p downloading)
        Shared UnitedStates2_WebOnly As String = "us2.vpnbook.com"           '(US VPN - optimized for fast web surfing; no p2p downloading)
        Shared Canada1_WebOnly As String = "ca1.vpnbook.com"            '(Canada VPN - optimized for fast web surfing; no p2p downloading)
        Shared Germany233_WebOnly As String = "de233.vpnbook.com"         '(Germany VPN - optimized for fast web surfing; no p2p downloading)
    End Structure

    Public Function AddVpnEntry() As Boolean
        Try
            PhoneBook.Open(PhoneBookPath)
            If RasEntry.Exists(VpnConnectionName, PhoneBookPath) Then
            Else
                Dim vpnEntry As RasEntry = RasEntry.CreateVpnEntry(VpnConnectionName, Server, RasVpnStrategy.PptpOnly, RasDevice.Create(VpnConnectionName, RasDeviceType.Vpn))
                PhoneBook.Entries.Add(vpnEntry)
            End If
        Catch
            Return False
        Finally
            If (PhoneBook IsNot Nothing) Then
                PhoneBook.Dispose()
            End If
        End Try
        Return True
    End Function
    Public Sub HangupConnections()
        Dim rcon As ReadOnlyCollection(Of RasConnection) = GetActiveConnections()
        If rcon IsNot Nothing Then
            For Each con In rcon
                con.HangUp()
            Next
            IsConnected = False
        End If
    End Sub
    Private Sub GetVpnActiveIp()
        Dim connections = GetActiveConnections()
        For Each connection In connections
            Dim ipAddress As RasIPInfo = connection.GetProjectionInfo(RasProjectionType.IP)
            If ipAddress IsNot Nothing Then
                ConnectedServerIp = ipAddress.ServerIPAddress.MapToIPv4()
                ConnectedClientIp = ipAddress.IPAddress.MapToIPv4()
            End If
        Next
    End Sub
    Public Function DeleteAllEntries() As Boolean
        HangupConnections()
        Try
            PhoneBook.Open(PhoneBookPath)
            For Each entry As RasEntry In From entry1 In PhoneBook.Entries Where entry1.EntryType = RasEntryType.Vpn
                If entry.Name = "OpenVPN-Temp" Then
                    entry.Remove()
                    Return True
                End If
            Next
        Catch ex As Exception
        End Try
        Return False
    End Function
    Public Function GetActiveConnections() As ReadOnlyCollection(Of RasConnection)
        Dim connections = RasConnection.GetActiveConnections()
        If connections.Count <> 0 Then
            IsConnected = True
            Return connections
        End If
        IsConnected = False
        Return Nothing
    End Function
    Public Function LogIn() As Boolean
        Dialer.EntryName = VpnConnectionName
        Dialer.PhoneBookPath = RasPhoneBook.GetPhoneBookPath(RasPhoneBookType.AllUsers)
        Try
            Dialer.Credentials = New NetworkCredential(Username, Password)
            ConnectionHandle = Dialer.DialAsync()
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.ToString())
            Return False
        End Try
        Return True
    End Function
    Sub New()
        If IsConnected Then
            GetVpnActiveIp()
        End If
    End Sub
    Private Sub Dialer_Error(sender As Object, e As ErrorEventArgs) Handles Dialer.Error
        Debug.WriteLine(e.GetException())
    End Sub
    Private Sub Dialer_StateChanged(ByVal sender As Object, ByVal e As StateChangedEventArgs) Handles Dialer.StateChanged
        Console.WriteLine("StateChanged: " & e.State.ToString())
    End Sub
    Private Sub Dialer_DialCompleted(ByVal sender As Object, ByVal e As DialCompletedEventArgs) Handles Dialer.DialCompleted
        If (e.Cancelled) Then
            Console.WriteLine("Cancelled!")
        ElseIf (e.TimedOut) Then
            Console.WriteLine("Connection attempt timed out!")
        ElseIf (e.Error IsNot Nothing) Then
            Console.WriteLine(e.Error.ToString())
        ElseIf (e.Connected) Then
            Console.WriteLine("Connection successful!")
        End If
        If (Not e.Connected) Then
            Console.WriteLine("Not connected :(")
        End If
    End Sub
End Class