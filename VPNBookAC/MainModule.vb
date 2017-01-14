
Imports System.Net
Imports System.Text.RegularExpressions
Module MainModule

    Private ReadOnly Vpn As New RasManager()

    Sub Main()

        Console.Title = "VPNBook AutoConnector"
        Console.WriteLine("VPNBook AutoConnector by MooreR" & vbCrLf)

        RemoveEntries()

        Dim strHtml As String = New WebClient().DownloadString("http://www.vpnbook.com/")

        If GetLoginData(strHtml) Then

            ShowServers()

            Dim selIndex = Console.ReadLine()

            If SelectServer(selIndex) Is Nothing Then
                Console.WriteLine("No valid selection was made. Try again? y/N")
                If Console.ReadLine().ToLower = "y" Then
                    Console.Clear()
                    ShowServers()
                Else
                    End
                End If
            Else
                If Vpn.AddVpnEntry() Then
                    Vpn.LogIn()
                End If
            End If

        End If

        RemoveEntries()

        Console.ReadKey()

    End Sub
    Private Sub ShowServers()
        Console.WriteLine("Please choose the server you'd like to use:" & vbCrLf)
        Dim index = 0
        For Each p As Reflection.FieldInfo In GetType(RasManager.VpnServers).GetFields()
            index += 1
            Console.WriteLine(index & ":" & p.Name)
        Next
    End Sub
    Private Function SelectServer(ByVal selIndex As String) As String
        Select Case selIndex
            Case "0"
                Return RasManager.VpnServers.Europe217
            Case "1"
                Return RasManager.VpnServers.Europe214
            Case "2"
                Return RasManager.VpnServers.UnitedStates1_WebOnly
            Case "3"
                Return RasManager.VpnServers.UnitedStates2_WebOnly
            Case "4"
                Return RasManager.VpnServers.Canada1_WebOnly
            Case "5"
                Return RasManager.VpnServers.Germany233_WebOnly
            Case Else
                Return Nothing
        End Select
    End Function
    Private Sub RemoveEntries()
        If Vpn.DeleteAllEntries() Then
            Console.WriteLine("Successfully removed existing VPN")
        End If
    End Sub
    Private Function GetLoginData(ByVal data As String)

        Dim user As String = ""
        Dim pass As String = ""

        Dim regUser As New Regex("Username:\s(.*?)<")
        Dim regPass As New Regex("Password:\s(.*?)<")

        If regUser.IsMatch(data) Then
            user = regUser.Match(data).Groups(1).Value
            Console.WriteLine("Username found: " & user)
        End If

        If regPass.IsMatch(data) Then
            pass = regPass.Match(data).Groups(1).Value
            Console.WriteLine("Password found: " & pass & vbCrLf)
        End If

        If pass <> "" And user <> "" Then

            Vpn.VpnConnectionName = "OpenVPN-Temp"
            Vpn.Username = user
            Vpn.Password = pass

            Return True

        End If

        Return False

    End Function

End Module
