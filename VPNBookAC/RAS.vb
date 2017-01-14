Imports System.Runtime.InteropServices
Public Class Ras

    Private Const RAS_MaxEntryName As Integer = 256
    Private Const RAS_MaxDeviceType As Integer = 16
    Private Const RAS_MaxDeviceName As Integer = 128
    Private Const MAX_PATH As Integer = 260
    Private Const ERROR_BUFFER_TOO_SMALL As Integer = 603
    Private Const ERROR_SUCCESS As Integer = 0

    <StructLayout(LayoutKind.Sequential, Pack:=4, CharSet:=CharSet.Auto)> _
    Public Structure RASCONN
        Public dwSize As Integer
        Public hrasconn As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> _
        Public szEntryName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceType + 1)> _
        Public szDeviceType As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxDeviceName + 1)> _
        Public szDeviceName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        Public szPhonebook As String
        Public dwSubEntry As Integer
        Public guidEntry As Guid
        Public dwSessionId As Integer
        Public dwFlags As Integer
        Public luid As Guid
    End Structure

    Private Const RAS_MaxPhoneNumber As Integer = 128
    Private Const UNLEN As Integer = 256
    Private Const PWLEN As Integer = 256
    Private Const DNLEN As Integer = 15
    Private Const RAS_MaxCallbackNumber As Integer = RAS_MaxPhoneNumber


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure RASDIALPARAMS
        Public dwSize As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> Public szEntryName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxPhoneNumber + 1)> Public szPhoneNumber As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxCallbackNumber + 1)> Public szCallbackNumber As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=UNLEN + 1)> Public szUserName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=PWLEN + 1)> Public szPassword As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=DNLEN + 1)> Public szDomain As String
    End Structure

    Public Delegate Function myrasdialfunc(ByVal unMsg As Integer, ByVal rasconnstate As Integer, ByVal dwError As Integer) As Integer

    Private Declare Auto Function RasDial Lib "rasapi32.dll" (ByVal ByVallpRasDialExtensions As IntPtr, _
    ByVal lpszPhonebook As String, ByRef lpRasDialParams As RASDIALPARAMS, ByVal dwNotifierType As Integer, _
    ByVal lpvNotifier As myrasdialfunc, ByRef lphRasConn As IntPtr) As Integer

    Private Declare Auto Function RasHangUp Lib "rasapi32.dll" (ByVal ByValhRasCon As IntPtr) As Integer

    <DllImport("RASAPI32")> _
    Friend Shared Function RasEnumConnections(<[In], Out> ByVal lpRasCon() As RASCONN, ByRef lpcb As Integer, ByRef lpcConnections As Integer) As Integer

    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure RASENTRYNAME
        Public dwSize As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=RAS_MaxEntryName + 1)> Public szEntryName As String
        Public dwFlags As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH + 1)> Public szPhonebookPath As String
    End Structure

    <DllImport("rasapi32.dll", EntryPoint:="RasEnumEntries", CharSet:=CharSet.Auto)> _
    Private Shared Function RasEnumEntries( _
     ByVal reserved As String, _
     ByVal lpszPhoneBook As String, _
     ByVal lpRasEntryName As IntPtr, _
     ByRef lpc As Int32, _
     ByRef lpcEntries As Int32) As Integer
    End Function

    Public Sub EnumEntries(ByRef mincoming() As RASENTRYNAME)
        Dim entries() As RASENTRYNAME
        Dim rasentrynamelen As Integer = Marshal.SizeOf(GetType(RASENTRYNAME))
        Dim lpcb As Integer = rasentrynamelen
        Dim lpcentries As Integer

        Dim parray As IntPtr = Marshal.AllocHGlobal(rasentrynamelen)
        Marshal.WriteInt32(parray, rasentrynamelen)

        Dim ret As Integer = RasEnumEntries(Nothing, Nothing, parray, lpcb, lpcentries)

        If ret = ERROR_BUFFER_TOO_SMALL Then
            parray = Marshal.ReAllocHGlobal(parray, New IntPtr(lpcb))
            Marshal.WriteInt32(parray, rasentrynamelen)
            ret = RasEnumEntries(Nothing, Nothing, parray, lpcb, lpcentries)
        End If

        If ret = 0 And lpcentries > 0 Then
            ReDim entries(lpcentries)
            Dim pentry As IntPtr = parray
            Dim i As Integer
            For i = 0 To lpcentries - 1
                entries(i) = Marshal.PtrToStructure(pentry, GetType(RASENTRYNAME))
                pentry = New IntPtr(pentry.ToInt32 + rasentrynamelen)
            Next
        End If

        Marshal.FreeHGlobal(parray)

        mincoming = entries
    End Sub
    Public Sub EnumConnections(ByRef mincoming() As RASCONN)

        Dim structtype As Type = GetType(RASCONN)
        Dim structsize As Integer = Marshal.SizeOf(GetType(RASCONN))
        Dim bufsize As Integer = structsize
        Dim realcount As Integer
        Dim TRasCon() As RASCONN

        Dim bufptr As IntPtr = Marshal.AllocHGlobal(bufsize)
        Marshal.WriteInt32(bufptr, structsize)

        Dim retcode As Integer = RasEnumConnections(TRasCon, bufsize, realcount)

        If retcode = ERROR_BUFFER_TOO_SMALL Then
            bufptr = Marshal.ReAllocHGlobal(bufptr, New IntPtr(structsize))
            Marshal.WriteInt32(bufptr, structsize)
            retcode = RasEnumConnections(TRasCon, bufsize, realcount)
        End If

        If (retcode = 0) And (realcount > 0) Then
            ReDim TRasCon(realcount - 1)
            Dim i As Integer
            Dim runptr As IntPtr = bufptr

            For i = 0 To (realcount - 1)
                TRasCon(i) = Marshal.PtrToStructure(runptr, structtype)
                runptr = New IntPtr(runptr.ToInt32 + structsize)
            Next
        End If

        Marshal.FreeHGlobal(bufptr)

        mincoming = TRasCon

    End Sub
    Public Function DialEntry(ByVal mentryname As String, ByVal musername As String, _
    ByVal mpassword As String, ByRef mcallback As myrasdialfunc) As IntPtr

        Dim objRASParams As RASDIALPARAMS
        Dim hRASConn As IntPtr

        Dim mystr() As String

        With objRASParams
            .szEntryName = mentryname
            .szPhoneNumber = ""
            .szCallbackNumber = "*"
            .szUserName = musername
            .szPassword = mpassword
            .szDomain = "*"
            .dwSize = Marshal.SizeOf(GetType(RASDIALPARAMS))
        End With

        Dim intRet As Integer = RasDial(IntPtr.Zero, Nothing, objRASParams, 0, mcallback, hRASConn)
        DialEntry = hRASConn

    End Function
    Public Sub HangEntry(ByVal mentryname As String)

        Dim structtype As Type = GetType(RASCONN)
        Dim structsize As Integer = Marshal.SizeOf(GetType(RASCONN))
        Dim bufsize As Integer = structsize
        Dim realcount As Integer
        Dim TRasCon() As RASCONN

        Dim bufptr As IntPtr = Marshal.AllocHGlobal(bufsize)
        Marshal.WriteInt32(bufptr, structsize)

        Dim retcode As Integer = RasEnumConnections(TRasCon, bufsize, realcount)

        If retcode = ERROR_BUFFER_TOO_SMALL Then
            bufptr = Marshal.ReAllocHGlobal(bufptr, New IntPtr(structsize))
            Marshal.WriteInt32(bufptr, structsize)
            retcode = RasEnumConnections(TRasCon, bufsize, realcount)
        End If

        If (retcode = 0) And (realcount > 0) Then
            ReDim TRasCon(realcount - 1)
            Dim i As Integer
            Dim runptr As IntPtr = bufptr

            For i = 0 To (realcount - 1)
                TRasCon(i) = Marshal.PtrToStructure(runptr, structtype)
                runptr = New IntPtr(runptr.ToInt32 + structsize)
            Next
        End If

        Marshal.FreeHGlobal(bufptr)

        Dim m As RASCONN
        For Each m In TRasCon
            If m.szEntryName = mentryname Then
                RasHangUp(m.hrasconn)
            End If
        Next

    End Sub
    Public Function IsConnected(ByVal mentryname As String) As Boolean


        Dim rascon() As RASCONN = {}
        Dim dwSize As Integer = Marshal.SizeOf(GetType(RASCONN))
        Dim totalConnections As Integer
        Dim result As Integer = RasEnumConnections(rascon, dwSize, totalConnections)
        Debug.WriteLine(result)

        For Each m In rascon
            If m.szEntryName = mentryname Then
                IsConnected = True
            End If
        Next

        'Dim structtype As Type = GetType(RASCONN)
        'Dim structsize As Integer = Marshal.SizeOf(GetType(RASCONN))
        'Dim bufsize As Integer = structsize
        'Dim realcount As Integer
        'Dim TRasCon() As RASCONN

        'Dim bufptr As IntPtr = Marshal.AllocHGlobal(bufsize)
        'Marshal.WriteInt32(bufptr, structsize)

        'Dim retcode As Integer = RasEnumConnections(TRasCon, bufsize, realcount)

        'If retcode = ERROR_BUFFER_TOO_SMALL Then
        '    bufptr = Marshal.ReAllocHGlobal(bufptr, New IntPtr(structsize))
        '    Marshal.WriteInt32(bufptr, structsize)
        '    retcode = RasEnumConnections(TRasCon, bufsize, realcount)
        'End If

        'If (retcode = 0) And (realcount > 0) Then
        '    ReDim TRasCon(realcount - 1)
        '    Dim i As Integer
        '    Dim runptr As IntPtr = bufptr

        '    For i = 0 To (realcount - 1)
        '        TRasCon(i) = Marshal.PtrToStructure(runptr, structtype)
        '        runptr = New IntPtr(runptr.ToInt32 + structsize)
        '    Next
        'End If

        'Marshal.FreeHGlobal(bufptr)

        'Dim m As RASCONN
        'For Each m In TRasCon
        '    If m.szEntryname = mentryname Then
        '        IsConnected = True
        '    End If
        'Next

    End Function

End Class