Attribute VB_Name = "modApiPing"
Option Explicit

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
Public Const PING_TIMEOUT = 200
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD And &H100 And &HFF& '///////////////////
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Public Type ICMP_OPTIONS
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
    Address As Long
    status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    DataPointer As Long
    Options As ICMP_OPTIONS
    Data As String * 250
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type


Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
(ByVal IcmpHandle As Long) As Long

Public Declare Function IcmpSendEcho Lib "icmp.dll" _
    (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
    (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long

Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
    (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long

Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
    (ByVal szHost As String) As Long

Public Declare Sub RtlMoveMemory Lib "KERNEL32" _
    (hpvDest As Any, _
    ByVal hpvSource As Long, _
    ByVal cbCopy As Long)


Public Function GetStatusCode(status As Long) As String

    Dim Msg As String
    
    Select Case status
        Case IP_SUCCESS: Msg = "ip success"
        Case IP_BUF_TOO_SMALL: Msg = "ip buf too_small"
        Case IP_DEST_NET_UNREACHABLE: Msg = "ip dest net unreachable"
        Case IP_DEST_HOST_UNREACHABLE: Msg = "ip dest host unreachable"
        Case IP_DEST_PROT_UNREACHABLE: Msg = "ip dest prot unreachable"
        Case IP_DEST_PORT_UNREACHABLE: Msg = "ip dest port unreachable"
        Case IP_NO_RESOURCES: Msg = "ip no resources"
        Case IP_BAD_OPTION: Msg = "ip bad option"
        Case IP_HW_ERROR: Msg = "ip hw_error"
        Case IP_PACKET_TOO_BIG: Msg = "ip packet too_big"
        Case IP_REQ_TIMED_OUT: Msg = "ip req timed out"
        Case IP_BAD_REQ: Msg = "ip bad req"
        Case IP_BAD_ROUTE: Msg = "ip bad route"
        Case IP_TTL_EXPIRED_TRANSIT: Msg = "ip ttl expired transit"
        Case IP_TTL_EXPIRED_REASSEM: Msg = "ip ttl expired reassem"
        Case IP_PARAM_PROBLEM: Msg = "ip param_problem"
        Case IP_SOURCE_QUENCH: Msg = "ip source quench"
        Case IP_OPTION_TOO_BIG: Msg = "ip option too_big"
        Case IP_BAD_DESTINATION: Msg = "ip bad destination"
        Case IP_ADDR_DELETED: Msg = "ip addr deleted"
        Case IP_SPEC_MTU_CHANGE: Msg = "ip spec mtu change"
        Case IP_MTU_CHANGE: Msg = "ip mtu_change"
        Case IP_UNLOAD: Msg = "ip unload"
        Case IP_ADDR_ADDED: Msg = "ip addr added"
        Case IP_GENERAL_FAILURE: Msg = "ip general failure"
        Case IP_PENDING: Msg = "ip pending"
        Case PING_TIMEOUT: Msg = "ping timeout"
        Case Else: Msg = "unknown msg returned"
    End Select
    
    GetStatusCode = CStr(status) & " [ " & Msg & " ]"

End Function


Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam And &H100 And &HFF&

End Function


Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function


Public Function Ping(szAddress As String) As Boolean

    Dim hPort As Long
    Dim dwAddress As Long
    Dim sDataToSend As String
    Dim iOpt As Long
    Dim ECHO As ICMP_ECHO_REPLY
    sDataToSend = "Echo This"
    dwAddress = AddressStringToLong(szAddress)
    
    Call SocketsInitialize
    hPort = IcmpCreateFile()
    
    If IcmpSendEcho(hPort, _
    dwAddress, _
    sDataToSend, _
    Len(sDataToSend), _
    0, _
    ECHO, _
    Len(ECHO), _
    PING_TIMEOUT) Then
    Ping = True
    Else: Ping = False
    End If
    
    Call IcmpCloseHandle(hPort)
    Call SocketsCleanup

End Function


Function AddressStringToLong(ByVal tmp As String) As Long

    Dim i As Integer
    Dim parts(1 To 4) As String
    
    i = 0
    
    'we have to extract each part of the
    '123.456.789.123 string, delimited by
    'a period
    While InStr(tmp, ".") > 0
    i = i + 1
    parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
    tmp = Mid(tmp, InStr(tmp, ".") + 1)
    Wend
    
    i = i + 1
    parts(i) = tmp
    
    If i <> 4 Then
    AddressStringToLong = 0
    Exit Function
    End If
    
    'build the long value out of the
    'hex of the extracted strings
    AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
    Right("00" & Hex(parts(3)), 2) & _
    Right("00" & Hex(parts(2)), 2) & _
    Right("00" & Hex(parts(1)), 2))

End Function


Public Function SocketsCleanup() As Boolean
    
    Dim X As Long
    
    X = WSACleanup()
    
    If X <> 0 Then
    MsgBox "Windows Sockets error " & Trim$(str$(X)) & _
    " occurred in Cleanup.", vbExclamation
    SocketsCleanup = False
    Else
    SocketsCleanup = True
    End If

End Function


Public Function SocketsInitialize() As Boolean

    Dim WSAD As WSADATA
    Dim X As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    
    If X <> 0 Then
    MsgBox "Windows Sockets for 32 bit Windows " & _
    "environments is not successfully responding."
    SocketsInitialize = False
    Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
    (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
    
    szHiByte = Trim$(str$(HiByte(WSAD.wVersion)))
    szLoByte = Trim$(str$(LoByte(WSAD.wVersion)))
    szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
    szBuf = szBuf & " is not supported by Windows " & _
    "Sockets for 32 bit Windows environments."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
    
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    szBuf = "This application requires a minimum of " & _
    Trim$(str$(MIN_SOCKETS_REQD)) & " supported sockets."
    MsgBox szBuf, vbExclamation
    SocketsInitialize = False
    Exit Function
    End If
    
    SocketsInitialize = True

End Function

