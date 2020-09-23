Attribute VB_Name = "Module1"
Option Explicit
'||================================||
'|| Remember to use:||
'|| WSACleanup in Form_Unload()||
'|| IP_Initialize in Form_Load() ||
'||================================||
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128
Const SOCKET_ERROR = 0

Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
    h_addr_list As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type



Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "Icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "Icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean


Private Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, _
    addrType As Long) As Long





Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
    'checks if string is valid IP address


Private Type IP_OPTION_INFORMATION
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Sub xListKillDupes(listbox As listbox)
On Error Resume Next
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

Function ImaPingJ00(targetip As String) As String
Dim hostname As String
    hostname = targetip
    Dim hFile As Long, lpWSAData As WSADATA
    Dim hHostent As HOSTENT, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSAData)
    If gethostbyname(hostname + String(64 - Len(hostname), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal gethostbyname(hostname + String(64 - Len(hostname), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        Form1.text1.Text = Form1.text1.Text & "Unable to ping " & hostname & vbCrLf
        Exit Function
    End If
    OptInfo.Ttl = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        'Form1.text1.Text = Form1.text1.Text & vbCrLf & "Time out on " & hostname
    End If
    If EchoReply.Status = 0 Then
        Form1.text1.Text = Form1.text1.Text & "Reply from " + hostname + " recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms" & vbCrLf
    Else
        Form1.text1.Text = Form1.text1.Text & "Failure ..." & vbCrLf
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function
