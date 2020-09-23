VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   Caption         =   "SMTP Mail Sender"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Subject"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   7095
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Hi Everyone"
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message"
      Height          =   2535
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   7095
      Begin VB.TextBox txtEmailBodyOfMessage 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "frmMail.frx":0000
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "To (""Name"" <Email Address>)"
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   7095
      Begin VB.TextBox txtToEmail 
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmMail.frx":001B
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "From (""Name"" <Email Address>)"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtFromEmail 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   """Sender Name"" <sender@emailaddress.com>"
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status:"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   7095
      Begin VB.TextBox txtStatus 
         Height          =   2205
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckMail 
      Index           =   0
      Left            =   7680
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_HOSTNAME_LEN = 132
Private Const MAX_DOMAIN_NAME_LEN = 132
Private Const MAX_SCOPE_ID_LEN = 260
Private Const MAX_ADAPTER_NAME_LENGTH = 260
Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const MIB_IF_TYPE_ETHERNET = 1
Private Const MIB_IF_TYPE_TOKENRING = 2
Private Const MIB_IF_TYPE_FDDI = 3
Private Const MIB_IF_TYPE_PPP = 4
Private Const MIB_IF_TYPE_LOOPBACK = 5
Private Const MIB_IF_TYPE_SLIP = 6

Private Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type
Private Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
'Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Private Const DNS_RECURSION As Byte = 1



Private Type DNS_HEADER
    qryID As Integer
    options As Byte
    response As Byte
    qdcount As Integer
    ancount As Integer
    nscount As Integer
    arcount As Integer
End Type

Private Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Const hostent_size = 16

Private Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Dim strMailBody As String
Private intConnections As Integer
Dim strFrom As String
Dim strFromName
Dim strSubject As String
Dim strImportance As String
Dim strMailer As String
Dim strMime As String
Dim strContentType As String
Dim strContentTrans As String
Dim strMailTo() As String
Dim strRcpt() As String
Dim strMailFrom As String
Dim strDNSServer As String
Dim lngOpen As Long 'Open connections
Dim lngEmails As Long
Dim blRunning As Boolean

Public Property Get Running() As Boolean
    Running = blRunning
End Property


Public Sub SendEmail(FromName As String, ToName As String, Subject As String, EmailBody As String)
    Dim lngEmail As Long
    Dim strEmailTo As String
    Dim strDateNow As String
    
    Dim dnsHead As DNS_HEADER
    Dim lngCount As Long
    Dim strDNS As String
    
    ' Query Variables
    Dim dnsQuery() As Byte
    Dim sQName As String
    Dim dnsQueryNdx As Integer
    Dim iTemp As Integer
    Dim iNdx As Integer
    Dim strDomainName As String
    
    'Don't run if currently running
    If blRunning Then Exit Sub
    lngOpen = 0
    blRunning = True
    
    ' Set the DNS parameters
    
    
    strMailTo = Split(ToName, vbCrLf)
    lngEmails = UBound(strMailTo)
    ReDim strRcpt(lngEmails + 1)
    
    strMailFrom = "Mail From: " & Mid(FromName, InStr(1, FromName, "<") + 1)
    strMailFrom = Left(strMailFrom, Len(strMailFrom) - 1) & vbCrLf
    strFromName = "From: " & FromName & vbCrLf ' Who's Sending
    strSubject = "Subject: " & Subject & vbCrLf ' Subject of E-Mail
    strImportance = "Importance: High" & vbCrLf   'Sets Importance(this will be customizible in newer version)
    strMime = "MIME-Version: 1.0" & vbCrLf   'Gives MIME Version
    strMailer = "X-Mailer: advInfoProc v 2.x" & vbCrLf   'Gives mail clients name
    strContentType = "Content-Type: text/html" & vbCrLf  'Gives content type
    strContentTrans = "Content-Transfer-Encoding: 7bit" & vbCrLf    'gives encoding (this will be customizible in newer version)
    strMailBody = EmailBody & vbCrLf  ' E-mail message body
    For lngEmail = 0 To lngEmails
        If intConnections < lngEmail Then
            Load Me.sckMail(lngEmail)
            intConnections = intConnections + 1
        End If
        If Len(strMailTo(lngEmail)) > 0 Then
            'Debug.Print "Winsock(" & lngEmail & ").State = " & Me.sckMail(lngEmail).State
            If Me.sckMail(lngEmail).State <> sckClosed Then Me.sckMail(lngEmail).Close
            If Me.sckMail(lngEmail).State = sckClosing Then
                Do Until Me.sckMail(lngEmail).State = sckClosed
                    DoEvents
                Loop
            End If
            strEmailTo = strMailTo(lngEmail)
            strRcpt(lngEmail) = "rcpt to: <" & Mid(strEmailTo, InStr(1, strEmailTo, "<") + 1) & vbCrLf
            strMailTo(lngEmail) = "TO: " & strEmailTo & vbCrLf ' Who it going to
            Me.sckMail(lngEmail).LocalPort = 0
            Me.sckMail(lngEmail).Protocol = sckTCPProtocol ' Set protocol for sending
            Me.sckMail(lngEmail).RemoteHost = strDNSServer
            
            dnsHead.qryID = htons(&H11DF)
            dnsHead.options = DNS_RECURSION
            dnsHead.qdcount = htons(1)
            dnsHead.ancount = 0
            dnsHead.nscount = 0
            dnsHead.arcount = 0
            dnsQueryNdx = 0
            ReDim dnsQuery(4000)
            MemCopy dnsQuery(dnsQueryNdx), dnsHead, 12
            dnsQueryNdx = dnsQueryNdx + 12
            
            ' Then the domain name (as a QNAME)
            strDomainName = Mid(strEmailTo, InStr(1, strEmailTo, "@") + 1) ' Set the server address
            strDomainName = Left(strDomainName, Len(strDomainName) - 1)
            sQName = MakeQName(strDomainName)
            iNdx = 0
            While (iNdx < Len(sQName))
                dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
                iNdx = iNdx + 1
            Wend
            Debug.Print strDomainName
            Debug.Print sQName
            dnsQueryNdx = dnsQueryNdx + Len(sQName)
            
            ' Null terminate the string
            dnsQuery(dnsQueryNdx) = &H0
            dnsQueryNdx = dnsQueryNdx + 1
            
            ' The type of query (15 means MX query)
            iTemp = htons(15)
            MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
            dnsQueryNdx = dnsQueryNdx + Len(iTemp)
            
            ' The class of query (1 means INET)
            iTemp = htons(1)
            MemCopy dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
            dnsQueryNdx = dnsQueryNdx + Len(iTemp)
            
            On Error Resume Next
            ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    
            'Need to get MX record for Address so we have correct mail server.
            Me.sckMail(lngEmail).Protocol = sckUDPProtocol
            Me.sckMail(lngEmail).RemotePort = 53 ' Set the SMTP Port
            Me.sckMail(lngEmail).Tag = "MXQUERY"
            'Send Query
            strDNS = ""
            For lngCount = 0 To dnsQueryNdx
                strDNS = strDNS & Chr(dnsQuery(lngCount))
            Next
            Debug.Print strDNS
            Me.sckMail(lngEmail).SendData dnsQuery
        End If
    Next
    
End Sub

Private Sub cmdSendMail_Click()
    SendEmail Me.txtFromEmail.Text, Me.txtToEmail.Text, Me.txtSubject.Text, Me.txtEmailBodyOfMessage.Text
End Sub

Private Sub Form_Load()
    strDNSServer = GetDNSinfo()
End Sub

Private Sub sckMail_Close(Index As Integer)
    'sckMail(Index).Close
    Debug.Print "Winsock " & Index & " Closed"
    lngOpen = lngOpen - 1
    If lngOpen = 0 Then blRunning = False
End Sub

Private Sub sckMail_Connect(Index As Integer)
    'sckMail(Index).SendData "HELO worldcomputers.com" & vbCrLf
    Debug.Print "Winsock " & Index & " Connected"
    lngOpen = lngOpen + 1
End Sub

Private Sub sckMail_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Handle Responses...
    Dim strBuffer As String
    Dim strMX As String
    Dim strBuff() As Byte
    ReDim strBuff(bytesTotal)
    strBuffer = Space(bytesTotal)
    If Me.sckMail(Index).Protocol = sckTCPProtocol Then
        Me.sckMail(Index).GetData strBuffer
    Else
        Me.sckMail(Index).GetData strBuff, vbArray + vbByte
    End If
    Me.txtStatus.Text = Me.txtStatus.Text & strBuffer
    Select Case Me.sckMail(Index).Tag
        Case "DNSCONNECT"
            Me.sckMail(Index).Tag = "MXQUERY"
            'Send DNS Query to Server
            Me.sckMail(Index).SendData vbCrLf
            
        Case "MXQUERY"
            'Got feedback from DNS Server
            Me.sckMail(Index).Tag = "CONNECT"
            Me.sckMail(Index).Close
            Do Until Me.sckMail(Index).State = sckClosed
                DoEvents
            Loop
            Me.sckMail(Index).Protocol = sckTCPProtocol
            Me.sckMail(Index).RemotePort = 25
            'Get Preferred Mail IP
            Dim iAnCount As Integer
            ' Get the number of answers
            MemCopy iAnCount, strBuff(6), 2
            iAnCount = ntohs(iAnCount)
            strMX = Trim(GetMXName(strBuff(), 12, iAnCount))
            Me.sckMail(Index).RemoteHost = strMX
            Debug.Print strMX
            Me.sckMail(Index).Connect
        Case "CONNECT"
            Me.sckMail(Index).Tag = "HELO"
            Me.sckMail(Index).SendData "HELO wpx-nw.com" & vbCrLf
        Case "HELO"
            Me.sckMail(Index).Tag = "MAILFROM"
            Debug.Print strMailFrom
            Me.sckMail(Index).SendData strMailFrom
        Case "MAILFROM"
            'Send RCPT
            Me.sckMail(Index).Tag = "RCPT"
            Me.sckMail(Index).SendData strRcpt(Index)
'            Pause 1
        Case "RCPT"
            If Len(strBuffer) > 0 Then
                If Left(strBuffer, 3) = "250" Then 'Good to go
                    'Send DATA
                    Me.sckMail(Index).Tag = "DATA"
        '            Debug.Print "DATA "
                    Me.sckMail(Index).SendData "DATA " & vbCrLf
        '            Debug.Print strFromName
                Else
                    Debug.Print strBuffer
                    Me.sckMail(Index).Close
                End If
            Else
                Me.sckMail(Index).Close
            End If
        Case "DATA"
            If Len(strBuffer) > 0 Then
                If Left(strBuffer, 3) = "354" Then 'Good to go
                    Me.sckMail(Index).SendData strFromName & strMailTo(Index) & strSubject & strImportance & strMime & strMailer & strContentType & strContentTrans
                    Pause 0.2
                    Me.sckMail(Index).Tag = "BODY"
                    Me.sckMail(Index).SendData strMailBody
                    Me.sckMail(Index).SendData "." & vbCrLf
                Else
                    Debug.Print strBuffer
                    Me.sckMail(Index).Close
                End If
            Else
                Me.sckMail(Index).Close
            End If
        Case "BODY"
            If Len(strBuffer) > 0 Then
                If Left(strBuffer, 3) = "250" Then 'Good to go let's close up
                    Me.sckMail(Index).Tag = "QUIT"
                    Me.sckMail(Index).SendData "QUIT" & vbCrLf
                Else
                    Debug.Print strBuffer
                    Me.sckMail(Index).Close
                End If
            Else
                Me.sckMail(Index).Close
            End If
        Case "QUIT"
            'It's closing down
        Case Else
            Debug.Print "?"
    End Select
End Sub

Private Sub sckMail_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print "Winsock " & Index & " Had Error " & Number & " - " & Description & vbCrLf & Source
    If sckMail(Index).Protocol = sckTCPProtocol Then
        lngOpen = lngOpen - 1
    End If
    sckMail(Index).Close
    If lngOpen = 0 Then blRunning = False
End Sub

Private Sub sckMail_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    'Debug.Print "Winsock " & Index & " Has sent " & bytesSent & " of " & bytesSent + bytesRemaining
End Sub

Public Sub Pause(Duration As Double) 'I needed this Sub to make intervals between the sending of commands
    Dim Current As Long 'Duration ican be change (i mean the amout of time)
    Current = Timer
    Do Until Timer - Current >= Duration 'Loops event until the current time matches the Duration defined
        DoEvents
    Loop
End Sub
Private Function GetDNSinfo() As String
    Dim error As Long
    Dim FixedInfoSize As Long
    Dim strDNS  As String
    Dim FixedInfo As FIXED_INFO
    Dim Buffer As IP_ADDR_STRING
    Dim FixedInfoBuffer() As Byte
    
    FixedInfoSize = 0
    error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If error <> 0 Then
        If error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetNetworkParams sizing failed with error: " & error
           Exit Function
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    

    error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If error = 0 Then
        CopyMemory FixedInfo, FixedInfoBuffer(0), Len(FixedInfo)
        strDNS = FixedInfo.DnsServerList.IpAddress
        strDNS = Replace(strDNS, vbCr, "")
        strDNS = Replace(strDNS, vbLf, "")
        strDNS = Replace(strDNS, vbNullChar, "")
        strDNS = Trim(strDNS)
        GetDNSinfo = strDNS
    End If
        
End Function

Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    Dim iChCount As Integer     ' Character counter
    Dim sTemp As String         ' Holds original query string

    Dim iMXLen As Integer
    Dim iBestPref As Integer    ' Holds the "best" preference number (lowest)
    Dim sBestMX As String       ' Holds the "best" MX record (the one with the lowest preference)

    iBestPref = -1

    ParseName dnsReply(), iNdx, sTemp
    ' Step over null
    iNdx = iNdx + 2

    ' Step over 6 bytes (not sure what the 6 bytes are, but all other
    '   documentation shows steping over these 6 bytes)
    iNdx = iNdx + 6

    'Dim xItem As ListItem

    On Error Resume Next
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer

            sName = ""
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6

            ' Read the MX data length specifier
            '              (not needed, hence why it's commented out)
            MemCopy iMXLen, dnsReply(iNdx), 2
            iMXLen = ntohs(iMXLen)

            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2

            MemCopy iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2

            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            Dim iNdx2 As Integer
            iNdx2 = iNdx
            ParseName dnsReply(), iNdx2, sName
            If (iBestPref = -1 Or iPref < iBestPref) Then
                iBestPref = iPref
                sBestMX = sName
            End If
            'Set xItem = fMainForm.ListView1.ListItems.Add(Text:=sName)
            'xItem.ListSubItems.Add Text:=iPref

            iNdx = iNdx + iMXLen + 1
            ' Step over 3 useless bytes
            'iNdx = iNdx + 3
        Else
            GetMXName = sBestMX
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend

    GetMXName = sBestMX
End Function


Private Function MakeQName(sDomain As String) As String
    Dim iQCount As Integer      ' Character count (between dots)
    Dim iNdx As Integer         ' Index into sDomain string
    Dim iCount As Integer       ' Total chars in sDomain string
    Dim sQName As String        ' QNAME string
    Dim sDotName As String      ' Temp string for chars between dots
    Dim sChar As String         ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
End Function

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    Dim iCompress As Integer        ' Compression index (index into original buffer)
    Dim iChCount As Integer         ' Character count (number of chars to read from buffer)
        
    ' While we didn't encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
End Sub


