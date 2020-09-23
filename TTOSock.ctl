VERSION 5.00
Begin VB.UserControl TTOSock 
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   816
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TTOSock.ctx":0000
   ScaleHeight     =   816
   ScaleWidth      =   1140
   ToolboxBitmap   =   "TTOSock.ctx":2C7A
   Begin VB.Image Image1 
      Height          =   384
      Left            =   0
      Picture         =   "TTOSock.ctx":2F8C
      Top             =   0
      Width           =   384
   End
End
Attribute VB_Name = "TTOSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyUCKey As String
Private UserEnvironment As Boolean

'********************************************************************
'
'                              Events
'
'********************************************************************

Public Event ConnectionRequest(ByVal FromListeningSocket As Long)
Public Event Connected(ByVal SocketID As Long)
Public Event DataArrival(ByVal SocketID As Long, sData As String)
Public Event Error(ByVal SocketID As Long, ByVal number As Integer, Description As String)
Public Event ConnectionsAlert(ByVal SocketID As Long)
Public Event PeerClosing(ByVal SocketID As Long)
Public Event SendComplete(ByVal SocketID As Long)
Public Event StateChanged(ByVal SocketID As Long)
'

'********************************************************************
'
'                           Properties
'
'********************************************************************

'Returns the Current State of the specified Socket
Public Property Get State(SocketID As Long) As Integer

  Dim x As Integer
  'Checks for the index in m_lngSocks() array. If the socket does not
  'exist, -1 is returned.
  x = GetIndexFromsID(SocketID)
  
  'If the socket does not exist, return -1.
  'Otherwise, return the current state of the socket.
  If x = -1 Then
    State = -1
  Else
    State = CurrentState(x)
  End If

End Property

'Since each socket array can only handle up to 64 simultaneous
'event calls per thread to WindowProc(), you can raise an event
'when you reach a specified number of connections. By default,
'the event will raise at 50 open sockets. This limitation is
'in the Winsock API.
Public Property Let MaxConnectionsAlert(v_intMaxConnectionsAlert As Integer)

  m_intConnectionsAlert = v_intMaxConnectionsAlert

End Property

'Returns the Max Connections Alert Value
Public Property Get MaxConnectionsAlert() As Integer

  MaxConnectionsAlert = m_intConnectionsAlert

End Property

'********************************************************************
'
'                              Methods
'
'********************************************************************

'This function allows you to Accept a connect from a listening socket,
'spawn a new Socket, and connect using this new socket.
Public Function AcceptConnection(FromListeningSocket As Long) As Long
  
  'I'm not sure why RC is used in the Winsock API as the error checking
  'variable, but it generally is, so I used it too.
  Dim RC As Long
  Dim i As Integer
  Dim ReadSockBuffer As sockaddr
      
  'Gets the next available index in the m_lngSocks() array.
  i = GetNextSocksIndex
  m_lngSocks(i) = accept(FromListeningSocket, ReadSockBuffer, Len(ReadSockBuffer))
  
  'If -1 is returned, then an error has occurred in the API accept
  'routine. If -1 is not returned, then the number of the new
  'socket is returned.
  If m_lngSocks(i) = -1 Then
    'Handle Error
    Exit Function
  End If
  
  'This sets the Asyncronous values to raise an event on.
  'See the Readme module for more on this.
  RC = WSAAsyncSelect(m_lngSocks(i), hWnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_READ Or FD_CLOSE Or FD_WRITE Or FD_OOB Or FD_CONNECT)
         
  'If -1 is returned, then an error has occurred in the API
  'WSAAsyncSelect routine. Another value is acceptable.
  If RC = -1 Then
    'Handle Error
    Exit Function
  End If
    
  m_intSocketAsync(i) = RC
    
  'For some reason, FD_CONNECT is not being raised in the WindowProc
  'fucntion, so I am raising here, after the accept has been completed.
  'I have a feeling that it may be because:
  '1) We do not accept connections on the listening socket, the
  '    accept connection is passed off to another socket;
  '2) Therefore, by setting the WSAAsyncSelect AFTER (which is when we
  '    are forced to set it) the accept, we are missing the connect
  '    event all together;
  '3) I will be treating this as a bug, and looking for a resolution to
  '    it.
  RaiseEvent Connected(m_lngSocks(i))
    
  'Return the number of the new socket
  AcceptConnection = m_lngSocks(i)
    
End Function

'This function allows you to open a socket, and attempt to connect to
'a server.
Public Function ConnectTo(IP As String, OnPort As Integer) As Long

  Dim RC As Long
  Dim SocketBuffer As sockaddr
  Dim lngSocket As Long
  Dim i As Integer
  Dim lngResolvedIP As Long
  
  'This function attempts to resolve the IP address that we will connect to.
  'A long Network Byte Order is returned.
  lngResolvedIP = ResolveIPtoNBO(IP)
  
  'If -1 is returned, then the IP could not be resolved.
  If lngResolvedIP = -1 Then
    'Handle Error
    Exit Function
  End If
  
  'When we connect on a socket, we need to tell the OS to call a
  'procedure that will act as our event handler. I'm not sure
  'exactly what the function returns, but I have seen that it
  'is used in a variety of projects, so I used it too.
  'hWnd - The handle of the usercontrol.
  'GWL_WNDPROC - Not too sure on this constant.
  'AddressOf WindowProg - returns the memory location of the
  'function that we will use to handle our events. This function
  'is in the API_Declarations module.
  OldWndProc = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, AddressOf WindowProc)
    
  'This aquires a new socket for us to use.
  lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
  'If there was an error creating the socket, lngSocket would contain
  'the value -1, otherwise it would contain the value of the socket handle.
  'We have not established that the socket will be used to connect
  'with yet, we have simply aquired a new socket.
  If lngSocket = -1 Then
    'Handle Error
    Exit Function
  End If
  
  'This defines several variable for the a new listening socket,
  'As per the Winsock API Type Definition
  SocketBuffer.sin_family = AF_INET
  SocketBuffer.sin_port = htons(OnPort)
  SocketBuffer.sin_addr = lngResolvedIP
  SocketBuffer.sin_zero = String$(8, 0)
  
  RC = Connect(lngSocket, SocketBuffer, Len(SocketBuffer))
  
  'If -1 is returned, then an error has occurred in the API accept
  'routine. If 0 is returned, then the socket is connected.
  If RC <> 0 Then
    'Handle Error
    closesocket CInt(lngSocket)
    Exit Function
  End If
    
  'This sets the Asyncronous values to raise an event on.
  'See the Readme module and http://ttosock.2y.net for more on this.
  RC = WSAAsyncSelect(lngSocket, hWnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_READ Or FD_CLOSE Or FD_WRITE Or FD_OOB Or FD_CONNECT)
  If RC <> 0 Then
    'Handle Errors
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
  End If
    
  'Gets the next available index in the m_lngSocks() array.
  i = GetNextSocksIndex
  m_lngSocks(i) = lngSocket
  m_intSocketAsync(i) = RC
  
  ConnectTo = m_lngSocks(i)
  
End Function


'Destroys the connection and references to this SocketID
Public Function Disconnect(SocketID As Long) As Boolean

  'Since a user could possibly attempt to destroy a socket that has
  'already been destroyed, we need to keep chugging along with the
  'routine.
  On Error Resume Next
  Dim x As Integer
  Dim RC As Integer
  
  'This is an API call to close the socket
  closesocket CInt(SocketID)
  
  'Remove Array Reference
  x = GetIndexFromsID(SocketID)
  
  m_lngSocks(x) = -1
  Disconnect = True
  
      
End Function

'Returns the local IP being used. If the socket is
'listening, the IP on which we are listing is returned.
'If the SocketID does not exist, "0.0.0.0" is returned.
Public Function LocalIP(SocketID As Long) As String
  
  Dim x As String
  Dim y As Integer
    
  'If the socket does not exist, return 0.0.0.0
  If Not IDExists(SocketID) Then
    LocalIP = "0.0.0.0"
    Exit Function
  End If
  
  'The peer address function returns the IP and the port in
  'one string, where the format of the string is:
  '      IP        Port
  '###.###.###.###:####
  'We just have to strip away the port.
  x = GetSockAddress(Int(SocketID))
  
  y = InStr(1, x, ":")
  x = Left$(x, y - 1)

  LocalIP = x
    
End Function

'Returns the local port being used. If the SocketID is connected
'to another IP, the port in use is returned. If the socket is
'listening, the port on which we are listing is returned.
'If the SocketID does not exist, -1 is returned.
Public Function LocalPort(SocketID As Long) As Variant

  Dim x As String
  Dim y As Integer
  
  'If the socket does not exist, return 0.0.0.0
  If Not IDExists(SocketID) Then
    LocalPort = 0
    Exit Function
  End If
    
  'The peer address function returns the IP and the port in
  'one string, where the format of the string is:
  '      IP        Port
  '###.###.###.###:####
  'We just have to strip away the port.
  x = GetSockAddress(Int(SocketID))

  y = InStr(1, x, ":")
  
  x = Right$(x, Len(x) - y)

  LocalPort = CLng(x)

End Function

'This function opens a port to listen on, as defined in the
'HostPort property
Public Function ListenNow(OnPort As Long) As Long
  
  Dim RC As Long
  Dim SocketBuffer As sockaddr
  Dim lngSocket As Long
  Dim i As Integer
  
  'When we listen on a socket, we need to tell the OS to call a
  'procedure that will act as our event handler. I'm not sure
  'exactly what the function returns, but I have seen that it
  'is used in a variety of projects, so I used it too.
  'hWnd - The handle of the usercontrol.
  'GWL_WNDPROC - Not too sure on this constant.
  'AddressOf WindowProg - returns the memory location of the
  'function that we will use to handle our events. This function
  'is in the API_Declarations module.
  OldWndProc = SetWindowLong(UserControl.hWnd, GWL_WNDPROC, AddressOf WindowProc)
  
  lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
  'If there was an error lngSocket would contain the value
  '-1, otherwise it would contain the value of the socket handle.
  'We have not established that the socket will be used to listen
  'on yet, we have simply aquired a new socket.
  If lngSocket = -1 Then
    'Handle Error
    Exit Function
  End If
    
  'This defines several variable for the a new listening socket,
  'As per the Winsock API Type Definition
  SocketBuffer.sin_family = AF_INET
  SocketBuffer.sin_port = htons(OnPort)
  SocketBuffer.sin_addr = 0
  SocketBuffer.sin_zero = String$(8, 0)
    
  'Here we bind the newly aquired socket with the data that we
  'have defined in SocketBuffer
  RC = bind(lngSocket, SocketBuffer, 16)
  
  'If RC = 0, the bind was successful. If any other value was
  'returned, the bind was a bust.
  If RC <> 0 Then
    'Deal with Error
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
  End If
    
  'Attempt to listen on the socket. The 5 specifies the maximum
  'number of connections to hold in the queue. That is, if 7 sockets
  'attempt to connect to our listening socket over say 1 second, and
  'our program does not accept the connections (for whatever reason)
  'we will reject the last two attempts to connect to us.
  'This value can be defined from 1 to 5. Any other value is
  'rounded to the nearest valid integer.
  'Also, I don't think that this value has any bearing in a
  'Windows 9x environment. I think that it is always 5, regardless
  'of what we define it as.
  RC = listen(ByVal lngSocket, ByVal 5)
  
  'If RC = 0, the bind was successful. If any other value was
  'returned, the listen was a bust.
  If RC <> 0 Then
    'Handle Errors
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
  End If
  
  'This sets the Asyncronous values to raise an event on.
  'See the Readme module for more on this.
  RC = WSAAsyncSelect(lngSocket, hWnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_CONNECT Or FD_ACCEPT)
  If RC <> 0 Then
    'Handle Errors
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
  End If

  'Gets the next available index in the m_lngSocks() array.
  i = GetNextSocksIndex
  m_lngSocks(i) = lngSocket
  m_intSocketAsync(i) = RC
  
  ListenNow = lngSocket
  
End Function

'This function releases the UCKey from the Public Functions module
'so that the instance can be closed cleanly.
'This was added to resolve a memory leak.
Public Function ReleaseInstance()

  CleanUp MyUCKey

End Function

'This function returns the remote IP of the socket that we
'specify.
Public Function RemoteIP(SocketID As Long) As String

  Dim x As String
  Dim y As Integer
    
  'If the socket does not exist, return 0.0.0.0
  If Not IDExists(SocketID) Then
    RemoteIP = "0.0.0.0"
    Exit Function
  End If
  
  'The peer address function returns the IP and the port in
  'one string, where the format of the string is:
  '      IP        Port
  '###.###.###.###:####
  'We just have to strip away the port.
  x = GetPeerAddress(Int(SocketID))
  
  y = InStr(1, x, ":")
  x = Left$(x, y - 1)

  RemoteIP = x

End Function

'This function returns the remote port of the socket that we
'specify.
Public Function RemotePort(SocketID As Long) As Long

  Dim x As String
  Dim y As Integer
  
  'If the socket does not exist, return 0.0.0.0
  If Not IDExists(SocketID) Then
    RemotePort = 0
    Exit Function
  End If
    
  'The peer address function returns the IP and the port in
  'one string, where the format of the string is:
  '      IP        Port
  '###.###.###.###:####
  'We just have to strip away the port.
  x = GetPeerAddress(Int(SocketID))

  y = InStr(1, x, ":")
  
  x = Right$(x, Len(x) - y)

  RemotePort = CLng(x)

End Function

'This function sends data out to the specified socket. If no
'socket is specified (or if socket 0 is specified) the data
'is sent to all sockets.
Public Function SendDataTo(DataToSend As String, Optional SocketID As Long = 0)

  Dim DummyDataToSend As String

  If SocketID = 0 Then
    
    Dim x As Integer
    
    'Sends the data to each socket
    For x = 1 To m_intMaxSockCount
      If m_lngSocks(x) <> -1 Then
        DummyDataToSend = DataToSend
        If ICanUseCryptionObject And IShouldUseCryptionObject(x) Then DummyDataToSend = CryptionObject.Encrypt(DataToSend, CryptionKey(x))
        SendToSock DummyDataToSend, m_lngSocks(x)
      End If
    Next x
    
  Else
    
    'If the socket does not exist, get out
    If Not IDExists(SocketID) Then
      Exit Function
    End If
    
    DummyDataToSend = DataToSend
    If ICanUseCryptionObject And IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then DummyDataToSend = CryptionObject.Encrypt(DataToSend, CryptionKey(GetIndexFromsID(SocketID)))
    SendToSock DummyDataToSend, SocketID
    
  End If

End Function

'This funtion allows us to specify a unique encryption key for each
'socket that we are connected to.
Public Function SetCryptionKey(sKey As String, SocketID As Long) As Boolean

  If Not IDExists(SocketID) Then
    SetCryptionKey = False
    Exit Function
  End If
  
  SetCryptionKey = True
  CryptionKey(GetIndexFromsID(SocketID)) = sKey

End Function

'This function allows us to use an external cryption object in our
'project if we choose to. The encryption process is being placed inside
'of the usercontrol for a couple of reasons:
'
'  1) We only need to add a couple of lines in our existing project
'     to add cryption to it.
'  2) We don't need to add a cryption line before every SendData call.
'
'This function returns one of three options:
'  0 - indicates that the cryption object has been tested sucessfully
'      and the object can now be used.
'  1 - indicates that after encrypting a test string and decrypting
'      the result, the test string did not equal the final result.
'  2 - indicates that an error occurred when attempting to access the
'      cryption object.
'
'Note: We can do this because when we declare something as an object,
'we have acces to that objects properties and methods.
'Object declarations are similar to Variant declarations, in the
'respect that a specific data type is not required, but more
'memory is needed to use them.
Public Function SetCryptionObject(CryptionObj As Object) As Integer

  On Error GoTo Error_Handle
  
  Dim TestString As String
  Dim EncryptReturnString As String
  Dim DecryptReturnString1 As String
  Dim DecryptReturnString2 As String
  Dim x As Integer
  
  'This sets a cryption key. If your cryption object does
  'use keys (or uses a static key), then the CryptionString
  'can be trashed in the cryption function, HOWEVER, it must
  'still be declared to avoid Object errors.
  'Try declaring your function as follows:
  'Public Function Encrypt(TestString as String, Optional CryptionString as String = "") as String
    
  'We need to test for all valid ascii characters to ensure that the
  'cryption process will be valid for any input. NOTE: TAB (ASCII 9) is
  'not considered to be a valid ascii character in the cryption process.
  TestString = vbCrLf
  
  For x = 32 To 126
    TestString = TestString & Chr$(x)
  Next x
    
  'This tests the cryption process useing a long key.
  EncryptReturnString = CryptionObj.Encrypt(TestString, "TestEncryptionString")
  DecryptReturnString1 = CryptionObj.Decrypt(EncryptReturnString, "TestEncryptionString")
  
  'This tests the cryption process useing a short key.
  EncryptReturnString = CryptionObj.Encrypt(TestString, "a")
  DecryptReturnString2 = CryptionObj.Decrypt(EncryptReturnString, "a")
    
  'This tests the cryption process useing a null key. This one is
  'particularily important because by default, the CryptionKey is null
  EncryptReturnString = CryptionObj.Encrypt(TestString, "")
  DecryptReturnString2 = CryptionObj.Decrypt(EncryptReturnString, "")
  
  'This checks to see if the object could encrypt/decrypt the function
  'to our satisfaction
  If TestString = DecryptReturnString1 And TestString = DecryptReturnString2 Then
    SetCryptionObject = 0
    ICanUseCryptionObject = True
    Set CryptionObject = CryptionObj
  Else
    SetCryptionObject = 1
  End If
  
  Exit Function
  
Error_Handle:
  'If there was an object error:
  SetCryptionObject = 2
  Exit Function
  'We need to include Resume Next as earlier versions of VB6 required
  'it whenever we used On Error
  Resume Next

End Function

'This returns a description of the state specified, based on socket,
'and state specified
Public Function StateDescription(Optional SocketID As Long = 0, Optional v_intSelState As Integer = -1) As String

  Dim x As Integer
  
  'If the state specifed is invalid (not 0-9), do this
  If v_intSelState < 0 Or v_intSelState > 9 Then
     
    'Get the SocketID stack index, based on the socket id
    x = GetIndexFromsID(SocketID)
  
    'If x = -1, the socked id does not exist.
    If x = -1 Then
      'Return an error.
      StateDescription = "Socket Does Not Exist"
    Else
      'Return the state of the socket
      StateDescription = WinsockStates(CurrentState(x))
    End If
    
  Else
    'If the state specified is valid (0-9), return its description
    StateDescription = WinsockStates(v_intSelState)
  End If

End Function


'This function allows us to turn cryption on and off for a
'specified socket.
'
'The OnOff variable in the declaration has three possible states:
'  0 - Attempts to turn off the cryption for the specified socket.
'  1 - Attempts to turn on the cryption for the specified socket.
'  Any Other Value - Returns the current state (enabled or
'      not enabled) for the specified socket's cryption.
'
'This function returns one of six options:
'  0 - indicates that the SocketID does not exist
'  1 - indicates that cryption for this socket has been turned off.
'  2 - indicates that cryption for this socket has been turned on.
'  3 - indicates that cryption for this socket could not be
'      turned on because no cryption object has been set.
'  4 - Returns that cryption is enabled for this socket.
'  5 - Returns that cryption is not enabled for this socket.
'
'Note: We can do this because when we declare something as an object,
'we have acces to that objects properties and methods.
Public Function UseCryption(SocketID As Long, Optional OnOff As Integer = -1) As Integer

  'Checks to see if the socketid exists
  If Not IDExists(SocketID) Then
    UseCryption = 0
    Exit Function
  End If
  
  'We are trying to turn the cryption off
  If OnOff = 0 Then
    
    IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = False
    UseCryption = 1
  
  'We are trying to turn the cryption on
  ElseIf OnOff = 1 Then
    
    'We will only turn the cryption on if a cryption object
    'has been set.
    If ICanUseCryptionObject Then
      IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = True
      UseCryption = 2
    Else
      IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = False
      UseCryption = 3
    End If
    
  'We want the current state of the socket's cryption
  Else
    
    If IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then
      UseCryption = 4
    Else
      UseCryption = 5
    End If
     
  End If

End Function

'********************************************************************
'
'                           More Methods
'                          ^^^^^^^^^^^^^^
'    These subs were added  to allow  the API_Declarations Module
'    access  to the  usercontrol's  events.  As such, the  events
'    are  also exposed  to the control's  user. If they  are  not
'    abused, this should be fine.
'
'********************************************************************

Public Sub RaiseConnected(ByVal SocketID As Long)

  RaiseEvent Connected(SocketID)

End Sub

Public Sub RaiseConnectionRequest(ByVal SocketID As Long)

  RaiseEvent ConnectionRequest(SocketID)

End Sub

Public Sub RaiseDataArrival(ByVal SocketID As Long, sData As String)

  RaiseEvent DataArrival(SocketID, sData)

End Sub

Public Sub RaiseError(ByVal SocketID As Long, ByVal number As Integer, Description As String)

  RaiseEvent Error(SocketID, number, Description)

End Sub

Public Sub RaisePeerClosing(ByVal SocketID As Long)

  Disconnect SocketID
  RaiseEvent PeerClosing(SocketID)

End Sub

Public Sub RaiseSendComplete(ByVal SocketID As Long)

  RaiseEvent SendComplete(SocketID)

End Sub

Public Sub RaiseStateChanged(ByVal SocketID As Long)

  RaiseEvent StateChanged(SocketID)

End Sub

'********************************************************************
'
'                    Private Functions and Subs
'
'********************************************************************

'This function returns an empty index in the socket stack. If there
'is no free socket index in the stack, a new entry is added.
Private Function GetNextSocksIndex() As Integer

  Dim x As Integer
  
  'Look through the stack and check for a socket ID that = -1
  For x = 1 To m_intMaxSockCount
    If m_lngSocks(x) = -1 Then
      'If we found one, return it's index and exit
      GetNextSocksIndex = x
      Exit Function
    End If
  Next x
  
  'If we did not find a free index in the socket stack, add new
  'space at the end of the arrays, and set their initial values.
  m_intMaxSockCount = m_intMaxSockCount + 1
  
  ReDim Preserve m_lngSocks(m_intMaxSockCount)
  ReDim Preserve m_intSocketAsync(m_intMaxSockCount)
  ReDim Preserve CurrentState(m_intMaxSockCount)
  ReDim Preserve IShouldUseCryptionObject(m_intMaxSockCount)
  ReDim Preserve CryptionKey(m_intMaxSockCount)
  
  m_lngSocks(m_intMaxSockCount) = -1
  m_intSocketAsync(m_intMaxSockCount) = -1
  CurrentState(m_intMaxSockCount) = -1
  IShouldUseCryptionObject(m_intMaxSockCount) = False
  CryptionKey(m_intMaxSockCount) = ""
  
  GetNextSocksIndex = m_intMaxSockCount
  
End Function

'This function checks to see if the SocketID exists in the
'socket stack.
Private Function IDExists(SocketID As Long) As Boolean

  Dim x As Integer
  
  For x = 1 To m_intMaxSockCount
    If m_lngSocks(x) = SocketID Then
      IDExists = True
      Exit Function
    End If
  Next x
  
  IDExists = False

End Function

'This function sends data to a specified socket
Private Sub SendToSock(msg As String, SocketID As Long)

  On Error Resume Next
  Dim RC As Long
    
  RC = SendData(SocketID, msg)
        
  'If RC = -1, then the data could not be sent.
  If RC = -1 Then
    'Handle Error
  Else
    RaiseEvent SendComplete(SocketID)
  End If
  
End Sub

'********************************************************************
'
'                         UserControl Events
'
'********************************************************************

Private Sub Initialize()
  
  'If we are in design mode, then we do not need to initialized anything.
  If Not UserEnvironment Then Exit Sub
  
  On Error Resume Next

  If Not WinsockStartedUp Then

    Dim RC As Long
    Dim StartupData As WSADataType

    'StartupData is returned to us. It contains various notes about
    'our system and the version of Winsock that we are using.
    RC = WSAStartup(&H101, StartupData)
  
    If RC = -1 Then
      'Handle Error
      Exit Sub
    End If
    
    WinsockStartedUp = True
    
    WinsockStates(0) = "Closed"
    WinsockStates(1) = "Open"
    WinsockStates(2) = "Listening"
    WinsockStates(3) = "Connection Pending"
    WinsockStates(4) = "Resolving Host"
    WinsockStates(5) = "Host Resolved"
    WinsockStates(6) = "Connecting"
    WinsockStates(7) = "Connected"
    WinsockStates(8) = "Peer Is Closing The Connection"
    WinsockStates(9) = "Error On Socket"
      
    m_intMaxSockCount = 1
    
    ReDim Preserve m_lngSocks(1)
    ReDim Preserve m_intSocketAsync(1)
    ReDim Preserve CurrentState(1)
    ReDim Preserve IShouldUseCryptionObject(1)
    ReDim Preserve CryptionKey(1)
    
    m_lngSocks(1) = -1
    m_intSocketAsync(1) = -1
    CurrentState(1) = -1
    IShouldUseCryptionObject(1) = False
    
    ICanUseCryptionObject = False
    
    CryptionKey(1) = ""
        
  End If

  'This function allows the API_Declarations to acces methods and
  'properties of this user control.
  MyUCKey = SetControlHost(Me)
  
  'Default value of when to trigger the ConnectionsAlert Event.
  m_intConnectionsAlert = 50

End Sub

Private Sub UserControl_Terminate()

  'If we are in design mode, then we do not need to clean up, as we
  'we have not initialized anything.
  If Not UserEnvironment Then Exit Sub
  
  On Error Resume Next
  
  Dim RC As Long
  Dim x As Integer
  
  'Closes any open sockets
  For x = 1 To m_intMaxSockCount
      
    closesocket CInt(m_lngSocks(x))
    RC = WSACancelAsyncRequest(m_intSocketAsync(x))
      
  Next x
  
  'Close winsock and clean it up.
  RC = WSCleanUp()

End Sub

'This just forces the usercontrol to remain a constant size
Private Sub UserControl_Resize()

  UserControl.Width = Image1.Width + 80
  UserControl.Height = Image1.Height + 60

End Sub

'This routine is called when properties are initialized, and only when the
'control is intialized. We use it set a boolean variable to determine
'if we are in Design Mode or Run Mode.
Private Sub UserControl_InitProperties()

  UserEnvironment = Ambient.UserMode
  
  Call Initialize
  
End Sub

'This routine is called when properties are read, and only when the
'control is intialized. We use it set a boolean variable to determine
'if we are in Design Mode or Run Mode.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserEnvironment = Ambient.UserMode
  
  Call Initialize
  
End Sub



