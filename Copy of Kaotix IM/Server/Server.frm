VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Server 
   Caption         =   "Instant Messenger Server - 192.168.0.5"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock ServiceSocket 
      Index           =   0
      Left            =   4320
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1008
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Log"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2775
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4895
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Server.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intMax As Integer

Private Sub Command1_Click()
   Unload Me
   End
End Sub

Private Sub Form_Load()
    intMax = 0
    ServiceSocket(0).Listen
End Sub

Private Sub ServiceSocket_Close(Index As Integer)
   
   ServiceSocket(Index).Close
 
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 Dim user As String
 conn.Open sConnString
 
 'Deletes user from ONLINE database when he signs off
 rs2.Open "SELECT username FROM online where oindex = " + _
            CStr(Index), conn
 If Not rs2.EOF Then
    user = rs2.Fields("username")
End If
  rs.Open "DELETE * FROM online WHERE oindex = " + _
                CStr(Index), conn

    Call update_status(2, user, Index)
   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = Now & ": Connected closed for " & ServiceSocket(Index).RemoteHostIP & vbCrLf
Set rs = Nothing
Set rs2 = Nothing
Set conn = Nothing
End Sub

Private Sub ServiceSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        intMax = intMax + 1
        
        'load new socket
        Load ServiceSocket(intMax)
        ServiceSocket(intMax).LocalPort = 0
        ServiceSocket(intMax).Accept requestID
        
        RichTextBox1.SelColor = vbBlue
        RichTextBox1.SelText = Now & ": New connection request from " & ServiceSocket(intMax).RemoteHostIP & vbCrLf
    End If
End Sub

Private Sub ServiceSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strIncoming As String
    Dim result As String
    'Store incoming data into strIncoming
    ServiceSocket(Index).GetData strIncoming
    
    'List of If statements looking at the FIRST word in strIncoming
    If Word(strIncoming, 1) = ".login" Then 'Compare username and pass to database
        result = login_user(Word(strIncoming, 2), Word(strIncoming, 3), Index)
        ServiceSocket(Index).SendData result
          
    ElseIf Word(strIncoming, 1) = ".getonlinebuddies" Then
        Call get_online_buddies(Word(strIncoming, 2), Index)
        
    ElseIf Word(strIncoming, 1) = ".updateStatus" Then 'Updates User status: online, away, etc
        Call update_status(Word(strIncoming, 2), Word(strIncoming, 3), Index)
    
    ElseIf Word(strIncoming, 1) = ".updateBuddy" Then
        Call update_buddy(Word(strIncoming, 2), Word(strIncoming, 3), Word(strIncoming, 4), Index)
        
    ElseIf Word(strIncoming, 1) = ".newBuddy" Then
        Call new_buddy(Word(strIncoming, 2), Word(strIncoming, 3), Index)
        
    ElseIf Word(strIncoming, 1) = ".delBuddy" Then
        Call del_buddy(Word(strIncoming, 2), Word(strIncoming, 3), Index)
        
    ElseIf Word(strIncoming, 1) = ".msg" Then
        Call send_message(Word(strIncoming, 2), Word(strIncoming, 3), strIncoming, Index)
        
    ElseIf Word(strIncoming, 1) = ".define" Or Word(strIncoming, 1) = ".spell" Then
        Call get_DefineSpell(Word(strIncoming, 1), Word(strIncoming, 2), Word(strIncoming, 3), Index)
        
    End If
    
End Sub

Private Function get_DefineSpell(df As String, buddy As String, wordtD As String, Index As Integer)
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=defin.mdb"
 Dim rs As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 Dim sendmsg As String
 Dim i As Integer
 conn.Open sConnString

sendmsg = df & " " & buddy & " " & wordtD

Select Case df
    Case ".define"
        rs.Open "SELECT type, definition FROM words WHERE word = '" + _
                    CStr(wordtD) + "' ", conn
        If Not rs.EOF Then
            i = 0
            sendmsg = sendmsg & " ..//.. " & rs.Fields("type") & " "
            While Not rs.EOF And i < 3
                i = i + 1
                sendmsg = sendmsg & i & ". " & rs.Fields("definition") & " "
                rs.MoveNext
            Wend
        Else
            sendmsg = sendmsg & " ..//.. undefined term"
        End If
        ServiceSocket(Index).SendData sendmsg
    Case ".spell"
        rs.Open "SELECT word FROM words WHERE word = '" + _
                    CStr(wordtD) + "' ", conn
        If Not rs.EOF Then
            sendmsg = sendmsg & " ..//.. " & wordtD & " is spelled correctly."
        Else
            sendmsg = sendmsg & " ..//.. " & wordtD & " appears to be misspelled."
        End If
        ServiceSocket(Index).SendData sendmsg
End Select
            
End Function

Private Function send_message(mfrom As String, mto As String, msg As String, Index As Integer)
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 Dim sendmsg As String
 Dim i As Long
 conn.Open sConnString
    rs.Open "SELECT oindex, ostatus FROM online WHERE username = '" + _
            CStr(mto) + "' ", conn
    If Not rs.EOF Then
        sendmsg = SplitString(msg, "..//..")
        If ServiceSocket(rs.Fields("oindex")).State = sckConnected Then
            ServiceSocket(rs.Fields("oindex")).SendData ".msg " & mfrom & " ..//.. " & sendmsg
            DoEvents: DoEvents
            ServiceSocket(Index).SendData ".msgOK 0"
        Else
            ServiceSocket(Index).SendData ".msgOK 1"
        End If
    Else
        ServiceSocket(Index).SendData ".msgOK 1"
    End If
Set rs = Nothing
Set conn = Nothing
End Function

Private Function del_buddy(user As String, delbuddy As String, Index As Integer)
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
    rs.Open "DELETE * FROM buddies WHERE user = '" + _
            CStr(user) + "' AND buddy = '" + _
            CStr(delbuddy) + "' ", conn
     ServiceSocket(Index).SendData ".statusUpdate 0 "
Set rs = Nothing
Set conn = Nothing
End Function

Private Function new_buddy(user As String, newbuddy As String, Index As Integer)
    
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
    
    rs2.Open "INSERT INTO buddies ([user], [buddy]) VALUES('" + _
        CStr(user) + "', '" + _
        CStr(newbuddy) + "') ", conn
    
  rs.Open "SELECT username, ostatus FROM online WHERE username = '" + _
            CStr(newbuddy) + "' ", conn
  If Not rs.EOF Then
     ServiceSocket(Index).SendData ".statusUpdate 1 " & newbuddy & " " & rs.Fields("ostatus")
  Else
     ServiceSocket(Index).SendData ".statusUpdate 1 " & newbuddy & " 2"
  End If
Set rs = Nothing
Set rs2 = Nothing
Set conn = Nothing
End Function

Private Function update_buddy(user As String, oldbuddy As String, newbuddy As String, Index As Integer)
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
 
  rs2.Open "UPDATE buddies SET buddy = '" + _
            CStr(newbuddy) + "' WHERE user = '" + _
            CStr(user) + "' AND buddy = '" + _
            CStr(oldbuddy) + "' ", conn

  rs.Open "SELECT username, ostatus FROM online WHERE username = '" + _
            CStr(newbuddy) + "' ", conn
  If Not rs.EOF Then
     ServiceSocket(Index).SendData ".statusUpdate 1 " & newbuddy & " " & rs.Fields("ostatus")
  Else
     ServiceSocket(Index).SendData ".statusUpdate 1 " & newbuddy & " 2"
  End If
Set rs = Nothing
Set rs2 = Nothing
Set conn = Nothing
End Function

Private Function login_user(user As String, pass As String, Index As Integer) As String
 Dim i As Integer
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
 
  rs.Open "SELECT username, password FROM users WHERE username = '" + _
            CStr(user) + "' " + _
            " AND password = '" + CStr(pass) + "' ", conn
  If Not rs.EOF Then
    login_user = ".goodlogin" 'Username and Passoword are correct
        'When .goodlogin, add user to ONLINE table
        rs2.Open "INSERT INTO online (username, signontime, oindex) VALUES('" + _
            CStr(user) + "', '" + _
            CStr(Now) + "', '" + _
            CStr(Index) + "') ", conn
  Else
    login_user = ".badlogin" 'Username or Password is wrong
  End If
Set rs = Nothing
Set rs2 = Nothing
Set conn = Nothing
End Function

Private Function get_online_buddies(user As String, Index As Integer)
 Dim OnlineBuddies As String
 Dim status As Integer
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
        'Get user's buddies that are online
        rs.Open "SELECT buddy FROM buddies WHERE user = '" + _
                    CStr(user) + "' ", conn
        While rs.EOF = False
            rs2.Open "SELECT username, ostatus FROM online WHERE username = '" + _
                CStr(rs.Fields("buddy")) + "' ", conn
            If Not rs2.EOF Then
                status = rs2.Fields("ostatus")
            Else
                status = 2
            End If
            OnlineBuddies = OnlineBuddies & " " & rs.Fields("buddy") & status
            rs2.Close
        rs.MoveNext
        Wend
        If OnlineBuddies <> "" Then
            ServiceSocket(Index).SendData ".showonline 1" & OnlineBuddies
        Else
            ServiceSocket(Index).SendData ".showonline 0"
        End If
Set rs = Nothing
Set rs2 = Nothing
Set conn = Nothing
End Function

Private Function update_status(status As Integer, user As String, Index As Integer)
 Dim sConnString  As String
 sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=imdb.mdb"
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim conn As New ADODB.Connection
 conn.Open sConnString
 If ServiceSocket(Index).State = sckConnected Then
        rs2.Open "UPDATE online SET ostatus = " + _
                CStr(status) + " WHERE username = '" + _
                CStr(user) + "' AND oindex = " + _
                CStr(Index), conn
 End If
                
        'Gets the ppl that have user as buddies that are currently online
        rs.Open "SELECT buddy, oindex FROM buddies, online WHERE buddies.buddy = '" + _
                    CStr(user) + "' AND buddies.user = online.username", conn
        While rs.EOF = False
        'Im not too sure about this part of the code...it seems it would work, not sure
        'if its the best code though
            For i = 1 To ServiceSocket.UBound
            testing = ServiceSocket(i).Index
                If ServiceSocket(i).Index = rs.Fields("oindex") Then
                    If ServiceSocket(i).State = sckConnected Then
                        ServiceSocket(i).SendData ".statusUpdate 1" & " " & user & " " & status 'send status of user to its buddies if their online
                    End If
                End If
            Next i
        rs.MoveNext
        Wend
 If ServiceSocket(Index).State = sckConnected Then
    ServiceSocket(Index).SendData ".statusUpdate 0"
 End If
 Set rs = Nothing
 Set rs2 = Nothing
 Set conn = Nothing
End Function
