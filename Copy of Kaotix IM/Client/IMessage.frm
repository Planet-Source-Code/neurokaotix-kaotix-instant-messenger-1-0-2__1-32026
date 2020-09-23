VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form IMessage 
   ClientHeight    =   3900
   ClientLeft      =   7440
   ClientTop       =   4185
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3645
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      ButtonWidth     =   582
      ButtonHeight    =   556
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Big Smile"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bored"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cool"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Dead"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Shocked"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sad"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Angry"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Blush"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Robot"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rolls Eyes"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Smile"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stick Out Tongue"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Wink"
            ImageIndex      =   14
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin RichTextLib.RichTextBox showmsg 
      Height          =   1575
      Left            =   60
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"IMessage.frx":1272
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
   Begin RichTextLib.RichTextBox typemsg 
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   2475
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   1085
      _Version        =   393217
      ScrollBars      =   2
      MaxLength       =   2048
      Appearance      =   0
      TextRTF         =   $"IMessage.frx":12EE
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1680
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":158E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":16A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":1AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbFonts 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   2100
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Decreace Font Size"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Increase Font Size"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbFonts 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":1F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":21AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":2410
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":2668
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":27F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":2A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":2CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":2F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3195
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3406
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":358D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3712
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3980
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3BF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":3E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":40B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "IMessage.frx":430B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/2048"
      Height          =   195
      Left            =   4720
      TabIndex        =   7
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   4080
      Picture         =   "IMessage.frx":454F
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image6 
      Height          =   450
      Left            =   3960
      Picture         =   "IMessage.frx":4AFE
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image5 
      Height          =   450
      Left            =   60
      Picture         =   "IMessage.frx":50B1
      Top             =   3135
      Width           =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5880
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   2760
      Picture         =   "IMessage.frx":5660
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   2640
      Picture         =   "IMessage.frx":5C0C
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   2520
      Picture         =   "IMessage.frx":61BD
      Top             =   4320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   4620
      Picture         =   "IMessage.frx":676E
      Top             =   3140
      Width           =   1050
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuFont 
      Caption         =   "&Format"
      Begin VB.Menu mnuFontBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFontItalic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFontUnderline 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuFontSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFontPT 
            Caption         =   "8 pt"
            Index           =   0
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "10 pt"
            Index           =   1
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "12 pt"
            Index           =   2
         End
         Begin VB.Menu mnuFontPT 
            Caption         =   "14 pt"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&People"
   End
End
Attribute VB_Name = "IMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWord As String

Private Sub cmbFonts_Click()
    typemsg.SelFontName = cmbFonts.Text
    typemsg.SetFocus
End Sub




Private Sub Form_Load()
Dim i As Integer
    'bAllowScroll = True
    'Call SetHook(showmsg.hwnd, True)
    For i = 1 To Screen.FontCount
        cmbFonts.AddItem Screen.Fonts(i)
    Next i
    cmbFonts.RemoveItem (0)
    cmbFonts.SelText = "Verdana"
End Sub

Private Sub Image1_Click()
typemsg.Text = Replace(typemsg.Text, vbCrLf, "")
If typemsg.Text <> "" And Len(typemsg) > 0 Then
    showmsg.SelStart = Len(showmsg.Text)
    showmsg.SelBold = True
    showmsg.SelColor = vbRed
    showmsg.SelText = YourSN & ": "
    
    
    showmsg.SelStart = Len(showmsg.Text)
    showmsg.SelBold = False
    showmsg.SelColor = vbBlack
    showmsg.SelText = typemsg.Text & vbCrLf
  
    sndPlaySound App.Path + "\sounds\imsend.wav", 1
    Client.Winsock1.SendData ".msg " & YourSN & " " & Word(Me.Caption, 1) & " ..//.. " & typemsg.Text
    Client.WaitFor (".msgOK")
    typemsg.Text = ""
End If
typemsg.Text = Replace(typemsg.Text, vbCrLf, "")
typemsg.SetFocus
typemsg.Text = Replace(typemsg.Text, vbCrLf, "")
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = Image2.Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.Picture = Image3.Picture
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Picture = Image6.Picture
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image5.Picture = Image7.Picture
End Sub





'Private Sub Form_Unload(Cancel As Integer)
    'Call SetHook(showmsg.hwnd, False)
'End Sub

Private Sub mnuFontBold_Click()

If typemsg.SelBold = True Then
   typemsg.SelBold = False
   tbFonts.Buttons(1).Value = tbrUnpressed
Else
   typemsg.SelBold = True
   tbFonts.Buttons(1).Value = tbrPressed
End If

End Sub

Private Sub mnuFontItalic_Click()

If typemsg.SelItalic = True Then
   typemsg.SelItalic = False
   tbFonts.Buttons(2).Value = tbrUnpressed
Else
   typemsg.SelItalic = True
   tbFonts.Buttons(2).Value = tbrPressed
End If

End Sub

Private Sub mnuFontUnderline_Click()

If typemsg.SelUnderline = True Then
   typemsg.SelUnderline = False
   tbFonts.Buttons(3).Value = tbrUnpressed
Else
   typemsg.SelUnderline = True
   tbFonts.Buttons(3).Value = tbrPressed
End If

End Sub

Private Sub mnuFontPT_Click(Index As Integer)
   typemsg.SelFontSize = Word(mnuFontPT(Index).Caption, 1)
End Sub


Private Function is_chars(x As String) As Boolean
Dim i As Integer
Dim flag As Integer
For i = 1 To Len(x)
    If (Asc(UCase(Mid(x, i, 1))) >= vbKeyA And Asc(UCase(Mid(x, i, 1))) <= vbKeyZ) Or Mid(x, i, 1) = " " Then
        flag = 0
    Else
        flag = 1
        Exit For
        is_chars = False
    End If
Next i

If flag = 0 Then
    is_chars = True
End If
    
End Function



Private Sub tbFonts_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
    Case 1
        mnuFontBold_Click
    Case 2
        mnuFontItalic_Click
    Case 3
        mnuFontUnderline_Click
    Case 5
        If typemsg.SelFontSize > 8 Then
            typemsg.SelFontSize = typemsg.SelFontSize - 2
        End If
    Case 6
        If typemsg.SelFontSize < 14 Then
            typemsg.SelFontSize = typemsg.SelFontSize + 2
        End If
End Select

End Sub

Private Sub typemsg_Change()
Label1 = Len(typemsg.Text) & "/2048"
End Sub

Private Sub typemsg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
    typemsg_Click
End If
End Sub

Private Sub typemsg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call Image1_Click
End If
End Sub

Private Sub typemsg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo HandleError
If Button = vbRightButton Then
        If typemsg.SelText <> "" And is_chars(Replace(Trim(typemsg.SelText), vbCrLf, "")) = True Then
            dWord = Replace(Trim(typemsg.SelText), vbCrLf, "")
        Else
            dWord = "- Please highlight a word."
        End If
        mnuSpell.Caption = "Spellcheck " & dWord
        mnuDefine.Caption = "Define " & dWord
    
    If typemsg.SelBold = True Then
       mnuFontBold.Checked = True
    Else
       mnuFontBold.Checked = False
    End If
    
    If typemsg.SelItalic = True Then
       mnuFontItalic.Checked = True
    Else
       mnuFontItalic.Checked = False
    End If
    
    If typemsg.SelUnderline = True Then
       mnuFontUnderline.Checked = True
    Else
       mnuFontUnderline.Checked = False
    End If
    PopupMenu mnuFont
End If

HandleError:
    'MsgBox Err.Number & " - " & Err.Description, vbOKOnly
End Sub

Private Sub typemsg_Click()
If cmbFonts.Text <> typemsg.SelFontName Then
    cmbFonts.Text = ""
    cmbFonts.SelText = typemsg.SelFontName
End If

    If typemsg.SelBold = True Then
       tbFonts.Buttons(1).Value = tbrPressed
    Else
       tbFonts.Buttons(1).Value = tbrUnpressed
       
    End If
    
    If typemsg.SelItalic = True Then
       tbFonts.Buttons(2).Value = tbrPressed
    Else
       tbFonts.Buttons(2).Value = tbrUnpressed
    End If
    
    If typemsg.SelUnderline = True Then
       tbFonts.Buttons(3).Value = tbrPressed
    Else
       tbFonts.Buttons(3).Value = tbrUnpressed
    End If
End Sub

