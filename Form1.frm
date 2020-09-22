VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Get Header v1.0 (Don Steele)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   30
      Left            =   120
      Top             =   2250
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   120
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5280
      Top             =   2280
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3525A&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Url goes here"
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3525A&
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5280
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      URL             =   "http://"
      RequestTimeout  =   6000
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   30
      Width           =   3105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "don4jc@saintmail.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.starsplace.net/teck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000080&
      Caption         =   "    OK"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "   About"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3525A&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "   Stop"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3525A&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "  Get It"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3525A&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'*   Get header V1.0 By Donald Steele                         *
'*                                                            *
'*  ever want to know what server software                    *
'*  a website is running? well here is your answer            *
'*  just type in a url and like magic you will                *
'*  receive the http header information that reveals          *
'*  all kinds of useful info like the os and server software  *
'*  being used                                                *
'*                                                            *
'* send any comments to don4jc@saintmail.net :)               *
'*                                                            *
'**************************************************************
'*            GET A REAL OS GET LINUX WWW.LINUX.COM           *
'**************************************************************
Dim running

Option Explicit


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Enum T_WindowStyle
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum

Dim crispy
Dim critter

Private Sub Form_Load()
crispy = 0
critter = 0
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
'ok lets check to see if were connected to the site
Select Case State
Case 3
'whoo hoo we are connected now lets turn on the timer to grab the data
Timer1.Enabled = True
End Select
End Sub
'ever get sick of buttons?
Private Sub Label1_Click()
On Error GoTo e
Timer1.Enabled = False
'close any open connections so we don't error out
Inet1.Cancel
'check for a return so we don't have to click the button every time
'it can be annoying after your 20th site
Text1.Text = ""
Label7.Caption = "Connecting...."
'clear the buggy inet control so it will be ready for a new site
Inet1.URL = ""
Inet1.Cancel
'open the url so we can get the nuggets of information we need
Inet1.OpenURL Text2.Text, 1

Text2.Enabled = False
GoTo last:
e:
Text1.Text = "there was an error please try your query again"
last:
End Sub

Private Sub Label2_Click()
'kill the timer so we won't try and proccess any data
Timer1.Enabled = False
'kill the active connection
Inet1.Cancel
'reset the control
Inet1.URL = ""
Text2.Enabled = True
Label7.Caption = ""
End Sub

Private Sub Label3_Click()
Timer2.Enabled = True
End Sub

Private Sub Label4_Click()
Text1.Top = 240
Text2.Left = 360
Text1.ForeColor = &HF3525A
Text1.BackColor = &H400000
Label4.Visible = False
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Text1.Text = ""
End Sub

Private Sub Label5_Click()
OpenInternet Me, "http://www.starsplace.net/teck", Normal

End Sub

Private Sub Label6_Click()
OpenInternet Me, "mailto:don4jc@saintmail.net", Normal
End Sub

Private Sub Text2_KeyUp(keycode As Integer, Shift As Integer)
'close any open connections so we don't error out
Inet1.Cancel
'check for a return so we don't have to click the button every time
'it can be annoying after your 20th site
If keycode = "13" Then
Text1.Text = ""
Label7.Caption = "Connecting...."
'clear the buggy inet control so it will be ready for a new site
Inet1.URL = ""
'open the url so we can get the nuggets of information we need
On Error Resume Next
Text2.Enabled = False
Inet1.OpenURL Text2.Text, 1


End If
End Sub

Private Sub Timer1_Timer()
DoEvents
On Error GoTo e
Dim h
'lets make sure we still don't have an open connection
'if we do the inet control will freek out and just continue to get the site
'for about 15 min (very very bad)
If Inet1.StillExecuting = False Then
'ok were all set lets read teh header information into a var.
h = Inet1.GetHeader
'let's make sure we didn't get a blank field
 If Not h = "" Then
 'dump what we got to a text box
Text1.Text = Text1.Text & h
'kill the timer so we don't have a loop going on here
Timer1.Enabled = False
Label7.Caption = ""
'set the input box for our next site
Text2.Enabled = True
'this is added to force the inet control to reset for some reason it would not reset
'correctly ????
Inet1.Cancel
GoTo last
End If
End If
GoTo last
e:
Text1.Text = "there was an error please submit your query again"

last:
End Sub
'that's all people
Private Sub Timer2_Timer()
critter = 0
crispy = crispy + 1
Select Case crispy
Case 20
Timer3.Enabled = True
Timer2.Enabled = False
crispy = 0
End Select
DoEvents
End Sub

Private Sub Timer3_Timer()
If Not Text2.Left < -5000 Then
Text2.Left = Text2.Left - 500
Else
If Not Text1.Top > 840 Then
Text1.Top = Text1.Top + 50
Else: critter = critter + 1
End If
End If
If critter = 1 Then Label1.Visible = False
If critter = 2 Then Label2.Visible = False
If critter = 3 Then Label3.Visible = False
If critter = 4 Then Text1.BackColor = &H80&
If critter = 5 Then Label4.Visible = True
If critter = 6 Then Text1.ForeColor = &H8080FF
If critter = 7 Then Text1.Text = "This program is designed to retreave http header information which will give you information on what type of server the website is running as well as the os that it is being run on if you have any questions or comments please contact me"
If critter = 8 Then Timer3.Enabled = False
End Sub
Public Sub OpenInternet(Parent As Form, URL As String, _
    WindowStyle As T_WindowStyle)
    ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub


Private Sub Timer4_Timer()
DoEvents
End Sub
