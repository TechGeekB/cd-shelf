VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00FFFCEC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hi"
   ClientHeight    =   3915
   ClientLeft      =   1875
   ClientTop       =   1155
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   8.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   1680
   End
   Begin VB.CommandButton OKbutton 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lbllink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Mail me : shai_ban007@yahoo.co.in"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   150
      MouseIcon       =   "Dialog.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2880
      Width           =   3990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   6480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFCEC&
      Caption         =   "  I heartly invites your suggestion and complaints( because no s/w is bug      free) about this  s/w. THANKS for using this s/w."
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   6375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "(0141) - 2501294    (on Jan 2003)"
      Height          =   225
      Left            =   1200
      TabIndex        =   8
      Top             =   2460
      Width           =   2820
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Phone :"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Jaipur - 302015"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Mahesh Nagar,"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   1725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "C-284,"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFCEC&
      Caption         =   "Prepared by :  SHAILESH BANSAL"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "SHAILCdShelf"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   45
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1050
      Left            =   -15
      TabIndex        =   1
      Top             =   120
      Width           =   6465
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

Option Explicit

Private Sub CancelButton_Click()

End Sub

Private Sub lblLink_Click()
On Error GoTo handle
'to_mail = "mailto: " & Text1.Text
ShellExecute hwnd, "open", "mailto:shai_ban007@yahoo.co.in", vbNullString, vbNullString, SW_SHOW
Exit Sub
handle:
MsgBox Err.Number & " : " & Err.Description, vbOKOnly + vbCritical, "!!!"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lbllink
    .ForeColor = vbRed
    .Font.Underline = False
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
' in case user unload the 'Hi' form by 'close' button instead of 'OK' button
Form1.Enabled = True

End Sub

Private Sub lbllink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lbllink
    .ForeColor = vbBlue
    .Font.Underline = True
End With
End Sub

Private Sub OKButton_Click()
Unload Me
Form1.Enabled = True

End Sub

Private Sub Timer1_Timer()
If Label1.Visible = True Then
    Label1.Visible = False
Else
    Label1.Visible = True
End If
End Sub
