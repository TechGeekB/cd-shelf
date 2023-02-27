VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFCEC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   8.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFCEC&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFCEC&
         Caption         =   "All Titles with Contents"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFCEC&
         Caption         =   "&All Titles(No contents)"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFCEC&
         Caption         =   "&Current Title Only"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "What do you want to Generate ?"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3120
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
On Error GoTo handle
'Unload frmPrint
'rs.Open "select f_name,given_date from friend where title='" & form1.txttitle.Text & "'",
If Option1.Value = True Then
    If Form1.IsCurrentTitleActive = False Then 'no title is active
        MsgBox "First complete the opeation." & vbCrLf _
        & "Cancel or save the record!", vbInformation + vbOKOnly, "SHAILCdShelf"
        Exit Sub
    End If
    If DataEnvironment1.rsCommand1.State = adStateOpen Then DataEnvironment1.rsCommand1.Close
    DataEnvironment1.Command1 Form1.txttitle.Text
    rptCurrentTitleOnly.Show 1
ElseIf Option2.Value = True Then
    If DataEnvironment1.rsCommand2.State = adStateOpen Then DataEnvironment1.rsCommand2.Close
    DataEnvironment1.command2 'I don't know what mean by this
    If DataEnvironment1.rsCommand2.RecordCount = 0 Then
        MsgBox "No Titles Available.", vbInformation + vbOKOnly, "SHAILCdShelf"
        Exit Sub
    End If
    rptAllTitlesNoContents.Show 1
Else
    If DataEnvironment1.rsCommand3.State = adStateOpen Then DataEnvironment1.rsCommand3.Close
    DataEnvironment1.command3 'I don't know what mean by this
    If DataEnvironment1.rsCommand3.RecordCount = 0 Then
        MsgBox "No Titles Available.", vbInformation + vbOKOnly, "SHAILCdShelf"
        Exit Sub
    End If
    rptAllTitlesWithContents.Show 1
End If

Exit Sub
handle:
    MsgBox Err.Description, vbInformation + vbOKOnly, "SHAILCdShelf"
End Sub

Private Sub Check2_Click()
Check2.Value = False
Unload Me
End Sub

