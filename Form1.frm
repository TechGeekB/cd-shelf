VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFCEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHAILCdShelf"
   ClientHeight    =   7830
   ClientLeft      =   465
   ClientTop       =   525
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   11.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H008080FF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "A&bout"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4320
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   3570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLocate 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Locate in the content"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6930
      Width           =   1095
   End
   Begin VB.ListBox lstfriend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   5250
      Width           =   2175
   End
   Begin VB.ListBox lstTitles 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3270
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Double click the title to see its full details."
      Top             =   430
      Width           =   2175
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   290
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPSearchDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   5490
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   6.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   255
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   255
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   19595267
      CurrentDate     =   37617
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFCEC&
      Caption         =   "Search Criteria :"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   240
      TabIndex        =   28
      Top             =   5010
      Width           =   8655
      Begin VB.CheckBox chkCase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFCEC&
         Caption         =   "Case Se&nsitive"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   1680
      End
      Begin VB.ListBox lstSearchTitles 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2190
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Select the title to see its full details."
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton OptContents 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFCEC&
         Caption         =   "&Keyword from contents"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton OptCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFCEC&
         Caption         =   "Categ&ory"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton OptSearchDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFCEC&
         Caption         =   "Da&te"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         MaskColor       =   &H00FF0000&
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Sea&rch"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdNewSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Ne&w Search"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtSearchContents 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1125
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   18
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ComboBox cmbSearchCategory 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "Form1.frx":0442
         Left            =   1920
         List            =   "Form1.frx":0467
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         X1              =   6120
         X2              =   6120
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFCEC&
         Caption         =   "Sea&rch Results(titles found)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6330
         TabIndex        =   29
         Top             =   150
         Width           =   2220
      End
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "This will delete your current CD title & its description."
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdModify 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1000
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1000
   End
   Begin VB.ComboBox Cmbfriend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      ItemData        =   "Form1.frx":04C7
      Left            =   6840
      List            =   "Form1.frx":04D1
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1410
      Width           =   1815
   End
   Begin VB.TextBox txtFriend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   930
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   5
      Top             =   2130
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPFriend 
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   3450
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   6.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483646
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   19595267
      CurrentDate     =   37617
   End
   Begin VB.ComboBox cmbCategory 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      ItemData        =   "Form1.frx":04DE
      Left            =   3960
      List            =   "Form1.frx":0503
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   290
      Width           =   1815
   End
   Begin VB.TextBox txtContents 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1050
      Width           =   6375
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   290
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483646
      CalendarTrailingForeColor=   16711680
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   19595267
      CurrentDate     =   37617
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFCEC&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   240
      TabIndex        =   36
      Top             =   90
      Width           =   8535
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFCEC&
         Caption         =   "Date"
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
         Left            =   5880
         TabIndex        =   39
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFCEC&
         Caption         =   "Category"
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
         Left            =   2880
         TabIndex        =   38
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFCEC&
         Caption         =   "Title"
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
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFCEC&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9120
      TabIndex        =   35
      Top             =   3735
      Width           =   1935
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "By:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   10080
      TabIndex        =   34
      Top             =   6930
      Width           =   270
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Shailesh Bansal"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   9480
      TabIndex        =   33
      Top             =   7170
      Width           =   1620
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "CD's given to friend"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9105
      TabIndex        =   32
      Top             =   4965
      Width           =   1965
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "Existing CD-titles"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   9210
      TabIndex        =   31
      Top             =   135
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000F&
      Height          =   2895
      Left            =   6720
      Top             =   1050
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFCEC&
      Caption         =   "CD's Contents"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   30
      Top             =   810
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFCEC&
      Caption         =   "Is given to friend"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFCEC&
      Caption         =   "Giving Date"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFCEC&
      Caption         =   "Friend (s) Name"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   26
      Top             =   1890
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5400
      TabIndex        =   24
      Top             =   330
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   9120
      Shape           =   2  'Oval
      Top             =   6690
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
   ByVal pszFile As String, _
   ByVal uCommand As Long, _
   dwData As Any) As Long

Const LB_SETHORIZONTALEXTENT = &H194
Public IsCurrentTitleActive As Boolean
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
'we can write two statement instead of single
'Dim db As ADODB.Connection
'Set db = New ADODB.Connection
'or we can use
'Dim conn1
'Set conn1 = CreateObject("ADODB.Connection") As Object
'--------------------------------------------------------
Dim modify As Boolean
Dim tit As String
Dim PIndex As Long
Dim IsToFri As Boolean
Dim pos As Double
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long


Sub ThreeDForm(frmForm As Form)
    Const cPi = 3.1415926
    Dim intLineWidth As Integer
    intLineWidth = 5
    ' 'save scale mode
    Dim intSaveScaleMode As Integer
    intSaveScaleMode = frmForm.ScaleMode
    frmForm.ScaleMode = 3
    Dim intScaleWidth As Integer
    Dim intScaleHeight As Integer
    intScaleWidth = frmForm.ScaleWidth
    intScaleHeight = frmForm.ScaleHeight
    ' 'clear form
    frmForm.Cls
    ' 'draw white lines
    frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
    frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
    ' 'draw grey lines
    frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
    frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
    ' 'draw triangles(actually circles) at corners
    Dim intCircleWidth As Integer
    intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
    frmForm.FillStyle = 0
    frmForm.FillColor = QBColor(15)
    frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), _
    -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
    frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), _
    -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
    ' 'draw black frame
    frmForm.Line (0, intScaleHeight)-(0, 0), 0
    frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
    frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmForm.ScaleMode = intSaveScaleMode
End Sub

Sub check(s As String, List1 As ListBox)
    If X < TextWidth(s & "  ") Then
    X = TextWidth(s & "  ")
End If
If ScaleMode = vbTwips Then
    X = X / Screen.TwipsPerPixelX      ' if twips change to pixels
    SendMessageByNum List1.hwnd, LB_SETHORIZONTALEXTENT, X, 0
End If

End Sub


Function CheckValidation()
    Dim reply As Byte
    Dim i As Double
    CheckValidation = True
    If Trim(txttitle.Text) = "" Then
        MsgBox "Please give the title of the CD.", vbOKOnly + vbInformation, "SHAILCdShelf"
        txttitle.Text = ""
        txttitle.SetFocus
        CheckValidation = False
        Exit Function
    ElseIf Len(txttitle.Text) > 50 Then
        MsgBox "This title is too long to store." & Chr(10) _
        & "Please make it small.", vbOKOnly + vbInformation, "SHAILCdShelf"
        txttitle.SetFocus
        SendKeys "{home}+{end}"
        CheckValidation = False
        Exit Function
    ElseIf cmbCategory.ListIndex = -1 Then
        MsgBox "Please select the category of the CD.", vbOKOnly + vbInformation, "SHAILCdShelf"
        cmbCategory.SetFocus
        SendKeys "{f4}"
        CheckValidation = False
        Exit Function
    End If
    If Trim(txtContents.Text) = "" Then
        reply = MsgBox("Are you sure about the contents of the CD.", vbYesNo + vbInformation, "SHAILCdShelf")
        If reply = vbNo Then
            txtContents.Text = ""
            txtContents.SetFocus
            CheckValidation = False
            Exit Function
        End If
    End If
    If Cmbfriend.ListIndex = 1 Then
        If Trim(txtFriend.Text) = "" Then
            MsgBox "Please give the friend name.", vbOKOnly + vbInformation, "SHAILCdShelf"
            txtFriend.SetFocus
            CheckValidation = False
            Exit Function
        ElseIf Len(txtFriend.Text) > 100 Then
            MsgBox "The friend name is too long to store." & Chr(10) _
            & "Please make it small.", vbOKOnly + vbInformation, "SHAILCdShelf"
            txtFriend.Text = ""
            txtFriend.SetFocus
            CheckValidation = False
            Exit Function

        End If
    End If
    
End Function

Sub PopulateListsWithTitles()
Dim sql_string As String
sql_string = "select title, to_friend from t_main   "
rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
If rs.EOF <> True Then
    rs.MoveLast
    rs.MoveFirst
End If
If rs.RecordCount > 0 Then
    While rs.EOF = False
        lstTitles.AddItem rs.Fields("title")
        
        If rs.Fields("to_friend") = True Then
            lstfriend.AddItem rs.Fields("title") 'populate in given to friend list
        End If
        rs.MoveNext
    Wend
    If lstfriend.ListCount = 0 Then
        lstfriend.AddItem "No CD to friend"
    End If
    rs.Close 'close the recordset
    'means at least one title exist, then unlock the controls
     RecordFound 'call procedure
    'set the controls with the first fields value
    sql_string = "select * from t_main where title='" & lstTitles.List(0) & "'"
    rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
    txttitle.Text = rs.Fields("title")
    cmbCategory.Text = rs.Fields("category")
    DTPDate.Value = rs.Fields("date")
    txtContents.Text = rs.Fields("description")
    If rs.Fields("to_friend") = False Then
        Cmbfriend.ListIndex = 0 'select no
        Cmbfriend.Locked = True
        'cmbFriend.ListIndex = 0 'select no
        'txtFriend.Text = ""
        
        'txtFriend.Locked = True
        'DTPFriend.Enabled = False
        rs.Close
    Else
        rs.Close
        sql_string = "select * from friend where title='" & lstTitles.List(0) & "'"
        rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
        Cmbfriend.ListIndex = 1
        Cmbfriend.Locked = True
        txtFriend.Text = rs.Fields("f_name")
        DTPFriend.Value = rs.Fields("given_date")
        rs.Close
        
    End If
Else 'no record found
    NoRecord 'call procedure
End If
'now check for horizontal scroll bar in list box
    txt = ""
    For i = 0 To lstTitles.ListCount - 1
        If Len(lstTitles.List(i)) > Len(txt) Then
            txt = lstTitles.List(i)
        End If
    Next
    check ((txt)), lstTitles
    '-----------------------------------------------
    txt = ""
    For i = 0 To lstfriend.ListCount - 1
        If Len(lstfriend.List(i)) > Len(txt) Then
            txt = lstfriend.List(i)
        End If
    Next
    check ((txt)), lstfriend

End Sub
Sub RecordFound()
    txttitle.Locked = True
    cmbCategory.Locked = True
    DTPDate.Enabled = False
    txtContents.Locked = True 'cd's contents
    'cmbFriend.Enabled = True 'is given to friend
    'cmbFriend.Locked = True
    'txtFriend.Locked = False   'friend's name
    'DTPFriend.Enabled = False 'given date
    cmdAdd.Enabled = True 'add
    cmdModify.Enabled = True 'modify
    cmdDelete.Enabled = True 'delele
    cmdSave.Enabled = False 'save
    cmdCancel.Enabled = False 'cancel
    

End Sub
Sub NoRecord()
    lstTitles.Clear
    lstfriend.Clear
    lstTitles.AddItem "No Title Available"
    lstfriend.AddItem "No CD to friend"
    txttitle.Text = "" 'Title
    txttitle.Locked = True
    cmbCategory.ListIndex = -1 'category
    cmbCategory.Locked = True
    txtContents.Text = ""
    txtContents.Locked = True 'cd's contents
    Cmbfriend.ListIndex = 0 ' is given to friend
    Cmbfriend.Locked = True
    txtFriend.Text = "" 'friend's name
    txtFriend.Locked = True
    DTPDate.Enabled = False 'Date
    DTPFriend.Value = Date 'given date
    DTPFriend.Enabled = False
    cmdModify.Enabled = False 'modify
    cmdDelete.Enabled = False 'delete
    cmdSave.Enabled = False 'save
    cmdCancel.Enabled = False 'cancel
    cmdAdd.Enabled = True
    cmdAdd.SetFocus 'add
    
End Sub

Sub DisableButtons()
cmdAdd.Enabled = False 'add
cmdModify.Enabled = False 'modify
cmdDelete.Enabled = False 'delete
End Sub
Sub EnableButtons()
cmdAdd.Enabled = True 'add
cmdModify.Enabled = True 'modify
cmdDelete.Enabled = True 'delete

End Sub

Private Sub cmbCategory_GotFocus()
'    If cmbCategory.Locked = False Then
        'SendKeys "{F4}"
    'End If
End Sub

Private Sub cmbFriend_Click()

If Cmbfriend.ListIndex = 0 Then
    txtFriend.Text = ""
    txtFriend.Enabled = False 'friend name
    DTPFriend.Enabled = False ' giving date
Else
    txtFriend.Enabled = True
    DTPFriend.Enabled = True
    txtFriend.SetFocus
        
End If

End Sub

Private Sub Cmbfriend_GotFocus()
    'If Cmbfriend.Locked = False Then
        'SendKeys "{f4}"
    'End If
End Sub

Private Sub cmbSearchCategory_Click()
    cmdSearch.SetFocus
End Sub


Private Sub cmdAdd_Click()
'If lstTitles.ListCount >= 5 Then
    'MsgBox "This is trial version(limited to 5 records)." & vbCrLf _
        ' & "To get full version see the help file" & vbCrLf _
         '& "or contact to me.", vbOKOnly + vbInformation, "SHAILCdShelf"
    
    'Call cmdHelp_Click
    'SendKeys "{down}", True
    'Exit Sub
'End If
    txttitle.Text = "" 'title
    txttitle.Locked = False
    cmbCategory.ListIndex = -1   'category
    cmbCategory.Locked = False
    DTPDate.Value = Date
    DTPDate.Enabled = True
    txtContents.Text = "Enter CD`s Contents here" 'contents
    txtContents.Locked = False
    Cmbfriend.ListIndex = 0 ' given to friend=NO
    Cmbfriend.Locked = False
    txtFriend.Locked = False
    DTPFriend.Value = Date
    cmdAdd.Enabled = False 'add
    cmdModify.Enabled = False 'modify
    cmdDelete.Enabled = False 'delete
    cmdSave.Enabled = True 'save
    cmdCancel.Enabled = True 'cancel
    txttitle.SetFocus
End Sub

Private Sub cmdHelp_Click()
    On Error GoTo handle:

    'Dim CMDFlags As Long
    'CMDFlags = &H3&
    'CommonDialog1.HelpCommand = CMDFlags
    'CommonDialog1.HelpFile = App.path & "\cdshelfhelp.hlp"
        'CommonDialog1.HelpFile = App.HelpFile

    'CommonDialog1.HelpContext = Val(App.path & "\cdshelfhelp.hlp")
    'CommonDialog1.HelpKey = HLPKey.Text
    'CommonDialog1.ShowHelp
    'Call HtmlHelp(0, App.path & "\shailcdshelf.chm", HH_HELP_CONTEXT, &HF)
    Dim hand As Long
    Dim winname As String
    winname = String(20, 0)
    winname = "SHAILCdShelf.chm"
    hand = FindWindow(vbNullString, winname)
    If hand = 0 Then ' means help file is not open
        Shell "hh.exe" & " " & App.path & "\shailcdshelf.chm", vbNormalFocus
    End If
    
    Exit Sub
handle:
    MsgBox "All help is included in the file 'SHAILCdShelf.chm' file." & Chr(10) _
        & "Please refer this file from the application folder.", vbOKOnly + vbInformation, "SHAILCdShelf"
        
End Sub

Private Sub cmdLocate_Click()
    
    'Static pos As Double
    Dim reply As Byte
    
    If lstTitles.List(0) = "No Title Available" Then
        MsgBox "There is not exists any title.", vbOKOnly + vbInformation, "SHAILCdShelf"
        cmdAdd.SetFocus
        Exit Sub
    End If
    If chkCase.Value = 1 Then
        pos = InStr(pos + 1, txtContents.Text, txtSearchContents.Text, vbBinaryCompare)
    Else
        pos = InStr(pos + 1, txtContents.Text, txtSearchContents.Text, vbTextCompare)
    End If

    If pos > 0 Then
    txtContents.SetFocus
    txtContents.SelStart = pos - 1
    txtContents.SelLength = Len(txtSearchContents.Text)
    'Else
        'reply = MsgBox("End of text is found. Do you want to search from start again ?", vbYesNo + vbInformation, "SHAILCdShelf")
        'If reply = vbYes Then
            'pos = 0
        'End If
    Else
        reply = MsgBox("End of text found." & Chr(10) _
        & "Want to start again ?", vbYesNo + vbInformation, "SHAILCdShelf")
        If reply = vbYes Then
            cmdLocate_Click
        End If
    End If

End Sub

Private Sub cmdModify_Click()
    If lstTitles.List(0) <> "No Title Available" Then
    'store the current record
        modify = True
        tit = txttitle.Text
        For i = 0 To lstTitles.ListCount - 1
            If lstTitles.List(i) = txttitle Then
                PIndex = i
                Exit For
            End If
        Next
        IsToFri = Cmbfriend.ListIndex
        'now current title has been saved
        txttitle.Locked = False
        cmbCategory.Locked = False
        DTPDate.Enabled = True
        txtContents.Locked = False
        Cmbfriend.Locked = False
        If Cmbfriend.Text = "Yes" Then
            txtFriend.Locked = False
            DTPFriend.Enabled = True
        End If
        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        
    End If
End Sub

Private Sub cmdDelete_Click()
On Error GoTo handle
Dim reply As Byte
Dim sql_string As String
Dim title As String
Dim to_friend As Boolean

If lstTitles.List(0) <> "No Title Available" Then
    reply = MsgBox("Want to delete the current title !", vbExclamation + vbYesNo + vbDefaultButton2, "SHAILCdShelf")
    If reply = vbYes Then
        db.BeginTrans
        title = txttitle.Text
        to_friend = Cmbfriend.ListIndex '0 means false, 1 means true
        sql_string = "select * from t_main where title='" & title & "'"
        rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
        rs.Delete
        rs.Update
        rs.Close
        'now remove this title from lsttitles
        For i = lstTitles.ListCount - 1 To 0 Step -1
            If lstTitles.List(i) = title Then
                lstTitles.RemoveItem i
                Exit For
            End If
        Next
        If lstTitles.ListCount = 0 Then
            lstTitles.AddItem "No Title Available"
        End If
        If to_friend = True Then
            sql_string = "select * from friend where title='" & title & "'"
            rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
            rs.Delete
            rs.Update
            rs.Close
            For i = lstfriend.ListCount - 1 To 0 Step -1
                If lstfriend.List(i) = title Then
                    lstfriend.RemoveItem i
                    Exit For
                End If
            Next
            If lstfriend.ListCount = 0 Then
                lstfriend.AddItem "No CD to friend"
            End If
        End If
        db.CommitTrans
        If lstTitles.List(0) <> "No Title Available" Then
            'means at least one title exist
            lstTitles.ListIndex = 0
            lstTitles_DblClick
            If lstTitles.List(0) = "No Title Available" Then
                cmdAdd.SetFocus
            End If
        Else 'No titles available
            NoRecord
        End If
        MsgBox "Title is deleted successfully", vbOKOnly + vbInformation, "SHAILCdshelf"
    End If
End If
If lstTitles.List(0) <> "No Title Available" Then
    lblTotal.Caption = "Total " & lstTitles.ListCount & " Titles"
Else
    lblTotal.Caption = "Total 0 Titles"
End If
Exit Sub

handle:
    MsgBox Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, "!!!"
    db.RollbackTrans
    
End Sub

Private Sub cmdCancel_Click()
    'if lstTitles has atleast one title, then select it and show its record
    ' otherwise call no record
    
    If lstTitles.List(0) <> "No Title Available" Then
        If modify = True Then
            reply = MsgBox("Want to discard the changes to current title.", vbYesNo + vbInformation + vbDefaultButton2, "SHAILCdShelf")
            If reply = vbNo Then
                Exit Sub
            Else
                modify = False
            End If
        End If
        'means at least one title exist
        lstTitles.ListIndex = 0
        lstTitles_DblClick
    Else 'No titles available
        NoRecord
    End If
End Sub

Private Sub cmdAbout_Click()
Form1.Enabled = False
Load About
About.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNewSearch_Click()
    Dim X
    X = MsgBox("This will clear your current search", vbInformation + vbOKCancel + vbDefaultButton2, "SHAILCdShelf")
    If X = vbOK Then
        OptSearchDate.Value = True 'date of creation
        DTPSearchDate.Value = Date
        txtSearchContents.Text = "" 'keyword from contents
        cmbSearchCategory.ListIndex = -1
        lstSearchTitles.Clear 'search results
        lstSearchTitles.AddItem "No Title found"
        
    End If
End Sub

Private Sub cmdPrint_Click()
If cmdModify.Enabled = False Then IsCurrentTitleActive = False
If cmdModify.Enabled = True Then IsCurrentTitleActive = True
frmPrint.Show 1
End Sub

Private Sub cmdSave_Click()
    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim sql_string As String
    Dim valid As Boolean
    On Error GoTo handle
    valid = CheckValidation
    If valid = False Then Exit Sub
    If modify = True Then 'means modify has been called
        For i = 0 To lstTitles.ListCount - 1
            If i <> PIndex And LCase(lstTitles.List(i)) = LCase(txttitle.Text) Then
                MsgBox "This title already exists," & Chr(10) _
                & "please change it.", vbOKOnly + vbInformation, "SHAILCdShelf"
                txttitle.SetFocus
                txttitle.SelStart = 0
                txttitle.SelLength = Len(txttitle.Text)
                Exit Sub
            End If
        Next
                
        db.BeginTrans
        For i = lstTitles.ListCount - 1 To 0 Step -1
            If lstTitles.List(i) = tit Then
                lstTitles.RemoveItem i
                Exit For
            End If
        Next
        'now delete the record from t_main table
        sql_string = "select * from t_main where title='" & tit & "'"
        rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
        rs.Delete
        rs.Update
        rs.Close
        'if this cd is given to friend, then delete this record also from friend
        If IsToFri = True Then
        'delete from lstfriend list box
            For i = lstfriend.ListCount - 1 To 0 Step -1
                If lstfriend.List(i) = tit Then
                    lstfriend.RemoveItem i
                    Exit For
                End If
            Next
            sql_string = "select * from friend where title='" & tit & "'"
            rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
            rs.Delete
            rs.Update
            rs.Close
        End If
        modify = False
        db.CommitTrans
    End If
    'now save the modified or new record
        For i = 0 To lstTitles.ListCount - 1
            If LCase(lstTitles.List(i)) = LCase(txttitle.Text) Then
                MsgBox "This title already exist." & Chr(10) _
                & "Please change it.", vbOKOnly + vbInformation, "SHAILCdShelf"
                txttitle.SetFocus
                txttitle.SelStart = 0
                txttitle.SelLength = Len(txttitle.Text)
                Exit Sub
            End If
        Next
        'Now save the record
        db.BeginTrans
        rs.Open "T_main", db, adOpenKeyset, adLockOptimistic
        rs.AddNew
        rs.Fields("Title") = txttitle.Text
        rs.Fields("Category") = cmbCategory.Text
        rs.Fields("Date") = DTPDate.Value
        If Trim(txtContents.Text) <> "" Then
            rs.Fields("Description") = txtContents.Text
        Else
            rs.Fields("Description") = ""
        End If
        If Cmbfriend.ListIndex = 1 Then 'yes
            rs.Fields("To_friend") = 1
        Else 'No
            rs.Fields("To_friend") = 0
        End If
        rs.Update
        rs.Close
        If Cmbfriend.ListIndex = 1 Then 'No
            rs.Open "friend", db, adOpenKeyset, adLockOptimistic
            rs.AddNew
            rs.Fields("title") = txttitle.Text
            rs.Fields("F_name") = txtFriend.Text
            rs.Fields("given_date") = DTPFriend.Value
            rs.Update
            rs.Close
        End If
        'Now record is saved successfully
        db.CommitTrans
        
        If lstTitles.List(0) = "No Title Available" Then
            lstTitles.Clear
        End If
        lstTitles.AddItem txttitle.Text
        check txttitle.Text, lstTitles
        If Cmbfriend.ListIndex = 1 Then
            If lstfriend.List(0) = "No CD to friend" Then
                lstfriend.Clear
            End If
            lstfriend.AddItem txttitle.Text
            check txttitle.Text, lstfriend
        End If
        lstTitles.ListIndex = 0
        lstTitles_DblClick
        MsgBox "Title is saved successfully.", vbOKOnly + vbInformation, "SHAILCdShelf"
    If lstTitles.List(0) <> "No Title Available" Then
        lblTotal.Caption = "Total " & lstTitles.ListCount & " Titles"
    Else
        lblTotal.Caption = "Total 0 Titles"
    End If
Exit Sub
handle:
    MsgBox Err.Number & " :- " & Err.Description, vbOKOnly + vbInformation, "SHAILCdShelf"
    db.RollbackTrans
End Sub

Private Sub cmdSearch_Click()
    Dim sql_string
    Dim rs As New ADODB.Recordset
    
    lstSearchTitles.Clear
    If OptSearchDate.Value = vbTrue Then
        sql_string = "select title from t_main where date=#" & DTPSearchDate.Value & "#"
        rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            lstSearchTitles.AddItem "No Title found"
        Else
            While rs.EOF = False
                lstSearchTitles.AddItem rs.Fields("title")
                rs.MoveNext
            Wend
        End If
        
    ElseIf OptCategory.Value = vbTrue Then
        If cmbSearchCategory.ListIndex = -1 Then
            MsgBox "Please select the search category.", vbOKOnly + vbInformation, "SHAILCdShelf"
            cmbSearchCategory.SetFocus
            SendKeys "{f4}"
            Exit Sub
        End If
        sql_string = "select title from t_main where category='" & cmbSearchCategory.Text & "'"
        rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            lstSearchTitles.AddItem "No Title found"
        Else
            While rs.EOF = False
                lstSearchTitles.AddItem rs.Fields("title")
                rs.MoveNext
            Wend
        End If
    Else 'optcontents is selected
        sql_string = "select title from t_main where description like '%" & CStr(txtSearchContents.Text) & "%'"
        rs.Open sql_string, db, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            lstSearchTitles.AddItem "No Title found"
        Else
            While rs.EOF = False
                lstSearchTitles.AddItem rs.Fields("title")
                rs.MoveNext
            Wend
        End If
    End If
    
    'Now check for horizontal scroll bar
    txt = ""
    For i = 0 To lstSearchTitles.ListCount - 1
        If Len(lstSearchTitles.List(i)) > Len(txt) Then
            txt = lstSearchTitles.List(i)
        End If
    Next
    check ((txt)), lstSearchTitles





End Sub


Private Sub DTPDate_GotFocus()
    'If DTPDate.Enabled = True Then
        'SendKeys "{F4}"
    'End If
End Sub

Private Sub DTPFriend_GotFocus()
'    If DTPFriend.Enabled = True Then
      '  SendKeys "{F4}"
    'End If
End Sub

Private Sub DTPSearchDate_Change()
    cmdSearch.SetFocus
End Sub

Private Sub DTPSearchDate_CloseUp()
    cmdSearch.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo handle
'Dim fsys As Object
Dim pwd As String
Dim path As String
If App.PrevInstance Then End
'Set fsys = CreateObject("scripting.filesystemobject")
CommonDialog1.CancelError = True
If Dir(App.path & "\" & "cd.dat") <> "" Then
    path = App.path & "\cd.dat"
Else
MsgBox "The file cd.dat didn't exist at its " & App.path & " path," & Chr(10) _
& "so please specify its path.", vbOKOnly + vbExclamation, "SHAILCdShelf"
    CommonDialog1.FileName = "cd.dat"
    CommonDialog1.Filter = "cd file|*.dat"
    CommonDialog1.InitDir = App.path
    CommonDialog1.Flags = cdlOFNFileMustExist And cdlOFNPathMustExist
    CommonDialog1.ShowOpen
    path = CommonDialog1.FileName
    
End If
'------open connection-----------------------------------------
Set db = New ADODB.Connection
'conn_string = "Driver={Microsoft Access Driver (*.mdb)};" & _
        '"Dbq=" & App.Path & "\cd.dat;" & _
        '"Uid=Admin; Pwd="
        'conn_string = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                   '"DBQ=;" & _
                   '"DefaultDir=" & path & ";" & _
                   '"UID=admin;PWD=;"
'conn_string = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & path
'conn_string = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" _
'& "Data Source=" & path & ";Jet OLEDB:Database Password=;"
conn_string = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Persist Security Info=False;Jet OLEDB:Database Password=pleasedon'topenit"
db.Open conn_string
'db.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & path & ";UserId=admin;Password=" & pwd & ";"

'Form1.Show
DTPDate.Value = Date 'date of creation
Me.Show
'==================== for lstTitles and lstSearchTitles  ======================
PopulateListsWithTitles 'populate lstTitles and lstSearchTitles
'If lstTitles.ListCount = 0 Then
If lstTitles.List(0) <> "No Title Available" Then
    lblTotal.Caption = "Total " & lstTitles.ListCount & " Titles"
Else
    lblTotal.Caption = "Total 0 Titles"
End If
'Else
    'lblTotal.Caption = "Total " & lstTitles.ListCount - 1 & " Titles"
'End If
lstSearchTitles.AddItem "No Title found"

ThreeDForm Form1
'code for accessing the registry values
'gets = GetSetting("SHAILCdShelf", "Sec", "SNO")
'MsgBox gets




Exit Sub
handle:
'3343=unrecognized database format
'32755=cancel selected
'3024=file not found
'If Err.Number = 3024 Or Err.Number = 32755 Or Err.Number = 3343 Then  'file not found
    MsgBox Err.Number & " : " & Err.Description, vbOKOnly + vbCritical, "!!!"

    Unload Form1
'Else
    'Resume Next
'End If


End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set db = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then ThreeDForm Form1
End Sub

Private Sub lstSearchTitles_Click()
    Dim reply As Byte
    Dim rs As New ADODB.Recordset
    Dim sql_string As String
    
    If lstSearchTitles.Text <> "No Title found" Then
        pos = 0 'clear the current Locate the Content
        If modify = True Then
            reply = MsgBox("Want to discard the changes to current title.", vbYesNo + vbInformation + vbDefaultButton2, "SHAILCdShelf")
            If reply = vbNo Then
                Exit Sub
            Else
                modify = False
            End If
        End If
        RecordFound 'call procedure
        'set the controls with the first fields value
        sql_string = "select * from t_main where title='" & lstSearchTitles.List(lstSearchTitles.ListIndex) & "'"
        rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
        txttitle.Text = rs.Fields("title")
        cmbCategory.Text = rs.Fields("category")
        DTPDate.Value = rs.Fields("date")
        txtContents.Text = rs.Fields("description")
        lstSearchTitles.SetFocus
        If rs.Fields("to_friend") = False Then
            Cmbfriend.ListIndex = 0 'select no
            Cmbfriend.Locked = True
            rs.Close
        Else
            rs.Close
            sql_string = "select * from friend where title='" & lstSearchTitles.List(lstSearchTitles.ListIndex) & "'"
            rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
            Cmbfriend.ListIndex = 1
            txtFriend.Locked = True
            DTPFriend.Enabled = False
            Cmbfriend.Locked = True
            txtFriend.Text = rs.Fields("f_name")
            DTPFriend.Value = rs.Fields("given_date")
            lstSearchTitles.SetFocus
            rs.Close
        End If
    End If
    
    
  
End Sub

Private Sub lstTitles_DblClick()
    Dim reply As Byte
    Dim rs As New ADODB.Recordset
    Dim sql_string As String
    
    If lstTitles.Text <> "No Title Available" Then
        If modify = True Then
            reply = MsgBox("Want to discard the changes to current title.", vbYesNo + vbInformation + vbDefaultButton2, "SHAILCdShelf")
            If reply = vbNo Then
                Exit Sub
            Else
                modify = False
            End If
        End If
        RecordFound 'call procedure
        'set the controls with the first fields value
        sql_string = "select * from t_main where title='" & lstTitles.List(lstTitles.ListIndex) & "'"
        rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
        txttitle.Text = rs.Fields("title")
        cmbCategory.Text = rs.Fields("category")
        DTPDate.Value = rs.Fields("date")
        txtContents.Text = rs.Fields("description")
        lstTitles.SetFocus
        If rs.Fields("to_friend") = False Then
            Cmbfriend.ListIndex = 0 'select no
            Cmbfriend.Locked = True
            rs.Close
        Else
            rs.Close
            sql_string = "select * from friend where title='" & lstTitles.List(lstTitles.ListIndex) & "'"
            rs.Open sql_string, db, adOpenForwardOnly, adLockOptimistic
            Cmbfriend.ListIndex = 1
            txtFriend.Locked = True
            DTPFriend.Enabled = False
            Cmbfriend.Locked = True
            txtFriend.Text = rs.Fields("f_name")
            DTPFriend.Value = rs.Fields("given_date")
            lstTitles.SetFocus
            rs.Close
        End If
    End If
    
    
    
    
End Sub

Private Sub optSearchDate_Click()
If OptSearchDate.Value = True Then
    OptCategory.Value = False 'category
    optCategory_Click
    OptContents.Value = False 'keyword
    optContents_Click
    DTPSearchDate.Enabled = True
    DTPSearchDate.SetFocus
    
Else
    
    DTPSearchDate.Enabled = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
    OptSearchDate.Value = False 'date
    OptCategory.Value = False 'category
    OptContents.Value = False 'keyword
End If
End Sub

Private Sub optCategory_Click()
If OptCategory.Value = True Then
    OptSearchDate.Value = False 'date
    optSearchDate_Click
    'optCategory.Value = False 'category
    OptContents.Value = False 'keyword
    optContents_Click
    cmbSearchCategory.Enabled = True
    cmbSearchCategory.SetFocus
    SendKeys "{F4}"
Else
    cmbSearchCategory.ListIndex = -1
    cmbSearchCategory.Enabled = False
End If
End Sub

Private Sub optContents_Click()
If OptContents.Value = True Then
    chkCase.Enabled = True
    OptSearchDate.Value = False 'date
    optSearchDate_Click
    OptCategory.Value = False 'category
    optCategory_Click
    'optContents.Value = False 'keyword
    txtSearchContents.Enabled = True
    txtSearchContents.SetFocus
    txtSearchContents.SelStart = 0
    txtSearchContents.SelLength = Len(txtSearchContents.Text)
Else
    chkCase.Enabled = False
    chkCase.Value = 0
    txtSearchContents.Enabled = False
    
End If
End Sub




Private Sub txtContents_GotFocus()
    If Len(txtContents.Text) <= 50 Then
        txtContents.SelStart = 0
        txtContents.SelLength = Len(txtContents.Text)
    End If
End Sub

Private Sub txtSearchContents_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSearch_Click
    End If
End Sub

Private Sub txttitle_Change()
    pos = 0 'clear the current Locate the Content
End Sub

Private Sub txttitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = Asc("`")
    End If
End Sub

Private Sub txtContents_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = Asc("`")
    End If
End Sub

Private Sub txtFriend_Change()
    If KeyAscii = Asc("'") Then
        KeyAscii = Asc("`")
    End If
End Sub

Private Sub txtSearchContents_Change()
    If KeyAscii = Asc("'") Then
        KeyAscii = Asc("`")
    End If
End Sub
