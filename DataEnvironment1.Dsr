VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   _ExtentX        =   5054
   _ExtentY        =   4974
   FolderFlags     =   1
   TypeLibGuid     =   "{BD1ACE79-66A7-4BDE-804F-54F7C2447ABA}"
   TypeInfoGuid    =   "{2AC56946-DF5C-4175-9F83-A9E9DB3C96D6}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "Command1"
      CommDispId      =   1004
      RsDispId        =   1020
      CommandText     =   "SELECT Title AS tit, Description AS content, `Date` AS dt, Category AS category FROM T_main WHERE (title = ?)"
      ActiveConnectionName=   "Connection1"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "tit"
         Caption         =   "tit"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "content"
         Caption         =   "content"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "dt"
         Caption         =   "dt"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "category"
         Caption         =   "category"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "?"
         UserName        =   "Tit"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Command2"
      CommDispId      =   1005
      RsDispId        =   1025
      CommandText     =   "T_main"
      ActiveConnectionName=   "Connection1"
      CommandType     =   2
      dbObjectType    =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "Category"
         Caption         =   "Category"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "To_friend"
         Caption         =   "To_friend"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Title"
         Caption         =   "Title"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Command3"
      CommDispId      =   1006
      RsDispId        =   1029
      CommandText     =   "T_main"
      ActiveConnectionName=   "Connection1"
      CommandType     =   2
      dbObjectType    =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "Category"
         Caption         =   "Category"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   7
         Name            =   "Date"
         Caption         =   "Date"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   536870910
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "To_friend"
         Caption         =   "To_friend"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Title"
         Caption         =   "Title"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()
    Dim conn_string As String
    
    'conn_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\cd.dat;Persist Security Info=False"
    conn_string = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\cd.dat" & ";Persist Security Info=False;Jet OLEDB:Database Password=pleasedon'topenit"

    ' Now establish the connction for the DataEnvironment1
    DataEnvironment1.Connection1.ConnectionString = conn_string

End Sub
