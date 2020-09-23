VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   13035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14205
   Icon            =   "form_options.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13035
   ScaleWidth      =   14205
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   3600
      TabIndex        =   7
      Top             =   6120
      Width           =   4695
      Begin VB.TextBox txtnbpongforping 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2160
         TabIndex        =   106
         Text            =   "2"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtmaxbogus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   95
         Text            =   "20"
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtmaxquery 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   94
         Text            =   "40"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtmaxping 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   93
         Text            =   "50"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtpushvaltime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   89
         Text            =   "30"
         Top             =   3000
         Width           =   735
      End
      Begin VB.CheckBox chkmycomputerinfo 
         Caption         =   "Don't send pong with my computer informations"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtmaxinittl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Text            =   "15"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtttl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkforwardping 
         Caption         =   "Don't forward ping"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkforwardquery 
         Caption         =   "Don't forward query"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkclip2reflector 
         Caption         =   "Don't forward ping an query on incoming connections"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox chkspyhit 
         Caption         =   "Spy all files passing throw queryhits"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label43 
         Caption         =   "Number of pong for a ping"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label37 
         Caption         =   "Max bogus per minute"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "Max query per minute"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label35 
         Caption         =   "Max ping per minute"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "Retry push every                    s"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Trash packets with initial TTL higher than"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "TTL"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "( max =255 but should not be upper 10 )"
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2895
      Left            =   8400
      TabIndex        =   82
      Top             =   9720
      Width           =   5415
      Begin VB.TextBox txtsizetrafficinfo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   104
         Text            =   "99"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtsizeknownfiles 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   103
         Text            =   "10000"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtsizeroutingtable 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   102
         Text            =   "10000"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtsizedescriptors 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   101
         Text            =   "10000"
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "Traffic info"
         Height          =   255
         Left            =   1560
         TabIndex        =   100
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label41 
         Caption         =   "Known Files"
         Height          =   255
         Left            =   1560
         TabIndex        =   99
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label40 
         Caption         =   "The routing table"
         Height          =   255
         Left            =   1560
         TabIndex        =   98
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "Your descriptors"
         Height          =   255
         Left            =   1560
         TabIndex        =   97
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label38 
         Caption         =   "Array size for"
         Height          =   255
         Left            =   240
         TabIndex        =   96
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   $"form_options.frx":0CCA
         Height          =   855
         Left            =   240
         TabIndex        =   87
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame6 
      Height          =   4095
      Left            =   7080
      TabIndex        =   53
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtminconnection 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   79
         Text            =   "0"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtand 
         Height          =   285
         Left            =   720
         TabIndex        =   58
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtor 
         Height          =   285
         Left            =   720
         TabIndex        =   57
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtnot 
         Height          =   285
         Left            =   720
         TabIndex        =   56
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox txtminfilesize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   55
         Text            =   "0"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtmaxfilesize 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   54
         Text            =   "0"
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "kb/s"
         Height          =   255
         Left            =   4200
         TabIndex        =   81
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label30 
         Caption         =   "Host should have a connection at least at"
         Height          =   375
         Left            =   600
         TabIndex        =   80
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label24 
         Caption         =   "All of the following words must be presents ( separator is ; ) :"
         Height          =   255
         Left            =   600
         TabIndex        =   63
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label23 
         Caption         =   "At least one of the followning words must be present :"
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label22 
         Caption         =   "All of the following words musn't  be present"
         Height          =   255
         Left            =   600
         TabIndex        =   61
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label21 
         Caption         =   "Min file size in kb :"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Max file size in kb  ( 0 if no limit ) :"
         Height          =   255
         Left            =   600
         TabIndex        =   59
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label25 
         Caption         =   $"form_options.frx":0DD1
         Height          =   735
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   51
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1320
      TabIndex        =   50
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   8400
      TabIndex        =   25
      Top             =   5760
      Width           =   5535
      Begin VB.TextBox txtminsharedfiles 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   52
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtminsharedkb 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   49
         Text            =   "0"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtmaxhosts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   32
         Text            =   "8"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtdumyip 
         Height          =   1335
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "form_options.frx":0E7B
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtminhosts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Text            =   "5"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtgnutellahosts 
         Height          =   1335
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "form_options.frx":0EB2
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label19 
         Caption         =   "They must share at least                         kb,                            files"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   4815
      End
      Begin VB.Label Label8 
         Caption         =   "Don't try direct connection with the following IP (push will be send)"
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Connect at least to                   hosts and at max to"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "Known Gnutella hosts  IP"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   240
      TabIndex        =   12
      Top             =   6600
      Width           =   4815
      Begin VB.TextBox txtretrydown 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   86
         Text            =   "30"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox chkloggooddown 
         Caption         =   "Log good downloader"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2880
         Width           =   1935
      End
      Begin VB.CheckBox chklogbaddown 
         Caption         =   "Log bad downloaders"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtnbmultipartdown 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Text            =   "3"
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkallowup 
         Caption         =   "Allow upload (you need to check this if you want to upload)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtmaxdown 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Text            =   "15"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtmaxup 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Text            =   "5"
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "Retry download every                    s"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "(default limeware max upload per IP : 3)"
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Number of parts for multi-parts downloads"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Max simultaneous downloads"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Max simultaneous uploads"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   7560
      TabIndex        =   67
      Top             =   4320
      Width           =   3375
      Begin VB.TextBox txtmyminsearch 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   73
         Text            =   "3"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtotherminsearch 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   70
         Text            =   "4"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label29 
         Caption         =   "chars"
         Height          =   255
         Left            =   2640
         TabIndex        =   72
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label28 
         Caption         =   "chars"
         Height          =   255
         Left            =   2640
         TabIndex        =   71
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label27 
         Caption         =   "My min search length"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Other min search length"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   6255
      Begin VB.TextBox txtmykbshared 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Text            =   "1000"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtmynbfiles 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "100"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkfilesharesimu 
         Caption         =   "File Sharing Simulation (because some soft don't allow freeloaders)"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label14 
         Caption         =   $"form_options.frx":0EE5
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   5055
      End
      Begin VB.Label Label16 
         Caption         =   "Size of shared files (in kb)"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Number of Files"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Height          =   3135
      Left            =   240
      TabIndex        =   33
      Top             =   360
      Width           =   6255
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   78
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   77
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   76
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   75
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chksubdirectories 
         Caption         =   "Incude sub directories of Shared directories"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox chklaunchserver 
         Caption         =   "Launch server (avoid other people to send push)"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   45
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkautoadddownload 
         Caption         =   "Automatically add downloaded files to shared files"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtmyip 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   41
         Text            =   "0.0.0.0"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkforceip 
         Caption         =   " Force IP"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5040
         TabIndex        =   36
         Text            =   "6346"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtnetconnectionspeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Text            =   "56"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   34
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label13 
         Caption         =   "Intermediate directory"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Download directory"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "On port"
         Height          =   255
         Left            =   4320
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Network connection speed in kb/s"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Shared directories"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6015
      Left            =   120
      TabIndex        =   74
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10610
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Your Computer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Hosts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Down/Up"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filters"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Network"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Coyotella"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBrowse_Click(index As Integer)
    Dim directory As String
    directory = GetFolderName()
    If directory <> "" Then Me.txtDirectory(index) = directory & "\"
End Sub

Private Sub Form_Load()
    Me.Height = 6960
    Me.Width = 6780
    Me.Frame2.Visible = False
    Me.Frame2.Top = 640
    Me.Frame2.Left = 960
    Me.Frame3.Visible = False
    Me.Frame3.Top = 1440
    Me.Frame3.Left = 960
    Me.Frame4.Visible = False
    Me.Frame4.Top = 1200
    Me.Frame4.Left = 600
    Me.Frame6.Visible = False
    Me.Frame6.Top = 480
    Me.Frame6.Left = 600
    Me.Frame7.Visible = False
    Me.Frame7.Top = 4680
    Me.Frame7.Left = 1560
    Me.Frame8.Visible = False
    Me.Frame8.Top = 1560
    Me.Frame8.Left = 600
    
    
    Call get_options_values
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdapply_Click()
    Call fill_options
End Sub

Private Sub get_options_values()
    On Error Resume Next
    Dim strtmp          As String
    Dim tmparray        As Variant
    Dim cpt             As Integer
    
    If launch_server Then Me.chklaunchserver.value = vbChecked
    Me.txtPort.Text = my_port
    Me.txtnetconnectionspeed.Text = my_speed
    If force_ip Then Me.chkforceip.value = vbChecked
    Me.txtmyip.Text = forced_ip
    
    strtmp = ""
    For cpt = 0 To UBound(myshared_directories)
        If CStr(myshared_directories(cpt)) <> "" Then
            strtmp = strtmp & myshared_directories(cpt) & ";"
        End If
    Next cpt
    Me.txtDirectory(0).Text = strtmp
    Me.txtDirectory(1).Text = my_download_directory
    Me.txtDirectory(2).Text = my_incomplete_directory
    If include_sub_dir Then Me.chksubdirectories.value = vbChecked
    If auto_add_down_to_shared_files Then Me.chkautoadddownload.value = vbChecked
    tmparray = Split(strknown_gnutella_server, ";")
    Me.txtgnutellahosts.Text = ""
    For cpt = 0 To UBound(tmparray)
        If CStr(tmparray(cpt)) <> "" Then
            Me.txtgnutellahosts.Text = Me.txtgnutellahosts.Text & CStr(tmparray(cpt)) & vbCrLf
        End If
    Next cpt
    strtmp = ""
    For cpt = 0 To UBound(dummy_ip)
        If dummy_ip(cpt) <> "" Then
            strtmp = strtmp & dummy_ip(cpt) & vbCrLf
        End If
    Next cpt
    Me.txtdumyip.Text = strtmp
    Me.txtminhosts.Text = min_dialing_hosts
    Me.txtmaxhosts.Text = max_dialing_hosts
    Me.txtminsharedfiles.Text = min_host_shared_file
    Me.txtminsharedkb.Text = min_host_shared_kb
    Me.txtminconnection.Text = my_min_speed
    Me.txtmaxdown.Text = max_download
    If allow_upload Then Me.chkallowup.value = vbChecked
    Me.txtmaxup.Text = max_upload
    Me.txtttl.Text = my_ttl
    If spy_all_query_hits Then Me.chkspyhit.value = vbChecked
    If forward_on_outgoing_only Then Me.chkclip2reflector.value = vbChecked
    Me.txtmaxinittl.Text = max_initial_ttl
    If Not forward_ping Then Me.chkforwardping.value = vbChecked
    If Not forward_query Then Me.chkforwardquery.value = vbChecked
    If sharing_simulation Then Me.chkfilesharesimu.value = vbChecked
    Me.txtmynbfiles.Text = simulation_nb_files
    Me.txtmykbshared.Text = simulation_size
    Me.txtminfilesize.Text = known_files_restrictions_array(0)
    Me.txtmaxfilesize.Text = known_files_restrictions_array(1)
    Me.txtand.Text = opt_all_and_words
    Me.txtor.Text = opt_all_or_words
    Me.txtnot.Text = opt_all_not_words
    Me.txtotherminsearch.Text = other_min_search
    Me.txtmyminsearch.Text = my_min_search_length
    If Not send_my_computers_info Then Me.chkmycomputerinfo.value = vbChecked
    
    Me.txtretrydown = retry_down_on_busy_server_every
    Me.txtsizedescriptors.Text = my_descriptors_ID_max_size
    Me.txtsizeroutingtable.Text = routing_table_max_size
    Me.txtsizeknownfiles.Text = known_files_max_size
    Me.txtsizetrafficinfo.Text = traffic_info.array_size
    Me.txtmaxbogus.Text = mymax_bogus_per_minute
    Me.txtmaxquery = mymax_query_per_minute
    Me.txtmaxping.Text = mymax_ping_per_minute
    Me.txtpushvaltime.Text = mypush_validity_time
    Me.txtnbpongforping.Text = nb_of_pong_for_a_ping
    If log_bad_downloaders Then Me.chklogbaddown.value = vbChecked
    If log_good_downloaders Then Me.chkloggooddown.value = vbChecked

End Sub


Private Sub fill_options()
    On Error Resume Next
    Dim tmparray            As Variant
    Dim array_size          As Integer
    Dim cpt                 As Integer
    Dim new_size            As Integer
    Dim delta               As Integer
    Dim strtmp              As String
    Dim lResult As Long

    If Me.chklaunchserver.value = vbChecked Then
        If launch_server = False Then Call make_server
        launch_server = True
    Else
        If launch_server = True Then
            lResult = MessageBox(Me.hWnd, "Do you want really close server ?" & vbCrLf & _
            "You won't be able to download files on firewalled hosts" & vbCrLf & _
            "If you want to reactivate server you may wait few seconds", Form_main.Caption, vbOKCancel)

            If lResult = vbOK Then
                Form_main.socket(0).Close
                launch_server = False
            End If
        End If
    End If
    If my_port <> Me.txtPort.Text Then
        my_port = Me.txtPort.Text
        'close server if already launched and create a new one
        Form_main.socket(0).Close
        Call make_server
    End If
    
    my_speed = Me.txtnetconnectionspeed.Text
    forced_ip = Me.txtmyip.Text
    If Me.chkforceip.value = vbChecked Then
        If Not force_ip Then
            my_ip = forced_ip
            force_ip = True
        End If
    Else
        If force_ip Then

            If Form_main.socket(1).state = sckConnected Then
                my_ip = Form_main.socket(1).LocalIP
            Else
                my_ip = "127.0.0.1"
            End If
            force_ip = False
        End If
    End If
    
    
    tmparray = Split(Me.txtDirectory(0).Text, ";")
    array_size = UBound(tmparray)
    new_size = array_size
    delta = 0
    If array_size > -1 Then
        ReDim myshared_directories(array_size)
        For cpt = 0 To array_size
            strtmp = Trim$(CStr(tmparray(cpt)))
            If strtmp <> "" Then
                If Right$(strtmp, 1) = "\" Then
                    myshared_directories(cpt - delta) = strtmp
                Else
                    myshared_directories(cpt - delta) = strtmp & "\"
                End If
            Else
                delta = delta + 1
                new_size = new_size - 1
                ReDim Preserve myshared_directories(new_size)
            End If
        Next cpt
    End If

    strtmp = Trim$(Me.txtDirectory(1).Text)
    If Len(strtmp) > 0 Then
        If Right$(strtmp, 1) = "\" Then
            my_download_directory = Me.txtDirectory(1).Text
        Else
            my_download_directory = Me.txtDirectory(1).Text & "\"
        End If
    Else
        my_download_directory = ""
    End If
    
    strtmp = Trim$(Me.txtDirectory(2).Text)
    If Len(strtmp) > 0 Then
        If Right$(strtmp, 1) = "\" Then
            my_incomplete_directory = Me.txtDirectory(2).Text
        Else
            my_incomplete_directory = Me.txtDirectory(2).Text & "\"
        End If
    Else
        my_incomplete_directory = ""
    End If

    If Not is_folder_existing(myshared_directories(0)) _
        Or Not is_folder_existing(my_download_directory) _
        Or Not is_folder_existing(my_incomplete_directory) _
        Then
            lResult = MessageBox(Me.hWnd, "Some directories are false" & vbCrLf & "Please update directories", Form_main.Caption, vbExclamation)
            Exit Sub
    End If
    ' refresh shared files
    Call share_files

    If Me.chksubdirectories.value = vbChecked Then
        include_sub_dir = True
    Else
        include_sub_dir = False
    End If
    If Me.chkautoadddownload.value = vbChecked Then
        auto_add_down_to_shared_files = True
    Else
        auto_add_down_to_shared_files = False
    End If
    
    tmparray = Split(Me.txtgnutellahosts.Text, vbCrLf)
    array_size = UBound(tmparray)
    strknown_gnutella_server = ""
    For cpt = 0 To array_size
        If tmparray(cpt) <> "" Then strknown_gnutella_server = strknown_gnutella_server & tmparray(cpt) & ";"
    Next cpt

    ' update known_gnutella_server
    fill_known_gnutella_server

    tmparray = Split(Me.txtdumyip.Text, vbCrLf)
    array_size = UBound(tmparray)
    ReDim dummy_ip(array_size + 1)
    new_size = array_size
    delta = 0
    If array_size > -1 Then
        ReDim dummy_ip(array_size + 1)
        For cpt = 0 To array_size
            If tmparray(cpt) <> "" Then
                dummy_ip(cpt - delta) = CStr(tmparray(cpt))
            Else
                new_size = new_size - 1
                delta = delta + 1
                ReDim Preserve dummy_ip(new_size)
            End If
        Next cpt
    End If

    min_dialing_hosts = Me.txtminhosts.Text
    max_dialing_hosts = Me.txtmaxhosts.Text
    min_host_shared_file = Me.txtminsharedfiles.Text
    min_host_shared_kb = Me.txtminsharedkb.Text
    my_min_speed = Me.txtminconnection.Text
    max_download = Me.txtmaxdown.Text
    
    
    If Me.chkallowup.value = vbChecked Then
        allow_upload = True
    Else
        allow_upload = False
    End If
    max_upload = Me.txtmaxup.Text
    my_ttl = Me.txtttl.Text
    If Me.chkspyhit.value = vbChecked Then
        spy_all_query_hits = True
    Else
        spy_all_query_hits = False
    End If
    If Me.chkclip2reflector.value = vbChecked Then
        forward_on_outgoing_only = True
    Else
        forward_on_outgoing_only = False
    End If
    max_initial_ttl = Me.txtmaxinittl.Text
    If Not Me.chkforwardping.value = vbChecked Then
        forward_ping = True
    Else
        forward_ping = False
    End If
    If Not Me.chkforwardquery.value = vbChecked Then
        forward_query = True
    Else
        forward_query = False
    End If
    If Me.chkfilesharesimu.value = vbChecked Then
        sharing_simulation = True
    Else
        sharing_simulation = False
    End If
    simulation_nb_files = Me.txtmynbfiles.Text
    simulation_size = Me.txtmykbshared.Text
    opt_all_and_words = Me.txtand.Text
    opt_all_or_words = Me.txtor.Text
    opt_all_not_words = Me.txtnot.Text
    
    other_min_search = Me.txtotherminsearch.Text
    my_min_search_length = Me.txtmyminsearch.Text

    If Me.chkmycomputerinfo.value = vbChecked Then
        send_my_computers_info = False
    Else
        send_my_computers_info = True
    End If
    
    
    
    
    
    retry_down_on_busy_server_every = Val(Me.txtretrydown)
    my_descriptors_ID_max_size = Val(Me.txtsizedescriptors.Text)
    routing_table_max_size = Val(Me.txtsizeroutingtable.Text)
    known_files_max_size = Val(Me.txtsizeknownfiles.Text)
    traffic_info.array_size = Val(Me.txtsizetrafficinfo.Text)
    mymax_bogus_per_minute = Val(Me.txtmaxbogus.Text)
    mymax_query_per_minute = Val(Me.txtmaxquery)
    mymax_ping_per_minute = Val(Me.txtmaxping.Text)
    mypush_validity_time = Val(Me.txtpushvaltime.Text)
    nb_of_pong_for_a_ping = Val(Me.txtnbpongforping.Text)
    If Me.chklogbaddown.value = vbChecked Then
        log_bad_downloaders = True
    Else
        log_bad_downloaders = False
    End If
    
    If Me.chkloggooddown.value = vbChecked Then
        log_good_downloaders = True
    Else
        log_good_downloaders = False
    End If
    
    'fill known_files_restrictions_array
    fill_known_files_restrictions_array Me.txtminfilesize.Text, Me.txtmaxfilesize.Text, opt_all_and_words, opt_all_or_words, opt_all_not_words

    
    Call update_ini
    
    Unload Me
End Sub

Private Sub TabStrip1_Click()
    Select Case Me.TabStrip1.SelectedItem.index
        Case 1
            Me.Frame1.Visible = True
            Me.Frame2.Visible = False
            Me.Frame3.Visible = False
            Me.Frame4.Visible = False
            Me.Frame5.Visible = True
            Me.Frame6.Visible = False
            Me.Frame7.Visible = False
            Me.Frame8.Visible = False
        Case 2
            Me.Frame1.Visible = False
            Me.Frame2.Visible = False
            Me.Frame3.Visible = False
            Me.Frame4.Visible = True
            Me.Frame5.Visible = False
            Me.Frame6.Visible = False
            Me.Frame7.Visible = False
            Me.Frame8.Visible = False
        Case 3
            Me.Frame1.Visible = False
            Me.Frame2.Visible = False
            Me.Frame3.Visible = True
            Me.Frame4.Visible = False
            Me.Frame5.Visible = False
            Me.Frame6.Visible = False
            Me.Frame7.Visible = False
            Me.Frame8.Visible = False
        Case 4
            Me.Frame1.Visible = False
            Me.Frame2.Visible = False
            Me.Frame3.Visible = False
            Me.Frame4.Visible = False
            Me.Frame5.Visible = False
            Me.Frame6.Visible = True
            Me.Frame7.Visible = True
            Me.Frame8.Visible = False
        Case 5
            Me.Frame1.Visible = False
            Me.Frame2.Visible = True
            Me.Frame3.Visible = False
            Me.Frame4.Visible = False
            Me.Frame5.Visible = False
            Me.Frame6.Visible = False
            Me.Frame7.Visible = False
            Me.Frame8.Visible = False
        Case 6
            Me.Frame1.Visible = False
            Me.Frame2.Visible = False
            Me.Frame3.Visible = False
            Me.Frame4.Visible = False
            Me.Frame5.Visible = False
            Me.Frame6.Visible = False
            Me.Frame7.Visible = False
            Me.Frame8.Visible = True
    End Select
End Sub
