VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_known_hosts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Known Hosts"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6510
   Icon            =   "form_known_hosts.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   6510
   Begin MSComctlLib.ListView ListView_known_hosts 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Shared Files"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Kbytes Shared"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWcascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile Horizontaly"
      End
      Begin VB.Menu mnuWtileV 
         Caption         =   "Tile Verticaly"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
End
Attribute VB_Name = "Form_known_hosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    On Error Resume Next
    Dim cpt
    For cpt = 0 To UBound(known_hosts) - 1
        With known_hosts(cpt)
            add_item_known_host .ip, .port, .nb_shared_files, .nb_shared_kbytes
        End With
    Next cpt
End Sub

Private Sub Form_Activate()
    activate_tab (Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    remove_tab (Me.hWnd)
End Sub

Private Sub ListView_known_hosts_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    With ListView_known_hosts
        .SortKey = ColumnHeader.index - 1
        
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub mnuWCascade_Click()
   Form_main.Arrange vbCascade
End Sub

Private Sub mnuWTileH_Click()
   Form_main.Arrange vbTileHorizontal
End Sub

Private Sub mnuWTileV_Click()
   Form_main.Arrange vbTileVertical
End Sub

Private Sub mnuWArrange_Click()
   Form_main.Arrange vbArrangeIcons
End Sub

