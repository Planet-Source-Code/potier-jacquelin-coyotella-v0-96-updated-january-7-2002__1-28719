VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_shared_files 
   Caption         =   "Your Shared Files"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6495
   Icon            =   "form_shared_files.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdRefresh 
      Height          =   495
      Left            =   120
      Picture         =   "form_shared_files.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Refresh"
      Top             =   0
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Full path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile Horizontaly"
      End
      Begin VB.Menu mnuWTileV 
         Caption         =   "Tile Verticaly"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
End
Attribute VB_Name = "Form_shared_files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Dim num_tab     As Integer
    num_tab = Form_main.TabStrip.Tabs.Count
    Form_main.CoolBar1.Bands(2).Visible = True
    Form_main.TabStrip.Tabs.Add num_tab + 1, , "Shared Files"
    Form_main.TabStrip.Tabs(num_tab + 1).Tag = Me.hWnd
    Form_main.TabStrip.Tabs(num_tab + 1).Selected = True
    
    Call show_my_shared_files
End Sub

Private Sub Form_Activate()
    activate_tab (Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    remove_tab (Me.hWnd)
End Sub

Private Sub form_resize()
    On Error Resume Next
    Dim size As Long
    size = Me.Height - 1000
    If size > 0 Then ListView1.Height = size
    size = Me.Width - 500
    If size > 0 Then ListView1.Width = size
End Sub

Private Sub cmdRefresh_Click()
    Call show_my_shared_files
End Sub


Private Sub show_my_shared_files()
    On Error Resume Next
    Dim cpt As Long
    'name,size,speed,ip,port,index,servent_id
    With Me.ListView1
        .ListItems.Clear
        For cpt = 0 To UBound(my_shared_files) - 1
            .ListItems.Add cpt + 1, , my_shared_files(cpt).file_name
            .ListItems(cpt + 1).SubItems(1) = Int(my_shared_files(cpt).file_size / 1000)
            .ListItems(cpt + 1).SubItems(2) = my_shared_files(cpt).full_path
        Next cpt
    End With
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error Resume Next
    With ListView1
        .Sorted = True
        .SortKey = ColumnHeader.index - 1
        
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        .Sorted = False 'for refresh (.sorted must be false when adding items)
    End With
End Sub
