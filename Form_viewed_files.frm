VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_viewed_files 
   Caption         =   "Known Files (files viewed throw queryhits)"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6735
   Icon            =   "Form_viewed_files.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdRefresh 
      Height          =   495
      Left            =   120
      Picture         =   "Form_viewed_files.frx":1782
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Refresh"
      Top             =   0
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   13150
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "speed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "port"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "servent_id"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileH 
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
Attribute VB_Name = "Form_viewed_files"
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
    Form_main.TabStrip.Tabs.Add num_tab + 1, , "Known Files"
    Form_main.TabStrip.Tabs(num_tab + 1).Tag = Me.hWnd
    Form_main.TabStrip.Tabs(num_tab + 1).Selected = True

    Call show_known_files
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
    Call show_known_files
End Sub


Private Sub show_known_files()
    On Error Resume Next
    Dim cpt              As Long
    Dim strtmp           As String
    Dim strtmp2          As String
    Dim nb_space         As Byte   '
    nb_space = 13                  '
    
    
    'name,size,speed,ip,port,index,servent_id
    With Me.ListView1
        .ListItems.Clear
        For cpt = 0 To UBound(known_files) - 1
            .ListItems.Add cpt + 1, , known_files(cpt).file_name
            strtmp = CStr(Int(known_files(cpt).file_size / 1000)) 'to kb
            strtmp2 = Space$(nb_space - Len(strtmp)) & strtmp
            .ListItems(cpt + 1).SubItems(1) = strtmp2
            .ListItems(cpt + 1).SubItems(2) = known_files(cpt).speed
            .ListItems(cpt + 1).SubItems(3) = known_files(cpt).ip
            .ListItems(cpt + 1).SubItems(4) = known_files(cpt).port
            strtmp = CStr(known_files(cpt).file_index)
            strtmp2 = Space$(nb_space - Len(strtmp)) & strtmp
            .ListItems(cpt + 1).SubItems(5) = strtmp2
            .ListItems(cpt + 1).SubItems(6) = known_files(cpt).servent_id
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
