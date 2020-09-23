VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_traffic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traffic"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   6420
   Icon            =   "Form_traffic.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   6420
   Begin MSComctlLib.ListView ListView_connected_ip 
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Connections"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Port"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "In/Out"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "State"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Speed (Kb/s)"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList_payload 
      Left            =   5520
      Top             =   3840
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":104A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":1D8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":2A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":2EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_traffic.frx":330E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView_current_payload 
      Height          =   3615
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList_payload"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payload descriptor"
         Object.Width           =   2681
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "from/to IP"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Recieved"
      Height          =   1935
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Width           =   2295
      Begin VB.Label in_push 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Push"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label in_bogus 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Bogus"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label in_queryhit 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label in_query 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label in_pong 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label in_ping 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Queryhit"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Query"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Pong"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Ping"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sent"
      Height          =   1575
      Left            =   3960
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
      Begin VB.Label out_push 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Push"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label out_queryhit 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label out_pong 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label out_query 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label out_ping 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Queryhit"
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "Pong"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Query"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Ping"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Menu mnuwindow 
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
   Begin VB.Menu my_popup 
      Caption         =   "my_popup"
      Visible         =   0   'False
      Begin VB.Menu Disconnect 
         Caption         =   "Disconnect this host"
      End
      Begin VB.Menu mnubanish 
         Caption         =   "Banish this host"
      End
   End
End
Attribute VB_Name = "Form_traffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function find_selected() As Integer
    On Error Resume Next
    Dim cpt             As Integer
    With ListView_connected_ip
        For cpt = 1 To .ListItems.Count 'find selection
            If .ListItems(cpt).Selected Then
                'ip,port,in/out,state
                find_selected = cpt
                Exit Function
            End If
        Next cpt
    End With
End Function

Private Sub Form_Load()



    On Error Resume Next
    Dim cpt         As Integer
    Dim last_pos    As Integer
    ' connections state
    For cpt = 0 To UBound(traffic_connected_ip) - 1
        With Form_traffic.ListView_connected_ip
            last_pos = .ListItems.Count + 1
            .ListItems.Add last_pos, , traffic_connected_ip(cpt).ip
            .ListItems(last_pos).SubItems(1) = traffic_connected_ip(cpt).port
            .ListItems(last_pos).SubItems(2) = traffic_connected_ip(cpt).incoming
            .ListItems(last_pos).SubItems(3) = traffic_connected_ip(cpt).state
            .ListItems(last_pos).Tag = traffic_connected_ip(cpt).num_socket
        End With
    Next cpt
    
    'nb of payload
    Form_traffic.in_ping.Caption = traffic_info.sent_ping
    Form_traffic.out_ping.Caption = traffic_info.rcv_ping
    Form_traffic.in_pong.Caption = traffic_info.sent_pong
    Form_traffic.out_pong.Caption = traffic_info.rcv_pong
    Form_traffic.in_push.Caption = traffic_info.sent_push
    Form_traffic.out_push.Caption = traffic_info.rcv_push
    Form_traffic.in_query.Caption = traffic_info.sent_query
    Form_traffic.out_query.Caption = traffic_info.rcv_query
    Form_traffic.in_queryhit.Caption = traffic_info.sent_queryhit
    Form_traffic.out_queryhit.Caption = traffic_info.rcv_queryhit
    Form_traffic.in_bogus.Caption = traffic_info.rcv_bogus
    
    ' payload list
    Dim strtmp          As String
    Dim num_img_liste   As Byte
    For cpt = 0 To traffic_info.array_size
        Select Case traffic_payload(cpt).payload_descriptor
            Case ping
                strtmp = "ping"
                num_img_liste = 1
            Case pong
                strtmp = "pong"
                num_img_liste = 2
            Case push
                strtmp = "push"
                num_img_liste = 3
            Case query
                strtmp = "query"
                num_img_liste = 4
            Case queryhit
                strtmp = "queryhit"
                num_img_liste = 5
            Case Else
                strtmp = "bogus payload=" & CStr(traffic_payload(cpt).payload_descriptor)
                num_img_liste = 6
        End Select
        If traffic_payload(cpt).info = "" Then
            Exit For 'not initialized
        End If
        With Form_traffic.ListView_current_payload
            .ListItems.Add 1, , strtmp, , num_img_liste
            .ListItems(1).SubItems(1) = traffic_payload(cpt).info
        End With
    Next cpt
        
End Sub

Private Sub Form_Activate()
    activate_tab (Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    remove_tab (Me.hWnd)
End Sub

Private Sub ListView_connected_ip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu my_popup
    End If
End Sub

Private Sub Disconnect_Click()
    On Error Resume Next
    'disonnect host with ip port
    Dim pos             As Integer
    Dim num_sock        As Integer
    
    pos = find_selected()
    num_sock = ListView_connected_ip.ListItems(pos).Tag
    Form_main.socket(num_sock).Close
    treat_socket_closing num_sock

End Sub

Private Sub mnubanish_Click()
    On Error Resume Next
    Dim pos_selected        As Integer
    Dim ip                  As String
    
    pos_selected = find_selected()
    ip = ListView_connected_ip.ListItems(pos_selected).Text
    ban_ip ip
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

