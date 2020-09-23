VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_download_upload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloads / Uploads"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   9375
   Icon            =   "Form_download_upload.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   9375
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":2282
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":2F5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":3C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":4916
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_download_upload.frx":51F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture_resize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      FillColor       =   &H80000015&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      Picture         =   "Form_download_upload.frx":5C7E
      ScaleHeight     =   135
      ScaleWidth      =   9135
      TabIndex        =   2
      Top             =   3480
      Width           =   9135
   End
   Begin MSComctlLib.ListView ListView_upload 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Speed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remaining Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView_download 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5106
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Range"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "C Speed"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Speed (kb/s)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Remaining Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   741
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Start/Resume"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Remove"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Up"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Down"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear Completed"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear Interrupted Uploads"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile Horizontaly"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "Tile Verticaly"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
   Begin VB.Menu mnu_download 
      Caption         =   "mnu_download"
      Visible         =   0   'False
      Begin VB.Menu mnuresume 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnupause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuresumeonotherhost 
         Caption         =   "Resume on other host"
      End
      Begin VB.Menu mnuremove 
         Caption         =   "Remove"
      End
   End
   Begin VB.Menu mnu_upload 
      Caption         =   "mnu_upload"
      Visible         =   0   'False
      Begin VB.Menu mnustop 
         Caption         =   "Stop this Upload"
      End
      Begin VB.Menu mnubanish 
         Caption         =   "Banish this IP"
      End
   End
End
Attribute VB_Name = "Form_download_upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private full_height As Long




Private Sub Form_Load()
    On Error Resume Next
    'find info to fill ListView_download in current_download
    ' and ListView_upload in current_upload
    
   
    Dim cpt As Integer
    Dim pos As Integer
    'show down info
    For cpt = 0 To UBound(current_download) - 1
        With Form_download_upload.ListView_download
            pos = .ListItems.Count + 1
            'file_name,file_range,connection speed,status ,speed, ip
            .ListItems.Add pos, , current_download(cpt).file_name
            .ListItems(pos).SubItems(1) = current_download(cpt).file_range
            .ListItems(pos).SubItems(2) = current_download(cpt).cspeed
            .ListItems(pos).SubItems(3) = current_download(cpt).status
            .ListItems(pos).SubItems(4) = current_download(cpt).speed
            .ListItems(pos).SubItems(5) = current_download(cpt).ip
            .ListItems(pos).Tag = cpt
        End With
    Next cpt
    'show up info
    For cpt = 0 To UBound(current_upload) - 1
        With Form_download_upload.ListView_upload
            pos = .ListItems.Count + 1
            'name, range, ip,status,speed
            .ListItems.Add pos, , current_upload(cpt).file_name
            .ListItems(pos).SubItems(1) = current_upload(cpt).file_range
            .ListItems(pos).SubItems(2) = current_upload(cpt).ip
            .ListItems(pos).SubItems(3) = current_upload(cpt).status
            .ListItems(pos).SubItems(4) = current_upload(cpt).speed
            .ListItems(pos).Tag = cpt
        End With
    Next cpt
    
    Me.Picture_resize.Top = Me.ListView_download.Top + Me.ListView_download.Height
    Me.ListView_upload.Top = Me.Picture_resize.Top + Me.Picture_resize.Height
    full_height = Me.ListView_download.Height + Me.ListView_upload.Height
End Sub

Private Sub Form_Activate()
    activate_tab (Me.hWnd)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    remove_tab (Me.hWnd)
End Sub

Private Sub ListView_download_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnu_download
    End If
End Sub

Private Sub ListView_upload_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnu_upload
    End If
End Sub



Private Sub Picture_resize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Picture_resize.MousePointer = 7
    If Button = 1 Then
        If Me.ListView_download.Height + y < full_height And Me.ListView_download.Height + y > 0 _
           And Me.ListView_upload.Height - y < full_height And Me.ListView_upload.Height - y > 0 _
        Then
            Me.ListView_download.Height = Me.ListView_download.Height + y
            Me.Picture_resize.Top = Me.ListView_download.Top + Me.ListView_download.Height
            Me.ListView_upload.Top = Me.Picture_resize.Top + Me.Picture_resize.Height
            Me.ListView_upload.Height = Me.ListView_upload.Height - y
        End If
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1
            Call mnuresume_Click
        Case 2
            Call mnupause_Click
        Case 3
            Call mnuremove_Click
        Case 4
            'move download up
            Call move_download_up
        Case 5
            'move download down
            Call move_download_down
        Case 6
            'clear completed download and upload
            Call clear_completed
        Case 7
            'clear interrupted uploads
            Call clear_interrupted
    End Select
End Sub

''''''''''''''menu functions
Private Sub mnubanish_Click() 'for upload only
    On Error Resume Next
    Dim pos_selected        As Integer
    Dim pos_current_up      As Integer
    pos_selected = find_selected_up()
    pos_current_up = ListView_upload.ListItems(pos_selected).Tag
    ban_ip current_upload(pos_current_up).ip
End Sub



'''''''''''
Private Sub mnustop_Click() 'for upload only
    On Error Resume Next
    Dim pos_selected        As Integer
    Dim pos_current_up      As Integer
    Dim num_socket          As Integer
    
    pos_selected = find_selected_up()
    pos_current_up = ListView_upload.ListItems(pos_selected).Tag
    num_socket = current_upload(pos_current_up).num_socket
    
    Form_main.socket(num_socket).Close
    treat_socket_closing num_socket
End Sub


Private Sub mnupause_Click() 'for download only
    On Error Resume Next
    '<---> disconnect
    Dim pos_selected        As Integer
    Dim pos_current_down    As Integer
    Dim pos_name_download   As Integer
    Dim file_name           As String
    Dim status              As String
    Dim cpt                 As Integer
    Dim size                As Integer

    
    stop_download pos_selected, pos_current_down, pos_name_download, file_name
    If pos_selected < 1 Then Exit Sub 'no item selected

    If InStr(1, current_download(pos_current_down).status, "Paused") > 0 Then Exit Sub 'paused already made
    If current_download(pos_current_down).status = "Completed" Then Exit Sub
    'remove from retry_download if necessary
    size = UBound(retry_download)
    For cpt = size - 1 To 0 Step -1
        If retry_download(cpt).pos_name_download = pos_name_download Then
            'remove
            retry_download(cpt) = retry_download(size - 1)
            ReDim Preserve retry_download(size - 1)
        End If
    Next cpt
    
    'remove from waiting giv if necessary
    size = UBound(waiting_giv)
    For cpt = size - 1 To 0 Step -1
        If waiting_giv(cpt).pos_name_download = pos_name_download Then
            'remove
            remove_from_waiting_giv cpt
        End If
    Next cpt
    
    If pos_name_download < 0 Then Exit Sub ' waiting
    
    status = "Paused " & name_download(pos_name_download).old_percent & "%" 'current_download(pos_current_down).status"
    
    current_download(pos_current_down).status = status
    ListView_download.ListItems(pos_selected).SubItems(3) = status
End Sub

Private Sub mnuremove_Click() 'download only
    'remove : stop (disconnect) and remove file
    On Error Resume Next
    Dim pos_selected        As Integer
    Dim pos_current_down    As Integer
    Dim pos_name_download   As Integer
    Dim file_name           As String
    Dim cpt                 As Integer
    Dim size                As Integer
    Dim saving_name         As String
    Dim other_part          As Boolean
    Dim strtmp              As String

    
    stop_download pos_selected, pos_current_down, pos_name_download, file_name
    If pos_selected < 1 Or pos_name_download = -1 Then Exit Sub 'no item selected
    
    'remove from retry_download if necessary
    size = UBound(retry_download)
    For cpt = size - 1 To 0 Step -1
        If retry_download(cpt).pos_name_download = pos_name_download Then
            'remove
            retry_download(cpt) = retry_download(size - 1)
            ReDim Preserve retry_download(size - 1)
        End If
    Next cpt
    
    'remove from waiting giv if necessary
    size = UBound(waiting_giv)
    For cpt = size - 1 To 0 Step -1
        If waiting_giv(cpt).pos_name_download = pos_name_download Then
            'remove
            remove_from_waiting_giv cpt
        End If
    Next cpt
    
    
    saving_name = name_download(current_download(pos_current_down).pos_name_download).saving_name
    'remove from name_download and waiting download if necessary
    If pos_name_download > -1 Then ' else download is finished don't need to remove from waiting download
        'remove from name_download
        remove_from_name_download pos_name_download
        
        For cpt = UBound(waiting_download) To 0 Step -1
            'remove from waiting download
            If waiting_download(cpt).position_name_download = pos_name_download Then
                remove_waiting_download cpt
            End If
        Next cpt
    End If


    'remove from current_download
    remove_from_current_download pos_current_down
    
    'remove from list view
    ListView_download.ListItems.Remove pos_selected

    'remove file and recovery file

    For cpt = 0 To UBound(current_download) - 1
        If current_download(cpt).pos_name_download > -1 Then
            If name_download(current_download(cpt).pos_name_download).saving_name = saving_name Then
                other_part = True
                Exit For
            End If
        End If
    Next cpt
    
    If other_part Then
        'case of multipart download ask to remove other downloads and the file or just to remove this part
        Form_remove_download.Tag = file_name 'passing saving name to the form
        Form_remove_download.Show vbModal
    Else
        strtmp = my_incomplete_directory & saving_name
        If is_file_existing(strtmp) Then
            Kill strtmp
        End If
        strtmp = strtmp & ".coy"
        If is_file_existing(strtmp) Then
            Kill strtmp
        End If
    End If
End Sub

Private Sub mnuresume_Click()
    On Error Resume Next
    ' download is already present in name_download, and current_download
    '.coy is already existing
    Dim data                As String
    Dim pos_selected        As Integer
    Dim pos_current_down    As Integer
    Dim strrange            As String
    Dim pos_name_download   As Integer
    Dim array_size          As Integer
    Dim cpt                 As Integer
    Dim strrange_end        As String
    Dim file_name           As String
    Dim size                As Integer
    
    stop_download pos_selected, pos_current_down, pos_name_download, file_name ' to be sure that the download is in pause
    If pos_selected < 1 Then Exit Sub 'no item selected
    If current_download(pos_current_down).status = "Completed" Then Exit Sub
    'remove from retry_download if necessary
    size = UBound(retry_download)
    For cpt = size - 1 To 0 Step -1
        If retry_download(cpt).pos_name_download = pos_name_download Then
            'remove
            retry_download(cpt) = retry_download(size - 1)
            ReDim Preserve retry_download(size - 1)
        End If
    Next cpt
    
    'remove from waiting giv if necessary
    size = UBound(waiting_giv)
    For cpt = size - 1 To 0 Step -1
        If waiting_giv(cpt).pos_name_download = pos_name_download Then
            'remove
            remove_from_waiting_giv cpt
        End If
    Next cpt
    
    strrange = current_download(pos_current_down).file_range

    current_download(pos_current_down).status = "Resuming"
    ListView_download.ListItems(pos_selected).SubItems(3) = "Resuming"

    'search the position in name download
    pos_name_download = current_download(pos_current_down).pos_name_download
    If pos_name_download = -1 Then Exit Sub
    
    With name_download(pos_name_download)
        If .range_end = 0 Then
            strrange_end = ""
        Else
            strrange_end = CStr(CLng(.range_end))
        End If
        If .need_to_push Then
            data = make_string_push_data(.servent_id, .file_index, .ip, .port)
        Else
            data = "GET /get/" & .file_index & "/" & .file_name & " HTTP/1.0" & vbCrLf _
                & "Range: bytes=" & CStr(CLng(.range_begin)) & "-" & strrange_end & vbCrLf _
                & "User-Agent: Coyotella" & vbCrLf _
                & vbCrLf
        End If
    End With

    array_size = UBound(waiting_download)
    waiting_download(array_size).data = data
    waiting_download(array_size).position_name_download = pos_name_download
    waiting_download(array_size).connection_tried = False
    waiting_download(array_size).num_socket = -1
    waiting_download(array_size).num_socket_queryhit = name_download(pos_name_download).num_socket_queryhit
    waiting_download(array_size).push = name_download(pos_name_download).need_to_push
    ReDim Preserve waiting_download(array_size + 1)

End Sub



Private Sub mnuresumeonotherhost_Click()
    On Error Resume Next
    ' show a new form search with possible known files
    '(the search form allow to find other results if we don't have enought)
    Dim pos_selected        As Integer
    Dim pos_current_down    As Integer
    Dim pos_name_download   As Integer
    Dim strrange            As String
    Dim strrange_end        As String
    Dim cpt                 As Integer
    Dim num_file            As Integer
    Dim nb_parts            As Single
    Dim file_name           As String
    Dim size                As Integer
    
    stop_download pos_selected, pos_current_down, pos_name_download, file_name ' to be sure that the download is in pause
    If pos_selected < 1 Then Exit Sub 'no item selected
    
    'remove from retry_download if necessary
    size = UBound(retry_download)
    For cpt = 0 To size - 1
        If retry_download(cpt).pos_name_download = pos_name_download Then
            'remove
            retry_download(cpt) = retry_download(size - 1)
            ReDim Preserve retry_download(size - 1)
        End If
    Next cpt
    
    'remove from waiting giv if necessary
    size = UBound(waiting_giv)
    For cpt = 0 To size - 1
        If waiting_giv(cpt).pos_name_download = pos_name_download Then
            'remove
            remove_from_waiting_giv cpt
        End If
    Next cpt

    If pos_name_download = -1 Then Exit Sub
    

    num_file = FreeFile()
    Open my_incomplete_directory & name_download(pos_name_download).saving_name & ".coy" For Random Access Read Shared As num_file Len = 4
        Get num_file, 1, nb_parts
    Close num_file
    new_host_search name_download(pos_name_download).file_name, name_download(pos_name_download).file_size, _
                    name_download(pos_name_download).range_begin + name_download(pos_name_download).bytes_recieved, name_download(pos_name_download).range_end, _
                    False, nb_parts, name_download(pos_name_download).num_part, _
                    name_download(pos_name_download).pos_current_down, name_download(pos_name_download).saving_name
End Sub

Private Sub clear_completed()
    On Error Resume Next
    'remove from interface and current_download current_upload arrays
    Dim cpt As Integer
    
    For cpt = Me.ListView_download.ListItems.Count To 1 Step -1
        If current_download(Me.ListView_download.ListItems(cpt).Tag).status = "Completed" Then
        'download completed
            remove_from_current_download Me.ListView_download.ListItems(cpt).Tag
            Me.ListView_download.ListItems.Remove cpt
        End If
    Next cpt
    '
    For cpt = Me.ListView_upload.ListItems.Count To 1 Step -1
        If current_upload(Me.ListView_upload.ListItems(cpt).Tag).status = "Completed" Then
        'upload completed
            remove_from_current_upload Me.ListView_upload.ListItems(cpt).Tag
            Me.ListView_upload.ListItems.Remove cpt
        End If
    Next cpt
End Sub

Private Sub clear_interrupted()
    On Error Resume Next
    'remove from interface and current_upload array
    Dim cpt As Integer
    For cpt = Me.ListView_upload.ListItems.Count To 1 Step -1
        If current_upload(Me.ListView_upload.ListItems(cpt).Tag).remaining_bytes > 0 _
            And current_upload(Me.ListView_upload.ListItems(cpt).Tag).num_socket = -1 Then
        'upload interrupted
            remove_from_current_upload Me.ListView_upload.ListItems(cpt).Tag
            Me.ListView_upload.ListItems.Remove cpt
        End If
    Next cpt
End Sub

Private Sub move_download_up()
    On Error Resume Next
    Dim pos_selected            As Integer
    Dim pos_current_download    As Integer
    Dim pos_name_down           As Integer
    Dim tmp                     As twaiting_download
    Dim tmp_listitem            As ListItem
    Dim cpt                     As Integer
    
    pos_selected = find_selected_down()
    If pos_selected < 1 Then Exit Sub
    If pos_selected = 1 Then Exit Sub 'item is the first
    
    'change position in waiting_download
    pos_current_download = Me.ListView_download.ListItems(pos_selected).Tag
    pos_name_down = current_download(pos_current_download).pos_name_download
    For cpt = 1 To UBound(waiting_download) - 1 'if in position 0 it is already the first next to be download
        If waiting_download(cpt).position_name_download = pos_name_down Then
            tmp = waiting_download(cpt - 1)
            waiting_download(cpt - 1) = waiting_download(cpt)
            waiting_download(cpt) = tmp
            Exit For
        End If
    Next cpt
    
    'change position on interface
     Set tmp_listitem = Me.ListView_download.ListItems(pos_selected)
     Me.ListView_download.ListItems.Remove (pos_selected)
     Me.ListView_download.ListItems.Add pos_selected - 1, , tmp_listitem.Text
     Me.ListView_download.ListItems(pos_selected - 1).SubItems(1) = tmp_listitem.SubItems(1)
     Me.ListView_download.ListItems(pos_selected - 1).SubItems(2) = tmp_listitem.SubItems(2)
     Me.ListView_download.ListItems(pos_selected - 1).SubItems(3) = tmp_listitem.SubItems(3)
     Me.ListView_download.ListItems(pos_selected - 1).SubItems(4) = tmp_listitem.SubItems(4)
     Me.ListView_download.ListItems(pos_selected - 1).SubItems(5) = tmp_listitem.SubItems(5)
End Sub

Private Sub move_download_down()
    On Error Resume Next
    Dim pos_selected            As Integer
    Dim pos_current_download    As Integer
    Dim pos_name_down           As Integer
    Dim tmp                     As twaiting_download
    Dim tmp_listitem            As ListItem
    Dim cpt                     As Integer
    
    pos_selected = find_selected_down()
    If pos_selected < 1 Then Exit Sub
    If pos_selected = Me.ListView_download.ListItems.Count Then Exit Sub  'item is the last
    
    'change position in waiting_download
    pos_current_download = Me.ListView_download.ListItems(pos_selected).Tag
    pos_name_down = current_download(pos_current_download).pos_name_download
    For cpt = 0 To UBound(waiting_download) - 2 'if in position UBound(waiting_download) - 2 it is already the last to be download
        If waiting_download(cpt).position_name_download = pos_name_down Then
            tmp = waiting_download(cpt + 1)
            waiting_download(cpt + 1) = waiting_download(cpt)
            waiting_download(cpt) = tmp
            Exit For
        End If
    Next cpt
    
    'change position on interface
     tmp_listitem = Me.ListView_download.ListItems(pos_selected + 1)
     Me.ListView_download.ListItems.Remove (pos_selected)
     Me.ListView_download.ListItems.Add pos_selected + 1, , tmp_listitem.Text
     Me.ListView_download.ListItems(pos_selected + 1).SubItems(1) = tmp_listitem.SubItems(1)
     Me.ListView_download.ListItems(pos_selected + 1).SubItems(2) = tmp_listitem.SubItems(2)
     Me.ListView_download.ListItems(pos_selected + 1).SubItems(3) = tmp_listitem.SubItems(3)
     Me.ListView_download.ListItems(pos_selected + 1).SubItems(4) = tmp_listitem.SubItems(4)
     Me.ListView_download.ListItems(pos_selected + 1).SubItems(5) = tmp_listitem.SubItems(5)
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

