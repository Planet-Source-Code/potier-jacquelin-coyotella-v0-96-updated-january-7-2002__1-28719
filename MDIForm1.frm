VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm Form_main 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Coyotella"
   ClientHeight    =   5805
   ClientLeft      =   4650
   ClientTop       =   6180
   ClientWidth     =   6960
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   6960
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinWidth1       =   4995
      MinHeight1      =   390
      Width1          =   4995
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Child2          =   "TabStrip"
      MinWidth2       =   6000
      MinHeight2      =   300
      Width2          =   1005
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin MSComctlLib.TabStrip TabStrip 
         Height          =   300
         Left            =   165
         TabIndex        =   3
         Top             =   450
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   688
         ButtonWidth     =   714
         ButtonHeight    =   688
         Style           =   1
         ImageList       =   "imlToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Connect to gnutella network"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.ToolTipText     =   "Disconnect from gnutella network"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New search"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Download / Upload informations"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Traffic"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Known Files"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Your Shared Files"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Known hosts"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Options"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Help"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer Timer_second 
      Interval        =   1000
      Left            =   120
      Top             =   3000
   End
   Begin VB.Timer Timer_minute 
      Interval        =   60000
      Left            =   120
      Top             =   2640
   End
   Begin MSComctlLib.ImageList imlmenu 
      Left            =   2880
      Top             =   2280
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   2880
      Top             =   1680
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1236
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":35D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":58DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":75E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7ECE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock socket 
      Index           =   0
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "MDIForm1.frx":87B6
            Text            =   "Disconnected"
            TextSave        =   "Disconnected"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWcascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWTileV 
         Caption         =   "Tile horizontaly"
      End
      Begin VB.Menu mnuWTileH 
         Caption         =   "Tile Verticaly"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
End
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''' socket events

Private Sub socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    On Error Resume Next
    Dim ip As String
    If index = 0 Then '0 <--> the server
        If current_nb_incoming < max_incoming_connection And Not disconnected_from_gnutella_network Then
                                                                'disconnected_from_gnutella_network allow to don't close server (because the freeing of socket is not immediate)
            ip = socket(0).RemoteHostIP 'ip of the host making the request
            'test if ip is not banished
            If is_ip_banished(ip) Then Exit Sub
            Dim num_socket As Integer
            'load or find a free socket
            num_socket = find_first_free_socket()
            If num_socket > 0 Then
                'accept the connection
                Form_main.socket(num_socket).LocalPort = 0
                Form_main.socket(num_socket).Accept requestID
                'increase current_nb_incoming and keep in memory the socket of the incomming connection
                current_nb_incoming = current_nb_incoming + 1

                socket_state(num_socket).connection_direction = incoming
                socket_state(num_socket).connection_type = not_define_yet
                'add to traffic interface
                add_to_connected_ip num_socket
            End If
        End If
    End If
End Sub

Private Sub socket_DataArrival(index As Integer, ByVal bytestotal As Long)
    On Error Resume Next
    Dim strdata As String

    socket_state(index).bytes_rcv = socket_state(index).bytes_rcv + bytestotal
    If Form_main.socket(index).state = sckConnected Then  'check if our socket has not been closed
        Form_main.socket(index).GetData strdata, vbString
        treat_data strdata, index, bytestotal
    End If
End Sub

Private Sub socket_Connect(index As Integer)
    On Error Resume Next
    Dim cpt     As Integer
    Dim pos     As Integer
    Dim found   As Boolean
    Dim data    As String

    socket_state(index).connection_direction = outgoing
    'add to traffic interface
    add_to_connected_ip index
    
    Select Case socket_state(index).connection_type
        Case dialing 'if we have a dialing socket, send gnutella header
            clarify_connection index, "Dialing"
            current_nb_dial = current_nb_dial + 1
            If current_nb_dial = 1 Then
                connected_to_gnutella_network = True
                Form_main.StatusBar1.Panels(1).Picture = Form_main.imlToolbar.ListImages(12).Picture
                Form_main.StatusBar1.Panels(1).Text = "Connected"
                Form_main.Toolbar1.Buttons(1).Enabled = False
                Form_main.Toolbar1.Buttons(2).Enabled = True
            End If
            'send protocol connection data
            Form_main.socket(index).SendData "GNUTELLA CONNECT/0.4" & vbLf & vbLf

        Case downloading
            'send the first download request of the list waiting_download
            For cpt = 0 To UBound(waiting_download) - 1
                If waiting_download(cpt).num_socket = index Then
                    found = True
                    pos = cpt
                    Exit For
                End If
            Next cpt
            If Not found Then 'error
                Form_main.socket(index).Close
                treat_socket_closing index
                Exit Sub
            End If
            'if found
            If send_data(waiting_download(pos).data, index) Then
                'remove first element of the liste
                waiting_download(pos).connection_tried = True
                clarify_connection index, "Downloading"
                remove_waiting_download (pos)
                update_status_download "Connected", index
            End If
        Case giving
            'search throught push_upload to find the corresponding socket
            For cpt = 0 To UBound(push_upload) - 1
                If push_upload(cpt).num_socket = index Then
                    data = push_upload(cpt).data
                    remove_from_push_upload cpt
                    Me.socket(index).SendData data ' send the push request
                    Exit Sub
                End If
            Next cpt
    End Select
End Sub

Private Sub socket_SendComplete(index As Integer)
    On Error Resume Next
    If socket_state(index).connection_type = uploading Then
        update_status_upload "Completed", index
        Form_main.socket(index).Close
        treat_socket_closing index
    End If
End Sub

Private Sub socket_SendProgress(index As Integer, ByVal bytesSent As Long, ByVal BytesRemaining As Long)
    On Error Resume Next
    Dim percent As String
    If BytesRemaining > 0 Then
        percent = CStr(Int(bytesSent / (bytesSent + BytesRemaining) * 100)) & " %"
        update_status_upload percent, index, BytesRemaining
    End If
    socket_state(index).bytes_rcv = socket_state(index).bytes_rcv + bytesSent
End Sub

Private Sub socket_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    treat_socket_closing index, True, Description, False, Number
End Sub

Private Sub socket_Close(index As Integer) 'when connection is closed by remote side
    treat_socket_closing index
End Sub



Private Sub Timer_minute_Timer()
    On Error Resume Next
    Dim cpt         As Integer
    For cpt = 0 To UBound(socket_state)
        If socket_state(cpt).connection_type = dialing Then
            socket_state(cpt).number_of_ping = 0
            socket_state(cpt).number_of_query = 0
            socket_state(cpt).number_of_bogus = 0
            send_a_ping cpt
        End If
    Next cpt
    
End Sub

Private Sub Timer_second_Timer()
    On Error Resume Next
    Dim cpt                     As Integer
    Dim cpt2                    As Integer
    Dim size                    As Integer
    Dim array_size              As Integer
    Dim num_socket              As Integer
    Dim pos_name_download       As Integer
    Dim pos_current_down        As Integer
    Dim strrange_end            As String
    Dim data                    As String

    'check push validity time
    For cpt = UBound(waiting_giv) - 1 To 0 Step -1
        waiting_giv(cpt).push_validity_time = waiting_giv(cpt).push_validity_time - 1
        'if validity time is <=0 then retry download
        If waiting_giv(cpt).push_validity_time <= 0 Then
            pos_name_download = waiting_giv(cpt).pos_name_download
            name_download(pos_name_download).need_to_push = True
            remove_from_waiting_giv cpt
            ' add to retry download
            size = UBound(retry_download)
            ReDim Preserve retry_download(size + 1)
            retry_download(size).remaining_time = retry_down_on_busy_server_every
            retry_download(size).pos_name_download = pos_name_download
        End If
    Next cpt

    ' check retry download
    size = UBound(retry_download)
    For cpt = 0 To size - 1
        If cpt > size - 1 Then Exit For
        retry_download(cpt).remaining_time = retry_download(cpt).remaining_time - 1
        
        If retry_download(cpt).remaining_time <= 0 Then
            'resume
            pos_name_download = retry_download(cpt).pos_name_download
            If pos_name_download < 0 Then Exit Sub
            name_download(pos_name_download).get_request_made = False
            name_download(pos_name_download).header_received = False
            
            If name_download(pos_name_download).need_to_push Then
                Dim strpush
                strpush = make_string_push_data(name_download(pos_name_download).servent_id, name_download(pos_name_download).file_index, my_ip, my_port)
                data = make_string_descriptor_data(my_servent_id, push, my_ttl, 0, 26)
                data = data & strpush
                
                'add file in waiting array
                array_size = UBound(waiting_download)
                waiting_download(array_size).connection_tried = False
                waiting_download(array_size).data = data
                waiting_download(array_size).num_socket = -1
                waiting_download(array_size).push = True
                waiting_download(array_size).num_socket_queryhit = name_download(pos_name_download).num_socket_queryhit
                waiting_download(array_size).position_name_download = pos_name_download
                ReDim Preserve waiting_download(array_size + 1)
                update_status_download2 "Waiting", name_download(pos_name_download).pos_current_down
            Else
                With name_download(pos_name_download)
                    If .range_end = 0 Then
                        strrange_end = ""
                    Else
                        strrange_end = CStr(CLng(.range_end))
                    End If
                    
                    data = "GET /get/" & .file_index & "/" & .file_name & " HTTP/1.0" & vbCrLf _
                        & "Range: bytes=" & CStr(CLng(.range_begin) + CLng(.bytes_recieved)) & "-" & strrange_end & vbCrLf _
                        & "User-Agent: Coyotella" & vbCrLf _
                        & vbCrLf
                
                End With
                array_size = UBound(waiting_download)
                waiting_download(array_size).connection_tried = False
                waiting_download(array_size).data = data
                waiting_download(array_size).num_socket = -1
                waiting_download(array_size).push = False
                waiting_download(array_size).num_socket_queryhit = name_download(pos_name_download).num_socket_queryhit
                waiting_download(array_size).position_name_download = pos_name_download
                ReDim Preserve waiting_download(array_size + 1)
                update_status_download2 "Waiting", name_download(pos_name_download).pos_current_down
            End If
            
            'remove from array retry_download
            For cpt2 = cpt + 1 To size - 1
                retry_download(cpt2 - 1) = retry_download(cpt2)
            Next cpt2
            
            ReDim Preserve retry_download(size - 1)
            size = size - 1
            cpt = cpt - 1
        Else
            update_status_download2 "Retring in " & CStr(retry_download(cpt).remaining_time) & "s (" & name_download(retry_download(cpt).pos_name_download).old_percent & "% completed)", _
                                    name_download(retry_download(cpt).pos_name_download).pos_current_down
        End If
    Next cpt
    
    If current_nb_download < max_download Then
        'check for waiting download
        Call check_for_waiting_download
    End If
    
    'connect to enougth people if necessary
    If current_nb_dial < min_dialing_hosts Then
        make_new_random_dial True, False, True
    End If
    'modify all connections rate
        
        For cpt = 0 To UBound(current_download) - 1
            If current_download(cpt).num_socket > -1 Then
                current_download(cpt).speed = Round(socket_state(current_download(cpt).num_socket).bytes_rcv / 1000, 4)
                'update download list
                update_speed_remain_download2 current_download(cpt).speed, cpt, current_download(cpt).remaining_bytes
            End If
        Next cpt
'        Form_download_upload.ListView_download.Refresh
    
        
        For cpt = 0 To UBound(current_upload) - 1
            If current_upload(cpt).num_socket > -1 Then
                current_upload(cpt).speed = Round(socket_state(current_upload(cpt).num_socket).bytes_rcv / 1000, 4)
                'update upload list
                update_speed_remain_upload2 current_upload(cpt).speed, cpt, current_upload(cpt).remaining_bytes
            End If
        Next cpt
'        Form_download_upload.ListView_upload.Refresh
        
        
        
        For cpt = 0 To UBound(traffic_connected_ip) - 1
            traffic_connected_ip(cpt).speed = Round(socket_state(traffic_connected_ip(cpt).num_socket).bytes_rcv / 1000, 4)
            'update form traffic
            update_traffic_speed traffic_connected_ip(cpt).speed, traffic_connected_ip(cpt).num_socket, cpt
            socket_state(traffic_connected_ip(cpt).num_socket).bytes_rcv = 0
        Next cpt
'        Form_traffic.ListView_connected_ip.Refresh

End Sub


Private Sub MDIForm_Load()
    On Error Resume Next
    Me.TabStrip.Tabs.Remove (1)
    Me.CoolBar1.Bands(2).Visible = False
    Call initialize_app
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call disconnect_from_gnutella_network
    quit_program = True
End Sub



Private Sub TabStrip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim cpt As Integer

    Select Case Me.TabStrip.SelectedItem.Caption
        Case "Down/Upload"
            Form_download_upload.Show
            Form_download_upload.SetFocus
        Case "Known Hosts"
            Form_known_hosts.Show
            Form_known_hosts.SetFocus
        Case "Options"
            Form_options.Show
            Form_options.SetFocus
        Case "Shared Files"
            Form_shared_files.Show
            Form_shared_files.SetFocus
        Case "Traffic"
            Form_traffic.Show
            Form_traffic.SetFocus
        Case "Known Files"
            Form_viewed_files.Show
            Form_viewed_files.SetFocus
        Case Else 'tab is for a search
            For cpt = 0 To UBound(Document_search)
                If Me.TabStrip.SelectedItem.Tag = Document_search(cpt).hWnd Then
                    Document_search(cpt).Show
                    Document_search(cpt).SetFocus
                    Exit For
                End If
            Next cpt
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim num_tab     As Integer

    Select Case Button.index
        Case 1
            'connect
            Toolbar1.Buttons(1).Enabled = False
            Toolbar1.Buttons(2).Enabled = True
            Call connect_to_gnutella_network
        Case 2
            'diconnect
            Toolbar1.Buttons(2).Enabled = False
            Toolbar1.Buttons(1).Enabled = True
            Call disconnect_from_gnutella_network
        Case 4
            'new search
            Call new_document_search
        Case 6
            'down/upload informations
            Form_download_upload.Show
            If Not is_form_in_tabstrip(Form_download_upload.hWnd) Then
                num_tab = Form_main.TabStrip.Tabs.Count
                Form_main.CoolBar1.Bands(2).Visible = True
                Form_main.TabStrip.Tabs.Add num_tab + 1, , "Down/Upload"
                Form_main.TabStrip.Tabs(num_tab + 1).Tag = Form_download_upload.hWnd
                Form_main.TabStrip.Tabs(num_tab + 1).Selected = True
            End If
        Case 7
            'traffic
            Form_traffic.Show
            If Not is_form_in_tabstrip(Form_traffic.hWnd) Then
                num_tab = Form_main.TabStrip.Tabs.Count
                Form_main.CoolBar1.Bands(2).Visible = True
                Form_main.TabStrip.Tabs.Add num_tab + 1, , "Traffic"
                Form_main.TabStrip.Tabs(num_tab + 1).Tag = Form_traffic.hWnd
                Form_main.TabStrip.Tabs(num_tab + 1).Selected = True
            End If
        Case 8
            'view known files
            Form_viewed_files.Show
        Case 9
            'show shared files
            Form_shared_files.Show
        Case 10
            'known hosts
            Form_known_hosts.Show
            If Not is_form_in_tabstrip(Form_known_hosts.hWnd) Then
                num_tab = Form_main.TabStrip.Tabs.Count
                Form_main.CoolBar1.Bands(2).Visible = True
                Form_main.TabStrip.Tabs.Add num_tab + 1, , "Known Hosts"
                Form_main.TabStrip.Tabs(num_tab + 1).Tag = Form_known_hosts.hWnd
                Form_main.TabStrip.Tabs(num_tab + 1).Selected = True
            End If
        Case 11
            'options
            Form_options.Show
        Case 12
            'help/about
            Form_about.Show
    End Select
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

