Attribute VB_Name = "socket"
Option Explicit



Public Sub treat_socket_closing(ByVal index As Integer, Optional ByVal error As Boolean = False, _
                                Optional ByVal error_description As String, _
                                Optional ByVal from_stop_download As Boolean = False, _
                                Optional ByVal error_num As Integer)
    On Error Resume Next
    Dim cpt               As Long
    
    remaining_data(index) = ""
    Select Case socket_state(index).connection_type
    
        Case dialing
            'not enought connection make a new one to a known host
            'if not enought known host send ping to all
            '--> all this is made in sub make_new_random_dial
            free_socket (index)
            
        Case uploading
            current_nb_upload = current_nb_upload - 1
            update_speed_remain_upload 0, index
            
            If error Then
                update_status_upload "Error: " & error_description, index
            Else
                If give_status_upload(index) <> "Completed" Then
                    update_status_upload "Transfer Interrupted", index
                End If
            End If
            free_socket (index)
            
            dissociate_current_upload_num_socket index
            
        Case downloading
            current_nb_download = current_nb_download - 1
            update_speed_remain_download 0, index
            
            Dim pos_name_download As Integer
            Dim size              As Integer
            size = UBound(name_download)
            For cpt = 0 To size - 1
                If name_download(cpt).num_socket = index Then
                    pos_name_download = cpt
                    Exit For
                End If
            Next cpt
            'remove file if it his in waiting_download (if connection hasn't be establish)
            remove_waiting_download_with_pos_name_download pos_name_download
            
'            If error Then 'firewall or host no more connected
'                If error_num = sckConnectionRefused Or error_num = sckConnectionReset Or error_num = sckNetworkUnreachable Then
'                   free_socket index
'                   'send a push
'
'                   update_status_download "Sending push", index
'                   current_download(name_download(pos_name_download).pos_current_down).num_socket = -1
'                   name_download(pos_name_download).num_socket = -1
'                   add_to_waiting_giv pos_name_download
'                   send_a_push pos_name_download
'                   Exit Sub
'                End If
            
'            Else

            'socket closed three possibilities : _
                 download completed, download closed, too many user downloading with the host _
                 in the two last cases need to retry connection
                'search throw name_download the corresponding element associated with this socket

           
                If download_completed(pos_name_download) Then
                'completed
                    update_status_download "Completed", index
                    If all_part_completed(name_download(pos_name_download).saving_name) Then
                        Dim file_name As String
                        file_name = name_download(pos_name_download).saving_name
                        'copy the file from incomplete dir to the download one
                        If is_file_existing(my_incomplete_directory & file_name) Then FileCopy my_incomplete_directory & file_name, my_download_directory & file_name
                        'remove the file in the incomplete dir and it resume file associated
                        If is_file_existing(my_incomplete_directory & file_name) Then Kill my_incomplete_directory & file_name
                        If is_file_existing(my_incomplete_directory & file_name & ".coy") Then Kill my_incomplete_directory & file_name & ".coy"
                        
                        'add to my shared files
                        Dim tmpsize         As Long
                        Dim strtmp          As String
                        Dim table_size      As Variant
                        Dim arr_dir_size    As Long
                        Dim need_to_add     As Boolean
                        If auto_add_down_to_shared_files Then
                                tmpsize = FileLen(my_download_directory & file_name)
                            If Not sharing_simulation Then
                                my_nb_kilobytes_shared = my_nb_kilobytes_shared + tmpsize / 1000
                                my_nb_shared_files = my_nb_shared_files + 1
                            End If
                            table_size = UBound(my_shared_files)
                            'check if file is not overwrite
                            need_to_add = True
                            strtmp = my_download_directory & file_name
                            For cpt = 0 To table_size - 1
                                If my_shared_files(cpt).full_path = strtmp Then
                                    need_to_add = False
                                    Exit For
                                End If
                            Next cpt
                            If need_to_add Then
                                With my_shared_files(table_size)
                                    .file_name = file_name
                                    .file_size = tmpsize
                                    .file_index = table_size
                                    .full_path = strtmp
                                End With
                                ReDim Preserve my_shared_files(table_size + 1)
                            End If
                        End If
                    End If
                    current_download(name_download(pos_name_download).pos_current_down).pos_name_download = -1
                    'remove the element from array name_download
                    remove_from_name_download (pos_name_download)
                Else 'download closed or too many user downloading with the host
                
                    If InStr(1, current_download(name_download(pos_name_download).pos_current_down).status, "Error") < 1 Then
                        Dim strtmp2 As String
                        strtmp2 = "Transfer Interrupted"
                        If InStr(1, current_download(name_download(pos_name_download).pos_current_down).status, "%") Then
                            strtmp2 = strtmp2 & " " & current_download(name_download(pos_name_download).pos_current_down).status
                        End If
                        update_status_download strtmp2, index
                    End If
                    'do the same as a pause
                    Dim pos_selected        As Integer
                    Dim pos_current_down    As Integer
                
                    If Not from_stop_download Then
                        stop_download pos_selected, pos_current_down, pos_name_download, file_name, True, index, True
                    End If
                    
'                    If InStr(1, current_download(name_download(pos_name_download).pos_current_down).status, "Busy") > 1 Then
                    'too many user
                        
                        size = UBound(retry_download)
                        ReDim Preserve retry_download(size + 1)
                        retry_download(size).remaining_time = retry_down_on_busy_server_every
                        retry_download(size).pos_name_download = pos_name_download
'                    End If
                End If
                
'            End If
            
            'clean info for the current socket
            dissociate_current_download_num_socket index
            free_socket (index)

        Case server
            'do nothing but don't free socket
            Exit Sub
        Case Else
            free_socket (index)
    End Select

End Sub

Private Sub free_socket(ByVal index As Integer)
    On Error Resume Next
    Dim cpt As Long
    If Form_main.socket(index).state <> sckClosed Then
        Form_main.socket(index).Close 'sometime socket state stay =disconnecting during few moments so it force them to realy been closed
    End If
    If socket_state(index).connection_type <> free Then
        'decrease nb_current_incomming_socket if necessary
        If socket_state(index).connection_direction = incoming Then
            current_nb_incoming = current_nb_incoming - 1
        End If
        If socket_state(index).connection_type = dialing Then 'avoid to send no sens messages
            For cpt = 0 To UBound(routing_table)
                If routing_table(cpt).num_socket = index Then
                    routing_table(cpt).num_socket = -2
                End If
            Next cpt
        End If
    End If
    remove_from_connected_ip index
    socket_state(index).connection_type = free
End Sub

Public Sub fill_known_gnutella_server()
    On Error Resume Next
    Dim tmp_array       As Variant
    Dim size            As Integer
    Dim cpt             As Integer
    Dim pos             As Integer

    tmp_array = Split(strknown_gnutella_server, ";")
    size = UBound(tmp_array)
    If CStr(tmp_array(size)) = "" Then size = size - 1
    If size < 0 Then
        ReDim tmp_array(0)
        tmp_array(0) = "gnutellahosts.com:6346"
        ReDim known_gnutella_server(0)
    Else
        ReDim known_gnutella_server(size)
    End If
    
    For cpt = 0 To size
        pos = InStr(tmp_array(cpt), ":")
        If pos > 7 Then 'x.x.x.x for ip address
            known_gnutella_server(cpt).ip = Mid$(tmp_array(cpt), 1, pos - 1)
            known_gnutella_server(cpt).port = Mid$(tmp_array(cpt), pos + 1)
        End If
    Next cpt
End Sub

Public Sub connect_to_gnutella_network()
    On Error Resume Next 'avoid no route to host at the .connect
    disconnected_from_gnutella_network = False
    Dim tmp_ip      As String
    Dim num_sock    As Integer
    Dim cpt         As Integer
    For cpt = 0 To UBound(known_gnutella_server)
        num_sock = find_first_free_socket()
        If num_sock > 0 Then
        
            With Form_main.socket(num_sock)
                .LocalPort = 0
                .RemoteHost = known_gnutella_server(cpt).ip
                .RemotePort = known_gnutella_server(cpt).port
                .Connect
                socket_state(num_sock).connection_type = dialing
                socket_state(num_sock).bytes_rcv = 0
                socket_state(num_sock).connection_direction = outgoing
            End With
        
        End If
    Next cpt

    If launch_server Then
        Call make_server
    End If
    Form_main.Timer_minute = True
    Form_main.Timer_second.Enabled = True
End Sub

Public Sub disconnect_from_gnutella_network()
    On Error Resume Next
    'close all sockets (execpted server) and unload them
    Dim lResult As Long
    Dim cpt As Integer
    Dim remaining_down_up As Boolean
    Dim stop_down_up    As Boolean
    disconnected_from_gnutella_network = True
    
    For cpt = UBound(socket_state) To 1 Step -1
        If socket_state(cpt).connection_type = downloading Or socket_state(cpt).connection_type = uploading Then
            remaining_down_up = True
            Exit For
        End If
    Next cpt
    If remaining_down_up Then
        lResult = MessageBox(0, "Do you want to close your upload and download too ?", Form_main.Caption, vbYesNo)
        If lResult = vbYes Then
            stop_down_up = True
        End If
    End If

    If stop_down_up Then
        For cpt = 1 To Form_main.socket.Count - 1
            Form_main.socket(cpt).Close
            treat_socket_closing cpt
        Next cpt
    Else
        For cpt = 1 To Form_main.socket.Count - 1
            If socket_state(cpt).connection_type <> downloading And socket_state(cpt).connection_type <> uploading Then
                Form_main.socket(cpt).Close
                treat_socket_closing cpt
            End If
        Next cpt
    End If
    Form_main.Timer_minute = False
End Sub

Public Sub make_server(Optional ByVal ip As String = "")
    On Error Resume Next
    'initialisation of the server part: allows other users to connect to us and file upload
    If Form_main.socket(0).state <> sckClosed Then Exit Sub
    With Form_main.socket(0)
        .LocalPort = my_port
        If ip = "" Then
            .Bind
        Else
            .Bind , ip
        End If
        .Listen
        If my_ip = "" Then my_ip = .LocalIP
        socket_state(0).connection_type = server
    
    End With
End Sub



'''''''''''''''''''''''''''''''''''functions for forwarding

Public Sub forward(ByVal full_descriptor As String, ByVal forbidden_num_socket As Integer, Optional payload_descriptor As Byte, Optional not_enough_known_hosts As Boolean = False)
    On Error Resume Next
    Dim cpt         As Integer
    Dim bforward    As Boolean

    If Not forward_ping And payload_descriptor = ping And Not not_enough_known_hosts Then Exit Sub
    If Not forward_query And payload_descriptor = query Then Exit Sub

    bforward = True
    For cpt = 1 To Form_main.socket.UBound
        If cpt <> forbidden_num_socket And Not Form_main.socket(cpt).RemoteHostIP <> my_ip _
         And Form_main.socket(cpt).state = sckConnected Then
            
            If forward_on_outgoing_only Then
                If socket_state(cpt).connection_direction = incoming Then
                    bforward = False
                End If
            End If
        
            If bforward Then
                Form_main.socket(cpt).SendData (full_descriptor)
                'add payload to traffic list
                add_to_payload_descriptor_list payload_descriptor, Form_main.socket(cpt).RemoteHostIP, False
            End If
        End If
    Next cpt
End Sub



'''''''''''''''''''''''''''''''''''function for routing

Public Sub route(ByVal full_descriptor As String, descriptor_or_servent_ID As String, ByVal payload_descriptor As Byte)
    On Error Resume Next
    'find the good number of socket to use
    'warning external reference to Form_main containing control array of socket
    Dim num_good_socket     As Integer
    'verify if we should route, and return the number of the socket or -1 if error
    num_good_socket = ask_routing_table(descriptor_or_servent_ID, payload_descriptor)
    If num_good_socket > -1 And num_good_socket <= Form_main.socket.UBound Then
        If Form_main.socket(num_good_socket).state = sckConnected Then
            Form_main.socket(num_good_socket).SendData (full_descriptor) 'send message
            'add payload to traffic list
            add_to_payload_descriptor_list payload_descriptor, Form_main.socket(num_good_socket).RemoteHostIP, False
        End If
    End If
End Sub



Public Sub add_to_my_descriptor_ID(ByVal descriptor_id As String)
    On Error Resume Next
    Dim size                As Long
    Dim place               As Long
    size = UBound(my_descriptors_ID)
    my_descriptors_ID_position = my_descriptors_ID_position + 1
    If size < my_descriptors_ID_max_size Then
        ReDim Preserve my_descriptors_ID(size + 1)
        place = size
    Else
        If my_descriptors_ID_position > my_descriptors_ID_max_size Then
            my_descriptors_ID_position = 0
        End If
        place = my_descriptors_ID_position
    End If
    
    my_descriptors_ID(place) = descriptor_id
End Sub


Public Sub add_to_routing_table(ByVal descriptor_or_servent_ID As String, ByVal payload_descriptor As Byte, ByVal num_socket As Integer)
    On Error Resume Next
        Dim size                As Long
        Dim place               As Long
        size = UBound(routing_table)
        routing_table_position = routing_table_position + 1
        If size < routing_table_max_size Then
            ReDim Preserve routing_table(size + 1)
            place = size
        Else
            If routing_table_position > routing_table_max_size Then
                routing_table_position = 0
            End If
            place = routing_table_position
        End If
        
        routing_table(place).ID = descriptor_or_servent_ID
        routing_table(place).payload = payload_descriptor
        routing_table(place).num_socket = num_socket
End Sub


Public Function ask_routing_table(ByVal descriptor_or_servent_ID As String, ByVal payload_descriptor As Byte) As Integer
    On Error Resume Next
    'on error returns -1 else return the socket number
    Dim cpt                 As Integer
    'invert payload descriptor (if we receive a pong we should look for a ping
    Select Case payload_descriptor
        Case pong
            payload_descriptor = ping 'signal we are looking at
        Case queryhit
            payload_descriptor = query
        Case push
            payload_descriptor = queryhit
    End Select
    ask_routing_table = -1 'suppose not found
    For cpt = 0 To UBound(routing_table)
        If routing_table(cpt).ID = descriptor_or_servent_ID _
           And routing_table(cpt).payload = payload_descriptor _
        Then
            ask_routing_table = routing_table(cpt).num_socket
            Exit Function
        End If
    Next cpt
End Function


Public Function routing_table_checked(descriptorID As String, payload_descriptor As Byte) As Boolean
    On Error Resume Next
    'False if we have already seen a ping or query
    'or if we haven't seen the corresponding ping/query/queryhit for the pong/queryhit/push
    'to kill the lost descriptors
    Dim cpt         As Long
    Dim mypayload   As Byte
    Select Case payload_descriptor
        Case ping, query  'if we have already seen them this means there's a loop on the network
                           'we must remove them from the network (no route,no forward)
            For cpt = 0 To UBound(routing_table) - 1
                If routing_table(cpt).ID = descriptorID Then
                    routing_table_checked = False
                    Exit Function
                End If
            Next cpt
            routing_table_checked = True
            Exit Function
        Case Else 'we must have seen the corresponding query (query or ping)
            Select Case payload_descriptor
                Case pong
                    mypayload = ping
                Case queryhit
                    mypayload = query
                Case push
                    mypayload = queryhit
                Case Else 'bogus packet
                    routing_table_checked = False
                    Exit Function
            End Select
            
            For cpt = 0 To UBound(routing_table) - 1
                If routing_table(cpt).ID = descriptorID And routing_table(cpt).payload = mypayload Then
                    routing_table_checked = True
                    Exit Function
                End If
            Next cpt
            'query not found
            routing_table_checked = False
            Exit Function
    End Select
    
End Function
                    




''''''''''''''''''''''''''''''''''''''''' sending functions

Public Function send_data(ByVal data As String, ByVal num_socket As Integer) As Boolean
    On Error Resume Next
    'return : true if ok,false on error
    If num_socket < 0 Then Exit Function
    If Form_main.socket(num_socket).state = sckConnected Then
        send_data = True
        Form_main.socket(num_socket).SendData data
        Exit Function
    Else
        send_data = False
    End If
End Function

Public Sub send_a_ping(num_socket As Integer, Optional ttl As Byte)
    On Error Resume Next
    Dim strtmp                  As String
    Dim str_descriptor_id       As String * 16
    str_descriptor_id = GetGUID() 'make a new dexcriptor_id
    If ttl < 1 Then ttl = my_ttl
    'send a ping
    strtmp = make_string_descriptor_data(str_descriptor_id, ping, my_ttl, 0, 0)
    add_to_my_descriptor_ID str_descriptor_id

    If Form_main.socket(num_socket).state = sckConnected Then
        Form_main.socket(num_socket).SendData strtmp
        'add payload to traffic list
        add_to_payload_descriptor_list ping, Form_main.socket(num_socket).RemoteHostIP, False
    End If

End Sub

Public Sub send_a_push(ByVal pos_name_download As Integer)
    On Error Resume Next
    Dim strdata         As String
    Dim num_socket_queryhit As Integer
    Dim strpush         As String
    Dim data            As String
    Dim lResult As Long
    
    num_socket_queryhit = name_download(pos_name_download).num_socket_queryhit
    
    If name_download(pos_name_download).num_socket_queryhit < 0 Then
        lResult = MessageBox(0, "File couldn't be download on this host," & vbCrLf & "you must make a resume on other host.", Form_main.Caption, vbExclamation)
    Else
        
        strpush = make_string_push_data(name_download(pos_name_download).servent_id, name_download(pos_name_download).file_index, my_ip, my_port)
        data = make_string_descriptor_data(my_servent_id, push, my_ttl, 0, 26)
        name_download(pos_name_download).get_request_made = False
        data = data & strpush
        
        'add payload to traffic list
        add_to_payload_descriptor_list push, Form_main.socket(num_socket_queryhit).RemoteHostIP, False, False, " for " & name_download(pos_name_download).file_name
        send_data data, num_socket_queryhit
        Exit Sub
    End If
End Sub

Public Sub send_to_all_dialing(ByVal data As String)
    On Error Resume Next
    'used only for search
    Dim cpt                     As Long
    'For cpt = 1 To UBound(socket_state)
    For cpt = 1 To Form_main.socket.UBound   'socket 0 is reserved for the server
        If socket_state(cpt).connection_type = dialing Then
            If (send_data(data, cpt)) Then
                ' add to traffic list
                add_to_payload_descriptor_list query, Form_main.socket(cpt).RemoteHostIP, False
            End If
        End If
    Next cpt
End Sub




Public Function find_first_free_socket() As Integer
    On Error Resume Next
    'return the number of the first free socket, and if no more free load another socket
    'return -1 if Form_main.socket.UBound < max_nb_socket
    Dim cpt As Integer
    For cpt = 1 To UBound(socket_state)
        If socket_state(cpt).connection_type = free And Form_main.socket(cpt).state = sckClosed Then
            remaining_data(cpt) = ""
            find_first_free_socket = cpt
            Exit Function
        End If
    Next cpt
    'no free socket
    Dim num_socket As Integer
    num_socket = Form_main.socket.UBound + 1
    If Form_main.socket.UBound < max_nb_socket Then
        Load Form_main.socket(num_socket)
        ReDim Preserve socket_state(UBound(socket_state) + 1)
        ReDim Preserve remaining_data(UBound(remaining_data) + 1)
    Else
        num_socket = -1
    End If
    find_first_free_socket = num_socket
End Function



Public Function can_i_dial_with(ip As String) As Boolean
    On Error Resume Next
    'search if we can dial with host
    'return false if we are dialing with this ip or if it is a forbidden ip
    Dim cpt                 As Integer
    Dim pos_star            As Byte
    Dim strtmp1             As String
    Dim strtmp2             As String
    

    If ip = "" Then
        can_i_dial_with = False
        Exit Function
    End If
    'check if it's not our ip
    If ip = Form_main.socket(0).LocalIP Then
        can_i_dial_with = False
        Exit Function
    End If
    
    'test if it's not a dummy ip
    If is_ip_forbidden(ip) Then
        can_i_dial_with = False
        Exit Function
    End If
    
    'test if it'not a banised ip
    If is_ip_banished(ip) Then
        can_i_dial_with = False
        Exit Function
    End If
    
    'test if we are already dialing with the host
    For cpt = 0 To UBound(socket_state) 'search through loaded sockets
        If socket_state(cpt).connection_type = dialing Then
            If ip = Form_main.socket(cpt).RemoteHostIP _
                Or ip = Form_main.socket(cpt).RemoteHost _
            Then
                can_i_dial_with = False
                Exit Function
            End If
        End If
    Next cpt

    can_i_dial_with = True
End Function

Public Function is_ip_forbidden(ByVal ip As String) As Boolean
    On Error Resume Next
    'true if forbidden
    'we don't try to connect to sample 0.0.0.0 192.168.*.*
    Dim cpt     As Integer
    Dim strtmp1 As String
    Dim strtmp2 As String
    Dim pos_star As Integer
    
    For cpt = 0 To UBound(dummy_ip) 'search through array dummy_ip
        strtmp2 = ip
        strtmp1 = dummy_ip(cpt)
        pos_star = InStr(1, strtmp1, "*")
        If pos_star > 0 Then ' no * in dummy ip
            strtmp2 = Mid$(ip, 1, pos_star - 1)
            strtmp1 = Mid$(dummy_ip(cpt), 1, pos_star - 1)
        End If
        
        If strtmp2 = strtmp1 Then
            is_ip_forbidden = True
            Exit Function
        End If
    Next cpt
    
End Function

Public Function is_ip_banished(ByVal ip As String) As Boolean
    On Error Resume Next
    'banish for upload
    'host will be rejected if it tries to connect to you and there will be no answer to its pushs
    Dim cpt As Integer
    For cpt = 0 To UBound(banished_ip) - 1
        If ip = banished_ip(cpt) Then
            is_ip_banished = True
            Exit Function
        End If
    Next cpt
    
End Function

Public Sub make_new_dial(ByVal ip As String, ByVal port As String)
    
    'load new socket if necessary
    'and make connection if it can be established
    On Error Resume Next ' must be present if we try to connect to stupid ip like 0.0.0.0
    Dim num_socket              As Integer
    If disconnected_from_gnutella_network Then Exit Sub
    num_socket = find_first_free_socket()
    If num_socket > 0 And can_i_dial_with(ip) Then
        With Form_main.socket(num_socket)
            .Close
            .RemoteHost = ip
            .RemotePort = port
            .LocalPort = 0
            If Form_main.socket(num_socket).state = sckClosed Then
                .Connect
            End If
        End With
        socket_state(num_socket).connection_type = dialing
    Else
        If num_socket > 0 Then 'socket loaded
            socket_state(num_socket).connection_type = free
        End If
    End If
End Sub

Public Sub make_new_random_dial(Optional send_ping_if_not_enought_known_host As Boolean = False, Optional connect_to_latest_known_host As Boolean = False, Optional really_random As Boolean = False)
    On Error Resume Next
    'make a dial with a known host we are not dialing with
    'if not enought known host send ping to all known host
    Dim cpt                 As Long
    Dim real_random         As Long
    Dim host                As thost
    Dim tsize               As Long
    
    
    If current_nb_dial >= min_dialing_hosts Then Exit Sub 'we are already dialing with enougth hosts
    
    If Not connect_to_latest_known_host Then
        host = preferently_connect_to_sharing_host() ' try connection
        If host.ip <> "" Then
            make_new_dial host.ip, host.port
            Exit Sub
        End If
    End If
    
    tsize = UBound(known_hosts)
    
    If really_random Then
        real_random = Int((tsize + 1) * Rnd)
        For cpt = real_random To tsize 'search through known_hosts (from the most recent to the oldest)
            If can_i_dial_with(known_hosts(cpt).ip) Then
                make_new_dial known_hosts(cpt).ip, known_hosts(cpt).port
                Exit Sub
            End If
        Next cpt
    Else
    
        For cpt = tsize To 0 Step -1 'search through known_hosts (from the most recent to the oldest)
            If can_i_dial_with(known_hosts(cpt).ip) Then
                make_new_dial known_hosts(cpt).ip, known_hosts(cpt).port
                Exit Sub
            End If
        Next cpt
    End If
    
    'if not enought known host send ping to all dialing sockets
    Dim strtmp                  As String
    Dim str_descriptor_id       As String * 16
    
    If send_ping_if_not_enought_known_host Then
        str_descriptor_id = GetGUID() 'make a new dexcriptor_id
                
        strtmp = make_string_descriptor_data(str_descriptor_id, ping, my_ttl, 0, 0)
        add_to_my_descriptor_ID str_descriptor_id
        'forward with no forbiden ip on outgoing connections
        forward strtmp, 0, ping, True
    End If
    
End Sub


Public Function preferently_connect_to_sharing_host() As thost
    'return the first ip with sharing files>0 we are not connected to
    'return "" if no ip found
    On Error Resume Next
    Dim cpt                     As Long
    For cpt = UBound(known_hosts) To 0 Step -1 ' begin with last known hosts
        If known_hosts(cpt).nb_shared_files > 0 Then
            If can_i_dial_with(known_hosts(cpt).ip) Then
                preferently_connect_to_sharing_host = known_hosts(cpt)
                Exit Function
            End If
        End If
    Next cpt
End Function

Public Sub ban_ip(ip As String)
    On Error Resume Next
    Dim position            As Integer
    'close all current connection and treat socket closing
    For position = 1 To Form_main.socket.Count - 1
        If Form_main.socket(position).RemoteHostIP = ip Then
            Form_main.socket(position).Close
            treat_socket_closing position
        End If
    Next position
    
    'add to banished_ip
    position = UBound(banished_ip)
    banished_ip(position) = ip
    ReDim Preserve banished_ip(position + 1)
End Sub

