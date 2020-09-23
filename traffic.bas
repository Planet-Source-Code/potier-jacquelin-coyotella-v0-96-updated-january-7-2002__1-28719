Attribute VB_Name = "traffic"
Option Explicit

Public Sub add_to_connected_ip(ByVal num_socket As Integer)
    On Error Resume Next
    'add item to Form_traffic.ListView_connected_ip
    Dim last_pos            As Integer 'position in the listview
    Dim pos                 As Integer 'position in the array
    Dim strtmp              As String
    
    pos = UBound(traffic_connected_ip)
    
    
        last_pos = Form_traffic.ListView_connected_ip.ListItems.Count + 1
        Form_traffic.ListView_connected_ip.ListItems.Add last_pos, , Form_main.socket(num_socket).RemoteHostIP
        Form_traffic.ListView_connected_ip.ListItems(last_pos).SubItems(1) = Form_main.socket(num_socket).RemotePort
        Form_traffic.ListView_connected_ip.ListItems(last_pos).Tag = num_socket 'allow to find elements faster

        traffic_connected_ip(pos).ip = Form_main.socket(num_socket).RemoteHostIP
        traffic_connected_ip(pos).port = Form_main.socket(num_socket).RemotePort
        traffic_connected_ip(pos).num_socket = num_socket
        
        If socket_state(num_socket).connection_direction = incoming Then
            Form_traffic.ListView_connected_ip.ListItems(last_pos).SubItems(2) = "in"
            traffic_connected_ip(pos).incoming = "in"
        Else
            Form_traffic.ListView_connected_ip.ListItems(last_pos).SubItems(2) = "out"
            traffic_connected_ip(pos).incoming = "out"
        End If
        Select Case socket_state(num_socket).connection_type
            'Case free
            Case dialing
                strtmp = "Dialing"
            Case uploading
                strtmp = "Uploading"
            Case downloading
                strtmp = "Downloading"
            'Case server
            Case not_define_yet
                strtmp = "Not define yet"
            Case giving
                strtmp = "Giving"
        End Select
        Form_traffic.ListView_connected_ip.ListItems(last_pos).SubItems(3) = strtmp
        traffic_connected_ip(pos).state = strtmp
    
    ReDim Preserve traffic_connected_ip(pos + 1)
End Sub


Public Sub remove_from_connected_ip(ByVal num_socket As Integer)
    On Error Resume Next
    'remove item from Form_traffic.ListView_connected_ip
    Dim cpt                 As Integer
    Dim pos                 As Integer
    Dim place               As Integer

    place = -1
    
    pos = UBound(traffic_connected_ip)
    If pos = 0 Then Exit Sub 'no connection
    'remove from array
    
    'find the place <-> tag item
    'put last element at the place of the removed one, and resize array
    For cpt = 0 To pos - 1
        If traffic_connected_ip(cpt).num_socket = num_socket Then
            traffic_connected_ip(cpt) = traffic_connected_ip(pos - 1) 'stupid if cpt=pos-1 but the next line is necessary
            place = cpt
            Exit For
        End If
    Next cpt

    If place < 0 Then Exit Sub ' error : not found (already removed)

    ReDim Preserve traffic_connected_ip(pos - 1)
    For cpt = Form_traffic.ListView_connected_ip.ListItems.Count To 1 Step -1
        If Form_traffic.ListView_connected_ip.ListItems.Item(cpt).Tag = num_socket Then
           Form_traffic.ListView_connected_ip.ListItems.Remove cpt
           Exit For
        End If
    Next cpt

    
    If socket_state(num_socket).connection_type = dialing Then
        current_nb_dial = current_nb_dial - 1
        If current_nb_dial <= 0 Then
            current_nb_dial = 0
            connected_to_gnutella_network = False
            Form_main.StatusBar1.Panels(1).Picture = Form_main.imlToolbar.ListImages(11).Picture
            Form_main.StatusBar1.Panels(1).Text = "Disconnected"
            Form_main.Toolbar1.Buttons(1).Enabled = True
            Form_main.Toolbar1.Buttons(2).Enabled = False
            If Not disconnected_from_gnutella_network Then
                'make new connection
                make_new_dial known_gnutella_server(0).ip, known_gnutella_server(0).port 'connect to the first known gnutella host
                make_new_random_dial True, True 'connect to latest known host
            End If
        End If
    End If

End Sub

Public Sub clarify_connection(num_socket As Integer, state As String)
    On Error Resume Next
    Dim cpt As Integer
    Dim cpt2 As Integer

    For cpt = 0 To UBound(traffic_connected_ip) - 1
        With traffic_connected_ip(cpt)
            If num_socket = .num_socket Then
                For cpt2 = 1 To Form_traffic.ListView_connected_ip.ListItems.Count
                    If Form_traffic.ListView_connected_ip.ListItems(cpt2).Tag = num_socket Then
                        Form_traffic.ListView_connected_ip.ListItems(cpt2).SubItems(3) = state
                        traffic_connected_ip(cpt).state = state
                        Exit Sub
                    End If
                Next cpt2
            End If
         End With
    Next cpt

End Sub


Public Sub add_to_payload_descriptor_list(payload_descriptor As Byte, ip As String, incomming_payload As Boolean, Optional killed_payload As Boolean = False, Optional optional_data As String = "")
'add item to Form_traffic.ListView_current_payload
    On Error Resume Next
    Dim strtmp              As String
    Dim strtmp2             As String
    Dim num_img_liste       As Integer
    
    add_to_traffic_nb_payload payload_descriptor, incomming_payload
    
    If incomming_payload Then
        strtmp = "from "
    Else
        strtmp = "to "
    End If
    Select Case payload_descriptor
        Case ping
            strtmp2 = "ping"
            num_img_liste = 1
        Case pong
            strtmp2 = "pong"
            num_img_liste = 2
        Case push
            strtmp2 = "push"
            num_img_liste = 3
        Case query
            strtmp2 = "query"
            num_img_liste = 4
        Case queryhit
            strtmp2 = "queryhit"
            num_img_liste = 5
        Case Else
            strtmp2 = "bogus payload=" & CStr(payload_descriptor)
            num_img_liste = 6
    End Select
    
    If killed_payload Then
        strtmp2 = strtmp2 & " (killed)"
        num_img_liste = 6
    End If
    
    
    
    'add changes to form
    With Form_traffic.ListView_current_payload
        .ListItems.Add 1, , strtmp2, , num_img_liste
        .ListItems(1).SubItems(1) = strtmp & ip & optional_data
    End With

    If traffic_info.should_remove Then
        Form_traffic.ListView_current_payload.ListItems.Remove traffic_info.array_size + 1
    End If
    
    'add to array
    traffic_payload(traffic_info.position).payload_descriptor = payload_descriptor
    traffic_payload(traffic_info.position).info = strtmp & ip & optional_data
    
    Dim cpt             As Integer
    If traffic_info.position = traffic_info.array_size - 1 Then
        traffic_info.position = 0
        traffic_info.should_remove = True
    Else
        traffic_info.position = traffic_info.position + 1
    End If
    
End Sub


Public Sub add_to_traffic_nb_payload(payload_descriptor As Byte, incomming_payload As Boolean)
    On Error Resume Next
    Select Case payload_descriptor
        Case ping
            If incomming_payload Then
                traffic_info.rcv_ping = traffic_info.rcv_ping + 1
                Form_traffic.in_ping.Caption = traffic_info.rcv_ping
            Else
                traffic_info.sent_ping = traffic_info.sent_ping + 1
                Form_traffic.out_ping.Caption = traffic_info.sent_ping
            End If
        Case pong
            If incomming_payload Then
                traffic_info.rcv_pong = traffic_info.rcv_pong + 1
                Form_traffic.in_pong.Caption = traffic_info.rcv_pong
            Else
                traffic_info.sent_pong = traffic_info.sent_pong + 1
                Form_traffic.out_pong.Caption = traffic_info.sent_pong
            End If
        Case push
            If incomming_payload Then
                traffic_info.rcv_push = traffic_info.rcv_push + 1
                Form_traffic.in_push.Caption = traffic_info.rcv_push
            Else
                traffic_info.sent_push = traffic_info.sent_push + 1
                Form_traffic.out_push.Caption = traffic_info.sent_push
            End If
        Case query
            If incomming_payload Then
                traffic_info.rcv_query = traffic_info.rcv_query + 1
                Form_traffic.in_query.Caption = traffic_info.rcv_query
            Else
                traffic_info.sent_query = traffic_info.sent_query + 1
                Form_traffic.out_query.Caption = traffic_info.sent_query
            End If
        Case queryhit
            If incomming_payload Then
                traffic_info.rcv_queryhit = traffic_info.rcv_queryhit + 1
                Form_traffic.in_queryhit.Caption = traffic_info.rcv_queryhit
            Else
                traffic_info.sent_queryhit = traffic_info.sent_queryhit + 1
                Form_traffic.out_queryhit.Caption = traffic_info.sent_queryhit
            End If
        Case Else
            traffic_info.rcv_bogus = traffic_info.rcv_bogus + 1
            Form_traffic.in_bogus.Caption = traffic_info.rcv_bogus
    End Select
End Sub

Public Sub update_traffic_speed(ByVal speed As Single, ByVal num_socket As Integer, ByVal pos_traffic_connected_ip As Integer)
    On Error Resume Next
    Dim cpt             As Integer
    With Form_traffic.ListView_connected_ip.ListItems
        For cpt = 1 To .Count
            If .Item(cpt).Tag = num_socket Then
                .Item(cpt).SubItems(4) = traffic_connected_ip(pos_traffic_connected_ip).speed
            End If
        Next cpt
    End With
End Sub
