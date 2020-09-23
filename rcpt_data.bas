Attribute VB_Name = "rcpt_data"
Option Explicit

Public Sub treat_data(strdata As String, ByVal index As Integer, ByVal bytestotal As Long)
    On Error Resume Next
    If Len(strdata) = 0 Then Exit Sub
    
    If UCase(Mid$(strdata, 1, 17)) = "GNUTELLA CONNECT/" Then
        'connection request
        socket_state(index).connection_type = dialing
        clarify_connection index, "Dialing"
        current_nb_dial = current_nb_dial + 1 'should be put here because current_nb_dial is decremented at the closing of the socket
        If current_nb_incoming <= max_incoming_connection Then
            
            If current_nb_dial = 1 Then
                connected_to_gnutella_network = True
                Form_main.StatusBar1.Panels(1).Picture = Form_main.imlToolbar.ListImages(12).Picture
                Form_main.StatusBar1.Panels(1).Text = "Connected"
                Form_main.Toolbar1.Buttons(1).Enabled = False
                Form_main.Toolbar1.Buttons(2).Enabled = True
            End If
        Else
            Form_main.socket(index).Close
            treat_socket_closing index
        End If
        Exit Sub
    End If
    
    If UCase(Mid$(strdata, 1, 11)) = "GNUTELLA OK" Then
        'response of one of our connection
        send_a_ping index, 1
        Exit Sub
    End If
    
    If Mid$(strdata, 1, 9) = "GET /get/" Then
        'asked for download
        If socket_state(index).connection_type = uploading Then     'not giving
            current_nb_upload = current_nb_upload + 1
        End If
        socket_state(index).connection_type = uploading
        clarify_connection index, "Uploading"
        
        If current_nb_upload < max_upload + 1 And allow_upload Then
            Dim info                    As file_get_info
            info = decode_get(strdata)
            'upload the file
            upload_file info.file_index, info.file_name, info.range, index
        Else
            Form_main.socket(index).SendData "HTTP 1.1 503 service unavailable" & vbCrLf 'the send_complete event will close the socket
        End If
        Exit Sub
    End If
    
    If socket_state(index).connection_type = downloading Then 'Mid$(strdata, 1, 13) =  "HTTP 200 OK"
    'response of one of our download request
        'write file
        decode_download strdata, index, bytestotal
        Exit Sub
    End If
    
    If Mid$(strdata, 1, 4) = "GIV " Then
        'make get request
        Dim str_get_request         As String
        Dim info2                   As file_giv_info
        Dim ret                     As Integer
        Dim strrange                As String
            
        info2 = decode_giv(strdata, Form_main.socket(index).RemoteHostIP)
        
        ret = asked_giv(info2.servent_id, info2.file_index)
        If ret > -1 Then
            If name_download(ret).range_end = 0 Then
                strrange = CStr(CLng(name_download(ret).range_begin)) & "-"
            Else
                strrange = CStr(CLng(name_download(ret).range_begin)) & "-" & CStr(CLng(name_download(ret).range_end))
            End If
            str_get_request = make_get_request(info2.file_index, info2.file_name, strrange)
            name_download(ret).get_request_made = True
            name_download(ret).num_socket = index
            current_download(name_download(ret).pos_current_down).num_socket = index
            'send get
            send_data str_get_request, index
            socket_state(index).connection_type = downloading

            current_nb_download = current_nb_download + 1
            clarify_connection index, "Downloading"
        Else
            Form_main.socket(index).Close
            treat_socket_closing (index)
        End If
        Exit Sub
    End If
    
    'in all over case
     find_descriptors strdata, index
End Sub


Private Sub find_descriptors(strdata As String, num_socket As Integer)
    'because one data arrival contain multipledescriptors
    On Error Resume Next
    
    If remaining_data(num_socket) <> "" Then
        strdata = remaining_data(num_socket) & strdata
        remaining_data(num_socket) = ""
    End If

    If Len(strdata) < 23 Then Exit Sub

    Dim descriptord                 As descriptor_data
    descriptord = decode_descriptor_data(Mid$(strdata, 1, 23))
    
    Dim bogus_payload_value         As Byte
    If descriptord.payload_length > 4294967294# Or unknown_payload_descriptor(descriptord.payload_descriptor) Then
        'bogus payload
        If descriptord.payload_descriptor <> 0 Then 'not initialized
            bogus_payload_value = descriptord.payload_descriptor
        Else
            bogus_payload_value = -1
        End If
'Debug.Print strdata
        add_to_payload_descriptor_list bogus_payload_value, Form_main.socket(num_socket).RemoteHostIP, True
        socket_state(num_socket).number_of_bogus = socket_state(num_socket).number_of_bogus + 1
        If socket_state(num_socket).number_of_bogus >= mymax_bogus_per_minute Then
            Form_main.socket(num_socket).Close
            treat_socket_closing num_socket
        End If
        Exit Sub
    End If
    
    'check if there's at least one full descriptor
    If descriptord.payload_length <= Len(strdata) - 23 Then
        'treat the descriptor
        decode_descriptor Mid$(strdata, 1, 23 + descriptord.payload_length), num_socket, descriptord
        'search if ther's other descriptors
        find_descriptors Mid$(strdata, 24 + descriptord.payload_length), num_socket
    Else
        'keep datas in memory to treat them after the next data arrival for this socket
        remaining_data(num_socket) = remaining_data(num_socket) & strdata
    End If

End Sub

Private Function unknown_payload_descriptor(num_payload As Byte) As Boolean
    On Error Resume Next
    unknown_payload_descriptor = True
    If num_payload = ping Or num_payload = pong Or num_payload = query _
        Or num_payload = queryhit Or num_payload = push _
    Then
        unknown_payload_descriptor = False
    End If
End Function
