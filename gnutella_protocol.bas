Attribute VB_Name = "gnutella_protocol"
Option Explicit

'for gnutella protocol
Public Const ping       As Byte = 0         '0x00
Public Const pong       As Byte = 1         '0x01
Public Const push       As Byte = 64        '0x40
Public Const query      As Byte = 128       '0x80
Public Const queryhit   As Byte = 129       '0x81

'all types are according to the gnutella protocol specification v0.4

Public Type str_descriptor_data
    descriptor_id         As String * 16  'guid number
    payload_descriptor    As String * 1        'message type:
                                                      '0x00=ping
                                                      '0x01=pong
                                                      '0x40=push
                                                      '0x80=query
                                                      '0x81=queryhit
    ttl                   As String * 1   'time to live
    hops                  As String * 1   'current number of hops
    payload_length        As String * 4   'parameter length
End Type



Public Type str_pong_data
    port                As String * 2
    ip                  As String * 4
    nb_shared_files     As String * 4
    nb_shared_kbytes    As String * 4
End Type

Public Type str_query_data
    min_speed           As String * 2
    search_criteria     As String
End Type

Public Type str_ttrailer
    vendor_code         As String * 4
    open_data_size      As String * 1
    open_data_flag1     As String * 1
    open_data_flag2     As String * 1
    private_data        As String 'undocumented in sp 0.4
End Type

Public Type str_queryhit_data 'could be changed for bearshare extra data (trailer field) (if you have specification)
    nb_hits             As String * 1
    port                As String * 2
    ip                  As String * 4
    speed               As String * 4
    all_str_result_set  As String
    trailer             As str_ttrailer
    servent_id          As String * 16
End Type


Public Type str_push_data
    servent_id          As String * 16
    file_index          As String * 4
    ip                  As String * 4
    port                As String * 2
End Type

Public Type str_result_set 'could be changed for gnotella extra data (if you have specification)
    file_index          As String * 4
    file_size           As String * 4
    file_name           As String
End Type

'for file transfer
Public Type str_file_get_info
    file_index          As Single
    file_name           As String
    range               As String
End Type

Public Type str_file_giv_info
    file_index          As String * 4
    servent_id          As String * 16
    file_name           As String
End Type



'same type but not in string


Public Type descriptor_data
  descriptor_id         As String * 16  'guid number
  payload_descriptor    As Byte         'message type:
                                                    '0x00=ping
                                                    '0x01=pong
                                                    '0x40=push
                                                    '0x80=query
                                                    '0x81=queryHit
  ttl                   As Byte         'time to live
  hops                  As Byte         'current number of hops
  payload_length        As Single       'parameter length
End Type



Public Type pong_data
    port                As Long
    ip                  As String
    nb_shared_files     As Single
    nb_shared_kbytes    As Single
End Type

Public Type query_data
    min_speed           As Long
    search_criteria     As String
End Type

Public Type ttrailer
    vendor_code         As String * 4
    open_data_size      As Byte   '=2 in sp 0.4
    open_data_flag1     As Byte   '
    open_data_flag2     As Byte   'should be private data but as we know open_data_size=2
    private_data        As String 'undocumented in sp 0.4
End Type

Public Type queryhit_data
    nb_hits             As Byte
    port                As Long
    ip                  As String
    speed               As Single
    all_result_set      As String
    trailer             As ttrailer
    servent_id          As String * 16
End Type

Public Type push_data
    servent_id          As String * 16
    file_index          As Single
    ip                  As String
    port                As Long
End Type

Public Type result_set 'could be changed for gnotella extra data (if you have specification)
    file_index          As Single
    file_size           As Single
    file_name           As String
End Type

'for file transfer
Public Type file_get_info
    file_index          As Single
    file_name           As String
    range               As String
End Type

Public Type file_giv_info
    file_index          As Single
    servent_id          As String * 16
    file_name           As String
    range               As String
End Type








''''''''''''''' decode & treat received messages
''''''''''''''' gnutella dialing main function

Public Sub decode_descriptor(descriptor As String, ByVal num_socket As Integer, descriptord As descriptor_data)
    On Error Resume Next
    Dim descriptorstr               As String
    Dim need_to_send                As Boolean
    'need_to_send : true if we need to send the message to other host(s)
    '(it means it's not one of our descriptor and its ttl is >0 )
    Dim one_of_my_descriptor        As Boolean
    Dim init_ttl                    As Integer
    Dim should_be_kill              As Boolean
   
   init_ttl = descriptord.ttl
   init_ttl = init_ttl + descriptord.hops
   If init_ttl > 255 Then Exit Sub 'bogus
    
    
    If descriptord.payload_descriptor = pong Or descriptord.payload_descriptor = queryhit Then
        'search if it is one of our descriptor or not
        one_of_my_descriptor = is_one_of_my_descriptor_ID(descriptord.descriptor_id)
    End If
    
    If Not one_of_my_descriptor Then
        'verify we need to treat the payload
        '(check if we have already received it or if we don't have seen the query for the answer or the queryhit for the push)
        If Not routing_table_checked(descriptord.descriptor_id, descriptord.payload_descriptor) Then
            'add the payload descriptor to the traffic form with the killed field
            should_be_kill = True
            add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True, True
        End If
    End If

    If should_be_kill Then
        need_to_send = False
    Else
        If one_of_my_descriptor Then
            need_to_send = False ' we don't need to forward or route it
        Else
            'check ttl
            If descriptord.ttl <= 0 Then
                need_to_send = False
            Else
                'decrease ttl and increase hops
                descriptord.ttl = descriptord.ttl - 1
                descriptord.hops = descriptord.hops + 1
                need_to_send = True
            End If
    
            'rebuild descriptor with new ttl and hop
            descriptorstr = make_string_descriptor_data(descriptord.descriptor_id, descriptord.payload_descriptor, descriptord.ttl, descriptord.hops, descriptord.payload_length)
        End If
    End If
    
    If need_to_send Then
        If init_ttl > max_initial_ttl Then
            need_to_send = False
        End If
    End If
    
    Select Case descriptord.payload_descriptor
        Case ping
            If Not should_be_kill Then
                'add the payload descriptor to the traffic form
                add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True
            
                'check if there's not to much traffic on this connection
                socket_state(num_socket).number_of_ping = socket_state(num_socket).number_of_ping + 1
                If socket_state(num_socket).number_of_ping >= mymax_ping_per_minute Then
                    ban_ip Form_main.socket(num_socket).RemoteHostIP
                    Exit Sub
                End If
                
                'add to routing_table (routed with descriptor_ID)
                add_to_routing_table descriptord.descriptor_id, descriptord.payload_descriptor, num_socket
                
                'treat ping
                treat_ping descriptord, num_socket
                
                If forward_ping Then
                    'forward ping if necessary
                    If need_to_send Then forward descriptorstr, num_socket, ping 'no data in a ping
                End If
            End If
        Case pong
            Dim pongd               As pong_data
            If Not should_be_kill Then
                'add the payload descriptor to the traffic form
                add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True
            End If
            'treat pong in any case
            
            'decode pong
            pongd = decode_pong_data(Mid$(descriptor, 24))
            
            If Not is_ip_forbidden(pongd.ip) Then
                'treat pong (get informations)
                treat_pong pongd, num_socket
                'route pong if necessary
                If need_to_send Then route descriptorstr + Mid$(descriptor, 24), descriptord.descriptor_id, descriptord.payload_descriptor
            End If
            
        Case push
            Dim pushd               As push_data
            If Not should_be_kill Then
                'add the payload descriptor to the traffic form
                add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True
            
                'decode push 'needed before adding to routing_table because this operation need sevent_ID
                pushd = decode_push_data(Mid$(descriptor, 24))
                
                'treat push (make connection if necessary)
                If pushd.servent_id = my_servent_id Then treat_push pushd
                
                'route push
                If need_to_send Then
                    route descriptorstr + Mid$(descriptor, 24), descriptord.descriptor_id, descriptord.payload_descriptor
                    'add to routing_table (routed with sevent_ID <--- WARNING )
                    add_to_routing_table pushd.servent_id, descriptord.payload_descriptor, num_socket
                End If
            End If
        Case query
            If Not should_be_kill Then
                If Not allow_upload Then Exit Sub 'no need to reply for refuse connection next
    
                Dim queryd                  As query_data
    
                'add to routing_table (routed with descriptor_ID)
                add_to_routing_table descriptord.descriptor_id, descriptord.payload_descriptor, num_socket
                
                If forward_query Then
                    'forward query if necessary (send descriptor with new ttl and the same query data)
                    If need_to_send Then forward descriptorstr + Mid$(descriptor, 24), num_socket, query
                End If
                
                'decode query
                If Len(descriptor) < 4 Then Exit Sub
                queryd = decode_query_data(Mid$(descriptor, 24))
                            
                'add the payload descriptor to the traffic form
                add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True, False, " for " & queryd.search_criteria
                
                'check if there's not to much traffic on this connection
                socket_state(num_socket).number_of_query = socket_state(num_socket).number_of_query + 1
                If socket_state(num_socket).number_of_query >= mymax_query_per_minute Then
                    ban_ip Form_main.socket(num_socket).RemoteHostIP
                    Exit Sub
                End If
                
                
                'reply with queryhit if speed>min speed and search is long enough
                If Len(queryd.search_criteria) < other_min_search Then Exit Sub
                If my_speed >= CLng(queryd.min_speed) And my_nb_shared_files > 0 Then
                    'create queryhit data
                    treat_query num_socket, descriptord, queryd
                End If
            End If
            
        Case queryhit
            Dim queryhitd2              As queryhit_data
            Dim result_setd()           As result_set
            If Not should_be_kill Then
                'add the payload descriptor to the traffic form
                add_to_payload_descriptor_list descriptord.payload_descriptor, Form_main.socket(num_socket).RemoteHostIP, True
            End If
            'treat queryhit in any case
            'decode queryhit
            If Len(descriptor) < 26 Then Exit Sub
            queryhitd2 = decode_queryhit_data(Mid$(descriptor, 24))

            'treat queryhit if we are concerned
            If one_of_my_descriptor Or spy_all_query_hits Then treat_queryhit queryhitd2, descriptord.descriptor_id, num_socket
                
            'route queryhit if necessary
            If need_to_send Then route descriptorstr + Mid$(descriptor, 24), descriptord.descriptor_id, descriptord.payload_descriptor
            
            
        Case Else 'error
            send_data "<html><body>ERROR: This is a gnutella client on port" & CStr(my_port) & "</body></html>", num_socket
    End Select
End Sub

'''''''''''''''''''''''''''''' little endian to  big endian & function used for encoding

Public Function ip_encode(ByVal ip As String) As String
    On Error Resume Next
    'ip is in little endian format
    Dim cpt             As Byte
    Dim last_position   As Byte
    Dim new_position    As Byte
    
    last_position = 1
    For cpt = 1 To 3
        'find the next byte
        new_position = InStr(last_position, ip, ".")
        'encode it and put it at the end
        ip_encode = ip_encode & Chr$(Val(Mid$(ip, last_position, new_position - last_position)))
        last_position = new_position + 1
    Next
    'for the last byte
    ip_encode = ip_encode & Chr$(Val(Mid$(ip, last_position)))
    
End Function

Public Function int_to_big_endian(ByVal Number As Long) As String
    On Error Resume Next
    'in fact it is int in C
    Dim lsb As Byte
    Dim hsb As Byte

    lsb = Number And 255
    hsb = (Number - lsb) / 256
    
    int_to_big_endian = Chr$(lsb) & Chr$(hsb)
    
End Function

Public Function long_to_big_endian(ByVal Number As Single) As String
    On Error Resume Next
    'in fact it should be an unsigned int
    Dim cpt             As Byte
    
    long_to_big_endian = Chr$(Number And 255) 'keep the last byte of the number
    
    For cpt = 1 To 3 '(we have 4 bytes and we have already treated the last one)
        Number = (Number - (Number And 255)) / 256
        long_to_big_endian = long_to_big_endian & Chr$(Number And 255) 'keep the last byte of the number
    Next cpt
    
End Function




'''''''''''''''''''''''''''''' big endian to little endian & functions used for decoding
'''''''''''''''''''''''''''''' used for decoding (that's why they all receive string)

Public Function ip_decode(ByVal ip As String) As String
    On Error Resume Next
    'ip is in little endian format
    If Len(ip) = 4 Then
        ip_decode = CStr(Asc(Mid$(ip, 1, 1))) & "." & CStr(Asc(Mid$(ip, 2, 1))) & "." & CStr(Asc(Mid$(ip, 3, 1))) & "." & CStr(Asc(Mid$(ip, 4, 1)))
    End If
End Function

Public Function byte_decode(ByVal Number As String) As Byte
    On Error Resume Next
    If Len(Number) = 1 Then
        byte_decode = CByte(Asc(Number))
    End If
End Function


Public Function int_to_little_endian(ByVal Number As String) As Long
    On Error Resume Next
    If Len(Number) = 2 Then
        int_to_little_endian = Asc(Mid$(Number, 2, 1))
        int_to_little_endian = int_to_little_endian * 256
        int_to_little_endian = int_to_little_endian + Asc(Mid$(Number, 1, 1))
    End If
End Function

Public Function long_to_little_endian(ByVal Number As String) As Single
    On Error Resume Next
    Dim cpt         As Integer
    Dim tmp         As Single
    Dim tmpres      As Single
    If Len(Number) = 4 Then
        For cpt = 4 To 1 Step -1
            tmp = Asc(Mid$(Number, cpt, 1))
            tmpres = tmpres + tmp * 256 ^ (cpt - 1)
        Next cpt
    End If
    long_to_little_endian = tmpres
End Function




'''''''''''''''''''''' make all header (little endian encoding)

Public Function make_descriptor_data(ByVal descriptor_id As String, ByVal payload_descriptor As Byte, ByVal ttl As Byte, ByVal hops As Byte, ByVal payload_length As Long) As str_descriptor_data
    On Error Resume Next
    With make_descriptor_data
        .descriptor_id = descriptor_id
        .payload_descriptor = Chr$(payload_descriptor)
        .ttl = Chr$(ttl)
        .hops = Chr$(hops)
        .payload_length = long_to_big_endian(payload_length)
    End With
End Function


Public Function make_pong_data(ByVal port As Long, ByVal ip As String, ByVal nb_shared_files As Long, ByVal nb_shared_kbytes As Long) As str_pong_data
    On Error Resume Next
    With make_pong_data
        .port = int_to_big_endian(port)
        .ip = ip_encode(ip)
        .nb_shared_files = long_to_big_endian(nb_shared_files)
        .nb_shared_kbytes = long_to_big_endian(nb_shared_kbytes)
    End With
End Function

Public Function make_query_data(ByVal minspeed As Long, ByVal search_criteria As String) As str_query_data
    With make_query_data
        .min_speed = long_to_big_endian(minspeed)
        .search_criteria = search_criteria + vbNullChar 'chr(0)
    End With
End Function

Public Function make_queryhit_data(ByVal nb_hits As Byte, ByVal port As Long, ByVal ip As String, ByVal speed As Single, ByVal all_str_result_set As String, ByVal servent_id As String) As str_queryhit_data
    On Error Resume Next
    With make_queryhit_data
        .nb_hits = Chr$(nb_hits)
        .port = long_to_big_endian(port)
        .ip = ip_encode(ip)
        .speed = long_to_big_endian(speed)
        .all_str_result_set = all_str_result_set
        .servent_id = servent_id
    End With
End Function

Public Function make_push_data(ByVal servent_id As String, ByVal file_index As Long, ByVal ip As String, ByVal port As Long) As str_push_data
    With make_push_data
        .servent_id = servent_id
        .file_index = long_to_big_endian(file_index)
        .ip = ip_encode(ip)
        .port = int_to_big_endian(port)
    End With
End Function

Public Function make_result_set(ByVal file_index As Long, ByVal file_size As Long, ByVal file_name As String) As str_result_set
    On Error Resume Next
    With make_result_set
        .file_index = long_to_big_endian(file_index)
        .file_size = long_to_big_endian(file_size)
        .file_name = file_name
    End With
End Function




''''''''''''''' all headers to string ---> ready to send

Public Function make_string_descriptor_data(ByVal descriptor_id As String, ByVal payload_descriptor As Byte, ByVal ttl As Byte, ByVal hops As Byte, ByVal payload_length As Long) As String
    On Error Resume Next
    Dim descriptord As str_descriptor_data
    descriptord = make_descriptor_data(descriptor_id, payload_descriptor, ttl, hops, payload_length)
    make_string_descriptor_data = descriptord.descriptor_id & descriptord.payload_descriptor & descriptord.ttl & descriptord.hops & descriptord.payload_length
End Function
Public Function make_string_pong_data(ByVal port As Long, ByVal ip As String, ByVal nb_shared_files As Long, ByVal nb_shared_kbytes As Long) As String
    On Error Resume Next
    Dim pongd As str_pong_data
    pongd = make_pong_data(port, ip, nb_shared_files, nb_shared_kbytes)
    make_string_pong_data = pongd.port & pongd.ip & pongd.nb_shared_files & pongd.nb_shared_kbytes
End Function

Public Function make_string_query_data(ByVal minspeed As Long, ByVal search_criteria As String) As String
    On Error Resume Next
    Dim queryd As str_query_data
    queryd = make_query_data(minspeed, search_criteria)
    make_string_query_data = queryd.min_speed & queryd.search_criteria
End Function

Public Function make_string_queryhit_data(ByVal nb_hits As Byte, ByVal port As Long, ByVal ip As String, ByVal speed As Single, ByVal all_str_result_set As String, ByVal servent_id As String) As String
    On Error Resume Next
    Dim queryhitd As str_queryhit_data
    queryhitd = make_queryhit_data(nb_hits, port, ip, speed, all_str_result_set, servent_id)
    make_string_queryhit_data = queryhitd.nb_hits & queryhitd.port & queryhitd.ip & queryhitd.speed & queryhitd.all_str_result_set & queryhitd.servent_id
End Function

Public Function make_string_push_data(ByVal servent_id As String, ByVal file_index As Long, ByVal ip As String, ByVal port As Long) As String
    On Error Resume Next
    Dim pushd As str_push_data
    pushd = make_push_data(servent_id, file_index, ip, port)
    make_string_push_data = pushd.servent_id & pushd.file_index & pushd.ip & pushd.port
End Function

Public Function make_string_result_set(ByVal file_index As Long, ByVal file_size As Long, ByVal file_name As String) As String
    On Error Resume Next
    Dim str_result_setd As str_result_set
    str_result_setd = make_result_set(file_index, file_size, file_name)
    make_string_result_set = str_result_setd.file_index & str_result_setd.file_size & str_result_setd.file_name & vbNullChar & vbNullChar
End Function



''''''''''''''''''''''''''' functions decoding each header

Public Function decode_descriptor_data(ByVal data As String) As descriptor_data
    On Error Resume Next
    With decode_descriptor_data
        .descriptor_id = Mid$(data, 1, 16)
        .payload_descriptor = byte_decode(Mid$(data, 17, 1))
        .ttl = byte_decode(Mid$(data, 18, 1))
        .hops = byte_decode(Mid$(data, 19, 1))
        .payload_length = long_to_little_endian(Mid$(data, 20, 4))
    End With
End Function

Public Function decode_pong_data(ByVal data As String) As pong_data
    On Error Resume Next
    With decode_pong_data
        .port = int_to_little_endian(Mid$(data, 1, 2))
        .ip = ip_decode(Mid$(data, 3, 4))
        .nb_shared_files = long_to_little_endian(Mid$(data, 7, 4))
        .nb_shared_kbytes = long_to_little_endian(Mid$(data, 11, 4))
    End With
End Function

Public Function decode_push_data(ByVal data As String) As push_data
    On Error Resume Next
    With decode_push_data
        .servent_id = Mid$(data, 1, 16)
        .file_index = long_to_little_endian(Mid$(data, 17, 4))
        .ip = ip_decode(Mid$(data, 21, 4))
        .port = int_to_little_endian(Mid$(data, 25, 2))
    End With
End Function

Public Function decode_query_data(ByVal data As String) As query_data
    On Error Resume Next
    If Len(data) < 4 Then Exit Function
    With decode_query_data
        .min_speed = int_to_little_endian(Mid$(data, 1, 2))
        .search_criteria = Mid$(data, 3, Len(data) - 3) 'remove the null char
    End With
End Function

Public Function decode_queryhit_data(data As String) As queryhit_data
    On Error Resume Next
    With decode_queryhit_data
        .nb_hits = byte_decode(Mid$(data, 1, 1))
        .port = int_to_little_endian(Mid$(data, 2, 2))
        .ip = ip_decode(Mid$(data, 4, 4))
        .speed = long_to_little_endian(Mid$(data, 8, 4))
        .all_result_set = Mid$(data, 12, Len(data) - 27) 'len(data)-size(all other field) we decode trailer in function decode_all_result_set
        .servent_id = Right$(data, 16)
    End With
End Function

Private Function decode_trailer(data As String) As ttrailer
    On Error Resume Next
    'called in decode_all_result_set
    If Len(data) > 6 Then
        With decode_trailer
            .vendor_code = Mid$(data, 1, 4)
            .open_data_size = CByte(Asc(Mid$(data, 5, 1)))  'specif 0.4 problem we give size even if we know it's 2
            .open_data_flag1 = CByte(Asc(Mid$(data, 6, 1))) '
            .open_data_flag2 = CByte(Asc(Mid$(data, 7, 1))) 'normally private data should begin here but we know we have 2 open data flag
            If Len(data) > 7 Then
                .private_data = Mid$(data, 8)
            End If
        End With
    End If
End Function

Public Function decode_all_result_set(ByRef queryhitd As queryhit_data) As result_set()
    On Error Resume Next
    'function charged to find all results of queryhit and call decode_trailer
    Dim array_results_set()     As result_set
    Dim all_results             As String
    Dim cpt                     As Byte
    
    all_results = queryhitd.all_result_set
    
    ReDim array_results_set(queryhitd.nb_hits - 1)
    For cpt = 0 To queryhitd.nb_hits - 1
        array_results_set(cpt) = decode_result_set(all_results) ' "all_results" is passed byref and we remove the
                                                                ' field discovered
    Next cpt
    'all_results should now contain only trailer data
    queryhitd.trailer = decode_trailer(all_results)
    
    
    decode_all_result_set = array_results_set
End Function

Private Function decode_result_set(ByRef strdata As String) As result_set
    On Error Resume Next
    Dim firstnull               As Integer
    With decode_result_set
        .file_index = long_to_little_endian(Mid$(strdata, 1, 4))
        .file_size = long_to_little_endian(Mid$(strdata, 5, 4))
        firstnull = InStr(9, strdata, vbNullChar)
        If firstnull > 9 Then
            .file_name = Mid$(strdata, 9, firstnull - 9)
        Else
            .file_name = Mid$(strdata, 9)
        End If
        If firstnull < Len(strdata) And firstnull > 0 Then
            firstnull = InStr(firstnull + 1, strdata, vbNullChar) 'second null char
            strdata = Mid$(strdata, firstnull + 1)
        Else
            strdata = ""
        End If
    End With
End Function








'''''''''''''''''''''''''''''''''''''''' treatement of data

Public Sub treat_ping(descriptord As descriptor_data, num_socket As Integer)
    On Error Resume Next
    'reply with a pong
    'create pong data
    Dim pongstr             As String
    Dim descriptorstr       As String
    Dim full_pong_data      As String
    Dim cpt
    Dim num_max
    Dim size
    
    size = UBound(known_hosts)
    full_pong_data = ""
    
    'create the answering descriptor
    descriptorstr = make_string_descriptor_data(descriptord.descriptor_id, pong, descriptord.hops, 0, 14)
    
    
    'give other hosts informations
    If 0 >= size - 1 - nb_of_pong_for_a_ping Then
        num_max = 0
    Else
        num_max = size - 1 - nb_of_pong_for_a_ping
    End If

    For cpt = size - 1 To num_max Step -1
        If known_hosts(cpt).ip <> "" And known_hosts(cpt).nb_shared_files < 2147483647 And known_hosts(cpt).nb_shared_kbytes < 2147483647 Then
            pongstr = make_string_pong_data(CLng(known_hosts(cpt).port), known_hosts(cpt).ip, CLng(known_hosts(cpt).nb_shared_files), CLng(known_hosts(cpt).nb_shared_kbytes))
        
            'add pong payload to traffic list
            add_to_payload_descriptor_list pong, Form_main.socket(num_socket).RemoteHostIP, False
            
            full_pong_data = full_pong_data + descriptorstr + pongstr
        End If
    Next cpt
    
    If send_my_computers_info Then
        'give your informations
        If Not sharing_simulation Then
            pongstr = make_string_pong_data(my_port, my_ip, my_nb_shared_files, my_nb_kilobytes_shared)
        Else
            pongstr = make_string_pong_data(my_port, my_ip, simulation_nb_files, simulation_size)
        End If
        
        'add pong payload to traffic list
        add_to_payload_descriptor_list pong, Form_main.socket(num_socket).RemoteHostIP, False
        
        full_pong_data = full_pong_data + descriptorstr + pongstr
    End If
    
    If full_pong_data = "" Then '<- thats means you don't want to send your informations
        ' --> send fake pong
        pongstr = make_string_pong_data("6346", "127.0.0.1", 100, 100) 'what ever you want
        'add pong payload to traffic list
        add_to_payload_descriptor_list pong, Form_main.socket(num_socket).RemoteHostIP, False
        full_pong_data = full_pong_data + descriptorstr + pongstr
    End If
    
    send_data full_pong_data, num_socket
                
            
    'If current_nb_dial <= max_dialing_hosts Then 'could be removed if you don't want to reduce ping traffic
    '    'if host is unknown, send ping to discover the host
    '    Dim known               As Boolean
    '    For cpt = 0 To UBound(known_hosts)
    '        If Form_main.socket(num_socket).RemoteHostIP = known_hosts(cpt).ip Then
    '            known = True
    '            Exit For
    '        End If
    '    Next cpt
    '    'check if it's not a forbidden ip
    '    For cpt = 0 To UBound(dummy_ip)
    '        If Form_main.socket(num_socket).RemoteHostIP = dummy_ip(cpt) Then
    '            known = True
    '            Exit For
    '        End If
    '    Next cpt
    '
    '    If Not known Then send_a_ping num_socket
    'End If
End Sub



Public Sub treat_pong(pongd As pong_data, num_socket As Integer)
    On Error Resume Next
    'add host to known hosts an dinterface
    If Not add_to_known_hosts(pongd.ip, pongd.port, pongd.nb_shared_files, pongd.nb_shared_kbytes) Then Exit Sub
    
    Dim host    As thost
    If current_nb_dial <= max_dialing_hosts Then
        host = preferently_connect_to_sharing_host()
        If host.ip = "" Then 'all host we are not connected with sharing 0kb
            make_new_dial pongd.ip, pongd.port
        Else
            make_new_dial host.ip, host.port
        End If
    End If
End Sub

Public Sub treat_query(num_socket As Integer, descriptord As descriptor_data, queryd As query_data)
    On Error Resume Next
    'search results
    Dim tmparray                As Variant
    Dim cpt
    Dim cpt2                    As Long
    Dim valid                   As Boolean
    Dim strtmp                  As String
    Dim descriptorstr           As String
    Dim queryhitd               As queryhit_data
    Dim queryhitstr             As String
    
    strtmp = LCase$(queryd.search_criteria)

    'prepare result descriptor
    queryhitd.ip = my_ip
    queryhitd.port = my_port
    queryhitd.servent_id = my_servent_id
    queryhitd.speed = my_speed
        
    ' split on space
    tmparray = Split(strtmp, " ")
    'search
    For cpt = 0 To UBound(my_shared_files) - 1
        valid = True
        For cpt2 = 0 To UBound(tmparray)
            If InStr(1, my_shared_files(cpt).file_name, tmparray(cpt2)) <= 0 Then
            'file name doesn't meet the query search criteria
                valid = False
                cpt2 = UBound(tmparray) 'to go out of the second for
            End If
        Next cpt2
        If valid Then

            queryhitd.nb_hits = queryhitd.nb_hits + 1
            queryhitd.all_result_set = queryhitd.all_result_set + make_string_result_set(my_shared_files(cpt).file_index, my_shared_files(cpt).file_size, my_shared_files(cpt).file_name)
            If queryhitd.nb_hits = 255 Then
            'number of hit is coded with 1 byte--> another queryhit is needed
            
                'send queryhit
                queryhitstr = make_string_queryhit_data(queryhitd.nb_hits, queryhitd.port, queryhitd.ip, queryhitd.speed, queryhitd.all_result_set, queryhitd.servent_id)
                'create the answering descriptor
                descriptorstr = make_string_descriptor_data(descriptord.descriptor_id, queryhit, descriptord.hops, 0, CLng(Len(queryhitstr)))
                'send queryhit
                send_data descriptorstr + queryhitstr, num_socket
                'add payload to traffic list
                add_to_payload_descriptor_list queryhit, Form_main.socket(num_socket).RemoteHostIP, False
                                

                'prepare data for a new queryhit
                queryhitd.nb_hits = 0
                queryhitd.all_result_set = ""
            End If
        End If
    Next cpt
    
    
    'treat the last queryhit
    If queryhitd.nb_hits > 0 Then 'don't route if no hit
        queryhitstr = make_string_queryhit_data(queryhitd.nb_hits, queryhitd.port, queryhitd.ip, queryhitd.speed, queryhitd.all_result_set, queryhitd.servent_id)
        'create the answering descriptor
        descriptorstr = make_string_descriptor_data(descriptord.descriptor_id, queryhit, descriptord.hops, 0, CLng(Len(queryhitstr)))
        'send queryhit
        send_data descriptorstr + queryhitstr, num_socket
        'add payload to traffic list
        add_to_payload_descriptor_list queryhit, Form_main.socket(num_socket).RemoteHostIP, False
                    
    End If
    
End Sub

Public Sub treat_queryhit(queryhitd As queryhit_data, ByVal descriptor_id As String, num_socket As Integer)
    'add servent id to table and all other informations
    On Error Resume Next
    Dim cpt                     As Long
    Dim cpt2                    As Long
    Dim already_exist           As Boolean
    Dim results()               As result_set
    Dim found                   As Boolean 'the form that made the query,if still exist
    Dim doc_number              As Integer 'number of the form that made the query, if still exist
    

    'decode results
    results = decode_all_result_set(queryhitd)
    

    'add result to the corresponding form_search
    For cpt = 0 To UBound(descriptorID_to_num_form_search, 2) - 1
        found = False
        If descriptorID_to_num_form_search(0, cpt) = descriptor_id Then
            doc_number = CInt(descriptorID_to_num_form_search(1, cpt))
            found = True
            Exit For
        End If
    Next cpt
    
    If found Then
        Document_search(doc_number).list_search_results.Sorted = False 'avoid troubles
        For cpt2 = 0 To UBound(results)
            add_result_to_form_search doc_number, results(cpt2).file_name, results(cpt2).file_size, queryhitd.ip, queryhitd.port, queryhitd.speed, results(cpt2).file_index, queryhitd.servent_id, num_socket, True
        Next cpt2
    Else 'we have close the corresponding form_search --> add to known_files only
        For cpt2 = 0 To UBound(results)
            add_to_known_files results(cpt2).file_index, results(cpt2).file_name, results(cpt2).file_size, queryhitd.ip, queryhitd.port, num_socket, queryhitd.speed, queryhitd.servent_id
        Next cpt2
    End If
Document_search(doc_number).list_search_results.Refresh
    
End Sub

Public Sub treat_push(pushd As push_data)
    'when we are firewalled with incoming connection --> we need to open an outgoing connection and send a giv
    On Error Resume Next 'error can occurs for bad values of file index (because file_index is the array position)
    Dim strdata         As String
    Dim strfile_name    As String
    Dim num_socket      As Integer
    Dim size            As Integer

    
    If current_nb_upload < max_upload And allow_upload Then ' the downloders should send push until connection is made
        If is_ip_banished(pushd.ip) Then Exit Sub
        current_nb_upload = current_nb_upload + 1
        strfile_name = my_shared_files(pushd.file_index).file_name
        If strfile_name <> "" Then
            strdata = make_giv(CStr(pushd.file_index), my_servent_id, strfile_name)
            'make outgoing connection
            num_socket = find_first_free_socket()
            If num_socket > -1 Then
                size = UBound(push_upload)
                push_upload(size).num_socket = num_socket
                push_upload(size).data = strdata
                ReDim Preserve push_upload(size + 1)
                Form_main.socket(num_socket).LocalPort = 0
                Form_main.socket(num_socket).RemoteHost = pushd.ip
                Form_main.socket(num_socket).RemotePort = pushd.port
                socket_state(num_socket).connection_direction = outgoing
                socket_state(num_socket).connection_type = giving
                Form_main.socket(num_socket).Connect
            Else
                current_nb_upload = current_nb_upload - 1
            End If
        End If
    End If
End Sub





'''''''''''''''''''''' get and giv functions
Public Function make_get_request(ByVal file_index As String, ByVal file_name As String, ByVal range As String) As String
    On Error Resume Next
    make_get_request = "GET /get/" & file_index & "/" & file_name & "/ HTTP/1.0" & vbCrLf _
                        & "Connection: Keep-Alive" & vbCrLf _
                        & "Range: bytes=" & range & vbCrLf _
                        & "User-Agent: Coyotella" & vbCrLf & vbCrLf
End Function


Public Function make_giv(ByVal file_index As String, ByVal servent_id As String, ByVal file_name As String) As String
    On Error Resume Next
    Dim str_hexa_id      As String
    Dim cpt As Byte
    str_hexa_id = ""
    For cpt = 1 To 16
        ' str_hexa_id = str_hexa_id & Hex$(Asc(Mid$(servent_id, cpt, 1))) won't work because because it dosen't return the first 0 ( return A instead of 0A )
        str_hexa_id = str_hexa_id & byte_to_hexa(Asc(Mid$(servent_id, cpt, 1)))
    Next cpt
    make_giv = "GIV " & file_index & ":" & str_hexa_id & "/" & file_name & vbLf & vbLf
End Function

Public Function decode_giv(ByVal data As String, ip As String) As file_giv_info
    On Error Resume Next
    Dim strtmp          As String
    Dim hexa_servent_id As String
    Dim servent_id      As String
    Dim tmp             As Integer
    Dim cpt             As Integer
    strtmp = Mid$(data, 5) 'remove "GIV "
    'find file_index
    tmp = InStr(1, strtmp, ":")
    decode_giv.file_index = Mid$(strtmp, 1, tmp - 1)
    strtmp = Mid$(strtmp, tmp + 1)
    'find servent_id
    tmp = InStr(1, strtmp, "/")
    hexa_servent_id = Trim$(Mid$(strtmp, 1, tmp - 1))
    servent_id = hexa_servent_id
    If Len(hexa_servent_id) = 32 Then
        servent_id = ""
        For cpt = 1 To 31 Step 2
            servent_id = servent_id & Chr$(Val("&H" & (Mid$(hexa_servent_id, cpt, 2))))
        Next cpt
    End If
    decode_giv.servent_id = servent_id
    strtmp = Mid$(strtmp, tmp + 1)
    'find file_name
    tmp = InStr(1, strtmp, vbLf & vbLf)
    decode_giv.file_name = Mid$(strtmp, 1, tmp - 1)

End Function

Public Function decode_get(ByVal data As String) As file_get_info
    On Error Resume Next
    Dim strtmp          As String
    Dim arr_line        As Variant
    Dim tmp             As Integer
    Dim cpt             As Byte
    
    arr_line = Split(data, vbCrLf)
    
    For cpt = 0 To UBound(arr_line)
        strtmp = CStr(arr_line(cpt))
        If Len(strtmp) > 7 Then
            If Mid$(strtmp, 1, 8) = "GET /get" Then
                strtmp = Mid$(data, 10) 'remove "GET /get/"
                'find file index (before first /)
                tmp = InStr(1, strtmp, "/")
                decode_get.file_index = CSng(Val(Mid$(strtmp, 1, tmp - 1)))
                'remove file index and /
                strtmp = Mid$(strtmp, tmp + 1)
                'search HTTP to remove the unused end of line
                tmp = InStr(1, strtmp, " HTTP")
                strtmp = Mid$(strtmp, 1, tmp - 1)
                'remove the / after filename if there is one
                tmp = InStr(1, strtmp, "/")
                If tmp > 0 Then
                    strtmp = Mid$(strtmp, 1, tmp - 1)
                End If
                decode_get.file_name = strtmp
            Else
                If Mid$(strtmp, 1, 6) = "Range:" Then
                    'find = of "bytes="
                    tmp = InStr(1, strtmp, "=")
                    strtmp = Mid$(strtmp, tmp + 1)
                    decode_get.range = strtmp
                End If
            End If
        End If
    Next cpt
    
End Function


Public Function is_one_of_my_descriptor_ID(ByVal ID As String) As Boolean
    On Error Resume Next
    'check if the ID is one of our descriptor
    Dim cpt                 As Long
    For cpt = 0 To UBound(my_descriptors_ID)
        If my_descriptors_ID(cpt) = ID Then
            is_one_of_my_descriptor_ID = True
            Exit Function
        End If
    Next cpt
End Function
