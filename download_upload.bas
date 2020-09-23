Attribute VB_Name = "download_upload"
Option Explicit

'''''''''''''''''''''''''' upload and download functions
Public Sub upload_file(ByVal file_index As String, ByVal file_name As String, ByVal range As String, ByVal num_socket As Integer)
    On Error Resume Next
    Dim strdata             As String
    Dim range_end           As Single
    Dim range_begin         As Single
    Dim full_file_name      As String
    
    'find range
    Dim strrange_begin      As String
    Dim strrange_end        As String
    Dim tmp                 As Integer
    tmp = InStr(1, range, "-")
    strrange_begin = Mid$(range, 1, tmp - 1)
    If Len(range) > tmp Then 'range_end is defined
        strrange_end = Mid$(range, tmp + 1)
    Else
        strrange_end = ""
    End If
    range_begin = Val(strrange_begin)
    range_end = Val(strrange_end)

    'read file
    Dim num_file            As Integer
    Dim buffer              As String
    Dim size                As Long
    Dim pos                 As Long

    'find the full path
    Dim cpt     As Long
    Dim found   As Boolean
    For cpt = 0 To UBound(my_shared_files)
        If file_name = my_shared_files(cpt).file_name And file_index = my_shared_files(cpt).file_index Then
            full_file_name = my_shared_files(cpt).full_path
            pos = cpt
            found = True
            Exit For
        End If
    Next cpt
     
    If Not found Or (range_end <= range_begin And range_end > 0) _
        Or range_begin > my_shared_files(pos).file_size _
        Or range_end > my_shared_files(pos).file_size _
    Then
        
        'log if we want to log bad downloaders ( time ip file_name range )
        If log_bad_downloaders Then
            num_file = FreeFile
            buffer = CStr(Now) & " Upload Error  ip:" & Form_main.socket(num_socket).RemoteHostIP & " file name: " & file_name & " Range: " & range
            Open downloader_log_file_name For Append Access Write As num_file
                Print #num_file, buffer
            Close num_file
        End If
        
        send_data "HTTP 403 Not Found" & vbCrLf & "Server: Coyotella" & vbCrLf & vbCrLf & _
                  "<HTML><HEAD><TITLE>403 Not Found</TITLE></HEAD><BODY><H1>Not Found</H1>" & _
                  "The requested file " & file_name & "with index " & file_index & _
                  " was not found on this Coyotella server</BODY></HTML>", num_socket
        treat_socket_closing num_socket
        Exit Sub
    End If
    
    'log if we want to log downloaders ( time ip file_name range )
    If log_good_downloaders Then
        num_file = FreeFile
        buffer = CStr(Now) & " ip:" & Form_main.socket(num_socket).RemoteHostIP & " file name: " & file_name & " Range: " & range
        Open downloader_log_file_name For Append Access Write As num_file
            Print #num_file, buffer
        Close num_file
    End If

    'start upload
    num_file = FreeFile
    
    If Not is_file_existing(full_file_name) Then
        'if file exist in another shared directory with same name and size, _
         we update the directory, else we remove the file from my_shared_files
        Dim file_exist As Boolean
        For cpt = 0 To UBound(myshared_directories)
           If is_file_existing(myshared_directories(cpt) & file_name) Then
                full_file_name = myshared_directories(cpt) & file_name
                my_shared_files(pos).full_path = full_file_name
                file_exist = True
                Exit For
           End If
        Next cpt
        
        If Not file_exist Then
            remove_from_my_shared_files (pos)
            send_data "HTTP 403 Not Found" & vbCrLf & "Server: Coyotella" & vbCrLf & vbCrLf & _
                      "<HTML><HEAD><TITLE>403 Not Found</TITLE></HEAD><BODY><H1>Not Found</H1>" & _
                      "The requested file " & file_name & "with index " & file_index & _
                      " was not found on this Coyotella server</BODY></HTML>", num_socket
            treat_socket_closing num_socket
        End If
    End If
    
    Open full_file_name For Binary Access Read Shared As num_file
        size = LOF(num_file)
        
        If range_end > size - 1 Or range_end = 0 Then
            range_end = size - 1
        End If
        range_begin = range_begin + 1 ' if file_size=n the correct ask is 0-(n-1) in http protocol
        range_end = range_end + 1     ' get use 1-n
        buffer = Space$(CLng(range_end - range_begin + 1))
        Get num_file, range_begin, buffer
    Close num_file
    '
    
    'make header
    strdata = "HTTP 200 OK" & vbCrLf
    strdata = strdata & "Server: Coyotella" & vbCrLf
    strdata = strdata & "Content-type: application/binary" & vbCrLf
    strdata = strdata & "Content-length: " & Len(buffer) & vbCrLf & vbCrLf

    'show upload
    range = strrange_begin & "-" & CStr(range_end)
    add_upload_to_interface file_name, range, Form_main.socket(num_socket).RemoteHostIP, "Sending", file_index, num_socket

    'send data
    Form_main.socket(num_socket).SendData strdata & buffer
End Sub


Public Function ask_for_download(ByVal ip As String, ByVal port As String, ByVal file_index As Single, _
                            ByVal file_name As String, ByVal range_begin As Single, ByVal range_end As Single, _
                            ByVal servent_id As String, ByVal file_size As Single, ByVal speed As String, _
                            ByVal num_part As Integer, Optional ByVal create_new_file As Boolean = False, _
                            Optional ByVal nb_parts As Integer = 1, Optional ByVal saving_name As String, _
                            Optional ByVal recovery As Boolean = False, Optional ByVal num_socket_queryhit As Integer = -1, _
                            Optional ByVal need_to_push As Boolean = False)
    
    On Error Resume Next 'avoid no route to host at the .connect
    Dim data                 As String
    Dim strrange_end         As String
    Dim file_name2           As String
    Dim position             As Integer
    Dim pos_name_download    As Integer
    Dim num_socket           As Integer
    Dim need_to_rename       As Boolean
    Dim strpush              As String
    
    If range_end = 0 Then
        strrange_end = ""
    Else
        strrange_end = CStr(CLng(range_end))
    End If
    
    port = Trim$(port)
            
    If create_new_file Then
        file_name2 = file_name 'depends if full path or not
        'if there is directory separator remove directories name
        file_name2 = get_file_name(file_name)
        ' verify if no file with same name exist in the temporary or download directory
        ' if there's already one file with the same name we ask for renaming or overwriting
        If is_file_existing(my_incomplete_directory & file_name) Then
            need_to_rename = True
        End If
        If is_file_existing(my_download_directory & file_name) Or need_to_rename Then
            renamed_file.old_name = file_name2
            renamed_file.overwrite = False
            ask_for_rename need_to_rename
            saving_name = renamed_file.new_name
        Else
            saving_name = file_name2
        End If
        init_prepare_resume saving_name, nb_parts, file_size
    End If

    If Not recovery Then
        prepare_resume_file_info_first saving_name, ip, CLng(port), file_index, servent_id, speed, file_name, nb_parts
    End If
    
    prepare_resume saving_name, num_part, range_begin
    If range_end = 0 Then
        update_resume_part_size saving_name, num_part, file_size
    Else
        update_resume_part_size saving_name, num_part, range_end
    End If
    
    'add file to name_download()

    position = UBound(name_download)
    With name_download(position)
        .file_name = file_name
        .saving_name = saving_name
        .range_begin = range_begin
        .range_end = range_end
        .num_socket = num_socket
        .bytes_recieved = 0
        .servent_id = servent_id
        .file_index = file_index
        .file_size = file_size
        .ip = ip
        .port = port
        .num_part = num_part
        .old_percent = 0
        .initial_size = 0
        .header_received = False
        .get_request_made = True
        .num_socket_queryhit = num_socket_queryhit
        .need_to_push = need_to_push
    End With
    ReDim Preserve name_download(position + 1)

    pos_name_download = position
    Dim lResult As Long
    
    If is_ip_forbidden(ip) Or need_to_push Then
        If name_download(position).num_socket_queryhit < 0 Then
            'we can't route the push
            lResult = MessageBox(0, "File couldn't be download on this host," & vbCrLf & "you must make a resume on other host.", Form_main.Caption, vbExclamation)
            add_download_to_interface saving_name, file_index, range_begin & "-" & _
                                        strrange_end, ip, speed, "Must be resumed on other host", -1, pos_name_download
        Else
            'send push
            strpush = make_string_push_data(name_download(position).servent_id, name_download(position).file_index, my_ip, my_port)
            data = make_string_descriptor_data(my_servent_id, push, my_ttl, 0, 26)
            name_download(position).get_request_made = False
            data = data & strpush

            name_download(pos_name_download).pos_current_down = add_download_to_interface(saving_name, file_index, range_begin & "-" & _
                                        strrange_end, ip, speed, "Waiting", -1, pos_name_download)

            'add file in waiting array
            position = UBound(waiting_download)
            waiting_download(position).data = data
            waiting_download(position).connection_tried = False
            waiting_download(position).num_socket = -1
            waiting_download(position).push = True
            waiting_download(position).num_socket_queryhit = name_download(pos_name_download).num_socket_queryhit
            waiting_download(position).position_name_download = pos_name_download
            ReDim Preserve waiting_download(position + 1)
        End If
    Else
        'send a get request
        data = "GET /get/" & Trim$(file_index) & "/" & Trim$(file_name) & " HTTP/1.0" & vbCrLf _
                & "Range: bytes=" & CStr(CLng(range_begin)) & "-" & strrange_end & vbCrLf _
                & "User-Agent: Coyotella" & vbCrLf _
                & vbCrLf
        
        'check if we can download now
        If current_nb_download < max_download Then
            num_socket = find_first_free_socket()
        Else
            num_socket = -1
        End If

        name_download(pos_name_download).num_socket = num_socket
        name_download(pos_name_download).pos_current_down = add_download_to_interface(saving_name, file_index, range_begin & "-" & _
                                        strrange_end, ip, speed, "Waiting", num_socket, pos_name_download)
        'add file in waiting array
        position = UBound(waiting_download)
        waiting_download(position).data = data
        waiting_download(position).num_socket = num_socket
        waiting_download(position).position_name_download = pos_name_download
        waiting_download(position).connection_tried = False
        waiting_download(position).push = False
        waiting_download(position).num_socket_queryhit = num_socket_queryhit
        ReDim Preserve waiting_download(position + 1)
    
        If num_socket > -1 Then ' we can download now so make a connection
            waiting_download(position).connection_tried = True
            current_nb_download = current_nb_download + 1
            update_status_download "Connecting", num_socket
            socket_state(num_socket).connection_type = downloading
            
            Form_main.socket(num_socket).LocalPort = 0
            Form_main.socket(num_socket).RemoteHost = ip
            Form_main.socket(num_socket).RemotePort = port
            Form_main.socket(num_socket).Connect
        End If
    
    End If
    ask_for_download = saving_name 'return the saving name for multiple part download
End Function
    
Public Sub decode_download(full_data As String, ByVal num_socket As Integer, ByVal bytes_received As Long)
    On Error Resume Next
    'decode header and write file
    Dim num_file            As Integer
    Dim cpt                 As Long
    Dim range_begin         As Long
    Dim range_end           As Long
    Dim file_name           As String
    Dim tmp_data            As String
    Dim strheader           As String
    
    Dim pos                 As Integer
    Dim pos1                As Integer
    Dim pos2                As Integer
    Dim content_length      As Single
    Dim range               As String
    Dim content_range       As Boolean
    Dim cr_range            As String
    Dim cr_range_begin      As Single
    Dim cr_range_end        As Single
    
    Dim place               As Long
    Dim strfile             As String
    Dim first_time          As Boolean
    Dim pos_name_down       As Long
    Dim range_end_not_respected As Boolean
    Dim small_data          As String

    For cpt = 0 To UBound(name_download) - 1
        If name_download(cpt).num_socket = num_socket Then 'we have found corresponding socket --> informations with name_download
            
            If name_download(cpt).header_received = False Then
                first_time = True
                
                tmp_data = remaining_data(num_socket) & full_data
                'find end of header
                place = InStr(1, tmp_data, vbCrLf & vbCrLf) + 4 ' position just after the vbcrlf
                
                If place <= 5 Then ' not full header
                    remaining_data(num_socket) = remaining_data(num_socket) & full_data
                    If Len(remaining_data(num_socket)) > 5000 Then '5000 should be enougth for header
                        Open App.Path & "\header_download_error.log" For Append Access Write As num_file
                            Print #num_file, remaining_data(num_socket)
                            Print #num_file, " " & Form_main.socket(num_socket).RemoteHostIP
                        Close num_file
                        treat_socket_closing num_socket
                    End If
                    Exit Sub
                End If
                
                'we have received a full header
                name_download(cpt).header_received = True
                remaining_data(num_socket) = ""
                strheader = Mid$(tmp_data, 1, place)
                small_data = LCase$(strheader)
                'check if download is accepted <--> http 200 ok is present
                If InStr(1, Mid$(small_data, 1, InStr(1, small_data, vbCrLf)), "http 200") Then
                    'search content length to be sure of the size we should receive
                     
                    pos = InStr(1, small_data, "content-length:")
                    If pos < 1 Then
                        pos1 = InStr(1, small_data, "content-range:")
                        content_range = True
                        pos = pos1
                    End If
                    If pos > 0 Then '"Content-length:" or "Content-range:" are present: --> we try to get value
                        
                        pos2 = InStr(pos, tmp_data, vbCrLf)
                        If pos2 > 0 Then ' full line: --> we try to get value
                            If Not content_range Then
                                content_length = CSng(Val(Trim$(Mid$(tmp_data, pos + 15, pos2 - pos - 16 + 1))))
                            Else 'content range min-max/val
                                cr_range = Trim$(Mid$(tmp_data, pos + 14, pos2 - pos - 15 + 1))
                                pos1 = InStr(1, cr_range, "/")
                                If pos1 > 1 Then
                                    cr_range = Mid$(cr_range, 1, pos1 - 1)
                                End If
                                pos1 = InStr(1, cr_range, "=")
                                If pos1 > 1 Then
                                    cr_range = Mid$(cr_range, pos1 + 1)
                                End If
                                pos1 = InStr(1, cr_range, "-")
                                If pos1 > 1 Then
                                    cr_range_begin = CSng(Val(Mid$(cr_range, 1, pos1 - 1)))
                                    cr_range_end = CSng(Val(Mid$(cr_range, pos1 + 1)))
                                    content_length = cr_range_end - cr_range_begin + 1
                                Else
                                    pos2 = -1
                                End If
                            End If
                            name_download(cpt).file_size = content_length 'this is not file size but the size of the downloading part
                            'the upper line was nice but some gnutella clones seem to not respect the range end so...
                            If name_download(cpt).range_end = 0 Or name_download(cpt).range_end = name_download(cpt).file_size Then
                                name_download(cpt).range_end = name_download(cpt).range_begin + content_length - 1
                                range = name_download(cpt).range_begin & "-" & name_download(cpt).range_end
                                update_range_download range, num_socket
    
                                update_resume_part_size name_download(cpt).saving_name, name_download(cpt).num_part, name_download(cpt).range_end
                            End If
                            name_download(cpt).file_size = name_download(cpt).range_end - name_download(cpt).range_begin + 1
                        End If
                    End If

                    range_begin = name_download(cpt).range_begin
                    name_download(cpt).bytes_recieved = Len(tmp_data) - place + 1 'place-1=header length
                
                
                Else 'there's not http 200 ok

                        'HTTP error messages
                        Dim strerror        As String
                        Dim posmax          As Long
                        posmax = InStr(1, tmp_data, vbCrLf)
                        If posmax > 0 Then
                            strerror = LCase$(Mid$(tmp_data, 1, posmax))
                            If InStr(1, strerror, "403") Or InStr(1, strerror, "not found") Then
                                update_status_download "Error: File not found on this server", num_socket
                                treat_socket_closing num_socket
                            Else
                                If InStr(1, strerror, "503") Or InStr(1, strerror, "busy") Then
                                    update_status_download "Error: Server Busy", num_socket
                                    treat_socket_closing num_socket
                                Else
                                    'another error
                                    update_status_download "Error: unknown error", num_socket
                                    treat_socket_closing num_socket
                                End If
                            End If
                        End If
                        Exit Sub
                
                End If
                
            Else 'header already recieved --> download has begun
                tmp_data = full_data
                file_name = name_download(cpt).saving_name
                'find the place we should write the next part
                range_begin = name_download(cpt).range_begin + name_download(cpt).bytes_recieved
                name_download(cpt).bytes_recieved = name_download(cpt).bytes_recieved + bytes_received
            End If
            pos_name_down = cpt
            Exit For
        End If
    Next cpt
    
    
    'verify if range end is respected
    If name_download(cpt).bytes_recieved + name_download(cpt).range_begin > name_download(cpt).range_end + 1 Then
        range_end_not_respected = True
    End If
    
    file_name = my_incomplete_directory & file_name
    If range_end_not_respected Then
        prepare_resume name_download(pos_name_down).saving_name, name_download(pos_name_down).num_part, name_download(pos_name_down).range_end + 1
    Else
        prepare_resume name_download(pos_name_down).saving_name, name_download(pos_name_down).num_part, name_download(pos_name_down).range_begin + name_download(pos_name_down).bytes_recieved
    End If

    
    'write data in corresponding file even if range end is not respected
    num_file = FreeFile
    file_name = my_incomplete_directory & name_download(pos_name_down).saving_name
    Open file_name For Binary Access Write Shared As num_file
    If first_time Then
        'remove header part
        Put num_file, range_begin + 1, Mid$(tmp_data, place)
    Else
        Put num_file, range_begin + 1, tmp_data
    End If
    Close num_file

    If range_end_not_respected Then
        'close socket
        treat_socket_closing num_socket
    End If
    
    
    'update percent info
    Dim percent             As Integer
    If range_end_not_respected Then
        update_status_download "Completed", num_socket
    Else
        If name_download(pos_name_down).range_end - name_download(pos_name_down).range_begin > 0 Then
            If name_download(pos_name_down).initial_size = 0 Then
                percent = Int(name_download(pos_name_down).bytes_recieved / _
                     (name_download(pos_name_down).range_end - name_download(pos_name_down).range_begin) * 100)
            Else
                percent = name_download(pos_name_down).old_percent + CByte(name_download(pos_name_down).bytes_recieved / name_download(pos_name_down).initial_size * 100)
            End If
            update_status_download CStr(percent) & " %", num_socket, name_download(pos_name_down).file_size - name_download(pos_name_down).bytes_recieved
        End If
    End If
End Sub

Public Sub remove_waiting_download_with_pos_name_download(ByVal position_name_download As Integer)
    On Error Resume Next
    'remove an element (if element is found) and keep other in order
    Dim cpt         As Integer
    Dim array_size  As Integer
    If position_name_download < 0 Then Exit Sub
    array_size = UBound(waiting_download)
    For cpt = position_name_download To array_size - 1
        If waiting_download(cpt).position_name_download = position_name_download Then
            remove_waiting_download cpt
            Exit For
        End If
    Next cpt
End Sub

Public Sub remove_waiting_download(ByVal index As Integer)
    On Error Resume Next
    'remove an element and keep other in order
    Dim cpt         As Integer
    Dim array_size  As Integer
    array_size = UBound(waiting_download)
    For cpt = index To array_size - 2
        waiting_download(cpt) = waiting_download(cpt + 1)
    Next cpt
    If array_size > 0 Then
        ReDim Preserve waiting_download(array_size - 1)
    End If
End Sub



Public Function download_completed(ByVal pos_name_download As Integer) As Boolean
    On Error Resume Next
    Dim cpt As Integer
    Dim size As Long
    If name_download(pos_name_download).bytes_recieved >= name_download(pos_name_download).file_size And name_download(pos_name_download).bytes_recieved > 0 Then
        'remember that .bytes_recieved is the number of recieved bytes- the header length
        'and file size  is not the file size but the size of the downloading part
        download_completed = True
    End If
End Function


Public Function all_part_completed(ByVal download_file_name As String) As Boolean
    On Error Resume Next
    Dim num_file        As Integer
    Dim type_len        As Byte
    type_len = 4 'single
    
    Dim num_record      As Single
    Dim nb_records      As Single
    Dim file_name       As String
    Dim next_byte       As Single
    Dim range_end       As Single
    
    Dim cpt             As Integer
    
    all_part_completed = True
    file_name = my_incomplete_directory & download_file_name & ".coy"

    num_file = FreeFile()
    'get number of record
    Open file_name For Random Shared As num_file Len = type_len
        Get num_file, 1, nb_records
        'check if
        For cpt = 3 To nb_records * 2 + 1 Step 2
            Get num_file, cpt, next_byte
            Get num_file, cpt + 1, range_end
            If next_byte < range_end Or range_end = 0 Then
                all_part_completed = False
                Exit For
            End If
        Next cpt
    Close num_file
End Function


                            ''''''''''''''''''''''''''''''''''
Private Sub ask_for_rename(ByVal need_to_rename As Boolean)
    On Error Resume Next
    If need_to_rename Then
        Form_rename.Lbltxt.Caption = "You are currently downloading a file with the same name," _
                                & vbCrLf & "so you should rename the file you try to download"
        Form_rename.Option2.Enabled = False
    End If
    Form_rename.Show vbModal 'code waiting for answer
    
    If renamed_file.overwrite = True Then
        renamed_file.new_name = renamed_file.old_name
        Exit Sub
    End If
    'else
    If is_file_existing(my_download_directory & renamed_file.new_name) Or is_file_existing(my_incomplete_directory & renamed_file.new_name) Or renamed_file.new_name = "" Then
        ask_for_rename need_to_rename
    End If

End Sub

Public Sub check_for_waiting_download()
    Dim cpt                 As Integer
    Dim new_num_socket      As Integer
    Dim pos                 As Integer
    Dim remaining_down      As Boolean
    Dim pos_waiting_down    As Integer
    Dim array_size          As Integer
    
    On Error Resume Next 'avoid no route to host at the .connect
    array_size = UBound(waiting_download)
    If array_size > 0 Then 'some files are not downloaded
        'find first request not made
        remaining_down = False
        For cpt = 0 To array_size - 1
            If Not waiting_download(cpt).connection_tried Then
                pos_waiting_down = cpt
                remaining_down = True
            End If
        Next cpt
        If Not remaining_down Then Exit Sub
        Dim lResult As Long
        
        If waiting_download(pos_waiting_down).push Then
            'send the corresponding push
            If waiting_download(pos_waiting_down).num_socket_queryhit = -1 Then
                lResult = MessageBox(0, "File couldn't be download on this host," & vbCrLf & "you must make a resume on other host.", Form_main.Caption, vbExclamation)
                remove_waiting_download pos_waiting_down
            Else
                send_data waiting_download(pos_waiting_down).data, waiting_download(pos_waiting_down).num_socket_queryhit
                update_status_download2 "Sending push", name_download(waiting_download(pos_waiting_down).position_name_download).pos_current_down
                remove_waiting_download pos_waiting_down
                add_to_waiting_giv waiting_download(pos_waiting_down).position_name_download
                add_to_payload_descriptor_list push, name_download(waiting_download(pos_waiting_down).position_name_download).ip, False, , " for " & name_download(waiting_download(pos_waiting_down).position_name_download).file_name
            End If
        Else
            new_num_socket = find_first_free_socket()
            waiting_download(pos_waiting_down).num_socket = new_num_socket
            
            If new_num_socket > -1 Then
                pos = waiting_download(pos_waiting_down).position_name_download 'give the position in array name_download to get info next
                name_download(pos).num_socket = new_num_socket
                'we should update current_download(x).num_socket wich is still=-1
                current_download(name_download(pos).pos_current_down).num_socket = new_num_socket
                
                socket_state(new_num_socket).connection_type = downloading
                socket_state(new_num_socket).connection_direction = outgoing
                
                Form_main.socket(new_num_socket).LocalPort = 0
                Form_main.socket(new_num_socket).RemoteHost = name_download(pos).ip
                Form_main.socket(new_num_socket).RemotePort = name_download(pos).port
    
                Form_main.socket(new_num_socket).Connect
                update_status_download "Connecting", new_num_socket
                waiting_download(pos_waiting_down).connection_tried = True
            End If
        End If
    End If
End Sub

Public Sub remove_from_name_download(ByVal pos_name_download As Integer)
    On Error Resume Next
    Dim pos     As Integer
    Dim cpt     As Integer
    pos = UBound(name_download)
    If pos > 0 Then
        If pos_name_download <> pos - 1 Then
            'swap last element with the pos_name_download's one
            name_download(pos_name_download) = name_download(pos - 1)
            'update new position in current_download
            current_download(name_download(pos_name_download).pos_current_down).pos_name_download = pos_name_download
            'update position in retry download
            Form_main.Timer_second.Enabled = False
            For cpt = 0 To UBound(retry_download)
                If retry_download(cpt).pos_name_download = pos - 1 Then
                    retry_download(cpt).pos_name_download = pos_name_download
                End If
            Next cpt
            Form_main.Timer_second.Enabled = True
            'update new position in waiting_download
            For cpt = 0 To UBound(waiting_download) - 1
                If waiting_download(cpt).position_name_download = pos - 1 Then
                    waiting_download(cpt).position_name_download = pos_name_download
                    Exit For
                End If
            Next cpt
            
            'update position in waiting_giv
            For cpt = 0 To UBound(waiting_giv) - 1
                If waiting_giv(cpt).pos_name_download = pos - 1 Then
                    waiting_giv(cpt).pos_name_download = pos_name_download
                    Exit For
                End If
            Next cpt
        End If
    End If
    If pos > 0 Then ReDim Preserve name_download(pos - 1)
End Sub

Public Sub remove_from_push_upload(ByVal pos_push_upload As Integer)
    On Error Resume Next
    Dim cpt As Integer
    Dim size As Integer
    
    size = UBound(push_upload)
    For cpt = 0 To size - 1
        If cpt = pos_push_upload Then
            push_upload(cpt) = push_upload(size - 1)
            ReDim Preserve push_upload(size - 1)
            Exit For
        End If
    Next cpt
End Sub


Public Sub stop_download(ByRef out_pos_selected As Integer, ByRef out_pos_current_down As Integer, ByRef out_pos_name_download As Integer, ByRef out_file_name As String, Optional ByVal without_interface As Boolean = False, Optional ByVal numsock As Integer, Optional ByVal come_from_treat_socket_closing As Boolean = False)
    On Error Resume Next
    Dim num_socket          As Integer
    Dim cpt                 As Integer

    If without_interface Then
        For cpt = 0 To UBound(current_download) - 1
            If current_download(cpt).num_socket = numsock Then
                out_pos_current_down = cpt
                Exit For
            End If
        Next cpt
        'find coresponding item
        For cpt = 1 To Form_download_upload.ListView_download.ListItems.Count
            If out_pos_current_down = Form_download_upload.ListView_download.ListItems.Item(cpt).Tag Then
                out_pos_selected = cpt
                Exit For
            End If
        Next cpt
        num_socket = numsock
    Else 'form event
        out_pos_selected = find_selected_down()
        If out_pos_selected = -1 Then Exit Sub
        out_pos_current_down = Form_download_upload.ListView_download.ListItems(out_pos_selected).Tag
        num_socket = current_download(out_pos_current_down).num_socket
    End If
    
    out_file_name = current_download(out_pos_current_down).file_name
    
    'give the position in name download
    out_pos_name_download = current_download(out_pos_current_down).pos_name_download
    
    If out_pos_name_download > -1 Then
        name_download(out_pos_name_download).num_socket = -1

        If name_download(out_pos_name_download).header_received Then
            If name_download(out_pos_name_download).initial_size = 0 Then
                name_download(out_pos_name_download).initial_size = name_download(out_pos_name_download).file_size
            End If
            name_download(out_pos_name_download).old_percent = name_download(out_pos_name_download).old_percent + CByte(name_download(out_pos_name_download).bytes_recieved / name_download(out_pos_name_download).initial_size * 100)
        End If
        name_download(out_pos_name_download).range_begin = name_download(out_pos_name_download).range_begin + name_download(out_pos_name_download).bytes_recieved
        name_download(out_pos_name_download).get_request_made = False
        name_download(out_pos_name_download).header_received = False
        name_download(out_pos_name_download).bytes_recieved = 0
    End If
    current_download(out_pos_current_down).num_socket = -1

    If num_socket > -1 Then ' download has begun
        If Not come_from_treat_socket_closing Then
            treat_socket_closing num_socket, , , True
        End If
    Else 'remove from waiting_download if necessary
        remove_waiting_download_with_pos_name_download out_pos_name_download
    End If
End Sub




''''''''''''''''''''''functions for resuming file even if Coyotella as been closed




'''''''''''''''''''''' download_file_name.coy structure
'''''
''''' field1  |field2   | field3              | field4        | ..... |field nb_parts*2+3
''''' (single) (single)           (single)             (single)          tpartfile_info1 | .. |   tpartfile_infon
'''''         |         |next byte to receive |real range end | ..... |                  | .. |                   |
''''' nb_parts|file size|( next range begin ) |               | ..... |                  | .. |
'''''         |         |         for part1   |    for part1  | ..... |                  | .. |
'''''


Public Sub init_prepare_resume(ByVal saving_name As String, ByVal number_parts As Single, ByVal file_size As Single)
    On Error Resume Next
    'fill fields 1 and 2 of the structure (nb_part and file_size)
    Dim num_file            As Integer
    Dim type_len            As Byte
    Dim file_name           As String
    
    type_len = 4 'single
    
    file_name = my_incomplete_directory & saving_name & ".coy"

    If is_file_existing(file_name) Then
        Kill file_name
    End If
    
    num_file = FreeFile()
    Open file_name For Random Shared As num_file Len = type_len
        Put num_file, 1, number_parts
        Put num_file, 2, file_size
    Close num_file
    
End Sub

Private Sub prepare_resume(ByVal saving_name As String, ByVal num_part As Single, ByVal next_byte_to_receive As Single)
    On Error Resume Next
    'update range_begin (or next byte to receive) . Called each time we receive data
    Dim file_name           As String
    Dim num_file            As Integer
    Dim type_len            As Byte
    Dim num_record          As Single
    
    type_len = 4 'single
    file_name = my_incomplete_directory & saving_name & ".coy"

    num_record = num_part * 2 + 1

    num_file = FreeFile()
    Open file_name For Random Shared As num_file Len = type_len
        Put num_file, num_record, next_byte_to_receive
    Close num_file
End Sub

Private Sub update_resume_part_size(ByVal download_file_name As String, ByVal num_part As Integer, ByVal part_range_end As Single)
    On Error Resume Next
    ' update the range end of the part
    Dim file_name           As String
    Dim num_record          As Single
    Dim num_file            As Integer
    Dim type_len            As Byte
    type_len = 4 'single
    file_name = my_incomplete_directory & download_file_name & ".coy"

    num_record = 2 * num_part + 2
    
    num_file = FreeFile()
    Open file_name For Random Shared As num_file Len = type_len
        Put num_file, num_record, part_range_end
    Close num_file
End Sub

Private Sub prepare_resume_file_info_first(ByVal saving_name As String, ByVal ip As String, ByVal port As Long, _
                                     ByVal index As Single, ByVal serventID As String, ByVal cspeed As Single, _
                                     ByVal name As String, ByVal nb_parts As Integer)
    On Error Resume Next
    'sub called only if it's NOT a resume or recovery else we need
    'to update fields we must find the field, overwrite it (keeping info placed after the replaced field)
    'all these info are useful only in case of  recovery for multiple hosts download
    Dim num_file            As Integer
    Dim buffer              As String
    Dim file_name           As String
    Dim type_len            As Byte
    Dim pos                 As Long
    Dim f_size              As Long
    
    
    type_len = 4
    file_name = my_incomplete_directory & saving_name & ".coy"
    
    pos = (2 * nb_parts + 2) * type_len ' we should let necessary place for field not fill yet
    
    buffer = long_to_big_endian(CSng(Len(name) + 34)) & ip_encode(ip) & int_to_big_endian(port) & long_to_big_endian(index) & serventID & long_to_big_endian(cspeed) & name

    num_file = FreeFile()
    Open file_name For Binary Shared As num_file
        f_size = LOF(num_file)
        If f_size > pos Then
            pos = f_size
        End If
        Put num_file, pos + 1, buffer
    Close num_file
End Sub



Public Sub make_recovery(ByVal download_file_name As String)
    On Error Resume Next
    Dim num_file            As Integer
    Dim type_len            As Byte
    Dim all_hosts_info      As String
    Dim all_host()          As tpartfile_info
    Dim num_record          As Single
    Dim nb_records          As Single
    Dim file_name           As String
    Dim file_size           As Single
    Dim range_begin         As Single
    Dim range_end           As Single
    Dim cpt                 As Integer
    Dim pos                 As Long

    type_len = 4 'single
    file_name = my_incomplete_directory & download_file_name & ".coy"

    'get the initial file name
    
    num_file = FreeFile()
    Open file_name For Binary Access Read Shared As num_file
        'get number of record
        Get num_file, 1, nb_records
        Get num_file, 5, file_size 'binary
        all_hosts_info = Space$(LOF(num_file) - (2 * nb_records + 2) * type_len)
        Get num_file, (2 * nb_records + 2) * type_len + 1, all_hosts_info 'binary
    Close num_file

    pos = 0
    ReDim all_host(nb_records - 1)
    For cpt = 0 To nb_records - 1

        all_host(cpt).field_size = long_to_little_endian(Mid$(all_hosts_info, pos + 1, 4))
        all_host(cpt).ip = ip_decode(Mid$(all_hosts_info, pos + 1 + 4, 4))
        all_host(cpt).port = int_to_little_endian(Mid$(all_hosts_info, pos + 1 + 8, 2))
        all_host(cpt).index = long_to_little_endian(Mid$(all_hosts_info, pos + 1 + 10, 4))
        all_host(cpt).servent_id = Mid$(all_hosts_info, pos + 1 + 14, 16)
        all_host(cpt).cspeed = long_to_little_endian(Mid$(all_hosts_info, pos + 1 + 30, 4))
        all_host(cpt).name = Mid$(all_hosts_info, pos + 1 + 34, all_host(cpt).field_size - 34)
        pos = pos + all_host(cpt).field_size
    Next cpt

    num_file = FreeFile()
    Open file_name For Random Shared As num_file Len = type_len

        'realise as much as download as unfinsihed parts
        For cpt = 3 To 2 * nb_records + 1 Step 2
            Get num_file, cpt, range_begin
            Get num_file, cpt + 1, range_end 'random access file
            pos = (cpt - 3) / 2

            'add to known_files (in case of resume on other host)
            add_to_known_files all_host(pos).index, all_host(pos).name, file_size, all_host(pos).ip, all_host(pos).port, -1, all_host(pos).cspeed, all_host(pos).servent_id, False
            
            If range_begin < range_end Then
                ask_for_download all_host(pos).ip, all_host(pos).port, all_host(pos).index, _
                             all_host(pos).name, range_begin, range_end, all_host(pos).servent_id, _
                             file_size, all_host(pos).cspeed, pos + 1, False, nb_records, download_file_name, True
            'Else 'show "already completed" in the download_upload form
            End If
        Next cpt
    Close num_file

End Sub

Public Sub check_recoveries()
    On Error Resume Next
    Dim cpt                 As Integer
    Dim cpt2                As Integer
    Dim nb_recoveries       As Integer
    Dim file_size           As Single
    Dim nb_records          As Single
    Dim type_len            As Byte
    Dim not_completed_size  As Single
    Dim percent_realized    As Byte
    Dim file_name           As String
    Dim num_file            As Integer
    Dim range_begin         As Single
    Dim range_end    As Single
    
    range_end = 0
    
    nb_recoveries = find_recoveries(my_incomplete_directory)
    For cpt = 0 To nb_recoveries - 1
        'search percent size completed
        file_name = my_incomplete_directory & recovery_files(cpt)
        not_completed_size = 0
        type_len = 4
        num_file = FreeFile()
        Open file_name For Random Shared As num_file Len = type_len
            'get number of record
            Get num_file, 1, nb_records
            Get num_file, 2, file_size
            
            For cpt2 = 3 To 2 * nb_records + 2 Step 2
                Get num_file, cpt2, range_begin
                Get num_file, cpt2 + 1, range_end
                If range_end = 0 Then 'only for the last range end
                    range_end = file_size - 1
                End If
                not_completed_size = not_completed_size + (range_end - range_begin + 1)
            Next cpt2
        Close num_file
        If not_completed_size <= -1 Then not_completed_size = 0
        not_completed_size = not_completed_size - 1 'avoid troubles with last range end
        Dim lResult As Long
        
        If file_size <= 0 Or not_completed_size < 0 Or file_size - not_completed_size < 0 Then 'error
            'delete
            lResult = MessageBox(0, "An error has occured in the recovery of " & file_name, Form_main.Caption, vbExclamation)
            'range end before=0
        Else
            percent_realized = (file_size - not_completed_size) / file_size * 100
            
            ask_recovery Mid$(recovery_files(cpt), 1, Len(recovery_files(cpt)) - 4), percent_realized
        End If
    Next cpt

End Sub
    
Private Sub ask_recovery(ByVal file_name As String, ByVal percent_completed As Byte)
    On Error Resume Next
    form_ask_for_recovery.label_name_value = file_name
    form_ask_for_recovery.label_percent_value = percent_completed & " %"
    form_ask_for_recovery.Show vbModal
End Sub


Public Function asked_giv(ByVal servent_id As String, ByVal file_index As Single) As Integer
    On Error Resume Next
    'if we have send a push <-> if we have in name_download the same servent id and the same file index
    'return position in name_download of a possible file to upload
    Dim cpt As Long
    asked_giv = -1
    Form_main.Timer_second.Enabled = False
    For cpt = UBound(waiting_giv) - 1 To 0 Step -1
        If name_download(waiting_giv(cpt).pos_name_download).servent_id = servent_id And _
            name_download(waiting_giv(cpt).pos_name_download).file_index = file_index And _
            name_download(waiting_giv(cpt).pos_name_download).get_request_made = False Then
            asked_giv = waiting_giv(cpt).pos_name_download
            remove_from_waiting_giv cpt
            Form_main.Timer_second.Enabled = True
            Exit Function
        End If
    Next cpt
    Form_main.Timer_second.Enabled = True
End Function

Public Sub add_to_waiting_giv(ByVal pos_name_download As Integer)
    On Error Resume Next
    Dim size As Integer
    size = UBound(waiting_giv)
    waiting_giv(size).push_validity_time = mypush_validity_time
    waiting_giv(size).pos_name_download = pos_name_download
    ReDim Preserve waiting_giv(size + 1)
End Sub

Public Sub remove_from_waiting_giv(ByVal pos As Integer)
    On Error Resume Next
    Dim size As Integer
    size = UBound(waiting_giv)
    waiting_giv(pos) = waiting_giv(size - 1)
    ReDim Preserve waiting_giv(size - 1)
End Sub
