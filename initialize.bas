Attribute VB_Name = "initialize"
Option Explicit

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

' files info
Public Type tfile
    file_name   As String
    file_size   As Single
    file_index  As Single
    full_path   As String
End Type
Public my_shared_files()        As tfile

Public Type tmyfile
    file_name           As String
    file_size           As Single
    file_index          As Single
    speed               As Long
    servent_id          As String
    ip                  As String
    port                As String
    num_socket_queryhit As Integer 'to send push only on the specified socket
    need_push           As Boolean
    have_uploaded       As Boolean
End Type
Public known_files()            As tmyfile 'limited to n elements, n as long _
                                            to change this search UBound(known_files) _
                                            and change counter type
Public known_files_max_size     As Long
Public known_files_position     As Long     ' allow to make a flushing

' download and upload
Public Type tdownload
    saving_name         As String 'allow to rename file if we have already one file with the same name
    file_name           As String
    file_index          As Single
    servent_id          As String
    range_begin         As Single
    range_end           As Single
    file_size           As Single
    num_socket          As Integer
    bytes_recieved      As Single
    ip                  As String
    port                As String
    num_part            As Integer
    old_percent         As Byte   ' for percent resume info
    initial_size            As Single ' for percent resume info
    pos_current_down    As Integer
    header_received     As Boolean
    get_request_made    As Boolean ' used only for multipart download on host with firewall (GIV)
    num_socket_queryhit As Integer
    need_to_push        As Boolean
End Type
Public name_download()          As tdownload 'contain information of file being downloading

Public Type twaiting_download
    num_socket             As Integer
    data                   As String
    position_name_download As Integer
    connection_tried       As Boolean
    push                   As Boolean
    num_socket_queryhit    As Integer
End Type
Public waiting_download()       As twaiting_download 'contain the request that will be sent for the download

Public Type tdownload_information
    file_name           As String ' the saving name
    file_range          As String
    file_index          As String
    size                As String
    cspeed              As String
    speed               As Single
    status              As String
    ip                  As String
    num_socket          As Integer
    remaining_bytes     As Single
    pos_name_download   As Integer
End Type
Public current_download() As tdownload_information 'contain information for form down_up

Public Type tupload_information
    file_name           As String
    file_range          As String
    file_index          As String
    size                As String
    speed               As Single
    status              As String
    ip                  As String
    num_socket          As Integer
    remaining_bytes     As Single
End Type
Public current_upload()  As tupload_information 'contain information for form down_up

Public Type tpartfile_info 'made to keep gnutella coherence
    field_size      As Single
    ip              As String
    port            As Long
    index           As Single
    servent_id      As String * 16
    cspeed          As Single
    name            As String
End Type

Public Type tother_host_info
    saving_name         As String
    file_name           As String
    range_begin         As Single
    range_end           As Single
    file_size           As Single
    num_part            As Integer
    pos_current_down    As Integer
    create_new_file     As Boolean
    old_nb_parts        As Integer
End Type

'information for socket

Public Type tsocket_state
    connection_type         As Byte 'free,dialing,uploading,downloading,server,not_defin_yet,ginving
    connection_direction    As Byte
    bytes_rcv               As Single ' allow to calculate the rate of each connection
    number_of_ping          As Long ' allow to remove connections pinging to frequently
    number_of_query         As Long ' allow to remove connections querying to frequently
    number_of_bogus         As Long ' allow to remove connections to much bogus packt (could be a bad synchronisation)
End Type
Public socket_state()           As tsocket_state  'specify the use of the socket by one of the following value 'use: socket_state(num_socket)
Public Const free               As Byte = 1 'tsocket_state.connection_type
Public Const dialing            As Byte = 2 '
Public Const uploading          As Byte = 3 '
Public Const downloading        As Byte = 4 '
Public Const server             As Byte = 5 '
Public Const not_define_yet     As Byte = 6
Public Const giving             As Byte = 7
Public Const incoming           As Byte = 1 'tsocket_state.connection_direction
Public Const outgoing           As Byte = 2 '
Public max_nb_socket            As Integer ' max number of socket to use
Public current_nb_dial          As Integer ' current number of socket used for dialing (ping pong query queryhit)
Public current_nb_upload        As Integer ' current number of socket used for uploading
Public current_nb_download      As Integer ' current number of socket used for downloading
Public current_nb_incoming      As Integer
Public max_dialing_hosts        As Integer
Public min_dialing_hosts        As Integer
Public max_upload               As Integer
Public max_download             As Integer
Public max_incoming_connection  As Integer
Public remaining_data()         As String


' network info
Public Type thost
    ip                As String
    port              As String
    nb_shared_files   As Single
    nb_shared_kbytes  As Single
End Type
Public known_hosts()            As thost

Public dummy_ip()       As String 'dummy ip like 0.0.0.0 or 127.0.0.1 --> send a push directly don't try a dummy connection before
Public banished_ip()                    As String 'ip in this array can not connect (and download files) from you

Public Type troute
    ID                As String 'descriptor_ID || servent_ID
    payload           As Byte
    num_socket        As Integer
End Type
Public routing_table()          As troute
Public routing_table_max_size   As Long    'maximum size of the routing table
Public routing_table_position   As Long    'after ubound(routing_table) has come up to routing_table_max_size
                                           'it make a flushing

Public Type ttraffic_info
    sent_ping           As Double
    sent_pong           As Double
    sent_query          As Double
    sent_queryhit       As Double
    sent_push           As Double
    rcv_ping            As Double
    rcv_pong            As Double
    rcv_query           As Double
    rcv_queryhit        As Double
    rcv_push            As Double
    rcv_bogus           As Double
    array_size          As Integer
    position            As Integer 'for flushing
    should_remove       As Boolean
End Type
Public traffic_info             As ttraffic_info

Public Type ttraffic_connected_ip
    ip                  As String
    port                As String
    incoming            As String
    state               As String
    num_socket          As Integer
    speed               As Single
End Type
Public traffic_connected_ip()   As ttraffic_connected_ip

Public Type ttraffic_payload
    payload_descriptor  As Byte
    info                As String
End Type
Public traffic_payload()        As ttraffic_payload

Public Type trename_file
    old_name            As String
    new_name            As String
    overwrite           As Boolean
End Type
Public renamed_file             As trename_file

Public Type tserver
    ip          As String
    port        As String
End Type
Public known_gnutella_server()   As tserver

Public strknown_gnutella_server  As String

Public Type tretry_down
    remaining_time              As Integer
    pos_name_download           As Integer
End Type
Public retry_download()         As tretry_down

Public Type twaiting_giv
    pos_name_download           As Integer
    push_validity_time          As Integer
End Type
Public waiting_giv()            As twaiting_giv

Public Type tpush_upload
    data            As String
    num_socket      As Integer
    ip              As String
    port            As Long
End Type
Public push_upload()            As tpush_upload


'your informations
Public my_ttl                   As Byte 'ttl defined in option
Public my_port                  As Long 'your local port to listen on defined in option
Public my_ip                    As String 'your ip
Public my_nb_shared_files       As Long
Public my_nb_kilobytes_shared   As Long
Public my_speed                 As Long 'your speed defined in option
Public my_min_speed             As Long
Public my_servent_id            As String * 16 'at the start of the program
Public myshared_directories()   As String
Public my_download_directory    As String
Public my_incomplete_directory  As String
Public mynumber_of_retry        As Integer
Public min_host_shared_file     As Long
Public min_host_shared_kb       As Long
Public my_min_search_length     As Byte
Public other_min_search         As Byte
Public mymax_ping_per_minute    As Long
Public mymax_query_per_minute   As Long
Public mymax_bogus_per_minute   As Long
Public mypush_validity_time     As Integer
Public max_initial_ttl          As Integer
Public spy_all_query_hits       As Boolean
Public allow_upload             As Boolean
Public forward_on_outgoing_only As Boolean 'clip2 reflector
Public forward_ping             As Boolean
Public forward_query            As Boolean
Public sharing_simulation       As Boolean
Public simulation_nb_files      As Long
Public simulation_size          As Long
Public nb_of_pong_for_a_ping    As Integer
Public opt_all_and_words        As String
Public opt_all_or_words         As String
Public opt_all_not_words        As String
Public ini_file_name            As String
Public send_my_computers_info   As Boolean
Public launch_server            As Boolean
Public nb_parts_for_download    As Integer
Public downloader_log_file_name As String
Public log_bad_downloaders      As Boolean
Public log_good_downloaders     As Boolean
Public recovery_files()         As String
Public known_files_restrictions_array() As Variant
Public retry_down_on_busy_server_every As Integer
Public force_ip                 As Boolean
Public forced_ip                As String
Public include_sub_dir          As Boolean
Public auto_add_down_to_shared_files As Boolean

Public my_descriptors_ID()         As String  'contain all descriptor ID of OUR messages
Public my_descriptors_ID_max_size  As Long
Public my_descriptors_ID_position  As Long

Public quit_program             As Boolean
Public disconnected_from_gnutella_network   As Boolean 'you WANT to be disconnected
Public connected_to_gnutella_network    As Boolean 'you ARE currently connected



 


Public Sub initialize_app()
    
    On Error Resume Next
    'for fill_known_files_restrictions_array
    Dim min_file_size As Long
    Dim max_file_size As Long
    Dim all_and_words As String
    Dim all_or_words  As String
    Dim all_not_words As String
    
    'init arrays
    ReDim Document_search(0)
    ReDim Document_search_deleted(0)
    ReDim descriptorID_to_num_form_search(1, 0)
    ReDim my_descriptors_ID(0)
    ReDim waiting_download(0)
    ReDim name_download(0)
    ReDim known_hosts(0)
    
    ReDim retry_download(0)
    ReDim dummy_ip(0)
    ReDim banished_ip(0)
    ReDim socket_state(0)
    my_servent_id = GetGUID
    ReDim traffic_connected_ip(0)
    ReDim current_download(0)
    ReDim current_upload(0)
    ReDim known_files(0)
    ReDim routing_table(0)
    ReDim myshared_directories(1)
    ReDim known_files_restrictions_array(4)
    ReDim push_upload(0)
    ReDim waiting_giv(0)
    known_files_restrictions_array(0) = 0
    known_files_restrictions_array(1) = 0
    known_files_restrictions_array(2) = 0
    known_files_restrictions_array(3) = 0
    known_files_restrictions_array(4) = 0
    my_ip = "127.0.0.1" 'for a recovery or when you don't make server
    ini_file_name = App.Path & "\coyotella.ini"
    
    'read ini values
    Call get_ini_value

    'init using .ini values
    
    ReDim traffic_payload(traffic_info.array_size)
    ReDim remaining_data(1)
    
    
     'remove the 2 following line to add ip to known_hosts with 0 shared files or kb
    'If min_host_shared_file <= 0 Then min_host_shared_file = 1
    'If min_host_shared_kb <= 0 Then min_host_shared_kb = 1
    
    Dim lResult As Long
label_verify:
    'check only the existence of first shared directory
    If Not is_folder_existing(myshared_directories(0)) _
        Or Not is_folder_existing(my_download_directory) _
        Or Not is_folder_existing(my_incomplete_directory) _
        Then
            lResult = MessageBox(0, "Some directories are false" & vbCrLf & "Please update directories", Form_main.Caption, vbExclamation)
            Form_options.Show vbModal
            GoTo label_verify
    End If
    
    'share files
    Call share_files


    'fill known_files_restrictions_array
    fill_known_files_restrictions_array min_file_size, max_file_size, all_and_words, all_or_words, all_not_words

    'check if there's recoveries
    Call check_recoveries

    If force_ip = True Then
        my_ip = forced_ip
    End If
End Sub
