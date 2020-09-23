Attribute VB_Name = "module_options"
'API Declaration for ini files

'Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'

Public Sub update_ini()
    On Error Resume Next
    Dim cpt As Integer
    Dim strtmp As String
    
        
    WritePrivateProfileString "Default", "launch_server", get_english_value(launch_server), ini_file_name
    WritePrivateProfileString "Default", "allow_upload", get_english_value(allow_upload), ini_file_name
    WritePrivateProfileString "Default", "forward_on_outgoing_only", get_english_value(forward_on_outgoing_only), ini_file_name
    WritePrivateProfileString "Default", "forward_query", get_english_value(forward_query), ini_file_name
    WritePrivateProfileString "Default", "forward_ping", get_english_value(forward_ping), ini_file_name
    WritePrivateProfileString "Default", "sharing_simulation", get_english_value(sharing_simulation), ini_file_name
    WritePrivateProfileString "Default", "simulation_nb_files", CStr(simulation_nb_files), ini_file_name
    WritePrivateProfileString "Default", "simulation_size", CStr(simulation_size), ini_file_name
    WritePrivateProfileString "Default", "my_port", CStr(my_port), ini_file_name
    WritePrivateProfileString "Default", "my_speed", CStr(my_speed), ini_file_name
    WritePrivateProfileString "Default", "my_min_speed", CStr(my_min_speed), ini_file_name
    WritePrivateProfileString "Default", "max_nb_socket", CStr(max_nb_socket), ini_file_name
    WritePrivateProfileString "Default", "max_dialing_hosts", CStr(max_dialing_hosts), ini_file_name
    WritePrivateProfileString "Default", "min_dialing_hosts", CStr(min_dialing_hosts), ini_file_name
    WritePrivateProfileString "Default", "max_upload", CStr(max_upload), ini_file_name
    WritePrivateProfileString "Default", "max_download", CStr(max_download), ini_file_name
    WritePrivateProfileString "Default", "max_incoming_connection", CStr(max_incoming_connection), ini_file_name
    WritePrivateProfileString "Default", "min_host_shared_file", CStr(min_host_shared_file), ini_file_name
    WritePrivateProfileString "Default", "min_host_shared_kb", CStr(min_host_shared_kb), ini_file_name
    WritePrivateProfileString "Default", "my_min_search_length", CStr(my_min_search_length), ini_file_name
    WritePrivateProfileString "Default", "other_min_search", CStr(other_min_search), ini_file_name
    WritePrivateProfileString "Default", "mymax_bogus_per_minute", CStr(mymax_bogus_per_minute), ini_file_name
    WritePrivateProfileString "Default", "mymax_query_per_minute", CStr(mymax_query_per_minute), ini_file_name
    WritePrivateProfileString "Default", "mymax_ping_per_minute", CStr(mymax_ping_per_minute), ini_file_name
    WritePrivateProfileString "Default", "spy_all_query_hits", get_english_value(spy_all_query_hits), ini_file_name
    WritePrivateProfileString "Default", "my_ttl", CStr(my_ttl), ini_file_name
    WritePrivateProfileString "Default", "mypush_validity_time", CStr(mypush_validity_time), ini_file_name
    WritePrivateProfileString "Default", "nb_of_pong_for_a_ping", CStr(nb_of_pong_for_a_ping), ini_file_name
    WritePrivateProfileString "Default", "max_initial_ttl", CStr(max_initial_ttl), ini_file_name
    WritePrivateProfileString "Default", "retry_down_on_busy_server_every", CStr(retry_down_on_busy_server_every), ini_file_name
    WritePrivateProfileString "Default", "nb_parts_for_download", CStr(nb_parts_for_download), ini_file_name
    WritePrivateProfileString "Default", "routing_table_max_size", CStr(routing_table_max_size), ini_file_name
    WritePrivateProfileString "Default", "known_files_max_size", CStr(known_files_max_size), ini_file_name
    WritePrivateProfileString "Default", "my_descriptors_ID_max_size", CStr(my_descriptors_ID_max_size), ini_file_name
    WritePrivateProfileString "Default", "traffic_info_array_size", CStr(traffic_info.array_size), ini_file_name
    WritePrivateProfileString "Default", "strknown_gnutella_server", strknown_gnutella_server, ini_file_name
    WritePrivateProfileString "Default", "forced_ip", forced_ip, ini_file_name
    WritePrivateProfileString "Default", "force_ip", get_english_value(force_ip), ini_file_name
    WritePrivateProfileString "Default", "include_sub_dir", get_english_value(include_sub_dir), ini_file_name
    WritePrivateProfileString "Default", "auto_add_down_to_shared_files", get_english_value(auto_add_down_to_shared_files), ini_file_name
    
    strtmp = ""
    For cpt = 0 To UBound(dummy_ip)
        If dummy_ip(cpt) <> "" Then strtmp = strtmp & dummy_ip(cpt) & ";"
    Next cpt
    WritePrivateProfileString "Default", "dummy_ip", strtmp, ini_file_name


    WritePrivateProfileString "Restrictions", "min_file_size", CStr(known_files_restrictions_array(0)), ini_file_name
    WritePrivateProfileString "Restrictions", "max_file_size", CStr(known_files_restrictions_array(1)), ini_file_name
    WritePrivateProfileString "Restrictions", "all_and_words", opt_all_and_words, ini_file_name
    WritePrivateProfileString "Restrictions", "all_or_words", opt_all_or_words, ini_file_name
    WritePrivateProfileString "Restrictions", "all_not_words", opt_all_not_words, ini_file_name

    WritePrivateProfileString "Directories", "log_bad_downloaders", get_english_value(log_bad_downloaders), ini_file_name
    WritePrivateProfileString "Directories", "log_good_downloaders", get_english_value(log_good_downloaders), ini_file_name
    WritePrivateProfileString "Directories", "downloader_log_file_name", downloader_log_file_name, ini_file_name
    WritePrivateProfileString "Directories", "my_incomplete_directory", my_incomplete_directory, ini_file_name
    WritePrivateProfileString "Directories", "my_download_directory", my_download_directory, ini_file_name
    strtmp = ""
    For cpt = 0 To UBound(myshared_directories)
        If myshared_directories(cpt) <> "" Then strtmp = strtmp & myshared_directories(cpt) & ";"
    Next cpt
    WritePrivateProfileString "Directories", "myshared_directories", strtmp, ini_file_name

End Sub
Private Function get_english_value(var As Boolean) As String
    If var Then ' to avoid language troubles
        get_english_value = "True"
    Else
        get_english_value = "False"
    End If
End Function
Public Sub get_ini_value()
    On Error Resume Next
    Dim retour              As String * 300
    Dim nb_char_retour      As Long
    Dim strtmp              As String
    Dim array_size          As Integer
    Dim tmp_array           As Variant
    Dim cpt                 As Integer
    Dim new_size            As Integer
    Dim delta               As Integer

    nb_char_retour = GetPrivateProfileString("Default", "launch_server", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            launch_server = True
        Else
            launch_server = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "allow_upload", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            allow_upload = True
        Else
            allow_upload = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "forward_on_outgoing_only", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            forward_on_outgoing_only = True
        Else
            forward_on_outgoing_only = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "forward_query", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            forward_query = True
        Else
            forward_query = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "forward_ping", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            forward_ping = True
        Else
            forward_ping = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "sharing_simulation", "False", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            sharing_simulation = True
        Else
            sharing_simulation = False
        End If

    simulation_nb_files = GetPrivateProfileInt("Default", "simulation_nb_files", 100, ini_file_name)
    simulation_size = GetPrivateProfileInt("Default", "simulation_size", 1000, ini_file_name)
    my_port = GetPrivateProfileInt("Default", "my_port", 6346, ini_file_name)
    my_speed = GetPrivateProfileInt("Default", "my_speed", 56, ini_file_name)
    my_min_speed = GetPrivateProfileInt("Default", "my_min_speed", 56, ini_file_name)
    max_nb_socket = GetPrivateProfileInt("Default", "max_nb_socket", 40, ini_file_name)
    max_dialing_hosts = GetPrivateProfileInt("Default", "max_dialing_hosts", 15, ini_file_name)
    min_dialing_hosts = GetPrivateProfileInt("Default", "min_dialing_hosts", 5, ini_file_name)
    max_upload = GetPrivateProfileInt("Default", "max_upload", 5, ini_file_name)
    max_download = GetPrivateProfileInt("Default", "max_download", 10, ini_file_name)
    max_incoming_connection = GetPrivateProfileInt("Default", "max_incoming_connection", 20, ini_file_name)
    min_host_shared_file = GetPrivateProfileInt("Default", "min_host_shared_file", 0, ini_file_name)
    min_host_shared_kb = GetPrivateProfileInt("Default", "min_host_shared_kb", 0, ini_file_name)
    my_min_search_length = GetPrivateProfileInt("Default", "my_min_search_length", 1, ini_file_name)
    other_min_search = GetPrivateProfileInt("Default", "other_min_search", 1, ini_file_name)
    mymax_bogus_per_minute = GetPrivateProfileInt("Default", "mymax_bogus_per_minute", 20, ini_file_name)
    mymax_query_per_minute = GetPrivateProfileInt("Default", "mymax_query_per_minute", 30, ini_file_name)
    mymax_ping_per_minute = GetPrivateProfileInt("Default", "mymax_ping_per_minute", 40, ini_file_name)
    nb_char_retour = GetPrivateProfileString("Default", "spy_all_query_hits", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            spy_all_query_hits = True
        Else
            spy_all_query_hits = False
        End If
    my_ttl = GetPrivateProfileInt("Default", "my_ttl", 5, ini_file_name)
    mypush_validity_time = GetPrivateProfileInt("Default", "mypush_validity_time", 30, ini_file_name)
    nb_of_pong_for_a_ping = GetPrivateProfileInt("Default", "nb_of_pong_for_a_ping", 2, ini_file_name)
    max_initial_ttl = GetPrivateProfileInt("Default", "max_initial_ttl", 30, ini_file_name)
    retry_down_on_busy_server_every = GetPrivateProfileInt("Default", "retry_down_on_busy_server_every", 30, ini_file_name)
    nb_parts_for_download = GetPrivateProfileInt("Default", "nb_parts_for_download", 3, ini_file_name)
    routing_table_max_size = GetPrivateProfileInt("Default", "routing_table_max_size", 10000, ini_file_name)
    known_files_max_size = GetPrivateProfileInt("Default", "known_files_max_size", 10000, ini_file_name)
    my_descriptors_ID_max_size = GetPrivateProfileInt("Default", "my_descriptors_ID_max_size", 10000, ini_file_name)
    traffic_info.array_size = GetPrivateProfileInt("Default", "traffic_info_array_size", 99, ini_file_name)
    nb_char_retour = GetPrivateProfileString("Default", "strknown_gnutella_server", "router.limewire.com:6346;gnutella.hostscache.com:6346;connect.newtella.net:6346", retour, 300, ini_file_name)
        strknown_gnutella_server = Mid$(retour, 1, nb_char_retour)
    Call fill_known_gnutella_server
    
    nb_char_retour = GetPrivateProfileString("Default", "forced_ip", "0.0.0.0", retour, 100, ini_file_name)
        forced_ip = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Default", "force_ip", "False", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            force_ip = True
        Else
            force_ip = False
        End If



    nb_char_retour = GetPrivateProfileString("Default", "include_sub_dir", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            include_sub_dir = True
        Else
            include_sub_dir = False
        End If

    nb_char_retour = GetPrivateProfileString("Default", "auto_add_down_to_shared_files", "False", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            auto_add_down_to_shared_files = True
        Else
            auto_add_down_to_shared_files = False
        End If
    nb_char_retour = GetPrivateProfileString("Default", "dummy_ip", "255.*.*.*;192.168.*.*;10.*.*.*;0.*.*.*;127.0.0.1", retour, 300, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        tmp_array = Split(strtmp, ";")
        array_size = UBound(tmp_array)
        new_size = array_size
        delta = 0
        If array_size > -1 Then
            ReDim dummy_ip(array_size)
            For cpt = 0 To array_size
                If tmp_array(cpt) <> "" Then
                    dummy_ip(cpt - delta) = CStr(tmp_array(cpt))
                Else
                    new_size = new_size - 1
                    delta = delta + 1
                    ReDim Preserve dummy_ip(new_size)
                End If
            Next cpt
        End If
    nb_char_retour = GetPrivateProfileString("Default", "send_my_computers_info", "False", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            send_my_computers_info = True
        Else
            send_my_computers_info = False
        End If

   'restrictions
    Dim min_file_size As Long
    Dim max_file_size As Long

    min_file_size = GetPrivateProfileInt("Restrictions", "min_file_size", 0, ini_file_name)
    max_file_size = GetPrivateProfileInt("Restrictions", "max_file_size", 0, ini_file_name)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_and_words", "", retour, 300, ini_file_name)
        opt_all_and_words = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_or_words", "", retour, 300, ini_file_name)
        opt_all_or_words = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Restrictions", "all_not_words", "", retour, 300, ini_file_name)
        opt_all_not_words = Mid$(retour, 1, nb_char_retour)
    'fill known_files_restrictions_array
    fill_known_files_restrictions_array min_file_size, max_file_size, opt_all_and_words, opt_all_or_words, opt_all_not_words
    
    
    'directories
    nb_char_retour = GetPrivateProfileString("Directories", "log_bad_downloaders", "True", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            log_bad_downloaders = True
        Else
            log_bad_downloaders = False
        End If
    nb_char_retour = GetPrivateProfileString("Directories", "log_good_downloaders", "False", retour, 6, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        If strtmp = "True" Then
            log_good_downloaders = True
        Else
            log_good_downloaders = False
        End If
    nb_char_retour = GetPrivateProfileString("Directories", "downloader_log_file_name", App.Path & "\" & "downloaders.log", retour, 100, ini_file_name)
        downloader_log_file_name = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Directories", "my_incomplete_directory", "my_incomplete_directory", retour, 300, ini_file_name)
        my_incomplete_directory = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Directories", "my_download_directory", "my_download_directory", retour, 300, ini_file_name)
        my_download_directory = Mid$(retour, 1, nb_char_retour)
    nb_char_retour = GetPrivateProfileString("Directories", "myshared_directories", "myshared_directories", retour, 100, ini_file_name)
        strtmp = Mid$(retour, 1, nb_char_retour)
        ReDim tmp_array(0)
        tmp_array = Split(strtmp, ";")
        array_size = UBound(tmp_array)
        new_size = array_size
        delta = 0
        If array_size > -1 Then
            ReDim myshared_directories(array_size)
            For cpt = 0 To array_size
                If tmp_array(cpt) <> "" Then
                    myshared_directories(cpt - delta) = CStr(tmp_array(cpt))
                Else
                    delta = delta + 1
                    new_size = new_size - 1
                    ReDim Preserve myshared_directories(new_size)
                End If
            Next cpt
        End If
End Sub

