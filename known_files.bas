Attribute VB_Name = "known_files_module"

Public Function belong_to_known_files(ByVal ip As String, ByVal file_index As Single, ByVal servent_id As String) As Long
    On Error Resume Next
    Dim position                   As Long
    belong_to_known_files = -1
    'check if file is already_existing (belonging to known_files with same ip and index)
    For position = 0 To UBound(known_files) - 1
        If known_files(position).ip = ip And known_files(position).file_index = file_index And known_files(position).servent_id = servent_id Then
            belong_to_known_files = position
            Exit Function
        End If
    Next position
End Function

Public Function add_to_known_files(ByVal file_index As Single, file_name As String, ByVal file_size As Single, ip As String, _
                                   port As Long, num_socket_queryhit As Integer, ByVal speed As Long, servent_id As String, _
                                   Optional ByVal check_filter As Boolean = True) As Long
    On Error Resume Next
    Dim position                   As Long
    Dim already_existing           As Boolean
    

    'check filtering conditions
    If check_filter Then
        If Not check_knownfiles_restrictions(file_name, file_size) Then
            add_to_known_files = -1
            Exit Function
        End If
    End If
    'adding file

    position = belong_to_known_files(ip, file_index, servent_id)

    If position = -1 Then
        position = UBound(known_files)
        
        If position < known_files_max_size Then
            ReDim Preserve known_files(position + 1)
        Else
            If known_files_position = known_files_max_size Then
                known_files_position = 0
                position = known_files_position
            Else
                known_files_position = known_files_position + 1
                position = known_files_position
            End If
        End If
    Else
        already_existing = True
    End If
    
    'update or fill file info
    With known_files(position)
        .file_index = file_index
        .file_name = file_name
        .file_size = file_size
        .ip = ip
        .port = port
        .num_socket_queryhit = num_socket_queryhit
        .servent_id = servent_id
        .speed = speed
    End With
    
    If already_existing Then
        add_to_known_files = -1 'to avoid to add another time to the form search
    Else
        add_to_known_files = position
    End If
End Function

Public Function check_knownfiles_restrictions(ByVal file_name As String, ByVal file_size As Single) As Boolean
    On Error Resume Next
    Dim cpt         As Integer
    Dim word_ok     As Boolean
    Dim pos         As Integer
    Dim begin       As Integer
    
    'check file size
        If known_files_restrictions_array(0) > file_size Then Exit Function 'min
        If known_files_restrictions_array(1) < file_size And known_files_restrictions_array(1) > 0 Then Exit Function 'max
        begin = 2
    'check and
        For cpt = 1 To known_files_restrictions_array(begin) 'known_files_restrictions_array(begin) contain the number of and words
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos < 1 Then Exit Function
        Next cpt
        begin = begin + known_files_restrictions_array(begin) + 1
    'check or
        For cpt = 1 To known_files_restrictions_array(begin)
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos > 0 Then word_ok = True
        Next cpt
        If known_files_restrictions_array(begin) > 0 And Not word_ok Then Exit Function
        begin = begin + known_files_restrictions_array(begin) + 1
    'check not
        For cpt = 1 To known_files_restrictions_array(begin)
            pos = InStr(1, file_name, known_files_restrictions_array(cpt + begin))
            If pos > 0 Then Exit Function
        Next cpt
        check_knownfiles_restrictions = True
End Function


Public Sub search_through_known_files(num_form_search As Integer, search_criteria As String, min_speed As Long)
    On Error Resume Next
    'search results
    Dim tmparray                As Variant
    Dim cpt                     As Long
    Dim cpt2                    As Long
    Dim valid                   As Boolean
    Dim strtmp                  As String

    strtmp = LCase$(search_criteria)


  'check search criteria
    'remove * from search
    Dim pos             As Long
    pos = InStr(1, strtmp, "*")
    Do While pos > 0 And Len(strtmp) > 0
        Select Case pos
            Case 1
                strtmp = Mid$(strtmp, 2)
            Case Len(search_criteria)
                strtmp = Mid$(strtmp, 1, Len(strtmp) - 1)
            Case Else
                strtmp = Mid$(strtmp, 1, pos - 1) _
                        & " " & Mid$(strtmp, pos + 1, Len(strtmp) - pos)
        End Select
        pos = InStr(1, strtmp, "*")
    Loop
    If Len(strtmp) <= 0 Then
        Exit Sub
    End If
    ' split on space
    tmparray = Split(strtmp, " ")
    'search
    
    Dim strtmp2         As String
    Document_search(num_form_search).list_search_results.Sorted = False 'avoid troubles
    For cpt = 0 To UBound(known_files) - 1
        strtmp2 = LCase$(known_files(cpt).file_name)
        valid = True
        For cpt2 = 0 To UBound(tmparray)
            If InStr(1, strtmp2, tmparray(cpt2)) <= 0 Then  'file name doesn't meet the query search criteria
                valid = False
                cpt2 = UBound(tmparray) 'to go out of for cpt2
            End If
        Next cpt2
        If valid Then
            If min_speed <= known_files(cpt).speed Then
                With known_files(cpt)
                    add_result_to_form_search num_form_search, .file_name, .file_size, .ip, .port, .speed, .file_index, .servent_id, .num_socket_queryhit, False, cpt
                End With
            End If
        End If
    Next cpt

End Sub




Public Sub fill_known_files_restrictions_array(ByVal min_kb_file_size As Long, ByVal max_kb_file_size As Long, ByVal all_and_words As String, ByVal all_or_words As String, ByVal all_not_words As String)
    On Error Resume Next
    Dim tmparray                As Variant
    Dim cpt                     As Integer
    Dim size                    As Integer
    Dim pos                     As Integer
    
    ReDim known_files_restrictions_array(2)
    known_files_restrictions_array(0) = min_kb_file_size 'in kb
    known_files_restrictions_array(1) = min_kb_file_size 'in kb

    pos = 2
    If all_and_words <> "" Then
        tmparray = Split(all_and_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos + 1)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
    
    If all_or_words <> "" Then
        tmparray = Split(all_or_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If
    
    If all_not_words <> "" Then
        tmparray = Split(all_not_words, ";")
        size = UBound(tmparray)
        known_files_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
        For cpt = 0 To size
            known_files_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve known_files_restrictions_array(pos)
        Next cpt
    Else
        known_files_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve known_files_restrictions_array(pos)
    End If

End Sub


