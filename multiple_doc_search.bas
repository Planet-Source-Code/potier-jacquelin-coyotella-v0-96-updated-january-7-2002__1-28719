Attribute VB_Name = "multiple_doc_search"
Option Explicit

Public descriptorID_to_num_form_search() As String 'corespondance between descriptor ID of a query and the mdi search form

Public Document_search_deleted()    As Boolean
Public Document_search()            As New Form_search

Public Sub new_document_search()
    On Error Resume Next
    Dim fIndex As Integer

    fIndex = FindFreeIndex()
    Document_search(fIndex).Caption = "Search " & CStr(fIndex)
    Document_search(fIndex).Framemultiplehostdown.Visible = False
    Document_search(fIndex).Lblnumberofresult.Top = 1300
    Document_search(fIndex).Labelnb_res.Top = 1300
    Document_search(fIndex).list_search_results.Top = 1600
    Document_search(fIndex).Tag = CStr(fIndex)
    Document_search(fIndex).Show
    
    Dim num_tab     As Integer
    num_tab = Form_main.TabStrip.Tabs.Count
    Form_main.CoolBar1.Bands(2).Visible = True
    Form_main.TabStrip.Tabs.Add num_tab + 1, , "Search " & CStr(fIndex)
    Form_main.TabStrip.Tabs(num_tab + 1).Tag = Document_search(fIndex).hWnd
    Form_main.TabStrip.Tabs(num_tab + 1).Selected = True
End Sub

Public Sub new_host_search(ByVal file_name As String, ByVal file_size As Single, ByVal range_begin As Single, _
                           ByVal range_end As Single, ByVal create_new_file As Boolean, _
                           Optional ByVal current_nbpart As Integer = 1, Optional ByVal num_part As Integer = 1, _
                           Optional pos_current_down As Integer = -1, Optional ByVal saving_name As String = "")
    On Error Resume Next
    Dim fIndex As Integer
    
    fIndex = FindFreeIndex()
    With Document_search(fIndex)
        .Caption = "Search " & CStr(fIndex) & " Finding hosts"
        .Tag = CStr(fIndex)
        .Combo_search.Text = file_name
        .lblfilesize = CStr(file_size)
        .my_popupmenu.Enabled = False
    End With
    'search through known files
    search_through_known_files fIndex, file_name, my_min_speed
    
    Document_search(fIndex).fill_info_for_other_host_down file_name, file_size, range_begin, range_end, _
                    create_new_file, current_nbpart, num_part, pos_current_down, saving_name
    
    Document_search(fIndex).Show
End Sub

Public Function FindFreeIndex() As Integer
    On Error Resume Next
    Dim i                   As Integer
    Dim ArrayCount          As Integer

    ArrayCount = UBound(Document_search)

    For i = 1 To ArrayCount
        If Document_search_deleted(i) Then
            FindFreeIndex = i
            Document_search_deleted(i) = False
            Exit Function
        End If
    Next
    
    ReDim Preserve Document_search(ArrayCount + 1)
    ReDim Preserve Document_search_deleted(ArrayCount + 1)
    FindFreeIndex = UBound(Document_search)
End Function

Public Sub add_result_to_form_search(ByVal num_form_search As Integer, file_name As String, file_size As Single, _
                                     ByVal ip As String, ByVal port As Long, ByVal speed As Single, _
                                     ByVal file_index As Single, ByVal serventID As String, num_socket As Integer, _
                                     Optional addtoknownfiles As Boolean = True, Optional place As Long)
                                     'note position must be set only if addtoknownfiles=false
    On Error Resume Next
    
        
    Dim last_pos         As Long
    Dim strtmp           As String 'allow a real sort for numbers
    Dim strtmp2          As String '
    Dim nb_space         As Byte   '
    nb_space = 13                  '
    
    Dim tag_value        As Long
    
    If addtoknownfiles = False Then 'result found in known_files
        tag_value = place
    Else
        tag_value = add_to_known_files(file_index, file_name, file_size, ip, port, num_socket, speed, serventID)
        If tag_value = -1 Then
            'file doesn't respect known_files filter or already exist --> could not be added to form search
            Exit Sub
        End If
    End If

    If Not Document_search(num_form_search).check_form_restrictions(file_name, file_size) Then
        'file doesn't respect form filter
        Exit Sub
    End If
    With Document_search(num_form_search).list_search_results
        'add to list
        last_pos = .ListItems.Count + 1
        .ListItems.Add last_pos, , file_name                'file_name
        strtmp = CStr(Int(file_size / 1000))                'to kb
        strtmp2 = Space$(nb_space - Len(strtmp)) & strtmp
        .ListItems(last_pos).SubItems(1) = strtmp2          'file_size
        strtmp = CStr(speed)
        strtmp2 = Space$(nb_space - Len(strtmp)) & strtmp
        .ListItems(last_pos).SubItems(2) = strtmp2          'speed
        .ListItems(last_pos).SubItems(3) = ip               'ip
        'if you want to put other column headers
        'strtmp = CStr(port)
        'strtmp2 = Space$(nb_space - Len(strtmp)) & strtmp
        '.ListItems(last_pos).SubItems(4) = strtmp2          'port
        '.ListItems(last_pos).SubItems(5) = CStr(file_index) 'file_index
        '.ListItems(last_pos).SubItems(6) = serventID        'servent_id
        
        .ListItems(last_pos).Tag = tag_value
    End With
    With Document_search(num_form_search).Labelnb_res
        .Caption = CLng(.Caption) + 1
    End With
End Sub
