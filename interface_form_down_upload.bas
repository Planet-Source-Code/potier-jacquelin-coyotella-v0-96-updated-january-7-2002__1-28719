Attribute VB_Name = "interface_form_down_upload"
Option Explicit

''''''''''''''''''''''''''''''''' function for interface (form_download_upload)
Public Function add_download_to_interface(ByVal file_name As String, ByVal file_index As String, ByVal file_range As String, ByVal ip As String, ByVal speed As String, ByVal status As String, ByVal num_socket As Integer, ByVal pos_name_download As Integer)
    On Error Resume Next
    Dim pos     As Integer
    Dim pos1    As Integer

    pos = UBound(current_download)
    
    With Form_download_upload.ListView_download
        pos1 = .ListItems.Count + 1
        .ListItems.Add pos1, , file_name
        .ListItems(pos1).SubItems(1) = file_range
        .ListItems(pos1).SubItems(2) = speed
        .ListItems(pos1).SubItems(3) = status
        .ListItems(pos1).SubItems(4) = ""
        .ListItems(pos1).SubItems(5) = ip
        .ListItems(pos1).Tag = pos
    End With
    
    With current_download(pos)
        .cspeed = speed
        .file_index = file_index
        .file_name = file_name
        .file_range = file_range
        .ip = ip
        .num_socket = num_socket
        .speed = 0
        .status = status
        .pos_name_download = pos_name_download
    End With
    ReDim Preserve current_download(pos + 1)
    add_download_to_interface = pos
End Function


Public Sub update_status_download(ByVal status As String, ByVal num_socket As Integer, Optional ByVal BytesRemaining As Single = 0)
    On Error Resume Next
    ' work if num_socket<>-1
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    pos1 = find_corresponding_down1(num_socket)
    If pos1 > -1 Then
        current_download(pos1).status = status
        current_download(pos1).remaining_bytes = BytesRemaining
        pos2 = find_corresponding_down2(pos1)
        If pos2 > 0 Then
            Form_download_upload.ListView_download.ListItems(pos2).SubItems(3) = status
'            Form_download_upload.ListView_download.Refresh
        End If
    End If
End Sub

Public Sub update_status_download2(ByVal status As String, ByVal pos_current_down As Integer)
    On Error Resume Next
    ' work if num_socket= -1
    Dim pos As Integer

    current_download(pos_current_down).status = status
    
    pos = find_corresponding_down2(pos_current_down)
    If pos > 0 Then
        Form_download_upload.ListView_download.ListItems(pos).SubItems(3) = status
'            Form_download_upload.ListView_download.Refresh
    End If
End Sub

Public Sub update_speed_remain_download(ByVal speed As Integer, ByVal num_socket As Integer, Optional ByVal BytesRemaining As Single = 0)
    On Error Resume Next
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    pos1 = find_corresponding_down1(num_socket)
    If pos1 > -1 Then
        current_download(pos1).speed = speed
        pos2 = find_corresponding_down2(pos1)
        If pos2 > 0 Then
            Form_download_upload.ListView_download.ListItems(pos2).SubItems(4) = speed
            If speed > 0.01 Then
                Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = s_to_hms(BytesRemaining / speed / 1000)
            Else
                If BytesRemaining = 0 Then
                    Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = "00:00:00"
                Else
                    Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = "Unknown"
                End If
            End If
        End If
    End If
End Sub


Public Sub update_speed_remain_download2(ByVal speed As Single, pos_current_down As Integer, Optional ByVal BytesRemaining As Single = 0) 'made with timer
    On Error Resume Next
    Dim pos2 As Integer

    pos2 = find_corresponding_down2(pos_current_down)
    If pos2 > 0 Then
        Form_download_upload.ListView_download.ListItems(pos2).SubItems(4) = speed
        If speed > 0.001 Then
            Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = s_to_hms(BytesRemaining / speed / 1000)
        Else
            If BytesRemaining = 0 Then
                Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = "00:00:00"
            Else
                Form_download_upload.ListView_download.ListItems(pos2).SubItems(6) = "Unknown"
            End If
        End If
    End If

End Sub

Public Sub update_range_download(ByVal range As String, num_socket As Integer)
    On Error Resume Next
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    pos1 = find_corresponding_down1(num_socket)
    If pos1 > -1 Then
        current_download(pos1).file_range = range
        pos2 = find_corresponding_down2(pos1)
        If pos2 > 0 Then
            Form_download_upload.ListView_download.ListItems(pos2).SubItems(1) = range
        End If
    End If
End Sub

Private Function find_corresponding_down1(num_socket) As Integer
    On Error Resume Next
    'return the position in current_download
    Dim cpt             As Integer
    
    For cpt = 0 To UBound(current_download) - 1
        If current_download(cpt).num_socket = num_socket Then
            find_corresponding_down1 = cpt
            Exit Function
        End If
    Next cpt

    find_corresponding_down1 = -1 'error
End Function

Private Function find_corresponding_down2(num_current_down As Integer) As Integer
    On Error Resume Next
    'return the position of the listindex for the position in current_download()
    Dim cpt             As Integer

    For cpt = 1 To Form_download_upload.ListView_download.ListItems.Count
        If Form_download_upload.ListView_download.ListItems(cpt).Tag = num_current_down Then
            find_corresponding_down2 = cpt
            Exit Function
        End If
    Next cpt

    find_corresponding_down2 = -1 'error
End Function


Public Sub add_upload_to_interface(ByVal file_name As String, ByVal file_range As String, ByVal ip As String, ByVal status As String, ByVal file_index As String, ByVal num_socket As Integer)
    On Error Resume Next
    Dim pos     As Integer
    Dim pos1    As Integer

    pos = UBound(current_upload)
    
    With Form_download_upload.ListView_upload
        pos1 = .ListItems.Count + 1
        .ListItems.Add pos1, , file_name
        .ListItems(pos1).SubItems(1) = file_range
        .ListItems(pos1).SubItems(2) = ip
        .ListItems(pos1).SubItems(3) = status
        .ListItems(pos1).SubItems(4) = ""
        .ListItems(pos1).Tag = pos
    End With
    
    With current_upload(pos)
        .file_index = file_index
        .file_name = file_name
        .file_range = file_range
        .ip = ip
        .speed = 0
        .status = status
        .num_socket = num_socket
    End With
    ReDim Preserve current_upload(pos + 1)
End Sub


Public Function give_status_upload(ByVal num_socket As Integer)
    On Error Resume Next
    Dim pos1    As Integer
    Dim pos2    As Integer
    
    pos1 = find_corresponding_up1(num_socket)
    If pos1 > -1 Then
        give_status_upload = current_upload(pos1).status
    End If
End Function

Public Sub update_status_upload(ByVal status As String, ByVal num_socket As Integer, Optional ByVal BytesRemaining As Single = 0)
    On Error Resume Next
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    pos1 = find_corresponding_up1(num_socket)
    If pos1 > -1 Then
        current_upload(pos1).status = status
        current_upload(pos1).remaining_bytes = BytesRemaining
        pos2 = find_corresponding_up2(pos1)
        If pos2 > 0 Then
            Form_download_upload.ListView_upload.ListItems(pos2).SubItems(3) = status
        End If
    End If
End Sub

Public Sub update_speed_remain_upload(ByVal speed As Integer, ByVal num_socket As Integer, Optional ByVal BytesRemaining As Single = 0)
    On Error Resume Next
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    pos1 = find_corresponding_up1(num_socket)
    If pos1 > -1 Then
        current_upload(pos1).speed = speed
        pos2 = find_corresponding_up2(pos1)
        If pos2 > 0 Then
            Form_download_upload.ListView_upload.ListItems(pos2).SubItems(4) = speed
            If speed > 0.01 Then
                Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = s_to_hms(BytesRemaining / speed / 1000)
            Else
                If BytesRemaining = 0 Then
                    Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = "00:00:00"
                Else
                    Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = "Unknown"
                End If
            End If
        End If
    End If
End Sub

Public Sub update_speed_remain_upload2(ByVal speed As Single, ByVal pos_current_up As Integer, Optional ByVal BytesRemaining As Single = 0) 'made with timer
    On Error Resume Next
    Dim pos2 As Integer

    pos2 = find_corresponding_up2(pos_current_up)
    If pos2 > 0 Then
        Form_download_upload.ListView_upload.ListItems(pos2).SubItems(4) = speed
        
        If speed > 0.001 Then
            Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = s_to_hms(BytesRemaining / speed / 1000)
        Else
            If BytesRemaining = 0 Then
                Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = "00:00:00"
            Else
                Form_download_upload.ListView_upload.ListItems(pos2).SubItems(5) = "Unknown"
            End If
        End If
    End If

End Sub

Private Function find_corresponding_up1(num_socket) As Integer
    On Error Resume Next
    'return the position in current_download
    Dim cpt             As Integer
    
    For cpt = 0 To UBound(current_upload) - 1
        If current_upload(cpt).num_socket = num_socket Then
            find_corresponding_up1 = cpt
            Exit Function
        End If
    Next cpt

    find_corresponding_up1 = -1 'error
End Function

Private Function find_corresponding_up2(num_current_up As Integer) As Integer
    On Error Resume Next
    'return the position of the listindex for the position in current_download()
    Dim cpt             As Integer

    For cpt = 1 To Form_download_upload.ListView_upload.ListItems.Count
        If Form_download_upload.ListView_upload.ListItems(cpt).Tag = num_current_up Then
            find_corresponding_up2 = cpt
            Exit Function
        End If
    Next cpt

    find_corresponding_up2 = -1 'error
End Function

Public Function find_selected_down() As Integer
    On Error Resume Next
    Dim cpt             As Integer
    find_selected_down = -1
    With Form_download_upload.ListView_download
        For cpt = 1 To .ListItems.Count 'find selection
            If .ListItems(cpt).Selected Then
                find_selected_down = cpt
                Exit For
            End If
        Next cpt
    End With
End Function

Public Function find_selected_up() As Integer
    On Error Resume Next
    Dim cpt             As Integer
    With Form_download_upload.ListView_upload
        For cpt = 1 To .ListItems.Count 'find selection
            If .ListItems(cpt).Selected Then
                find_selected_up = cpt
                Exit For
            End If
        Next cpt
    End With
End Function

Public Sub dissociate_current_download_num_socket(ByVal num_socket As Integer)
    On Error Resume Next
    Dim cpt         As Integer
    For cpt = 0 To UBound(current_download) - 1
        If current_download(cpt).num_socket = num_socket Then
            current_download(cpt).num_socket = -1
            Exit For
        End If
    Next cpt
End Sub

Public Sub dissociate_current_upload_num_socket(ByVal num_socket As Integer)
    On Error Resume Next
    Dim cpt         As Integer
    For cpt = 0 To UBound(current_upload) - 1
        If current_upload(cpt).num_socket = num_socket Then
            current_upload(cpt).num_socket = -1
            Exit For
        End If
    Next cpt
End Sub

Public Sub remove_from_current_download(ByVal pos_current_down As Integer)
    On Error Resume Next
    Dim cpt                     As Integer
    Dim array_size              As Integer
    Dim tmparray()              As Integer
    Dim pos_name_download       As Integer
    array_size = UBound(current_download)
    ReDim tmparray(array_size - 1)

    For cpt = 1 To Form_download_upload.ListView_download.ListItems.Count
        tmparray(Form_download_upload.ListView_download.ListItems.Item(cpt).Tag) = cpt
    Next cpt
    
    For cpt = pos_current_down To array_size - 2
        current_download(cpt) = current_download(cpt + 1)
        'update name_download info
        pos_name_download = current_download(cpt).pos_name_download
        If pos_name_download > -1 Then
            name_download(pos_name_download).pos_current_down = cpt
        End If
        'update tag info
        Form_download_upload.ListView_download.ListItems.Item(tmparray(cpt + 1)).Tag = cpt
    Next cpt
    
   
    ReDim Preserve current_download(array_size - 1)
End Sub

Public Sub remove_from_current_upload(ByVal pos_current_down As Integer)
    On Error Resume Next
    Dim cpt         As Integer
    Dim array_size  As Integer
    Dim tmparray()   As Integer
    array_size = UBound(current_upload)
    ReDim tmparray(array_size - 1)
    
    For cpt = 1 To Form_download_upload.ListView_upload.ListItems.Count
        tmparray(Form_download_upload.ListView_upload.ListItems.Item(cpt).Tag) = cpt
    Next cpt
    
    For cpt = pos_current_down To array_size - 2
        current_upload(cpt) = current_upload(cpt + 1)
        'update tag info
        Form_download_upload.ListView_upload.ListItems.Item(tmparray(cpt + 1)).Tag = cpt
    Next cpt
    ReDim Preserve current_upload(array_size - 1)
End Sub
