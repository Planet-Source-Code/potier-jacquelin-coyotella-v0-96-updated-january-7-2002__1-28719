Attribute VB_Name = "knowhosts"
Option Explicit

Public Function add_to_known_hosts(ByVal ip As String, ByVal port As String, ByVal nb_shared_files As Single, ByVal nb_shared_kbytes As Single) As Boolean
    On Error Resume Next
    'return true if added
    Dim already_exist           As Boolean
    Dim found                   As Boolean
    Dim cpt                     As Long
    Dim size
    
    
     If nb_shared_files < min_host_shared_file Or nb_shared_kbytes < min_host_shared_kb Then Exit Function
    
    'If is_ip_forbidden Then Exit Sub 'don't show all known host if you let this
    
    'search if ip is already present in the array known_host
    already_exist = False
    size = UBound(known_hosts)
    For cpt = 0 To size - 1
        If known_hosts(cpt).ip = ip And known_hosts(cpt).port = port Then
            already_exist = True
            Exit For
        End If
    Next cpt
    If Not already_exist Then
        'add to interface (should be before adding to array because it causes the loading of the form)
        add_item_known_host ip, port, nb_shared_files, nb_shared_kbytes
        ' add to array
        known_hosts(size).ip = ip
        known_hosts(size).port = port
        known_hosts(size).nb_shared_files = CStr(nb_shared_files)
        known_hosts(size).nb_shared_kbytes = CStr(nb_shared_kbytes)
        ReDim Preserve known_hosts(size + 1)
    Else 'update host information
        'update array
        known_hosts(cpt).nb_shared_files = nb_shared_files
        known_hosts(cpt).nb_shared_kbytes = nb_shared_kbytes
        'update interface
            'find item
        For cpt = 1 To Form_known_hosts.ListView_known_hosts.ListItems.Count
            If Form_known_hosts.ListView_known_hosts.ListItems.Item(cpt) = ip Then
                found = True
                Exit For
            End If
        Next cpt
            'update item
        If found Then
            Form_known_hosts.ListView_known_hosts.ListItems(cpt).SubItems(2) = CStr(nb_shared_files)
            Form_known_hosts.ListView_known_hosts.ListItems(cpt).SubItems(3) = CStr(nb_shared_kbytes)
        End If
    End If
    add_to_known_hosts = True
End Function


Public Sub remove_from_known_hosts(ByVal ip As String)
    On Error Resume Next
    Dim cpt                 As Long
    Dim cpt2                As Long
    Dim size
    size = UBound(known_hosts)
    'remove from array
    For cpt = 0 To size
        If known_hosts(cpt).ip <> ip Then
                known_hosts(cpt2) = known_hosts(cpt)
            cpt2 = cpt2 + 1
        End If
    Next cpt
    ReDim Preserve known_hosts(size - 1)
    'remove from  interface
    remove_item_from_known_hosts ip
End Sub

Public Sub add_item_known_host(ByVal ip As String, ByVal port As String, ByVal nb_shared_files As String, ByVal nb_shared_kbytes As String)
    On Error Resume Next
    Dim last_pos            As Long
    last_pos = Form_known_hosts.ListView_known_hosts.ListItems.Count + 1
    With Form_known_hosts.ListView_known_hosts
        .ListItems.Add last_pos, , ip
        .ListItems(last_pos).SubItems(1) = port
        .ListItems(last_pos).SubItems(2) = nb_shared_files
        .ListItems(last_pos).SubItems(3) = nb_shared_kbytes
    End With
End Sub


Private Sub remove_item_from_known_hosts(ByVal ip As String)
    On Error Resume Next
    Dim cpt                 As Long
    With Form_known_hosts.ListView_known_hosts
    For cpt = 1 To .ListItems.Count
        If .ListItems.Item(cpt) = ip Then
           .ListItems.Remove (cpt)
        End If
    Next cpt
    End With
End Sub
