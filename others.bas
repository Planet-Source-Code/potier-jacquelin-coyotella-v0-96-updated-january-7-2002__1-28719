Attribute VB_Name = "others"
Public Function s_to_hms(seconds As Single) As String
    On Error Resume Next
    Dim tmp1 As Single
    Dim tmp2 As Single
    tmp1 = seconds Mod 60
    s_to_hms = ":" & CStr(tmp1)
    tmp2 = (seconds - tmp1) / 60
    tmp1 = tmp2 Mod 60
    s_to_hms = ":" & CStr(tmp1) & s_to_hms
    tmp2 = (tmp2 - tmp1) / 60
    tmp1 = tmp2 Mod 60
    s_to_hms = CStr(tmp1) & s_to_hms
    s_to_hms = Format(s_to_hms, "hh:mm:ss")
End Function

Public Function byte_to_hexa(ByVal value As Byte) As String
    On Error Resume Next
    Dim lsb As Byte
    Dim msb As Byte
    lsb = value Mod 16
    msb = ((value - lsb) / 16) Mod 16
    byte_to_hexa = Hex$(msb) & Hex$(lsb)
End Function


Public Sub remove_from_my_shared_files(ByVal place As Long)
    On Error Resume Next
    Dim size As Long
    size = UBound(my_shared_files)
    my_nb_kilobytes_shared = my_nb_kilobytes_shared - my_shared_files(place).file_size / 1000
    my_nb_shared_files = my_nb_shared_files - 1
    my_shared_files(place) = my_shared_files(size - 1)
    ReDim Preserve my_shared_files(size - 1)
End Sub



Public Sub remove_tab(ByVal window_handle As Long)
    On Error Resume Next
    Dim cpt As Integer
    For cpt = 1 To Form_main.TabStrip.Tabs.Count
        If Form_main.TabStrip.Tabs(cpt).Tag = window_handle Then
            Form_main.TabStrip.Tabs.Remove cpt
            Exit For
        End If
    Next cpt
    If Form_main.TabStrip.Tabs.Count = 0 Then Form_main.CoolBar1.Bands(2).Visible = False
End Sub

Public Sub activate_tab(ByVal window_handle As Long)
    On Error Resume Next
    Dim cpt As Integer
    For cpt = 1 To Form_main.TabStrip.Tabs.Count
        If Form_main.TabStrip.Tabs(cpt).Tag = window_handle Then
            Form_main.TabStrip.Tabs(cpt).Selected = True
            Exit For
        End If
    Next cpt
End Sub

Public Sub update_caption_tab(ByVal my_caption As String, ByVal window_handle As Long)
    On Error Resume Next
    Dim cpt As Integer
    For cpt = 1 To Form_main.TabStrip.Tabs.Count
        If Form_main.TabStrip.Tabs(cpt).Tag = window_handle Then
            Form_main.TabStrip.Tabs(cpt).Caption = my_caption
            Exit For
        End If
    Next cpt
End Sub

Public Function is_form_in_tabstrip(ByVal handle As Long) As Boolean
    On Error Resume Next
    'true if is in tabstrip
    Dim cpt As Integer
    For cpt = 1 To Form_main.TabStrip.Tabs.Count
        If Form_main.TabStrip.Tabs(cpt).Tag = handle Then
            is_form_in_tabstrip = True
            Exit For
        End If
    Next cpt
End Function


