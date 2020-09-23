VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_search 
   Caption         =   "Search"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   6075
   Icon            =   "form_search.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdrefinesearch 
      Caption         =   "Refine Search"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtttl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "7"
      Top             =   960
      Width           =   375
   End
   Begin VB.ComboBox Combo_search 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.TextBox txtminspeed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "50"
      Top             =   600
      Width           =   735
   End
   Begin MSComctlLib.ListView list_search_results 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size (Kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Speed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   252
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1092
   End
   Begin VB.Frame Framemultiplehostdown 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Text            =   "0"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmddownonthesehost 
         Caption         =   "Download"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "bytes"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "+ or -"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblfilesize 
         Caption         =   "lblfilesize"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "File size "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Download on all selected host(s)"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Label Lblnumberofresult 
      Caption         =   "Number of results :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Labelnb_res 
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Query TTL"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Min speed in kB/s"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWtileH 
         Caption         =   "Tile Horizontaly"
      End
      Begin VB.Menu mnuWTileV 
         Caption         =   "Tile Verticaly"
      End
      Begin VB.Menu mnuWArrange 
         Caption         =   "Arrange"
      End
   End
   Begin VB.Menu my_popupmenu 
      Caption         =   "my_popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuselect 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnudeselect 
         Caption         =   "Deselect all"
      End
      Begin VB.Menu mnuinvertselection 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnudownload 
         Caption         =   "Download all selected"
      End
      Begin VB.Menu mnumultipartdownload 
         Caption         =   "Multipart Download all selected"
      End
      Begin VB.Menu mnumultihostdownload 
         Caption         =   "Multi Host Download first selected"
      End
   End
End
Attribute VB_Name = "Form_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private need_to_remove          As Boolean
Private on_other_host_infos     As tother_host_info
Private my_restrictions_array() As Variant
' min_file_size,max_file_size,nb_and,and_value1,...and_valuen,nb_or,or_val1,..or_valn,nb_not,not_val1,..not_valn

Private Sub cmdSearch_Click()
    On Error Resume Next
    Dim lResult As Long
    If Len(Combo_search.Text) < my_min_search_length Then
        lResult = MessageBox(Me.hWnd, "Your search is to short !" & vbCrLf & "If you want to make a such search, modify options", Form_main.Caption, vbExclamation)
        Exit Sub
    End If
    
    Me.Caption = "Searching for " & Combo_search.Text
    
    Me.Labelnb_res = 0
    Me.list_search_results.ListItems.Clear
    
    update_caption_tab Combo_search.Text, Me.hWnd
    
    
    
    'add to combo
    Combo_search.AddItem Combo_search.Text
    'search through known files
    search_through_known_files Me.Tag, Me.Combo_search.Text, Me.txtminspeed.Text
    
    Dim strdata         As String
    Dim query_guid      As String * 16
    'make query before descriptor because descriptor needs to know the size of query
    strdata = make_string_query_data(txtminspeed.Text, Combo_search.Text)
    'make guid for the descriptor
    query_guid = GetGUID()
    'add to descriptorID_to_num_form_search
    If need_to_remove Then
        Call remove_from_descriptorID_to_num_form_search
    End If
    descriptorID_to_num_form_search(0, UBound(descriptorID_to_num_form_search, 2)) = query_guid
    descriptorID_to_num_form_search(1, UBound(descriptorID_to_num_form_search, 2)) = Me.Tag
    ReDim Preserve descriptorID_to_num_form_search(1, UBound(descriptorID_to_num_form_search, 2) + 1)
    need_to_remove = True
    'add to my_descriptors_ID
    add_to_my_descriptor_ID query_guid
    'make descriptor
    strdata = make_string_descriptor_data(query_guid, query, CByte(Me.txtttl.Text), 0, Len(strdata)) + strdata
    'send data
    send_to_all_dialing strdata

End Sub


Private Sub remove_from_descriptorID_to_num_form_search()
    On Error Resume Next
    'remove reference from descriptorID_to_num_form_search()
    Dim cpt             As Integer
    Dim cpt2            As Integer
    For cpt = 0 To UBound(descriptorID_to_num_form_search, 2) - 1
        If descriptorID_to_num_form_search(1, cpt) <> Me.Tag Then
            descriptorID_to_num_form_search(0, cpt2) = descriptorID_to_num_form_search(0, cpt)
            descriptorID_to_num_form_search(1, cpt2) = descriptorID_to_num_form_search(1, cpt)
            cpt2 = cpt2 + 1
        End If
    Next cpt
    ReDim Preserve descriptorID_to_num_form_search(1, UBound(descriptorID_to_num_form_search, 2) - 1)
    need_to_remove = False
End Sub

Private Sub Combo_search_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If
End Sub



Private Sub list_search_results_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And Me.my_popupmenu.Enabled Then
        Me.PopupMenu my_popupmenu
    End If
End Sub

Private Sub list_search_results_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    With list_search_results
        .Sorted = True
        .SortKey = ColumnHeader.index - 1
        
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub



'form events
Private Sub Form_Load()
    On Error Resume Next

    Dim size As Long
    Dim cpt  As Byte
    For cpt = 1 To Me.list_search_results.ColumnHeaders.Count
        size = size + Me.list_search_results.ColumnHeaders(cpt).Width
    Next cpt
    size = size + 400
    Me.Width = size
    Me.txtminspeed.Text = my_min_speed
    
    ' init info in case of refine search
    ReDim my_restrictions_array(4)
    my_restrictions_array(0) = 0
    my_restrictions_array(1) = 0
    my_restrictions_array(2) = 0
    my_restrictions_array(3) = 0
    my_restrictions_array(4) = 0
End Sub

Private Sub form_resize()
    On Error Resume Next
    Dim size           As Long
    Dim column_size(4) As Long
    Dim percent(4)     As Long
    Dim cpt            As Byte
    
    'move button search
    size = Me.Width - 400 - cmdSearch.Width
    If size > 3600 Then
        cmdSearch.Left = size
        'redim Combo_search
        Combo_search.Width = Me.Width - 800
        'redim list_search_results
        list_search_results.Width = Me.Width - 350
        'redim list column
        size = 0
        For cpt = 1 To 4
            column_size(cpt) = Me.list_search_results.ColumnHeaders(cpt).Width
            size = size + column_size(cpt)
        Next cpt
        If size = 0 Then Exit Sub
        For cpt = 1 To 4
            percent(cpt) = column_size(cpt) / size * 100
        Next cpt
        size = Int(Me.list_search_results.Width)
        For cpt = 1 To 4
            Me.list_search_results.ColumnHeaders(cpt).Width = Int((percent(cpt) * (size - 400)) / 100)
        Next cpt
        
    End If
    If Framemultiplehostdown.Visible = False Then
        size = Me.Height - 2100
    Else
        size = Me.Height - 2750
    End If
    If size > 0 Then list_search_results.Height = size

End Sub








Private Sub txtminspeed_Change()
    If Val(txtminspeed) < 1 Then txtminspeed = 1
End Sub

Private Sub txtttl_Change()
    If Val(txtttl) > 255 Then txtttl = 255
    If Val(txtttl) < 1 Then txtttl = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    remove_tab (Me.hWnd)

    If need_to_remove Then
        Call remove_from_descriptorID_to_num_form_search
    End If
    If Not quit_program Then
        'for window indexing
        Document_search_deleted(Me.Tag) = True
    End If
End Sub

Private Sub Form_Activate()
    activate_tab (Me.hWnd)
End Sub

' menu sub
Private Sub mnuinvertselection_Click()
    On Error Resume Next
    Dim cpt As Long
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            list_search_results.ListItems(cpt).Selected = False
        Else
            list_search_results.ListItems(cpt).Selected = True
        End If
    Next cpt
End Sub



Private Sub mnuselect_Click()
    On Error Resume Next
    Dim cpt As Long
    For cpt = 1 To list_search_results.ListItems.Count
        list_search_results.ListItems(cpt).Selected = True
    Next cpt
End Sub

Private Sub mnudeselect_Click()
    On Error Resume Next
    Dim cpt As Long
    For cpt = 1 To list_search_results.ListItems.Count
        list_search_results.ListItems(cpt).Selected = False
    Next cpt
End Sub

Private Sub mnudownload_Click()
    On Error Resume Next
    Dim cpt As Long
    Dim pos As Long
    
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            pos = list_search_results.ListItems(cpt).Tag
            'init download
            ask_for_download known_files(pos).ip, known_files(pos).port, _
                             known_files(pos).file_index, known_files(pos).file_name, _
                             0, 0, known_files(pos).servent_id, _
                             known_files(pos).file_size, known_files(pos).speed, 1, True, _
                             , , , known_files(pos).num_socket_queryhit, known_files(pos).need_push
        End If
    Next cpt
End Sub

Private Sub mnumultipartdownload_Click()
    On Error Resume Next
    Dim cpt             As Long
    Dim pos             As Long
    Dim cpt2            As Integer
    Dim part_size       As Single
    Dim saving_name     As String
    
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            pos = list_search_results.ListItems(cpt).Tag
            'init download
            part_size = Int(known_files(pos).file_size / nb_parts_for_download)
            'for the first part
            saving_name = ask_for_download(known_files(pos).ip, known_files(pos).port, _
                             known_files(pos).file_index, known_files(pos).file_name, _
                             0, part_size, known_files(pos).servent_id, _
                             known_files(pos).file_size, known_files(pos).speed, 1, True, nb_parts_for_download, saving_name, False, _
                             known_files(pos).num_socket_queryhit, known_files(pos).need_push)
            For cpt2 = 2 To nb_parts_for_download - 1
                ask_for_download known_files(pos).ip, known_files(pos).port, _
                                 known_files(pos).file_index, known_files(pos).file_name, _
                                 (cpt2 - 1) * part_size + 1, cpt2 * part_size, known_files(pos).servent_id, _
                                 known_files(pos).file_size, known_files(pos).speed, cpt2, False, nb_parts_for_download, saving_name, False, _
                                 known_files(pos).num_socket_queryhit, known_files(pos).need_push
            Next cpt2
            'for the last part
            ask_for_download known_files(pos).ip, known_files(pos).port, _
                             known_files(pos).file_index, known_files(pos).file_name, _
                             (nb_parts_for_download - 1) * part_size + 1, 0, known_files(pos).servent_id, _
                             known_files(pos).file_size, _
                             known_files(pos).speed, nb_parts_for_download, False, nb_parts_for_download, saving_name, False, _
                             known_files(pos).num_socket_queryhit, known_files(pos).need_push
        End If
    Next cpt
End Sub

Private Sub mnumultihostdownload_Click()
    On Error Resume Next
    'find first selected item only
    Dim pos_known_files     As Integer
    Dim cpt                 As Long
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            pos_known_files = list_search_results.ListItems(cpt).Tag
            new_host_search known_files(pos_known_files).file_name, known_files(pos_known_files).file_size, _
                            0, 0, True
            Exit For
        End If
    Next cpt
    
End Sub

Private Sub cmddownonthesehost_Click()
    On Error Resume Next
    Dim cpt                 As Long
    Dim cpt2                As Integer
    Dim pos                 As Long
    Dim nb_parts            As Integer
    Dim part_size           As Single
    Dim all_nbparts         As Integer
    Dim num_file            As Integer
    Dim rbegin              As Single
    Dim rend                As Single
    Dim type_len            As Byte
    Dim pos_modified        As Single
    Dim last_range_end      As Single
    Dim rpos                As Integer
    Dim field_begin         As Single
    Dim field_size          As Single
    Dim saving_name         As String
    
    Dim str_before_part     As String
    Dim str_after_part      As String
    Dim str_updated_part    As String
            
    Dim array_range()       As Single
    Dim all_hosts_info      As String
    type_len = 4
    
    'find the number of hosts selected to get the number of parts
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            nb_parts = nb_parts + 1
        End If
    Next cpt
    
    'get the size of a part
    If on_other_host_infos.range_end >= on_other_host_infos.range_begin And on_other_host_infos.range_end > 0 Then
        part_size = Int((on_other_host_infos.range_end - on_other_host_infos.range_begin + 1) / nb_parts)
    Else
        part_size = Int((on_other_host_infos.file_size - on_other_host_infos.range_begin) / nb_parts) 'file_size=range end +1
    End If
    If part_size = 0 Then 'downloaded part too small
        nb_parts = 1
        part_size = on_other_host_infos.file_size - on_other_host_infos.range_begin
    End If
    If part_size < 0 Then Exit Sub


    If on_other_host_infos.create_new_file Then
    '--> download on multiple hosts
        all_nbparts = nb_parts
    
    Else
    '--> resume on other hosts
        all_nbparts = nb_parts + on_other_host_infos.old_nb_parts - 1
        ReDim all_host(on_other_host_infos.old_nb_parts)
        
        'remove from form_download_upload.listview_download
        For cpt = 1 To Form_download_upload.ListView_download.ListItems.Count
            If Form_download_upload.ListView_download.ListItems(cpt).Tag = on_other_host_infos.pos_current_down Then
                Form_download_upload.ListView_download.ListItems.Remove (cpt)
                Exit For
            End If
        Next cpt
        'remove from name_download
        remove_from_name_download (current_download(on_other_host_infos.pos_current_down).pos_name_download)
        'remove from current download
        remove_from_current_download (on_other_host_infos.pos_current_down)
    
        'update the file structure by adding and moving fields if necessary
        num_file = FreeFile()
        Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary Access Read Shared As num_file
            'get number of record
            all_hosts_info = Space$(LOF(num_file) - (2 * on_other_host_infos.old_nb_parts + 2) * type_len)
            Get num_file, (2 * on_other_host_infos.old_nb_parts + 2) * type_len + 1, all_hosts_info 'binary
        Close num_file
    
        field_begin = 1
        ReDim all_host(on_other_host_infos.old_nb_parts - 1)
        For cpt = 0 To on_other_host_infos.old_nb_parts - 1
        'extract part file infos before and after the modified part
            If cpt + 1 = on_other_host_infos.num_part Then
                str_before_part = Mid$(all_hosts_info, 1, field_begin)
                pos_modified = field_begin
                field_size = long_to_little_endian(Mid$(all_hosts_info, field_begin, 4))
                field_begin = field_begin + field_size
                str_after_part = Mid$(all_hosts_info, field_begin)
                Exit For
            End If
            field_size = long_to_little_endian(Mid$(all_hosts_info, field_begin, 4))
            field_begin = field_begin + field_size
        Next cpt

        If nb_parts > 1 Then
        'we need to modify
        '   the number of parts of the .coy file
        '   the place of range_begin and range_end of field placed after the resuming one
        '   the parts info
        '   the num_part in name_download() for these fields
        
            ReDim array_range(2 * on_other_host_infos.old_nb_parts)
            num_file = FreeFile()
            Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Random As num_file Len = type_len
                'update number of parts of the .coy file
                Put num_file, 1, all_nbparts
                'put in memory range_begin and range_end for part after num_part
                For cpt = on_other_host_infos.num_part + 2 To 2 * on_other_host_infos.old_nb_parts Step 2
                    Get num_file, cpt, rbegin
                    Get num_file, cpt + 1, rend
                    array_range(cpt * 2) = rbegin
                    array_range(cpt * 2 + 1) = rend
                Next cpt
                    
                rpos = 0
                'writing new part range begin and range end
                For cpt = on_other_host_infos.num_part * 2 + 2 To (on_other_host_infos.num_part + nb_parts) * 2 + 2 - 2 Step 2
                    Select Case rpos
                        Case 0
                            Put num_file, cpt, on_other_host_infos.range_begin
                            Put num_file, cpt, on_other_host_infos.range_begin + (rpos + 1) * part_size
                        Case nb_parts - 1
                            Put num_file, cpt, on_other_host_infos.range_begin + rpos * part_size + 1
                            Put num_file, cpt, on_other_host_infos.range_begin + (rpos + 1) * part_size
                        Case Else
                            Put num_file, cpt, on_other_host_infos.range_begin + rpos * part_size + 1
                            Put num_file, cpt, on_other_host_infos.range_end
                    End Select
                    rpos = rpos + 1
                Next cpt
                'writing other part range begin and range end at the good place
                For cpt = (on_other_host_infos.num_part) * 2 + 2 + 1 To (on_other_host_infos.old_nb_parts) * 2 + 2
                    Put num_file, cpt, array_range(cpt)
                Next cpt
                'change num_part in name download
                For cpt = 0 To UBound(name_download) - 1
                    If name_download(cpt).saving_name = on_other_host_infos.saving_name _
                       And name_download(cpt).num_part > on_other_host_infos.num_part Then
                        name_download(cpt).num_part = name_download(cpt).num_part + nb_parts - 1
                    End If
                Next cpt
            Close num_file
        End If ' nb_part >1
        'write first part file infos
        num_file = FreeFile()
        Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary As num_file
            Put num_file, (2 * all_nbparts + 2) * type_len + 1, str_before_part 'binary
        Close num_file
    End If
    
    'launch the corresponding downloads
    cpt2 = 0
    saving_name = on_other_host_infos.saving_name
    For cpt = 1 To list_search_results.ListItems.Count
        If list_search_results.ListItems(cpt).Selected Then
            pos = list_search_results.ListItems(cpt).Tag
            cpt2 = cpt2 + 1
            Select Case cpt2
                Case 1
                'NOTE:in a select case only the first corresponding case is done
                    If nb_parts > 1 Then
                        'for the first part
                        saving_name = ask_for_download(known_files(pos).ip, known_files(pos).port, _
                                         known_files(pos).file_index, known_files(pos).file_name, _
                                         on_other_host_infos.range_begin, on_other_host_infos.range_begin + part_size, _
                                         known_files(pos).servent_id, _
                                         known_files(pos).file_size, known_files(pos).speed, on_other_host_infos.num_part, _
                                         on_other_host_infos.create_new_file, all_nbparts, on_other_host_infos.saving_name, _
                                         Not on_other_host_infos.create_new_file, known_files(pos).num_socket_queryhit, known_files(pos).need_push)
                        If Not on_other_host_infos.create_new_file Then 'resume on other hosts
                            'write part file info
                            str_updated_part = long_to_big_endian(CSng(Len(known_files(pos).file_name) + 34)) _
                                                & ip_encode(known_files(pos).ip) & int_to_big_endian(known_files(pos).port) _
                                                & long_to_big_endian(known_files(pos).file_index) _
                                                & known_files(pos).servent_id & long_to_big_endian(known_files(pos).speed) _
                                                & known_files(pos).file_name
                            num_file = FreeFile()
                            Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary As num_file
                                Put num_file, pos_modified, str_updated_part
                            Close num_file
                            
                            pos_modified = pos_modified + Len(str_updated_part)
                        End If
                    Else 'nb_parts=1
                        If on_other_host_infos.create_new_file Then
                            'for the last part of a multiplehost down
                            last_range_end = 0
                        Else 'last part of a resume on other host
                            last_range_end = on_other_host_infos.range_end
                        End If
                        ask_for_download known_files(pos).ip, known_files(pos).port, _
                                         known_files(pos).file_index, known_files(pos).file_name, _
                                         on_other_host_infos.range_begin, last_range_end, known_files(pos).servent_id, _
                                         known_files(pos).file_size, _
                                         known_files(pos).speed, on_other_host_infos.num_part + nb_parts - 1, _
                                         on_other_host_infos.create_new_file, all_nbparts, saving_name, _
                                         Not on_other_host_infos.create_new_file, known_files(pos).num_socket_queryhit, known_files(pos).need_push
                        If Not on_other_host_infos.create_new_file Then 'resume on other hosts
                            'write part file info & str_after_part
                            str_updated_part = long_to_big_endian(CSng(Len(known_files(pos).file_name) + 34)) _
                                                & ip_encode(known_files(pos).ip) & int_to_big_endian(known_files(pos).port) _
                                                & long_to_big_endian(known_files(pos).file_index) _
                                                & known_files(pos).servent_id & long_to_big_endian(known_files(pos).speed) _
                                                & known_files(pos).file_name
                            num_file = FreeFile()
                            Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary As num_file
                                Put num_file, pos_modified, str_updated_part & str_after_part
                            Close num_file
                        End If
                    End If
                Case nb_parts
                    If on_other_host_infos.create_new_file Then
                        'for the last part of a multiplehost down
                        last_range_end = 0
                    Else 'last part of a resume on other host
                        last_range_end = on_other_host_infos.range_end
                    End If
                    ask_for_download known_files(pos).ip, known_files(pos).port, _
                                     known_files(pos).file_index, known_files(pos).file_name, _
                                     on_other_host_infos.range_begin + (cpt2 - 1) * part_size + 1, _
                                     last_range_end, known_files(pos).servent_id, _
                                     known_files(pos).file_size, _
                                     known_files(pos).speed, on_other_host_infos.num_part + nb_parts - 1, False, _
                                     all_nbparts, saving_name, Not on_other_host_infos.create_new_file, _
                                     known_files(pos).num_socket_queryhit, known_files(pos).need_push
                        If Not on_other_host_infos.create_new_file Then 'resume on other hosts
                            'write part file info & str_after_part
                            str_updated_part = long_to_big_endian(CSng(Len(known_files(pos).file_name) + 34)) _
                                                & ip_encode(known_files(pos).ip) & int_to_big_endian(known_files(pos).port) _
                                                & long_to_big_endian(known_files(pos).file_index) _
                                                & known_files(pos).servent_id & long_to_big_endian(known_files(pos).speed) _
                                                & known_files(pos).file_name
                            num_file = FreeFile()
                            Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary As num_file
                                Put num_file, pos_modified, str_updated_part & str_after_part
                            Close num_file
                        End If
                Case Else
                    ask_for_download known_files(pos).ip, known_files(pos).port, _
                                     known_files(pos).file_index, known_files(pos).file_name, _
                                     on_other_host_infos.range_begin + (cpt2 - 1) * part_size + 1, _
                                     on_other_host_infos.range_begin + cpt2 * part_size, known_files(pos).servent_id, _
                                     known_files(pos).file_size, known_files(pos).speed, _
                                     on_other_host_infos.num_part + cpt2 - 1, False, all_nbparts, saving_name, _
                                     Not on_other_host_infos.create_new_file, _
                                     known_files(pos).num_socket_queryhit, known_files(pos).need_push
                                     
                If Not on_other_host_infos.create_new_file Then 'resume on other hosts
                    'write part file info
                    str_updated_part = long_to_big_endian(CSng(Len(known_files(pos).file_name) + 34)) _
                                        & ip_encode(known_files(pos).ip) & int_to_big_endian(known_files(pos).port) _
                                        & long_to_big_endian(known_files(pos).file_index) _
                                        & known_files(pos).servent_id & long_to_big_endian(known_files(pos).speed) _
                                        & known_files(pos).file_name
                    num_file = FreeFile()
                    Open my_incomplete_directory & on_other_host_infos.saving_name & ".coy" For Binary As num_file
                        Put num_file, pos_modified, str_updated_part
                    Close num_file
                    
                    pos_modified = pos_modified + Len(str_updated_part)
                End If
            End Select
        End If
    Next cpt

End Sub


Public Sub fill_info_for_other_host_down(ByVal file_name As String, ByVal file_size As Single, ByVal range_begin As Single, ByVal range_end As Single, _
                ByVal create_new_file As Boolean, Optional ByVal old_nb_parts As Integer = 1, Optional ByVal num_part As Integer = 1, Optional ByVal pos_current_down As Integer = -1, _
                Optional ByVal saving_name As String = "")
    On Error Resume Next
    'fill on_other_host_infos
    With on_other_host_infos
        .file_name = file_name
        .file_size = file_size
        .range_begin = range_begin
        .range_end = range_end
        .create_new_file = create_new_file
        .old_nb_parts = old_nb_parts
        .num_part = num_part
        .pos_current_down = pos_current_down
        .saving_name = saving_name
    End With
End Sub


Public Sub fill_restriction_array(ByVal min_file_size As String, ByVal max_file_size As String, ByVal all_and_words As String, ByVal all_or_words As String, ByVal all_not_words As String)
    On Error Resume Next
    Dim tmparray                As Variant
    Dim cpt                     As Integer
    Dim size                    As Integer
    Dim pos                     As Integer
    
    ReDim my_restrictions_array(2)
    my_restrictions_array(0) = CSng(Val(min_file_size) * 1000) 'in kb
    my_restrictions_array(1) = CSng(Val(max_file_size) * 1000) 'in kb

    pos = 2
    If all_and_words <> "" Then
        tmparray = Split(all_and_words, ";")
        size = UBound(tmparray)
        my_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos + 1)
        For cpt = 0 To size
            my_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve my_restrictions_array(pos)
        Next cpt
    Else
        my_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos)
    End If
    
    If all_or_words <> "" Then
        tmparray = Split(all_or_words, ";")
        size = UBound(tmparray)
        my_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos)
        For cpt = 0 To size
            my_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve my_restrictions_array(pos)
        Next cpt
    Else
        my_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos)
    End If
    
    If all_not_words <> "" Then
        tmparray = Split(all_not_words, ";")
        size = UBound(tmparray)
        my_restrictions_array(pos) = size + 1
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos)
        For cpt = 0 To size
            my_restrictions_array(pos) = tmparray(cpt)
            pos = pos + 1
            ReDim Preserve my_restrictions_array(pos)
        Next cpt
    Else
        my_restrictions_array(pos) = 0
        pos = pos + 1
        ReDim Preserve my_restrictions_array(pos)
    End If

    'make a search throught known_files with new filter
    search_through_known_files CInt(Me.Tag), Me.Combo_search.Text, CLng(Me.txtminspeed)

    'the new incoming answers will respect filters -> nothing to do here
End Sub


Public Function check_form_restrictions(ByVal file_name As String, ByVal file_size As Single) As Boolean
    On Error Resume Next
    Dim cpt         As Integer
    Dim word_ok     As Boolean
    Dim pos         As Integer
    Dim begin       As Integer
    
    'check file size
        If my_restrictions_array(0) > file_size Then Exit Function 'min
        If my_restrictions_array(1) < file_size And my_restrictions_array(1) > 0 Then Exit Function 'max
        begin = 2
    'check and
        For cpt = 1 To my_restrictions_array(begin)
            pos = InStr(1, file_name, my_restrictions_array(cpt + begin))
            If pos < 1 Then Exit Function
        Next cpt
        begin = begin + my_restrictions_array(begin) + 1
    'check or
        For cpt = 1 To my_restrictions_array(begin)
            pos = InStr(1, file_name, my_restrictions_array(cpt + begin))
            If pos > 0 Then word_ok = True
        Next cpt
        If my_restrictions_array(begin) > 0 And Not word_ok Then Exit Function
        begin = begin + my_restrictions_array(begin) + 1
    'check not
        For cpt = 1 To my_restrictions_array(begin)
            pos = InStr(1, file_name, my_restrictions_array(cpt + begin))
            If pos > 0 Then Exit Function
        Next cpt
        check_form_restrictions = True
End Function

Private Sub cmdrefinesearch_Click()
    Form_refine_search.Tag = Me.Tag
    Form_refine_search.Show vbModal
End Sub

Private Sub mnuWCascade_Click()
   Form_main.Arrange vbCascade
End Sub

Private Sub mnuWTileH_Click()
   Form_main.Arrange vbTileHorizontal
End Sub

Private Sub mnuWTileV_Click()
   Form_main.Arrange vbTileVertical
End Sub

Private Sub mnuWArrange_Click()
   Form_main.Arrange vbArrangeIcons
End Sub

