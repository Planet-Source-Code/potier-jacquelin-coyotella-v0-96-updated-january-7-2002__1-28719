Attribute VB_Name = "harddrive_access"
Option Explicit
Private array_directories() As String

Public Sub share_files()
    On Error Resume Next
    'share files in shared files
    Dim cpt As Long
    If Not sharing_simulation Then
        my_nb_shared_files = 0
        my_nb_kilobytes_shared = 0
    End If
    ReDim my_shared_files(0)
    Dim lResult As Long
    For cpt = 0 To UBound(myshared_directories)
        If myshared_directories(cpt) <> "" Then
            If is_folder_existing(myshared_directories(cpt)) Then
                find_my_shared_files myshared_directories(cpt), include_sub_dir
            Else
                lResult = MessageBox(0, "The folder " & myshared_directories(cpt) & vbCrLf & "was not found, files in this directory won't be shared", Form_main.Caption, vbExclamation)
            End If
        End If
    Next cpt
    'share files in the download directory not recursively because incomplete directory could be inside
    find_my_shared_files my_download_directory, False
End Sub

Public Sub find_my_shared_files(directory As String, Optional search_in_subdirectories As Boolean = True)
    'search files from directories and add them to my_shared_files
    On Error Resume Next
    If search_in_subdirectories Then
        ReDim array_directories(1)
        array_directories(1) = directory
    
        Dim last_dir As String
        Do While UBound(array_directories) > 0
            last_dir = array_directories(UBound(array_directories))
            ReDim Preserve array_directories(UBound(array_directories) - 1)
            find_through_directory last_dir
        Loop
    Else
        find_through_directory directory
    End If
End Sub

Private Sub find_through_directory(ByVal directory As String)
    On Error Resume Next
    Dim myname          As String
    Dim tmpsize         As Long
    Dim table_size      As Variant
    Dim arr_dir_size    As Long

    myname = Dir(directory, vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
    
    Do While myname <> ""   ' Start the loop.
        ' Ignore the current directory and the encompassing directory.
        If myname <> "." And myname <> ".." Then
            ' Use bitwise comparison to make sure myname is a directory.
            If (GetAttr(directory & myname) And vbDirectory) = vbDirectory Then
                arr_dir_size = UBound(array_directories) + 1
                ReDim Preserve array_directories(arr_dir_size)
                array_directories(arr_dir_size) = directory & myname & "\"
            Else
              'If (GetAttr(directory & myname) And vbArchive) = vbArchive _
                 Or (GetAttr(directory & myname) And vbHidden) = vbHidden _
                 Or (GetAttr(directory & myname) And vbNormal) = vbNormal _
                 Or (GetAttr(directory & myname) And vbReadOnly) = vbReadOnly _
                 Or (GetAttr(directory & myname) And vbSystem) = vbSystem _
               Then
              'add file to my_shared_files
                tmpsize = FileLen(directory & myname)
                If Not sharing_simulation Then
                    my_nb_kilobytes_shared = my_nb_kilobytes_shared + tmpsize / 1000
                    my_nb_shared_files = my_nb_shared_files + 1
                End If
                table_size = UBound(my_shared_files)
                With my_shared_files(table_size)
                    .file_name = LCase$(myname)
                    .file_size = tmpsize
                    .file_index = table_size
                    .full_path = directory & myname
                End With
                ReDim Preserve my_shared_files(table_size + 1)
            End If
        End If
        myname = Dir ' Get next entry.
    Loop

End Sub


Public Function find_recoveries(directory As String) As Integer
    On Error Resume Next
    Dim myname          As String
    ReDim recovery_files(0)
    myname = Dir(directory & "*.coy")
    Do While myname <> ""
        recovery_files(find_recoveries) = myname
        find_recoveries = find_recoveries + 1
        ReDim Preserve recovery_files(find_recoveries)
        myname = Dir ' Get next entry
    Loop
End Function

''''''''''''''''''''''' file system object functions
Public Function is_file_existing(full_path As String) As Boolean
    Dim fso As New FileSystemObject
    is_file_existing = fso.FileExists(full_path)
End Function

Public Function is_folder_existing(full_path As String) As Boolean
    Dim fso As New FileSystemObject
    is_folder_existing = fso.FolderExists(full_path)
End Function

Public Function get_file_name(full_path As String) As String
    Dim fso As New FileSystemObject
    get_file_name = fso.GetFileName(full_path)
End Function

Public Function get_extension(full_path As String) As String
    Dim fso As New FileSystemObject
    get_extension = fso.GetExtensionName(full_path)
End Function
