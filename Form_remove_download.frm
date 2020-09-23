VERSION 5.00
Begin VB.Form Form_remove_download 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove download"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "Form_remove_download.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdallparts 
      Caption         =   "Stop all parts and remove files"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdonlythisdownload 
      Caption         =   "Remove only this part"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form_remove_download.frx":08CA
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"Form_remove_download.frx":0D0C
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form_remove_download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub cmdallparts_Click()
    On Error Resume Next
    'stop all downloads having the same saving name and delete saving_name and saving_name.coy
    Dim strtmp              As String
    Dim file_name           As String
    Dim cpt                 As Integer
    Dim cpt2                As Integer
    Dim pos_name_download   As Integer
    Dim nb_part_removed     As Integer
    Dim size                As Integer
    file_name = Me.Tag
    
    Form_main.Timer_second.Enabled = False
    
    nb_part_removed = 0
    'search throught current download if saving name is corresponding
    For cpt = UBound(current_download) - 1 To 0 Step -1
        If current_download(cpt).file_name = file_name Then
            pos_name_download = current_download(cpt).pos_name_download
            
            'remove from name_download and waiting download if necessary
            If pos_name_download > -1 Then ' else download is finished don't need to remove from waiting download
                'remove from name_download
                remove_from_name_download pos_name_download
                
                For cpt2 = UBound(waiting_download) To 0 Step -1
                    'remove from waiting download
                    If waiting_download(cpt2).position_name_download = pos_name_download Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        remove_waiting_download cpt2 - nb_part_removed 'because remove_waiting_download change index
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        nb_part_removed = nb_part_removed + 1
                    End If
                Next cpt2
                
                'remove from retry_download if necessary
                size = UBound(retry_download)
                For cpt2 = 0 To size - 1
                    If retry_download(cpt2).pos_name_download = pos_name_download Then
                        'remove
                        retry_download(cpt2) = retry_download(size - 1)
                        ReDim Preserve retry_download(size - 1)
                    End If
                Next cpt2
                
                'remove from waiting giv if necessary
                size = UBound(waiting_giv)
                For cpt2 = 0 To size - 1
                    If waiting_giv(cpt2).pos_name_download = pos_name_download Then
                        'remove
                        remove_from_waiting_giv cpt2
                    End If
                Next cpt2
            End If
            
            'remove from current_download
            remove_from_current_download cpt
            'remove from list view
            For cpt2 = Form_download_upload.ListView_download.ListItems.Count To 1 Step -1
                If Form_download_upload.ListView_download.ListItems(cpt2).Tag = cpt Then
                    Form_download_upload.ListView_download.ListItems.Remove cpt2
                    Exit For
                End If
            Next cpt2

        End If
    Next cpt
    
    strtmp = my_incomplete_directory & file_name
    If is_file_existing(strtmp) Then
        Kill strtmp
    End If
    strtmp = strtmp & ".coy"
    If is_file_existing(strtmp) Then
        Kill strtmp
    End If
    
    Form_main.Timer_second.Enabled = True
    
    Unload Me
End Sub

Private Sub cmdonlythisdownload_Click()
    'already down nothing to do here
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Remove download :" & Me.Tag
End Sub
