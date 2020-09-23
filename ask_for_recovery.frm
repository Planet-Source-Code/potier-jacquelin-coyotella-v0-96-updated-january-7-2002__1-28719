VERSION 5.00
Begin VB.Form form_ask_for_recovery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recovery"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "ask_for_recovery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton delete 
      Caption         =   "Delete it"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton recover_later 
      Caption         =   "Recover it later"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton recover_now 
      Caption         =   "Recover it now"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label label_percent_value 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label label_name_value 
      Caption         =   "Label4"
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Percent Completed :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "File name :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "An incomplete file was found in your intermediate directory. What do you want to do :"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "form_ask_for_recovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delete_Click()
    Dim file_name As String
    file_name = my_incomplete_directory & Me.label_name_value
    If is_file_existing(file_name) Then
        Kill file_name
    End If
    file_name = my_incomplete_directory & Me.label_name_value & ".coy"
    If is_file_existing(file_name) Then
        Kill file_name
    End If
    Unload Me
End Sub

Private Sub recover_later_Click()
    Unload Me
End Sub

Private Sub recover_now_Click()
    make_recovery Me.label_name_value
    Unload Me
End Sub
