VERSION 5.00
Begin VB.Form Form_rename 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Form_rename.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Over Write"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rename"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3135
      End
   End
   Begin VB.Label Lbltxt 
      Alignment       =   2  'Center
      Caption         =   "Warning a file with the same name is already existing. What do you want to do ?"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form_rename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Option2.value = True Then 'overwrite
        renamed_file.overwrite = True
    Else 'rename
        renamed_file.new_name = Text1.Text
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = renamed_file.old_name
End Sub

