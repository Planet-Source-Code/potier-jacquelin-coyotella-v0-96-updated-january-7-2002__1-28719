VERSION 5.00
Begin VB.Form Form_refine_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refine Search"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "Form_refine_search.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtmaxfilesize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtminfilesize 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtnot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtand 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Max file size in kb ( 0 if no limit ) :"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Min file size in kb :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "All of the following words musn't  be present"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "At least one of the followning words must be present :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "All of the following words must be presents ( separator is ; ) :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form_refine_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Me.Hide
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    'clear listview
    Document_search(Me.Tag).list_search_results.ListItems.Clear
    Document_search(Me.Tag).Labelnb_res = 0
    Document_search(Me.Tag).fill_restriction_array Me.txtminfilesize.Text, _
                                 Me.txtmaxfilesize.Text, Me.txtand.Text, _
                                 Me.txtor.Text, Me.txtnot.Text
    Me.Hide
End Sub
