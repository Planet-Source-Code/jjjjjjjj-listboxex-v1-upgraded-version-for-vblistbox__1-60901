VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "ListBoxEx Test Form"
   ClientHeight    =   6465
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbAlline 
      Height          =   360
      ItemData        =   "frmTest.frx":0000
      Left            =   2760
      List            =   "frmTest.frx":000D
      TabIndex        =   16
      Text            =   "vbLeftJustify"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox chkIcon 
      Caption         =   "Icon Focus"
      Height          =   240
      Left            =   2760
      TabIndex        =   15
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtRemove 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   4320
      TabIndex        =   14
      Text            =   "-1"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Set Selected"
      Height          =   360
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtSelect 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   4320
      TabIndex        =   12
      Text            =   "30"
      Top             =   3480
      Width           =   375
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Items"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtIndex 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   4320
      TabIndex        =   10
      Text            =   "-1"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CheckBox chkAscending 
      Caption         =   "Sort Ascending"
      Height          =   240
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin Project1.ListBoxEX ListBoxEX1 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2535
      _extentx        =   4471
      _extenty        =   7011
      picture         =   "frmTest.frx":003A
      listicon        =   "frmTest.frx":10F4
      font            =   "frmTest.frx":1490
   End
   Begin VB.CheckBox chkBorder 
      Caption         =   "Border"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkAppear 
      Caption         =   "3D - Appearence"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkStrech 
      Caption         =   "Strech Icon"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkFull 
      Caption         =   "Full Row Select"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to"
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtAdd 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Test Add"
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Label lbCount 
      AutoSize        =   -1  'True
      Caption         =   "ListCount"
      Height          =   240
      Left            =   240
      TabIndex        =   19
      Top             =   5400
      Width           =   780
   End
   Begin VB.Label lbSelItem 
      AutoSize        =   -1  'True
      Caption         =   "Selected Item"
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label lbSeltext 
      AutoSize        =   -1  'True
      Caption         =   "Sel Text"
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   5760
      Width           =   705
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAppear_Click()
    ListBoxEX1.Appearance = chkAppear
End Sub

Private Sub chkAscending_Click()
    If chkAscending = 1 Then
        ListBoxEX1.SortOrder = Ascending
    Else
        ListBoxEX1.SortOrder = Desending
    End If
End Sub

Private Sub chkBorder_Click()
    ListBoxEX1.BorderStyle = chkBorder
End Sub

Private Sub chkFull_Click()
    ListBoxEX1.FullRowSelect = chkFull
End Sub

Private Sub chkIcon_Click()
    ListBoxEX1.IconFocus = chkIcon
End Sub

Private Sub chkSort_Click()
    ListBoxEX1.SortItems = chkSort
End Sub

Private Sub chkStrech_Click()
    ListBoxEX1.StrechIcon = chkStrech
End Sub

Private Sub cmbAlline_Click()
    ListBoxEX1.TextAlignment = cmbAlline.ListIndex
End Sub

Private Sub cmdAdd_Click()
    ListBoxEX1.AddItem txtAdd, txtIndex
End Sub

Private Sub cmdBackPic_Click()
    Set ListBoxEX1.Picture = imgBack.Picture
End Sub

Private Sub cmdClear_Click()
    ListBoxEX1.Clear
End Sub

Private Sub cmdIcon_Click()
    Set ListBoxEX1.ListIcon = imgIcon.Picture
End Sub

Private Sub cmdRemove_Click()
    ListBoxEX1.Remove txtRemove
End Sub

Private Sub cmdSelect_Click()
    ListBoxEX1.SelectedItem = txtSelect
End Sub

Private Sub Form_Load()
Dim X As Long
    For X = 1 To 50
        ListBoxEX1.AddItem "Item Text " & X
    Next X
End Sub

Private Sub ListBoxEX1_SelChange()
    lbCount = "List Count = " & ListBoxEX1.ListCount
    lbSelItem = "Sel Item = " & ListBoxEX1.SelectedItem
    lbSeltext = "Sel Text = " & ListBoxEX1.SelectedText
End Sub

