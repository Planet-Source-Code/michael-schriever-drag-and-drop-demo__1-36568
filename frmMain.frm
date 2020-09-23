VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Drag and Drop Demo"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9015
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ImageList iml 
      Left            =   8070
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3105
      Left            =   240
      TabIndex        =   1
      Top             =   2130
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5477
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      ImageList       =   "iml"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1635
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   2884
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EMail"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7585
      EndProperty
   End
   Begin VB.Label lblDaD 
      Height          =   225
      Left            =   8100
      TabIndex        =   2
      Top             =   2850
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*                  Drag and Drop Demo                      *
'*                         by                               *
'*             MSoftware, Michael Schriever                 *
'*            webmaster@michael-schriever.de                *
'*               www.michael-schriever.de                   *
'*                                                          *
'************************************************************

Option Explicit

Private Sub Form_Load()
    Call createTVW
    Call createLVW
    lvw.FullRowSelect = True
    lvw.SelectedItem.Selected = False
    lblDaD.Visible = False
End Sub

Private Sub createLVW()
    Dim li As ListItem
    
    Set li = lvw.ListItems.Add(, , "webmaster@michael-schriever.de")
    li.SubItems(1) = "Michael Schriever"
    li.Tag = li.SubItems(1) + "," + li.Text
    Set li = lvw.ListItems.Add(, , "eric.warden@hotmail.com")
    li.SubItems(1) = "Eric Warden"
    li.Tag = li.SubItems(1) + "," + li.Text
End Sub

Private Sub createTVW()
    Dim nodX As Node
    
    Set nodX = tvw.Nodes.Add(, , "root", "root", 2)
    Set nodX = tvw.Nodes.Add("root", tvwChild, "n1", "Node1", 1)
    Set nodX = tvw.Nodes.Add("root", tvwChild, "n2", "Node2", 1)
    Set nodX = tvw.Nodes.Add("root", tvwChild, "n3", "Node3", 1)
    
    nodX.EnsureVisible
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim li As ListItem
    
    If Button <> vbLeftButton Then Exit Sub
    
    Set li = lvw.HitTest(x, y)
    
    If li Is Nothing Then Exit Sub
    
    lblDaD.Left = lvw.Left + x
    lblDaD.Top = lvw.Top + y
    lblDaD.Width = 1000
    lblDaD.Tag = li.Tag
    lblDaD.Drag
    
    Set li = Nothing
End Sub

Private Sub tvw_DragDrop(Source As Control, x As Single, y As Single)
    Dim nodX As Node
    Dim s As String
    Dim sName As String
    Dim sEMail As String
    Dim pos As Long
    
    If Not (TypeOf Source Is Label) Then Exit Sub
    If Source.Name <> "lblDaD" Then Exit Sub
    
    Set nodX = tvw.HitTest(x, y)
    
    If nodX Is Nothing Then Exit Sub
    If nodX.Key = "root" Then
        Set nodX = Nothing
        Exit Sub
    End If
    s = Source.Tag
        
    pos = InStr(1, s, ",")
    sName = Left(s, pos - 1)
    sEMail = Mid(s, pos + 1)
    
    nodX.Text = "EmailAdress of " + sName + " is: " + sEMail
    Set nodX = Nothing
End Sub
