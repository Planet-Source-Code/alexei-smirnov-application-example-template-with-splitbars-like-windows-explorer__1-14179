VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Example"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   2280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3015
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   450
      Width           =   50
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6375
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11245
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"Main.frx":0354
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   50
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   15
      ScaleWidth      =   2145
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0440
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7740
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15266
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5318
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "file"
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   5760
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   5760
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   5760
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mFOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mFS1 
         Caption         =   "-"
      End
      Begin VB.Menu mFExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mHContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mHS1 
         Caption         =   "-"
      End
      Begin VB.Menu mHAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ToolBar_Resize()

Dim tx&, y12&, y22&, fw&

On Error Resume Next

tx = 40
y12 = 10
y22 = y12 + tx + Toolbar1.Height + 10

fw = Me.ScaleWidth

With Toolbar1
    .Top = 40
    .Left = 30
End With
With Line1(0)
    .X1 = 0
    .Y1 = 0
    .X2 = fw
    .Y2 = 0
End With
With Line1(1)
    .X1 = 0
    .Y1 = y12
    .X2 = fw
    .Y2 = y12
End With
With Line1(2)
    .X1 = 0
    .Y1 = y22
    .X2 = fw
    .Y2 = y22
End With

End Sub

Private Sub Form_Load()

On Error Resume Next

Picture1.BorderStyle = vbBSNone
Picture2.BorderStyle = vbBSNone

TreeView1.Width = 2500

Load_Tree

End Sub

Private Sub Load_Tree()

Dim nodX As Node

With TreeView1
   Set nodX = .Nodes.Add(, , "R", "Root", 1)
   Set nodX = .Nodes.Add("R", tvwChild, "C1", "Child 1", 1)
   Set nodX = .Nodes.Add("R", tvwChild, "C2", "Child 2", 1)
   Set nodX = .Nodes.Add("R", tvwChild, "C3", "Child 3", 1)
   Set nodX = .Nodes.Add("R", tvwChild, "C4", "Child 4", 1)
   nodX.EnsureVisible
End With

End Sub

Private Sub RightSide_Resize()

On Error Resume Next

With RichTextBox1
    .Top = TreeView1.Top + 400
    .Left = TreeView1.Width + 100
    .Width = Me.ScaleWidth - .Left - 10
    .Height = Me.ScaleHeight - StatusBar1.Height - .Top + 12
End With

With Picture2
    .Height = Me.ScaleHeight - StatusBar1.Height - .Top + 12
    .Left = TreeView1.Left + TreeView1.Width
End With

End Sub

Private Sub Form_Resize()

ToolBar_Resize
LeftSide_Resize
RightSide_Resize

End Sub

Private Sub LeftSide_Resize()

Dim y32&

On Error Resume Next

y32 = 60 + Toolbar1.Height

With TreeView1
    .Top = y32 + 40
    .Left = 30
End With

With Picture1
    .Top = TreeView1.Top + TreeView1.Height
    .Width = TreeView1.Width
End With

With ListView1
    .Width = TreeView1.Width + 10
    .Left = 25
    .Top = Picture1.Top + Picture1.Height
    .Height = Me.ScaleHeight - StatusBar1.Height - .Top
End With

End Sub

Private Sub mFExit_Click()

Unload Me
End

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Picture1.BackColor = &H0&

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
If Button = 1 Then Picture1.Top = Picture1.Top + Y

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim pt&

On Error Resume Next

pt = Picture1.Top

If pt < 1000 Then
    pt = 1000
    Picture1.Top = pt
End If
If pt > (Me.ScaleHeight - StatusBar1.Height - 1000) Then
    pt = Me.ScaleHeight - StatusBar1.Height - 1000
    Picture1.Top = pt
End If

TreeView1.Height = pt - TreeView1.Top - 5

With ListView1
    .Top = pt + Picture1.Height
    .Height = Me.ScaleHeight - StatusBar1.Height - .Top
End With

Picture1.BackColor = &H8000000F

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Picture2.BackColor = &H0&

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
If Button = 1 Then Picture2.Left = Picture2.Left + X

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim pl&

On Error Resume Next

pl = Picture2.Left

If pl < 1000 Then
    pl = 1000
    Picture2.Left = pl
End If
If pl > (Me.ScaleWidth - 1000) Then
    pl = Me.ScaleWidth - 1000
    Picture2.Left = pl
End If

TreeView1.Width = pl - TreeView1.Left - 5
ListView1.Width = pl - ListView1.Left - 5
Picture1.Width = TreeView1.Width

RichTextBox1.Left = TreeView1.Width + 100
RichTextBox1.Width = Me.ScaleWidth - RichTextBox1.Left - 10

Picture2.BackColor = &H8000000F

End Sub
