VERSION 5.00
Object = "{1292FDC1-6231-407E-A10D-F419BBFDA432}#3.0#0"; "ButtonXp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form configuration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "configuration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin ButtonXP.XPButton XPButton5 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Clear background"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ButtonXP.XPButton XPButton4 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Delete from registry"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ButtonXP.XPButton XPButton3 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Caption         =   "Write on registry"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5040
      Top             =   1920
   End
   Begin ButtonXP.XPButton XPButton2 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "configuration.frx":48FA
      Left            =   4440
      List            =   "configuration.frx":4913
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2160
      Top             =   2640
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "configuration.frx":494B
      Left            =   1440
      List            =   "configuration.frx":4961
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin ButtonXP.XPButton XPButton1 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Change background"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcuts font color"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   2760
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   2760
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Windows startup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Background style"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next

Image1.Picture = LoadPicture(App.Path & "\background.jpg")

If principal.ListView1.PictureAlignment = lvwBottomLeft Then
Combo1.Text = "Bottom Left"
End If
If principal.ListView1.PictureAlignment = lvwBottomRight Then
Combo1.Text = "Bottom Right"
End If
If principal.ListView1.PictureAlignment = lvwCenter Then
Combo1.Text = "Center"
End If
If principal.ListView1.PictureAlignment = lvwTile Then
Combo1.Text = "Tile"
End If
If principal.ListView1.PictureAlignment = lvwTopLeft Then
Combo1.Text = "Top Left"
End If
If principal.ListView1.PictureAlignment = lvwTopRight Then
Combo1.Text = "Top Right"
End If

If principal.ListView1.ForeColor = &H80000008 Then
Combo2.Text = "Black"
End If
If principal.ListView1.ForeColor = &HFFFFFF Then
Combo2.Text = "White"
End If
If principal.ListView1.ForeColor = &HFFFF00 Then
Combo2.Text = "Light blue"
End If
If principal.ListView1.ForeColor = &HFF0000 Then
Combo2.Text = "Blue"
End If
If principal.ListView1.ForeColor = &HC000& Then
Combo2.Text = "Green"
End If
If principal.ListView1.ForeColor = &HFFFF& Then
Combo2.Text = "Yellow"
End If
If principal.ListView1.ForeColor = &HFF& Then
Combo2.Text = "Red"
End If

End Sub

Private Sub Timer1_Timer()
If Combo1.Text = "Bottom Left" Then
principal.ListView1.PictureAlignment = lvwBottomLeft
End If
If Combo1.Text = "Bottom Right" Then
principal.ListView1.PictureAlignment = lvwBottomRight
End If
If Combo1.Text = "Center" Then
principal.ListView1.PictureAlignment = lvwCenter
End If
If Combo1.Text = "Tile" Then
principal.ListView1.PictureAlignment = lvwTile
End If
If Combo1.Text = "Top Left" Then
principal.ListView1.PictureAlignment = lvwTopLeft
End If
If Combo1.Text = "Top Right" Then
principal.ListView1.PictureAlignment = lvwTopRight
End If
End Sub

Private Sub Timer2_Timer()
If Combo2.Text = "Black" Then
principal.ListView1.ForeColor = &H80000008
End If
If Combo2.Text = "White" Then
principal.ListView1.ForeColor = &HFFFFFF
End If
If Combo2.Text = "Light blue" Then
principal.ListView1.ForeColor = &HFFFF00
End If
If Combo2.Text = "Blue" Then
principal.ListView1.ForeColor = &HFF0000
End If
If Combo2.Text = "Green" Then
principal.ListView1.ForeColor = &HC000&
End If
If Combo2.Text = "Yellow" Then
principal.ListView1.ForeColor = &HFFFF&
End If
If Combo2.Text = "Red" Then
principal.ListView1.ForeColor = &HFF&
End If
End Sub

Private Sub XPButton1_Click()
On Error GoTo errsub
With CommonDialog1
.FileName = ""
.DialogTitle = "Change background"
.Filter = "Image files (.jpg, .bmp, .png, .gif)|*.jpg;*.bmp;*.png;*.gif"
.ShowOpen
If .FileTitle <> "" Then
Image1.Picture = LoadPicture(.FileName)
SavePicture Image1.Picture, App.Path & "\background.jpg"
principal.ListView1.Picture = LoadPicture(App.Path & "\background.jpg")
End If
End With
Exit Sub
errsub:
MsgBox "Selected image is invalid. Please choose another one", vbExclamation
End Sub

Private Sub XPButton2_Click()
Unload Me
End Sub

Private Sub XPButton3_Click()
On Error Resume Next

Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")

El_Objeto.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Virtual Shortcuts", App.Path & "\" & App.EXEName & ".exe"

MsgBox "Windows start entry has been written", vbInformation
End Sub

Private Sub XPButton4_Click()
On Error Resume Next

Dim El_Objeto As Object
Set El_Objeto = CreateObject("WScript.Shell")

El_Objeto.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Virtual Shortcuts"

MsgBox "Windows start entry has been deleted", vbInformation
End Sub

Private Sub XPButton5_Click()
On Error Resume Next
principal.ListView1.Picture = LoadPicture("")
Image1.Picture = LoadPicture("")
Kill App.Path & "\background.jpg"
End Sub
