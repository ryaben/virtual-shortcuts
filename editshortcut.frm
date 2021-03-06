VERSION 5.00
Object = "{1292FDC1-6231-407E-A10D-F419BBFDA432}#3.0#0"; "ButtonXp.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form editshortcut 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit shortcut"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "editshortcut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin ButtonXP.XPButton XPButton3 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "Set route"
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
   Begin ButtonXP.XPButton XPButton2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "Set icon"
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
   Begin ButtonXP.XPButton XPButton1 
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Edit"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You can modify any value of the selected icon."
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "editshortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const APPLICATION As String = "Data"

'Funci?n api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
    Private Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

Private Sub Form_Load()
On Error Resume Next
Dim retrieve_icon As String
Dim retrieve_path As String
Dim retrieve_title As String

retrieve_icon = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & principal.ListView1.SelectedItem.Index, 0)
retrieve_path = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Path" & principal.ListView1.SelectedItem.Index, 0)
Image1.Picture = LoadPicture(retrieve_icon)
Text2.Text = retrieve_path
Text1.Text = principal.ListView1.SelectedItem.Text
End Sub

Private Sub XPButton1_Click()
Dim subelemento As ListItem

If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "You need to fill all the fields (included icon)", vbExclamation
Else

principal.ListView1.SelectedItem.Text = Text1.Text
principal.ListView1.SelectedItem.ListSubItems(1).Text = Text2.Text
principal.ListView1.SelectedItem.ListSubItems(2).Text = Text3.Text
principal.ListView1.SelectedItem.Icon = LoadPicture(Text3.Text)

Unload Me
End If
End Sub

Private Sub XPButton2_Click()
On Error GoTo errsub
With CommonDialog1
.FileName = ""
.DialogTitle = "Choose an image icon"
.InitDir = App.Path
.Filter = "Icon files (.gif, .ico)|*.gif;*.ico|"
.ShowOpen
If .FileName <> "" Then
Image1.Picture = LoadPicture(.FileName)
Text3.Text = .FileName
End If
End With
Exit Sub
errsub:
MsgBox "Selected image is invalid. Please choose another one", vbExclamation
End Sub

Private Sub XPButton3_Click()
With CommonDialog1
.FileName = ""
.DialogTitle = "Choose an application or file"
.InitDir = App.Path
.Filter = "All files|*.*|"
.ShowOpen
If .FileName <> "" Then
Text2.Text = .FileName
End If
End With
End Sub
