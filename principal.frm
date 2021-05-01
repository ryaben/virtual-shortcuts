VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C3967F87-FD47-4E87-B007-06264CBD1A36}#2.0#0"; "systray.ocx"
Begin VB.Form principal 
   Caption         =   "Virtual Shortcuts"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9015
   Icon            =   "principal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin IconSystray.sysTray sysTray1 
      Left            =   6120
      Top             =   600
      _ExtentX        =   1376
      _ExtentY        =   1376
      ToolTipText     =   "Virtual Shortcuts"
      IconPicture     =   "principal.frx":48FA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11668
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Items"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Route"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Image_Route"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu save 
         Caption         =   "&Save shortcuts list"
         Shortcut        =   ^S
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close Virtual Shortcuts"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu add 
         Caption         =   "&Add shortcut"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete selected shortcut"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu clear 
         Caption         =   "&Clear shortcuts list"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu configuration_menu 
      Caption         =   "&Tools"
      Begin VB.Menu optionsbutton 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu twitterpage 
         Caption         =   "Rama Studios Website"
         Shortcut        =   ^W
      End
      Begin VB.Menu bar82 
         Caption         =   "-"
      End
      Begin VB.Menu aboutbutton 
         Caption         =   "&About ..."
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu elementmenu 
      Caption         =   "elementmenu"
      Begin VB.Menu openshortcut 
         Caption         =   "Open shortcut"
      End
      Begin VB.Menu deleteshortcut 
         Caption         =   "Delete shortcut"
      End
   End
   Begin VB.Menu iconmenu 
      Caption         =   "iconmenu"
      Visible         =   0   'False
      Begin VB.Menu restore 
         Caption         =   "Restore Virtual Shortcuts"
      End
      Begin VB.Menu bar78 
         Caption         =   "-"
      End
      Begin VB.Menu invisible 
         Caption         =   "invisible"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu bar54 
         Caption         =   "-"
      End
      Begin VB.Menu closevs 
         Caption         =   "Close Virtual Shortcuts"
      End
   End
End
Attribute VB_Name = "principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1

Private sngListViewX As Single
Private sngListViewY As Single

Const APPLICATION As String = "Data"

Dim form_windowstate As String
Dim form_backgroundstyle As Single
Dim form_forecolor As String

Dim shortcut_count As Single
Dim shortcut_start As Single
Dim shortcut_total As Single
Dim shortcut_icon As Single
Dim shortcut_index As Single
Dim shortcut_name As String
Dim shortcut_path As String
Dim icon_path As String

'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    
Sub Agregar(TextoDeMenu As String, QueMenu As Object)

Dim indice As Integer

indice = QueMenu.Count

Load QueMenu(indice)

QueMenu(indice).Caption = TextoDeMenu
QueMenu(indice).Visible = True

End Sub

Sub Eliminar(TextoDeMenu As String, QueMenu As Object)

Dim cMenu As Object

For Each cMenu In QueMenu

    If cMenu.Caption = TextoDeMenu Then
       Unload cMenu
    End If

Next

End Sub

Private Sub GuardarArchivo()

Open App.Path & "\shortcuts_table.shrt" For Output As #1

Print #1, ""
Close #1

End Sub

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

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Private Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String

    WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI

End Function

Public Sub EjecutarArchivos(ruta As String)
Dim ejecutarShell As Variant
On Error GoTo errsub
ejecutarShell = Shell("rundll32.exe url.dll,FileProtocolHandler " & (ruta), 1)
Exit Sub
errsub: MsgBox Err.Description, vbCritical
End Sub

Private Sub aboutbutton_Click()
about.Show 1
End Sub

Private Sub add_Click()
addshortcut.Show 1
End Sub

Private Sub clear_Click()
If MsgBox("Are you really sure you want to clear the shortcuts list?", vbExclamation + vbYesNo, "Clear shortcuts list") = vbYes Then
ListView1.ListItems.clear
Call GuardarArchivo
End If
End Sub

Private Sub close_Click()
sysTray1.RemoverSystray
Unload Me
End Sub

Private Sub closevs_Click()
sysTray1.RemoverSystray
Unload Me
End Sub

Private Sub delete_Click()
If ListView1.ListItems.Count > 0 Then
If MsgBox("Are you sure you want to delete selected shortcut?", vbExclamation + vbYesNo, "Delete selected") = vbYes Then
Call GuardarArchivo
ListView1.ListItems.Remove ListView1.SelectedItem.Index
shortcut_count = ListView1.ListItems.Count
For shortcut_start = 1 To shortcut_count
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon" & shortcut_start, ListView1.ListItems(shortcut_start).Icon)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Index" & shortcut_start, ListView1.ListItems(shortcut_start).Index)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, ListView1.ListItems(shortcut_start).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(1).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(2).Text)
Next shortcut_start
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Total", ListView1.ListItems.Count)
End If
Else
MsgBox "There are no shortcuts in the desktop.", vbExclamation
End If
End Sub

Private Sub deleteshortcut_Click()
If MsgBox("Are you sure you want to delete selected shortcut?", vbExclamation + vbYesNo, "Delete selected") = vbYes Then
ListView1.ListItems.Remove ListView1.SelectedItem.Index
Call GuardarArchivo
shortcut_count = ListView1.ListItems.Count
For shortcut_start = 1 To shortcut_count
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon" & shortcut_start, ListView1.ListItems(shortcut_start).Icon)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Index" & shortcut_start, ListView1.ListItems(shortcut_start).Index)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, ListView1.ListItems(shortcut_start).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(1).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(2).Text)
Next shortcut_start
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Total", ListView1.ListItems.Count)
ListView1.ListItems.Remove ListView1.ListItems.Count
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim fso As Object
Dim subelemento As ListItem

If App.PrevInstance = True Then
MsgBox "Virtual Shortcuts is already open", vbExclamation
End
End If

form_windowstate = Leer_Ini(App.Path & "\config.ini", "Windowstate", 0)
form_backgroundstyle = Leer_Ini(App.Path & "\config.ini", "Backgroundstyle", 4)
form_forecolor = Leer_Ini(App.Path & "\config.ini", "Forecolor", &H80000008)

Me.WindowState = form_windowstate
ListView1.PictureAlignment = form_backgroundstyle
ListView1.ForeColor = form_forecolor

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(App.Path & "\background.jpg") Then
ListView1.Picture = LoadPicture(App.Path & "\background.jpg")
End If

If fso.FileExists(App.Path & "\shortcuts_table.shrt") Then
shortcut_total = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Total", 0)
For shortcut_start = 1 To shortcut_total
icon_path = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & shortcut_start, "")
shortcut_icon = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Icon" & shortcut_start, 0)
Image2.Picture = LoadPicture(icon_path)
ImageList1.ListImages.add shortcut_icon, , Image2.Picture
shortcut_index = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Index" & shortcut_start, 0)
shortcut_name = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, 0)
shortcut_path = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Path" & shortcut_start, 0)
Set subelemento = ListView1.ListItems.add(shortcut_index, , shortcut_name, shortcut_icon)
subelemento.SubItems(1) = shortcut_path
subelemento.SubItems(2) = icon_path
Next shortcut_start

For shortcut_start = 1 To shortcut_total
shortcut_name = Leer_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, 0)
Call Agregar(shortcut_name, invisible)
Next shortcut_start
If ListView1.ListItems.Count <> 0 Then
End If

End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
sysTray1.PonerSystray
Me.Visible = False
End If

ListView1.Width = Me.Width
ListView1.Height = Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call Grabar_Ini(App.Path & "\config.ini", "Windowstate", Me.WindowState)
Call Grabar_Ini(App.Path & "\config.ini", "Backgroundstyle", ListView1.PictureAlignment)
Call Grabar_Ini(App.Path & "\config.ini", "Forecolor", ListView1.ForeColor)

shortcut_count = ListView1.ListItems.Count
For shortcut_start = 1 To shortcut_count
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon" & shortcut_start, ListView1.ListItems(shortcut_start).Icon)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Index" & shortcut_start, ListView1.ListItems(shortcut_start).Index)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, ListView1.ListItems(shortcut_start).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(1).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(2).Text)
Next shortcut_start
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Total", ListView1.ListItems.Count)

sysTray1.RemoverSystray

End Sub

Private Sub invisible_Click(Index As Integer)
On Error GoTo errsub
Call EjecutarArchivos(ListView1.ListItems(Index).ListSubItems(1).Text)
Exit Sub
errsub:
MsgBox "Selected item no longer exists", vbCritical
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Item As ListItem

' verifica que se presionó el botón derecho
If Button = vbRightButton Then
For Each Item In principal.ListView1.ListItems
Item.Selected = False
Next Item
' HitTest devuelve la referencia al item, a partir de las coordenadas del mouse
Set Item = ListView1.HitTest(x, y)
' chequea que haya un item seleccionado
If Not Item Is Nothing Then
' Selecciona el elemento
Set ListView1.SelectedItem = Item
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Acá colocar el código para desplegar el menú popup.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ' texto del elemento seleccionado
            
            ' despliega el menú
            PopupMenu elementmenu
End If
End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As _
    Integer, x As Single, y As Single)
    sngListViewX = x
    sngListViewY = y
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
Dim lListItem As ListItem
Set lListItem = ListView1.HitTest(sngListViewX, _
        sngListViewY)
    If (lListItem Is Nothing) Then
    Else
Call EjecutarArchivos(ListView1.SelectedItem.ListSubItems(1).Text)
End If
    Set lListItem = Nothing
Exit Sub
End Sub

Private Sub openshortcut_Click()
On Error Resume Next
Dim lListItem As ListItem
Set lListItem = ListView1.HitTest(sngListViewX, _
        sngListViewY)
    If (lListItem Is Nothing) Then
    Else
Call EjecutarArchivos(ListView1.SelectedItem.ListSubItems(1).Text)
End If
    Set lListItem = Nothing
Exit Sub
End Sub

Private Sub optionsbutton_Click()
configuration.Show 1
End Sub

Private Sub restore_Click()
Me.WindowState = vbMaximized
Me.Visible = True
sysTray1.RemoverSystray
End Sub

Private Sub save_Click()
shortcut_count = ListView1.ListItems.Count
For shortcut_start = 1 To shortcut_count
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon" & shortcut_start, ListView1.ListItems(shortcut_start).Icon)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Index" & shortcut_start, ListView1.ListItems(shortcut_start).Index)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Name" & shortcut_start, ListView1.ListItems(shortcut_start).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(1).Text)
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Icon_Path" & shortcut_start, ListView1.ListItems(shortcut_start).ListSubItems(2).Text)
Next shortcut_start
Call Grabar_Ini(App.Path & "\shortcuts_table.shrt", "Total", ListView1.ListItems.Count)
MsgBox "Shortcuts list has been saved", vbInformation
End Sub

Private Sub sysTray1_DblClick(Button As Integer)
Me.WindowState = vbMaximized
Me.Visible = True
sysTray1.RemoverSystray
End Sub

Private Sub sysTray1_MouseUP(Button As Integer)
If Button = vbRightButton Then
PopupMenu iconmenu
End If
End Sub

Private Sub twitterpage_Click()
Dim Z
Z = ShellExecute(Me.hwnd, "Open", "http://adf.ly/Kk0PI", &O0, &O0, SW_NORMAL)
End Sub
