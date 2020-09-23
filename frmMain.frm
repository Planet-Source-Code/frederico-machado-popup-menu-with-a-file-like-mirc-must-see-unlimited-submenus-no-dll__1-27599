VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB Popups"
   ClientHeight    =   2385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Right click on the form to view the popup menu"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is an example how to make a popup menu using a file like mIRC
' REMEMBER: the modules and class modules of this example
' aren't mine, I've got them from VBAccelerator.com
' PLEASE VOTE FOR ME IF YOU LIKE THE CODE

'Main variables
Dim PopupFile As String
Dim Path As String

Private WithEvents cP As cPopupMenu
Attribute cP.VB_VarHelpID = -1

Private Sub Form_Load()
  Set cP = New cPopupMenu ' set cP variable to cPopupMenu
  cP.hWndOwner = Me.hwnd ' set the owner form
  cP.GradientHighlight = True ' if you don't like gradient highlight then set it to false
  
  Path = App.Path ' set the Path variable to the Application Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\" ' It happens when the Application is in a Folder
  PopupFile = Path & "popup.ini" ' set the PopupFile variable to the full path of Popup file
  ParsePopupFile PopupFile ' It parses the Popup file to add menus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button And 2 Then ' Mouse right click
    cP.Restore "MainPopup" ' Restore the popup to show it
    i = cP.ShowPopupMenu(X, Y) ' show the popup
    If i > 0 Then ParsePopupCommand cP.ItemKey(i) ' Parse the command
  End If
End Sub

Sub ParsePopupFile(PopupFile As String)
  Dim X, sCaption, sCommand
  Dim Header As Boolean
  Dim SubMenu(0 To 100) As Integer ' Number of submenus
  cP.Clear ' Clear the popup
  Open PopupFile For Input As #1 ' Open the popup file to get lines
    While Not EOF(1)
      X = 0 ' it is very important set X = 0
      Line Input #1, strLine ' it gets a line from popup file
      If left$(strLine, 1) <> ";" And strLine <> "" Then ' if the first chr = ; then it is a comment
        Do While left$(strLine, 1) = "." ' it gets the submenu number
          X = X + 1
          strLine = Mid$(strLine, 2)
        Loop
        If InStr(strLine, ":") = 0 Then Header = True Else Header = False ' Headers DON'T has command, it has submenus
        If Header = False Then  ' if it isn't a Header then ...
          sCaption = left$(strLine, InStr(strLine, ":") - 1) ' sets the caption
          sCommand = Mid$(strLine, InStr(strLine, ":") + 1) ' sets the command
        Else
          sCaption = strLine ' if it is a Header then it sets only the caption
          sCommand = "" ' REMEMBER: Headers DON'T has command, it has submenus
        End If
        If X = 0 Then ' X = 0 when it isn't a submenu
          SubMenu(0) = cP.AddItem(sCaption, , , , , , , sCommand) ' it creates a MENU item
        Else
          SubMenu(X) = cP.AddItem(sCaption, , , SubMenu(X - 1), , , , sCommand) ' it creates a  SUBMENU item
        End If
      End If
    Wend
  Close #1 ' close the popup file
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' If you want to use a background picture then set the propertie BackgroundPicture to you Picture '
  ' Set cP.BackgroundPicture = Picture1.Picture                                                     '
  ' That is the command, don't forget the "Set" before
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
  cP.Store "MainPopup" ' it "saves" the final popup
End Sub

' It parses the command
Sub ParsePopupCommand(PopCommand As String)
  Select Case LCase(PopCommand)
    Case "minimize"
      WindowState = 1
    Case "maximize"
      WindowState = 2
    Case "showabout"
      frmAbout.Show 1
    Case "exitapp"
      End
    Case "message"
      MsgBox "Test Message!"
  End Select
End Sub
