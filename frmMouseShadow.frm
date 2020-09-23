VERSION 5.00
Begin VB.Form frmMouseShadow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Shadow"
   ClientHeight    =   0
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "frmMouseShadow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   0
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTest 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmMouseShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''draw a frame if the mouse is clicked
'Call GetCursorPos(MousePos)
'Call DrawFrame(MousePos.x, MousePos.y, True, True)
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''redraw the mouse frame
'Call GetCursorPos(MousePos)
'Call DrawFrame(MousePos.x, MousePos.y, Button, True)
'End Sub
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'remove bitmaps from memory before exiting program
Call DeleteBitmap(CursorBmp.hDcMemory, CursorBmp.hDcBitmap, CursorBmp.hDcPointer)
Call DeleteBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer)
Call DrawFrame(0, 0, False, False, True)
End
End Sub

