Attribute VB_Name = "modMouseShadow"
Const ShadowHeightX = 6
Const ShadowHeightY = 4
Const PressHeightX = 2
Const PressHeightY = 2

Public CursorBmp As BitmapStruc
Public BackBmp As BitmapStruc
Dim MousePos As PointAPI

Public Sub Main()
Dim hCursor As Long
Dim Result As Long

CursorBmp.Area.Bottom = 32
CursorBmp.Area.Right = 32
BackBmp.Area = CursorBmp.Area

'create the cursor and the background bitmaps
Call CreateNewBitmap(CursorBmp.hDcMemory, CursorBmp.hDcBitmap, CursorBmp.hDcPointer, CursorBmp.Area, frmMouseShadow.hdc, vbWhite, InPixels)
Call CreateNewBitmap(BackBmp.hDcMemory, BackBmp.hDcBitmap, BackBmp.hDcPointer, BackBmp.Area, frmMouseShadow.hdc, vbWhite, InPixels)

'load the form so I can tell when to exit the program
Load frmMouseShadow
frmMouseShadow.Show

While True
    Call SetAndDraw
    DoEvents
Wend
End Sub

Public Sub SetAndDraw()
'find the mouses position
Call GetCursorPos(MousePos)

Call DrawFrame(MousePos.X, MousePos.Y, False, True)
End Sub

Public Sub DrawFrame(ByVal X As Integer, ByVal Y As Integer, Optional ByVal MousePressed As Boolean, Optional OnTopOfAll As Boolean, Optional ByVal Unloading As Boolean = False)
'This will draw the mouse cursor onto the screen

Static LastX As Integer
Static LastY As Integer
Static Started As Boolean
Static LasthCursor As Long
Static LastCursorBmp As BitmapStruc

Dim OffsetX As Integer
Dim OffsetY As Integer
Dim BackOffsetX As Integer
Dim BackOffsetY As Integer
Dim MouseOffsetX As Integer
Dim MouseOffsetY As Integer
Dim Result As Long
Dim TotalBackBmp As BitmapStruc
Dim TempCurBmp As BitmapStruc
Dim TempMaskBmp As BitmapStruc
Dim hDcSurphase As Long

'if unloading the program, then remove static bitmaps from memory
If Unloading Then
    Call DeleteBitmap(LastCursorBmp.hDcMemory, LastCursorBmp.hDcBitmap, LastCursorBmp.hDcPointer)
End If

DoEvents

' Draw the cursor's mask to a picturebox
hCursor = GetCursor
If hCursor <> LasthCursor Then
    Result = DrawIconEx(CursorBmp.hDcMemory, 0, 0, hCursor, 0, 0, 0, 0, 1)
End If

'set the previous cursor picture if one hasn't been loaded.
If LasthCursor = 0 Then
    LastCursorBmp.Area.Right = 32
    LastCursorBmp.Area.Bottom = 32
    Call CreateNewBitmap(LastCursorBmp.hDcMemory, LastCursorBmp.hDcBitmap, LastCursorBmp.hDcPointer, LastCursorBmp.Area, frmMouseShadow.hdc, vbWhite, InPixels)
    
    'remember the mouse cursor bitmap used so we can delete the shadow next time.
    Result = BitBlt(LastCursorBmp.hDcMemory, 0, 0, LastCursorBmp.Area.Right, LastCursorBmp.Area.Bottom, CursorBmp.hDcMemory, 0, 0, SRCCOPY)
End If
LasthCursor = hCursor

'get the surphase to draw on
If OnTopOfAll Then
    'get the top surphase
    hDcSurphase = GetDC(0)
Else
    'only get the forms surphase
    hDcSurphase = frmMouseShadow.hdc
End If

'adjust for the shadow height
If MouseKeyPressed(MouseLeft) Or MouseKeyPressed(MouseMiddle) Or MouseKeyPressed(MouseRight) Or MousePressed Then
    'a mouse key is pressed
    X = X + PressHeightX
    Y = Y + PressHeightY
Else
    'no mouse key is pressed
    X = X + ShadowHeightX
    Y = Y + ShadowHeightY
End If

'calculate the difference in position
OffsetX = X - LastX
OffsetY = Y - LastY

If OffsetX > 0 Then
    BackOffsetX = 0
    MouseOffsetX = OffsetX
Else
    BackOffsetX = -OffsetX
    MouseOffsetX = 0
End If
If OffsetY > 0 Then
    BackOffsetY = 0
    MouseOffsetY = OffsetY
Else
    BackOffsetY = -OffsetY
    MouseOffsetY = 0
End If

'create the bitmap
'set the bitmap size
TotalBackBmp.Area.Top = 0
TotalBackBmp.Area.Left = 0
If Started Then
    TotalBackBmp.Area.Right = Abs(OffsetX) + 32
    TotalBackBmp.Area.Bottom = Abs(OffsetY) + 32
Else
    'only do this once
    TotalBackBmp.Area.Right = 32
    TotalBackBmp.Area.Bottom = 32
End If
Call CreateNewBitmap(TotalBackBmp.hDcMemory, TotalBackBmp.hDcBitmap, TotalBackBmp.hDcPointer, TotalBackBmp.Area, frmMouseShadow.hdc, 0, InPixels)

'create a temperory cursor sized bitmap used for a cursor mask and
'to hold the result.
TempCurBmp.Area = CursorBmp.Area
TempMaskBmp.Area = CursorBmp.Area
Call CreateNewBitmap(TempCurBmp.hDcMemory, TempCurBmp.hDcBitmap, TempCurBmp.hDcPointer, TempCurBmp.Area, frmMouseShadow.hdc, vbWhite, InPixels)
Call CreateNewBitmap(TempMaskBmp.hDcMemory, TempMaskBmp.hDcBitmap, TempMaskBmp.hDcPointer, TempMaskBmp.Area, frmMouseShadow.hdc, vbWhite, InPixels)

'capture the screen area
Result = BitBlt(TotalBackBmp.hDcMemory, 0, 0, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, hDcSurphase, X - MouseOffsetX, Y - MouseOffsetY, SRCCOPY)

If Not Started Then
    'put the background into the background picture - only do this once
    Result = BitBlt(BackBmp.hDcMemory, 0, 0, 32, 32, hDcSurphase, X, Y, SRCCOPY)
    Started = True
Else
    'copy the old background over where the mouse used to be
    
    'copy a cursor-shpaed section of the old background onto where the
    'cursor shadow is on the screen. This means that we can leave as much
    'of the current screenshot as untouched as we can. Only a cursor-
    'shaped section of it will not be as current. See the procedure
    'MergeBitmaps to see how this works.
    Call MergeBitmaps(TempMaskBmp.hDcMemory, TotalBackBmp.hDcMemory, BackBmp.hDcMemory, LastCursorBmp.hDcMemory, 0, 0, BackOffsetX, BackOffsetY, 0, 0, CursorBmp.Area.Right, CursorBmp.Area.Bottom, InPixels)
    
    'update the screen shot
    Result = BitBlt(TotalBackBmp.hDcMemory, BackOffsetX, BackOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, TempMaskBmp.hDcMemory, 0, 0, SRCCOPY)
    
    'TotalBackBmp should now contain a clean picture with no cursor.
    'copy a section of this as the background for next time.
    Result = BitBlt(BackBmp.hDcMemory, 0, 0, 32, 32, TotalBackBmp.hDcMemory, MouseOffsetX, MouseOffsetY, SRCCOPY)
    
    'remember the mouse cursor bitmap used so we can delete the shadow next time.
    Result = BitBlt(LastCursorBmp.hDcMemory, 0, 0, LastCursorBmp.Area.Right, LastCursorBmp.Area.Bottom, CursorBmp.hDcMemory, 0, 0, SRCCOPY)
End If

'draw the mouse cursor onto TotalBackBmp
Result = BitBlt(TotalBackBmp.hDcMemory, MouseOffsetX, MouseOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, CursorBmp.hDcMemory, 0, 0, SRCAND)

'copy the drawn picture onto the screen
Result = BitBlt(hDcSurphase, X - MouseOffsetX, Y - MouseOffsetY, TotalBackBmp.Area.Right, TotalBackBmp.Area.Bottom, TotalBackBmp.hDcMemory, 0, 0, SRCCOPY)

LastX = X
LastY = Y

'remove the temperory bitmap holding the new and old backgrounds
Call DeleteBitmap(TotalBackBmp.hDcMemory, TotalBackBmp.hDcBitmap, TotalBackBmp.hDcPointer)
Call DeleteBitmap(TempCurBmp.hDcMemory, TempCurBmp.hDcBitmap, TempCurBmp.hDcPointer)
Call DeleteBitmap(TempMaskBmp.hDcMemory, TempMaskBmp.hDcBitmap, TempMaskBmp.hDcPointer)

If OnTopOfAll Then
    'let go of the top screen
    Result = ReleaseDC(0, hDcSurphase)
End If
End Sub


