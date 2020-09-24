Attribute VB_Name = "Module1"
Option Explicit
Option Base 0

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hobj As Long) As Integer
Public Declare Function GetObjectA Lib "gdi32" (ByVal hobj As Long, ByVal buffsize As Integer, ByRef buff As bitmap) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdcd As Long, ByVal xd As Long, ByVal yd As Long, ByVal widthd As Long, ByVal heightd As Long, ByVal hdcs As Long, ByVal xs As Long, ByVal ys As Long, ByVal widths As Long, ByVal heights As Long, ByVal opr As Long) As Integer
Public Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal pfilename As String, ByVal typeimg As Long, ByVal width As Long, ByVal height As Long, ByVal flag As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hdcd As Long, ByVal xd As Long, ByVal yd As Long, ByVal widthd As Long, ByVal heightd As Long, ByVal hdcs As Long, ByVal xs As Long, ByVal ys As Long, ByVal opr As Long) As Integer
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Integer
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Integer
Public Declare Function DeleteObject Lib "gdi32" (ByVal hobj As Long) As Integer
  
  
'MS bitmap's structure !
Type bitmap
 btype As Long
 bwidth As Long
 bheight As Long
 bwidthbytes As Long
 bplanes As Integer
 bbitpixels As Integer
 bbits As Integer
End Type

'my frames' structure hehehe
Type Pframe
 fx As Integer
 fy As Integer
 fwidth As Integer
 fheight As Integer
End Type

'constant  p.s don't ask me where it's came from !
Public Const IMAGE_BITMAP = &O0
Public Const LR_LOADFROMFILE = 16

'number for the frames (i just use 6 + 1), more means better but more complex calculation
'if anyone can give me better solution email me ! very@gobytown.com
Public Const maxframes = 6

Public picinfo As bitmap
Public picframe(maxframes) As Pframe 'array for the frames

Public Const xs = 112 'start position for  the picture
Public Const ys = 72
Public Const tborder = 40 'distance for the edge of the picture
Public Const lborder = 40

'procedure to swap the positon of the frames in the array
'1rst --> last one , the other just follow the flow

Public Sub swapframe()
 Dim a As Integer
 Dim picframetemp As Pframe
 
 'use temporary variabel
 picframetemp = picframe(0)
 
 'flow it !
 For a = 1 To maxframes - 1
  picframe(a) = picframe(a + 1)
 Next a
 
 'don't forget to move it !
 picframe(maxframes) = picframetemp
 
End Sub



