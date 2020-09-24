VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Don't Forget The Title !"
   ClientHeight    =   6180
   ClientLeft      =   2685
   ClientTop       =   1455
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   4500
      Left            =   120
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   446
      TabIndex        =   3
      Top             =   480
      Width           =   6750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End Program"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   240
      Top             =   5400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'some times the animation runs like shit ! (slowly)
'don't know why but often happen at early minutes program runs
'after a few minutes it's run very smooth ! no flick ! with standard mouse pointer
'(i tried it on  Intel PII 333 MHz mmx,  MS Win 98 SE, with winamp running, 64 Mb RAM and 4 Mb AGP Adapter [it's AGP technology takes any part in this 2D animation ???])
'any solution to make it stabil please email me
' very@gobytown.com
'any question just email me  !!!
'hope you like it !

Private himg As Long, hpic As Long
Private himgdump As Long
Private himgtemp As Long
Dim err As Integer
Dim xadder As Integer, yadder As Integer

Private Sub Command1_Click()

Timer1.Interval = 50
Timer1.Enabled = True
   
End Sub

Private Sub Command2_Click()

Timer1.Enabled = False

End Sub

Private Sub Command3_Click()
 
 'release the resource (hope it's rite too)
  err = DeleteDC(himgtemp)
  'MsgBox CStr(err)
  err = DeleteDC(himg)
  'MsgBox CStr(err)
  err = DeleteObject(hpic)
  'MsgBox CStr(err)
  
End

End Sub



Private Sub Form_GotFocus()
 
 Timer1.Enabled = True
 
End Sub

Private Sub Form_Load()
Dim a As Integer

 'preparation for the picture and bla bla bla bla ...
 'load the picture
 hpic = LoadImageA(0, App.Path & "/winamp.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
 'MsgBox CStr(hpic)
 himg = CreateCompatibleDC(Picture1.hdc)
 'msgbox cstr(himg)
 'copy the picture to device context
 err = SelectObject(himg, hpic)
 'MsgBox CStr(err)
 
 'get the info about the picture (size, bitmaps bits etc)
 err = GetObjectA(hpic, Len(picinfo), picinfo)
 'MsgBox CStr(err)
 
 'standard procedure to make it ++ in appearance !
 Picture1.BorderStyle = 0
 Picture1.BackColor = vbBlack
 Picture1.ScaleMode = 3
 Picture1.AutoSize = False

 'ok, the dump dc must be prepare
 himgdump = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
 'MsgBox CStr(himgdump)
 
 'the memory dc is about to be created !!!
 himgtemp = CreateCompatibleDC(0)
 'MsgBox CStr(himgtemp)
 err = SelectObject(himgtemp, himgdump)
 'MsgBox CStr(err)
 'free the resource ! (hopes it's work!)
 'give a better solution : very@gobytown.com
 err = DeleteObject(himgdump)
 
 'prepare the frames
 For a = 0 To maxframes
  picframe(a).fx = xs
  picframe(a).fy = ys
  picframe(a).fheight = picinfo.bheight
  picframe(a).fwidth = picinfo.bwidth
 Next a
 
 'I pick this randomly
 xadder = 4
 yadder = 7
 
 Timer1.Interval = 50
 
End Sub

Private Sub Form_LostFocus()
 
 Timer1.Enabled = False
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  'release the resource (hope it's rite too)
  err = DeleteDC(himgtemp)
  'MsgBox CStr(err)
  err = DeleteDC(himg)
  'MsgBox CStr(err)
  err = DeleteObject(hpic)
  
End Sub

Private Sub Picture1_Click()
 
 MsgBox "Leave a souvenir here heheheh", vbOKOnly, "very@gobytown.com"
 
End Sub

Private Sub Timer1_Timer()
Dim a As Integer
Dim n As Integer

 'copy the background to the memory dc
 err = StretchBlt(himgtemp, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy)
 'MsgBox CStr(err)
 
 'make a right animation in the memory dc
 For a = 0 To maxframes
  'simple bounce algorithm
  'send me better one please ! very@gobytown.com
  If picframe(maxframes).fx > Picture1.ScaleWidth - picinfo.bwidth - 40 Then xadder = -4
  If picframe(maxframes).fx < lborder Then xadder = 4
  If picframe(maxframes).fy > Picture1.ScaleHeight - picinfo.bheight - 40 Then yadder = -7
  If picframe(maxframes).fy < tborder Then yadder = 7
   
   n = maxframes - a
   If n < 0 Then n = 0
   
   'draw it to the memory dc and .....
   picframe(a).fx = picframe(a).fx - (5 * n) + xadder
   picframe(a).fy = picframe(a).fy - (5 * n) + yadder
   '(don't forget to arrange the size)
   picframe(a).fwidth = picframe(a).fwidth + (5 * n * 2)
   picframe(a).fheight = picframe(a).fheight + (5 * n * 2)
   err = StretchBlt(himgtemp, picframe(a).fx, picframe(a).fy, picframe(a).fwidth, picframe(a).fheight, himg, 0, 0, picinfo.bwidth, picinfo.bheight, vbSrcCopy)
   'MsgBox CStr(err)
   
   'if the most outer frames left reach the edges then
   'change it as the last frame
   If picframe(0).fx < 0 Then
    'don't forget to arrannge the size
    picframe(0).fx = picframe(maxframes).fx
    picframe(0).fy = picframe(maxframes).fy
    picframe(0).fwidth = picinfo.bwidth
    picframe(0).fheight = picinfo.bheight
    swapframe
   End If
 Next a
 
 'copy it to the picture1 dc !
 err = BitBlt(Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, himgtemp, 0, 0, vbSrcCopy)
 'MsgBox CStr(err)
 
End Sub
