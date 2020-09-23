VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6885
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WOODS4.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2865
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   0
      Width           =   4035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " THIS IS ONLY DONE        ONCE TO MAKE THE       BITMAP FILE                     ""PopSeqRed.bin"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2835
      TabIndex        =   1
      Top             =   915
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Woods + by Robert Rayment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   2580
      TabIndex        =   0
      Top             =   570
      Width           =   2415
   End
   Begin VB.Menu mnuAnimStart 
      Caption         =   "START"
   End
   Begin VB.Menu mnuAnimStop 
      Caption         =   "STOP"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuTreeStart1 
         Caption         =   "Tree Starts = 1"
      End
      Begin VB.Menu mnuTreeStarts2 
         Caption         =   "Tree Starts = 2"
      End
      Begin VB.Menu mnuTreeStarts3 
         Caption         =   "Tree Starts = 3"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeSpread1 
         Caption         =   "Tree Spread = 1"
      End
      Begin VB.Menu mnuTreeSpread2 
         Caption         =   "Tree Spread = 2"
      End
      Begin VB.Menu mnuTreeSpread3 
         Caption         =   "Tree Spread = 3"
      End
      Begin VB.Menu mnuTreeSpread4 
         Caption         =   "Tree Spread = 4"
      End
      Begin VB.Menu mnuTreeSpread8 
         Caption         =   "Tree Spread = 8"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Woods4.frm  by  Robert Rayment 23/2/01 -

'Version 2 7/3/01

'Improved creation of smaller bitmaps thanks to having been goaded
'by Paul Bower.  Also some simple sound added.

'NB This needs compiling to EXE to get full speed!!!

Option Base 1

DefInt A-T  'A.. to T.. integers
DefSng U-Z  'U.. to Z.. singles


Private Sub Form_Load()

Done = True       'For Loop exit
KeyPreview = True 'Allows Form_Key... to pick up keying first
keyright = False
keyleft = False

'Set form up
ScaleMode = vbPixels
WindowState = vbNormal
AutoRedraw = False
Top = 500
Left = 1000

'Match form size to draw surface size
Width = (512 + 8) * Screen.TwipsPerPixelX  'ie * 15
Height = (512 + 19 + 8) * Screen.TwipsPerPixelY 'ie * 15

Show
DoEvents

'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

'Drawing surfaces
ReDim BackSurf(512, 512)
ReDim FrontSurf(512, 512)

'Fill BITDATA structure
'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
FillBITStructures      'Sub in BAS module

'Move cursor so that HourGlass shows
res& = SetCursorPos(100, 200)

MousePointer = vbHourglass
'============================================
'LOAD TREE BMP INTO LARGEST TREE ARRAY .. Tree19(128,512)

'also

'get palette from BMP and place in bm.Colors(n).rgbBlue .rgbGreen  .rgbRed
'NB it is essential that the color number 255 is RGB = 255,255,255 (White)
'White will be a transparent area. No masks are needed.

Caption = " LOADING BITMAPS"

'SET UP TREE DATA

NTreeStarts = 3  'Maximum number of tree starts

SetupTreeArrays      'Sub in BAS module
'includes
'NTreeSizes = 19  'Number of tree sizes
'NTreeTypes = 3   'Number of tree types

'Load car left, centre & right
'Global LCar() As Byte, HCar() As Byte, RCar() As Byte
ReDim LCar(48, 48), HCar(48, 48), RCar(48, 48)
LoadCar PathSpec$ & "LCar.bmp", 1
LoadCar PathSpec$ & "HCar.bmp", 2
LoadCar PathSpec$ & "RCar.bmp", 3
'============================================
'DEVELOP OTHER TREE SIZES & STORE IN REMAINING TREE ARRAYS

If Dir$(PathSpec$ & "PopSeqRed.bin") = "" Then
   
   'This is run once to produce the bit map file "PopSeqRed.bin"
   
   LoadMainTree PathSpec$ & "PopTree.bmp", 1 'Sub in BAS module
   LoadMainTree PathSpec$ & "SeqTree.bmp", 2
   LoadMainTree PathSpec$ & "RedTree.bmp", 3
   
   DevelopTrees
   
Else  'Load from "PopSeqRed.bin"
   
   'Load a tree to get palette
   LoadMainTree PathSpec$ & "PopTree.bmp", 1 'Sub in BAS module
   
   LoadTrees

End If

MousePointer = vbDefault

'Fix array for Trees' data on form for all sizes
ReDim ITX(NTreeStarts, NTreeSizes)  'X-pos of tree
ReDim ITD(NTreeStarts, NTreeSizes)  'Tree's left or right direction
ReDim ITT(NTreeStarts, NTreeSizes)  'Tree type 1,2,3,,

'Change screen to 640 x 480
Width = (640 + 8) * Screen.TwipsPerPixelX  'ie * 15
Height = (480 + 19 + 8) * Screen.TwipsPerPixelY 'ie * 15

DoEvents

FIW& = Width: FIH& = Height

WhitenBackSurface  'Sub in BAS module

Caption = " Woods  by Robert Rayment"
Done = False   'Allow coords to show with Form_MouseMove


'##########################################################

'SET INITIAL OPTIONS

NTreeStarts = 1
Spread = 3        'Trees x-pos +/- Spread as they move forward
StretchMode = 1
mnuTreeStart1.Checked = True
mnuTreeSpread3.Checked = True

'ANIMATIONLOOP

End Sub

Private Sub ANIMATIONLOOP()

'Set all trees off screen by making x = -99
For i = 1 To NTreeStarts
For j = 1 To NTreeSizes
   ITX(i, j) = -99
   ITD(i, j) = 0
   ITT(i, j) = 1
Next j
Next i


'Set background sky
For iy = 256 To 512
For ix = 1 To 512
    BackSurf(ix, iy) = 185
Next ix
Next iy

Randomize   'Seed random numbers, used for fixing Tree1 positions

'Set background ground
For iy = 8 To 248 Step 8
For ix = 4 To 508 Step 4
    If Rnd < 0.75 Then
      BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 142
    Else
      BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 142 + 2 * Rnd
    End If
Next ix
Next iy

Counter# = 0       'Hit counter
pathcentre = 256   'Trees x-pos < 256 will move left
                   '& >=256 will move rightTreeType = 1
'Start speed
TickDifference = 10   '1, 10 or 100
Speed$ = " Medium"    'Slow, Medium or Fast


'Animation loop

Done = False   'Will loop until Done = True
   
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech2.wav", 0, wFlags%)

Do
    'Timing
   CurrentTick = GetTickCount()
   If CurrentTick - LastTick > TickDifference Then
   
      'Get current form width & height
      FIW& = Me.Width \ Screen.TwipsPerPixelX - 8
      FIH& = Me.Height \ Screen.TwipsPerPixelY - 8
   
      'Clear largest trees at the front of screen
      TreeNum = NTreeSizes
      For N = 1 To NTreeStarts
         ITX(N, TreeNum) = -99
         ITD(N, TreeNum) = 0
         ITT(N, TreeNum) = 1
      Next N
   
      'Move trees (NTreeSizes - 1) to 1 forward & sideways to (NTreeSizes - 2) to 2
      For TreeNum = NTreeSizes - 1 To 1 Step -1
      For N = 1 To NTreeStarts
         If ITX(N, TreeNum) <> -99 Then
            'Move forward & sideways according to direction ITD & Spread
            If ITD(N, TreeNum) = -1 Then
               ITX(N, TreeNum + 1) = ITX(N, TreeNum) - Spread
               If ITX(N, TreeNum + 1) + TreeW(TreeNum + 1) < 1 Then ITX(N, TreeNum + 1) = -99
            ElseIf ITD(N, TreeNum) = 1 Then
               ITX(N, TreeNum + 1) = ITX(N, TreeNum) + Spread
               If ITX(N, TreeNum + 1) > 511 Then ITX(N, TreeNum + 1) = -99
            End If
            'Move direction & type forward
            ITD(N, TreeNum + 1) = ITD(N, TreeNum)
            ITT(N, TreeNum + 1) = ITT(N, TreeNum)
         End If
      Next N
      Next TreeNum
   
      'Place first NTreeStarts trees randomly
      TreeNum = 1
   
      For N = 1 To NTreeStarts
         ityp = 1 + Int(Rnd * NTreeTypes)
         If ityp > NTreeTypes Then ityp = NTreeTypes
         ITT(N, TreeNum) = ityp
         ITX(N, TreeNum) = 1 + Rnd * 511
         If ITX(N, TreeNum) < (pathcentre) Then
            ITD(N, TreeNum) = -1 'Tree tends to left
         ElseIf ITX(N, TreeNum) >= (pathcentre) Then
            ITD(N, TreeNum) = 1  'Tree tends to right
         End If
      Next N
   
      '---------- CHECK KEYING --------------
      If keyleft = True Then
         'Show current settings & Hits
         pic1.Cls
         pic1.Print "Left     Speed" & Speed$ & "  Hits = " & Str$(Counter#)
      
         'Set background ground
         For iy = 1 To 248 Step 8
         For ix = 1 To 508 Step 4
            If Rnd < 0.75 Then
               BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 255
            Else
               BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 142 + 2 * Rnd
            End If
         Next ix
         Next iy
         
         'Move all trees to right
         For TreeNum = 1 To NTreeSizes
            For N = 1 To NTreeStarts
               If ITX(N, TreeNum) <> -99 Then
                  ITX(N, TreeNum) = ITX(N, TreeNum) + 12
                  If ITX(N, TreeNum) > 512 Then ITX(N, TreeNum) = -99
               End If
            Next N
         Next TreeNum
   
      ElseIf keyright = True Then
         'Show current settings & Hits
         pic1.Cls
         pic1.Print "Right    Speed" & Speed$ & "  Hits = " & Str$(Counter#)
      
         'Set background ground
         For iy = 1 To 248 Step 8
         For ix = 1 To 508 Step 4
            If Rnd < 0.75 Then
               BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 255
            Else
               BackSurf(ix + Rnd * 3, iy + Rnd * 7) = 142 + 2 * Rnd
            End If
         Next ix
         Next iy
      
         'Move all trees to left
         For TreeNum = 1 To NTreeSizes
            For N = 1 To NTreeStarts
               If ITX(N, TreeNum) <> -99 Then
                  ITX(N, TreeNum) = ITX(N, TreeNum) - 16
                  If ITX(N, TreeNum) + TreeW(TreeNum) < 1 Then ITX(N, TreeNum) = -99
               End If
            Next N
         Next TreeNum
      Else
         pic1.Cls
         pic1.Print "Forward  Speed" & Speed$ & "  Hits = " & Str$(Counter#)
      End If
      '---------- KEYING ENDS --------------
   
   
      'Fill FrontSurf from BackSurf
      CopyMemory FrontSurf(1, 1), BackSurf(1, 1), 512& * 512&
   
    
      'COPY non-white bits to FrontSurface, for each tree
      For TreeNum = 1 To NTreeSizes

         For N = 1 To NTreeStarts
      
            If ITX(N, TreeNum) <> -99 Then
      
               TreeType = ITT(N, TreeNum)
      
               ty = 1
               For iy = 256 - (TreeH(TreeNum) \ 2) + 1 To 256 + (TreeH(TreeNum) \ 2) - 1
               tx = 1
               For ix = ITX(N, TreeNum) To ITX(N, TreeNum) + TreeW(TreeNum) - 1
               Select Case TreeNum
               Case 1: bpp = Tree1(tx, ty, TreeType)
               Case 2: bpp = Tree2(tx, ty, TreeType)
               Case 3: bpp = Tree3(tx, ty, TreeType)
               Case 4: bpp = Tree4(tx, ty, TreeType)
               Case 5: bpp = Tree5(tx, ty, TreeType)
               Case 6: bpp = Tree6(tx, ty, TreeType)
               Case 7: bpp = Tree7(tx, ty, TreeType)
               Case 8: bpp = Tree8(tx, ty, TreeType)
               Case 9: bpp = Tree9(tx, ty, TreeType)
               Case 10: bpp = Tree10(tx, ty, TreeType)
         
               Case 11: bpp = Tree11(tx, ty, TreeType)
               Case 12: bpp = Tree12(tx, ty, TreeType)
               Case 13: bpp = Tree13(tx, ty, TreeType)
               Case 14: bpp = Tree14(tx, ty, TreeType)
               Case 15: bpp = Tree15(tx, ty, TreeType)
               Case 16: bpp = Tree16(tx, ty, TreeType)
               Case 17: bpp = Tree17(tx, ty, TreeType)
               Case 18: bpp = Tree18(tx, ty, TreeType)
               Case 19: bpp = Tree19(tx, ty, TreeType)
               End Select
      
               If bpp <> 255 And ix > 0 And ix < 512 Then
                  FrontSurf(ix, iy) = bpp
               End If
         
               tx = tx + 1
               Next ix
               ty = ty + 1
               Next iy
      
            End If
      
         Next N
   
      Next TreeNum
   
      'Put on car
      If keyleft = True Then
         'LCar
         For iy = 1 To 48
         For ix = 1 To 48
            bpp = LCar(ix, iy)
            If bpp <> 255 Then
               FrontSurf(ix + 256 - 24, iy + 60) = bpp
            End If
         Next ix
         Next iy
         'Shadow under LCar
         For iy = 56 To 58
         For ix = -8 To 8
            FrontSurf(ix + 256, iy) = 22
         Next ix
         Next iy
      ElseIf keyright = True Then
         'RCar
         For iy = 1 To 48
         For ix = 1 To 48
            bpp = RCar(ix, iy)
            If bpp <> 255 Then
               FrontSurf(ix + 256 - 24, iy + 60) = bpp
            End If
         Next ix
         Next iy
         'Shadow under RCar
         For iy = 56 To 58
         For ix = -8 To 8
            FrontSurf(ix + 256, iy) = 22
         Next ix
         Next iy
      Else
         'HCar
         For iy = 48 To 1 Step -1
         For ix = 1 To 48
            bpp = HCar(ix, iy)
            If bpp <> 255 Then
               FrontSurf(ix + 256 - 24, iy + 60) = bpp
            End If
         Next ix
         Next iy
         'Shadow under HCar
         For iy = 56 To 58
         For ix = -12 To 12
            FrontSurf(ix + 256, iy) = 22
         Next ix
         Next iy
      End If
   
      'Check collision at last trees
      TreeNum = NTreeSizes
      For N = 1 To NTreeStarts
         TreeCen = ITX(N, TreeNum) + (TreeW(TreeNum) / 2)
         If TreeCen > 242 And TreeCen < 270 Then   'Collision width, sensitivity
            Counter# = Counter# + 1
         End If
      Next N
   
      'Display FrontSurf
   
      succ& = StretchDIBits(Me.hdc, _
      0, 0, _
      FIW&, FIH&, _
      0, 0, _
      512&, 512&, _
      ByVal bmp.ptrFrontSurf, bm, _
      DIB_RGB_COLORS, SRCCOPY)
   
      LastTick = GetTickCount()
   
   End If   'GetTickCount()
   
   DoEvents
   
Loop Until Done


res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Changes direction & speed

kleft = 37
kright = 39
kup = 38
kdn = 40

Select Case KeyCode
Case kleft
   keyleft = True
   keyright = False
   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech.wav", 0, wFlags%)
Case kright
   keyright = True
   keyleft = False
   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech.wav", 0, wFlags%)
Case kup
   If TickDifference > 1 Then
      TickDifference = TickDifference / 10
      Select Case TickDifference
      Case 1: Speed$ = "  Fast "
      Case 10: Speed$ = " Medium"
      Case 100: Speed$ = "   Slow"
      End Select
   End If
   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech2.wav", 0, wFlags%)
Case kdn
   If TickDifference < 100 Then
      TickDifference = 10 * TickDifference
      Select Case TickDifference
      Case 1: Speed$ = "  Fast "
      Case 10: Speed$ = " Medium"
      Case 100: Speed$ = "   Slow"
      End Select
   End If
   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech2.wav", 0, wFlags%)
Case Else
   keyleft = False
   keyright = False
   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech2.wav", 0, wFlags%)
End Select


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Cancel keyleft & keyright

kleft = 37
kright = 39
kup = 38
kdn = 40

Select Case KeyCode
Case kleft
   keyleft = False
Case kright
   keyright = False
Case Else
'   keyleft = False
'   keyright = False
End Select

   res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
   wFlags% = SND_LOOP Or SND_ASYNC
   res& = PlaySound("HiTech2.wav", 0, wFlags%)

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Temporary for Mouse position
'If Done = False Then
'   FormWidth& = Form1.Width \ Screen.TwipsPerPixelX
   'pathcentre = Int(X * (512 / FormWidth&))
'   If pathcentre < 70 Then pathcentre = 70
'   If pathcentre > 460 Then pathcentre = 460
'   Caption = Str$(X) + Str$(Y) & Str$(pathcentre)
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Over-done exit !!?
stopper = True
Done = True    'end loop
res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
Erase BackSurf, FrontSurf
DoEvents
Unload Me
End
End Sub


Private Sub DevelopTrees()
   Caption = "CREATING BITMAPS FOR SMALLER TREES"
   Label2.Visible = True
   Refresh
   
   For TreeType = 1 To NTreeTypes
      DevelopOtherTrees TreeType  'Sub in BAS module
   Next TreeType
   
   'Save to "PopSeqRed.bin"
   Open PathSpec$ & "PopSeqRed.bin" For Binary Access Write As #1
       Put #1, , Tree1()   'Will do all TreeTypes for TreeN()
       Put #1, , Tree2()   'etc TreeN(ix,iy, TreeType)
       Put #1, , Tree3()
       Put #1, , Tree4()
       Put #1, , Tree5()
       Put #1, , Tree6()
       Put #1, , Tree7()
       Put #1, , Tree8()
       Put #1, , Tree9()
       Put #1, , Tree10()
       Put #1, , Tree11()
       Put #1, , Tree12()
       Put #1, , Tree13()
       Put #1, , Tree14()
       Put #1, , Tree15()
       Put #1, , Tree16()
       Put #1, , Tree17()
       Put #1, , Tree18()
       Put #1, , Tree19()
   Close

Label2.Visible = False

End Sub

Private Sub LoadTrees()
   Close
   Open PathSpec$ & "PopSeqRed.bin" For Binary Access Read As #1
      
       Get #1, , Tree1()   'ie for each Get TreeN(ix,iy, TreeType)
       Get #1, , Tree2()
       Get #1, , Tree3()
       Get #1, , Tree4()
       Get #1, , Tree5()
       Get #1, , Tree6()
       Get #1, , Tree7()
       Get #1, , Tree8()
       Get #1, , Tree9()
       Get #1, , Tree10()
       Get #1, , Tree11()
       Get #1, , Tree12()
       Get #1, , Tree13()
       Get #1, , Tree14()
       Get #1, , Tree15()
       Get #1, , Tree16()
       Get #1, , Tree17()
       Get #1, , Tree18()
       Get #1, , Tree19()
   Close
End Sub

Private Sub mnuAnimStart_Click()
ANIMATIONLOOP
End Sub

Private Sub SpreadClrChecks()
mnuTreeSpread1.Checked = False
mnuTreeSpread2.Checked = False
mnuTreeSpread3.Checked = False
mnuTreeSpread4.Checked = False
mnuTreeSpread8.Checked = False
End Sub
Private Sub StartClrChecks()
mnuTreeStart1.Checked = False
mnuTreeStarts2.Checked = False
mnuTreeStarts3.Checked = False
End Sub

Private Sub mnuAnimStop_Click()
Done = True
res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
End Sub


Private Sub mnuOptions_Click()
Done = True
res& = PlaySound("", 0, SND_PURGE Or SND_NODEFAULT)
End Sub

Private Sub mnuTreeSpread1_Click()
Done = True
SpreadClrChecks
Spread = 1
mnuTreeSpread1.Checked = True
End Sub

Private Sub mnuTreeSpread2_Click()
Done = True
SpreadClrChecks
Spread = 2
mnuTreeSpread2.Checked = True
End Sub

Private Sub mnuTreeSpread3_Click()
Done = True
SpreadClrChecks
Spread = 3
mnuTreeSpread3.Checked = True
End Sub

Private Sub mnuTreeSpread4_Click()
Done = True
SpreadClrChecks
Spread = 4
mnuTreeSpread4.Checked = True
End Sub

Private Sub mnuTreeSpread8_Click()
Done = True
SpreadClrChecks
Spread = 8
mnuTreeSpread8.Checked = True
End Sub

Private Sub mnuTreeStart1_Click()
Done = True
StartClrChecks
NTreeStarts = 1
mnuTreeStart1.Checked = True
End Sub

Private Sub mnuTreeStarts2_Click()
Done = True
StartClrChecks
NTreeStarts = 2
mnuTreeStarts2.Checked = True
End Sub

Private Sub mnuTreeStarts3_Click()
Done = True
StartClrChecks
NTreeStarts = 3
mnuTreeStarts3.Checked = True
End Sub
