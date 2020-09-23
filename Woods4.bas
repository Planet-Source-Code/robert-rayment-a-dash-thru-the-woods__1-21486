Attribute VB_Name = "Module1"
'Woods4.bas by Robert Rayment

'VERSION 2

Option Base 1

DefInt A-T  'A.. to T.. integers
DefSng U-Z  'U.. to Z.. singles


Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long


Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


'Copy one array to another of same number of bytes
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)


'Structures for StretchDIBits
Public Type BITMAPINFOHEADER '40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type

Public bm As BITMAPINFO


'For transferring drawing in byte array to Form
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal SrcW As Long, ByVal SrcH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Set alternative actions for StretchDIBits (not used here)
'Seems to have little effect on 8-bit color
Public Declare Function SetStretchBltMode Lib "gdi32" _
(ByVal hdc As Long, ByVal nStretchMode As Long) As Long
'nStretchMode
Public Const STRETCH_ANDSCANS = 1    'default
Public Const STRETCH_ORSCANS = 2
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_HALFTONE = 4


'Structure for surface info
Public Type BITDATA
   bmWidth As Long         ' Pixel width
   bmHeight As Long        ' Pixel height
   ptrBackSurf As Long     ' Pointer to drawing surface
   ptrFrontSurf As Long     ' Pointer to rendering surface
   ptrBMPSurf As Long     ' Pointer to BMP surface
   ptrCopySurf As Long     ' Pointer to BMP surface
End Type

Public bmp As BITDATA

'Constants for StretchDIBits
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046
Public Const SRCAND = &H880C6
Public Const SRCPAINT = &HEE086

'-------------------------------------------------------
Global FrontSurf() As Byte
Global BackSurf() As Byte

'Number of Tree arrays = NTreeSizes
Global Tree19() As Byte
Global Tree18() As Byte
Global Tree17() As Byte
Global Tree16() As Byte
Global Tree15() As Byte
Global Tree14() As Byte
Global Tree13() As Byte
Global Tree12() As Byte
Global Tree11() As Byte
Global Tree10() As Byte
Global Tree9() As Byte
Global Tree8() As Byte
Global Tree7() As Byte
Global Tree6() As Byte
Global Tree5() As Byte
Global Tree4() As Byte
Global Tree3() As Byte
Global Tree2() As Byte
Global Tree1() As Byte
Global FIW&, FIH&       'StretchDIBits Destination Width & Height on Form
Global TreeW(), TreeH()  'Tree surfaces' width & height
Global NTreeStarts   'Number of tree starts on each loop
Global NTreeSizes    'Number of tree sizes
Global NTreeTypes    'Number of tree types
Global Counter#      'Hit counter
Global Spread        'Trees x-pos +/- Spread as they move forward
Global pathcentre    'Trees x-pos < 256 will move left
Global LCar() As Byte, HCar() As Byte, RCar() As Byte
'Tree data
Global ITX()             'A tree's x-pos & size
Global ITD()             '-1 left, +1 right, tree horizontal movement
Global ITT()             'Tree type 1,2,3,

'Global bpp As Byte, bpAND As Byte, bpOR As Byte     'For picking up color number
Global PathSpec$        'App path
Global Done As Boolean  'To exit loop
Global bpp As Byte      'color numbers picked up

Global keyright As Boolean    'save key state
Global keyleft As Boolean
Global TickDifference As Long 'timing
Global LastTick As Long
Global CurrentTick As Long

Global Speed$                 'speed text Slow, Medium, Fast

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_PURGE = &H40

'- SND_SYNC specifies that the sound is played synchronously and the
'  function does not return until the sound ends.

'- SND_ASYNC specifies that the sound is played asynchronously and the
'  function returns immediately after beginning the sound.

'- SND_NODEFAULT specifies that if the sound cannot be found, the
'  function returns silently without playing the default sound.

'- SND_LOOP specifies that the sound will continue to play continuously
'  until PlaySound is called again with the lpszSoundName$ parameter
'  set to null. You must also specify the SND_ASYNC flag to loop sounds.

'- SND_NOSTOP specifies that if a sound is currently playing, the
'  function will immediately return False without playing the requested
'  sound.

'_ SND_PURGE Stop playback

' The PlaySound function returns True (-1) if the sound is played,
'  otherwise it returns False (0).

Global stopper


Public Sub FillBITStructures()
'Fill BITDATA structure
bmp.bmWidth = 512          'Pixel width of draw surfaces
bmp.bmHeight = 512         'Pixel height of draw surfaces
bmp.ptrBackSurf = VarPtr(BackSurf(1, 1))   'Pointer to back surface
bmp.ptrFrontSurf = VarPtr(FrontSurf(1, 1)) 'Pointer to front surface
'.....

'Fill BITMAPINFO.BITMAPINFOHEADER FOR StretchDIBits
bm.bmiH.biSize = 40
bm.bmiH.biwidth = 512  'Size of the source surface
bm.bmiH.biheight = 512
bm.bmiH.biPlanes = 1
bm.bmiH.biBitCount = 8
bm.bmiH.biCompression = 0
bm.bmiH.biSizeImage = 0  ' not needed here
bm.bmiH.biXPelsPerMeter = 0
bm.bmiH.biYPelsPerMeter = 0
bm.bmiH.biClrUsed = 0
bm.bmiH.biClrImportant = 0

End Sub

Public Sub LoadCar(BMPSpec$, N)
On Error GoTo CarError
Open BMPSpec$ For Input As #2
Line Input #2, a$
a$ = ""
Close
Open BMPSpec$ For Binary As #2

'LOAD CARS
Seek #2, 1079  'start of picture color indices
For iy = 1 To 48
For ix = 1 To 48
   Select Case N
   Case 1: Get #2, , LCar(ix, iy)
   Case 2: Get #2, , HCar(ix, iy)
   Case 3: Get #2, , RCar(ix, iy)
   End Select
Next ix
Next iy
Close
On Error GoTo 0

Exit Sub
'==========
CarError:
   res = MsgBox(BMPSpec$ & " missing")
   Close
   End
End Sub


Public Sub LoadMainTree(BMPSpec$, TreeType)

Dim RED As Byte, GREEN As Byte, BLUE As Byte
Dim bpp As Byte

'Test if file exists without creating
'a 0 size binary file
'BMPSpec$ = PathSpec$ & "RRTree.bmp"
On Error GoTo FileError
Open BMPSpec$ For Input As #2
Line Input #2, a$
a$ = ""
Close

Open BMPSpec$ For Binary As #2
'CHECK ITS 8-BIT BMP
Seek #2, 29    'bits per pixel
Get #2, , bpp
If bpp <> 8 Then
   res = MsgBox("Not a 256 color 8-bit BMP")
   Close
   End
End If
   
   'READ PALETTE
   Seek #2, 55 'Start of palette bytes
   For N = 0 To 255
      B$ = Input$(1, 2): g$ = Input$(1, 2): r$ = Input$(1, 2)
      If B$ <> "" Then BLUE = Asc(B$) Else BLUE = 0
      If g$ <> "" Then GREEN = Asc(g$) Else GREEN = 0
      If r$ <> "" Then RED = Asc(r$) Else RED = 0
      bm.Colors(N).rgbBlue = BLUE
      bm.Colors(N).rgbGreen = GREEN
      bm.Colors(N).rgbRed = RED
      bm.Colors(N).rgbReserved = 0
      d$ = Input$(1, 2)
   Next N
   
'LOAD PICTURE TO LARGEST Tree Array
Seek #2, 1079  'start of picture color indices
For iy = 1 To 512
For ix = 1 To 128
   Get #2, , Tree19(ix, iy, TreeType)
Next ix
Next iy
Close
On Error GoTo 0
Exit Sub
'==========
FileError:
   res = MsgBox(BMPSpec$ & " missing")
   Close
   End
End Sub

Public Sub DevelopOtherTrees(TreeType)

   'This uses StretchDIBits to compress BMP to smaller
   'sizes on screen (NB Does a fair job of this)
   '& store these in the Tree arrays

   'Transfer large tree to FrontSurf
   For iy = 1 To 512
   For ix = 1 To 128
      FrontSurf(ix, iy) = Tree19(ix, iy, TreeType)
   Next ix
   Next iy

   FIW& = TreeW(NTreeSizes): FIH& = TreeH(NTreeSizes)  'StretchDIBits dest size
   ShowFrontSurfOnce       'Show original BMP

   '1 sec delay to see large tree
   T! = Timer
   Do: Loop Until (Timer - T!) > 1

   'Taking FrontSurf as full size source show compressed images
   'Compressed to FIW&, FIH& depending on TreeType
   For TreeSize = NTreeSizes - 1 To 1 Step -1
      FIW& = TreeW(TreeSize): FIH& = TreeH(TreeSize)
      ShowFrontSurfOnce 'Compress large BMP to FIW&,FIH&
      DoEvents
      CopyPic TreeSize, TreeType 'Read colors off screen for compressed image
                              '& store in tree arrays
   Next TreeSize

End Sub

Public Sub ShowFrontSurfOnce()
   
   'Size of FrontSurf
   bm.bmiH.biwidth = 512&
   bm.bmiH.biheight = 512&

   succ& = StretchDIBits(Form1.hdc, _
   0, 0, _
   FIW&, FIH&, _
   0, 0, _
   128&, 512&, _
   ByVal bmp.ptrFrontSurf, bm, _
   DIB_RGB_COLORS, SRCCOPY)
   
End Sub

Public Sub CopyPic(TreeSize, TreeType)
'AMENDED

'Maps large tree color numbers to smaller tree arrays
'No averaging done but seems adequate

For iy = 0 To FIH& - 1
For ix = 0 To FIW& - 1
      'Get source color number
      nxs = 1 + (ix - 1) * (127 / FIW&)
      nys = 1 + (iy - 1) * (511 / FIH&)
      If nxs >= 1 And nxs <= 128 And nys >= 1 And nys <= 512 Then
         N = Tree19(nxs, nys, TreeType)
      Else
         N = 255
      End If
         
      Select Case TreeSize
      Case 18: Tree18(ix + 1, iy + 1, TreeType) = N
      Case 17: Tree17(ix + 1, iy + 1, TreeType) = N
      Case 16: Tree16(ix + 1, iy + 1, TreeType) = N
      Case 15: Tree15(ix + 1, iy + 1, TreeType) = N
      Case 14: Tree14(ix + 1, iy + 1, TreeType) = N
      Case 13: Tree13(ix + 1, iy + 1, TreeType) = N
      Case 12: Tree12(ix + 1, iy + 1, TreeType) = N
      Case 11: Tree11(ix + 1, iy + 1, TreeType) = N
      Case 10: Tree10(ix + 1, iy + 1, TreeType) = N
         
      Case 9: Tree9(ix + 1, iy + 1, TreeType) = N  'Upside down
      Case 8: Tree8(ix + 1, iy + 1, TreeType) = N
      Case 7: Tree7(ix + 1, iy + 1, TreeType) = N
      Case 6: Tree6(ix + 1, iy + 1, TreeType) = N
      Case 5: Tree5(ix + 1, iy + 1, TreeType) = N
      Case 4: Tree4(ix + 1, iy + 1, TreeType) = N
      Case 3: Tree3(ix + 1, iy + 1, TreeType) = N
      Case 2: Tree2(ix + 1, iy + 1, TreeType) = N
      Case 1: Tree1(ix + 1, iy + 1, TreeType) = N
      End Select
   
Next ix
Next iy
End Sub

Public Sub WhitenBackSurface()
'Make BackSurf white & whiten screen

For iy = 1 To 512
For ix = 1 To 512
   BackSurf(ix, iy) = 255
Next ix
Next iy

FormWidth& = Form1.Width \ Screen.TwipsPerPixelX
FormHeight& = Form1.Height \ Screen.TwipsPerPixelY
   
bm.bmiH.biwidth = 512&
bm.bmiH.biheight = 512&

'Stretch byte-array to Form
'NB The ByVal is critical in this! Otherwise big memory leak!

   succ& = StretchDIBits(Form1.hdc, _
   0, 0, _
   FormWidth& - 8, FormHeight& - 8, _
   0, 0, _
   512&, 512&, _
   ByVal bmp.ptrBackSurf, bm, _
   DIB_RGB_COLORS, SRCCOPY)
   
End Sub

Public Sub SetupTreeArrays()

'NTreeStarts = 1  'Number of tree starts on each loop
                  'can be changed in Form_Load
NTreeSizes = 19  'Number of tree sizes
NTreeTypes = 3   'Number of tree types
ReDim TreeW(NTreeSizes), TreeH(NTreeSizes)  'Tree surfaces' width & height
TreeW(1) = 8:   TreeH(1) = 32
TreeW(2) = 12:  TreeH(2) = 40
TreeW(3) = 16:  TreeH(3) = 48
TreeW(4) = 20:  TreeH(4) = 56
TreeW(5) = 24:  TreeH(5) = 64
TreeW(6) = 28:  TreeH(6) = 72
TreeW(7) = 32:  TreeH(7) = 88
TreeW(8) = 36:  TreeH(8) = 104
TreeW(9) = 40:  TreeH(9) = 120
TreeW(10) = 44: TreeH(10) = 136
TreeW(11) = 48: TreeH(11) = 152
TreeW(12) = 56: TreeH(12) = 168
TreeW(13) = 64: TreeH(13) = 200
TreeW(14) = 72: TreeH(14) = 232
TreeW(15) = 80: TreeH(15) = 264
TreeW(16) = 88: TreeH(16) = 328
TreeW(17) = 96: TreeH(17) = 392
TreeW(18) = 112: TreeH(18) = 456
TreeW(19) = 128: TreeH(19) = 512

'Trees, Global bytes
ReDim Tree1(TreeW(1), TreeH(1), NTreeTypes)
ReDim Tree2(TreeW(2), TreeH(2), NTreeTypes)
ReDim Tree3(TreeW(3), TreeH(3), NTreeTypes)
ReDim Tree4(TreeW(4), TreeH(4), NTreeTypes)
ReDim Tree5(TreeW(5), TreeH(5), NTreeTypes)
ReDim Tree6(TreeW(6), TreeH(6), NTreeTypes)
ReDim Tree7(TreeW(7), TreeH(7), NTreeTypes)
ReDim Tree8(TreeW(8), TreeH(8), NTreeTypes)
ReDim Tree9(TreeW(9), TreeH(9), NTreeTypes)
ReDim Tree10(TreeW(10), TreeH(10), NTreeTypes)
ReDim Tree11(TreeW(11), TreeH(11), NTreeTypes)
ReDim Tree12(TreeW(12), TreeH(12), NTreeTypes)
ReDim Tree13(TreeW(13), TreeH(13), NTreeTypes)
ReDim Tree14(TreeW(14), TreeH(14), NTreeTypes)
ReDim Tree15(TreeW(15), TreeH(15), NTreeTypes)
ReDim Tree16(TreeW(16), TreeH(16), NTreeTypes)
ReDim Tree17(TreeW(17), TreeH(17), NTreeTypes)
ReDim Tree18(TreeW(18), TreeH(18), NTreeTypes)
ReDim Tree19(TreeW(19), TreeH(19), NTreeTypes)

End Sub

