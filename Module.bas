Attribute VB_Name = "Module"
'*******IDENTIFICATION BLOCK*******
'
'PROGRAMMERS: ZAC ANDREW, KEVIN LAWRENCE, STEVEN PEARCE
'PROJECT NAME: HTMLVILLE
'COURSE CODE: TIK 200-C
'TEACHER: MRS. SINTZEL
'ROOM: 208
'
'*******SPECIFICATION*******
'
'THIS MODULE IS USED FOR VARIOUS PARTS OF THE FORMS, SUCH AS PRINTING,
'ADDING ON TO QUIZ VARIABLE VALUES, ETC.
'
'*******PROGRAM CODE*******

Option Explicit

Public file As String
Type question
    intQno As String * 10 'ADD ONTO QUIZ VARIABLES
    ques As String * 80
    ans As Integer
    ans1 As String * 80
    ans2 As String * 80
    ans3 As String * 80
    ans4 As String * 80
End Type

Public fMainForm As frmEditor
   Private Type Rect
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
   End Type

   Private Type CharRange
     cpMin As Long 'CHARACTER RANGE MIN
     cpMax As Long 'CHARACTER RANGE MAX
   End Type

   Private Type FormatRange
     hdc As Long
     hdcTarget As Long
     rc As Rect
     rcPage As Rect
     chrg As CharRange
   End Type

   Private Const WM_USER As Long = &H400
   Private Const EM_FORMATRANGE As Long = WM_USER + 57
   Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
   Private Const PHYSICALOFFSETX As Long = 112
   Private Const PHYSICALOFFSETY As Long = 113

   Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hdc As Long, ByVal nIndex As Long) As Long
   Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" _
      (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
      lp As Any) As Long
   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
      (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
      ByVal lpOutput As Long, ByVal lpInitData As Long) As Long


'PRINT CODE FOR EDITOR
   Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, _
      TopMarginHeight, RightMarginWidth, BottomMarginHeight)
      Dim LeftOffset As Long, TopOffset As Long
      Dim LeftMargin As Long, TopMargin As Long
      Dim RightMargin As Long, BottomMargin As Long
      Dim fr As FormatRange
      Dim rcDrawTo As Rect
      Dim rcPage As Rect
      Dim TextLength As Long
      Dim NextCharPosition As Long
      Dim r As Long

      'START PRINT JOB
      Printer.Print Space(1)
      Printer.ScaleMode = vbTwips

      'GET PRINT OFFSET
      LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETX), vbPixels, vbTwips)
      TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
         PHYSICALOFFSETY), vbPixels, vbTwips)

      'GET MARGINS
      LeftMargin = LeftMarginWidth - LeftOffset
      TopMargin = TopMarginHeight - TopOffset
      RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
      BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset

      'GET PRINTABLE AREA
      rcPage.Left = 0
      rcPage.Top = 0
      rcPage.Right = Printer.ScaleWidth
      rcPage.Bottom = Printer.ScaleHeight

      'GET MARGINS
      rcDrawTo.Left = LeftMargin
      rcDrawTo.Top = TopMargin
      rcDrawTo.Right = RightMargin
      rcDrawTo.Bottom = BottomMargin

      'PRINT INSTRUCTIONS
      fr.hdc = Printer.hdc
      fr.hdcTarget = Printer.hdc  'GO TO PRINTER
      fr.rc = rcDrawTo 'GET AREA ON PAGE
      fr.rcPage = rcPage 'GET SIZE OF PAGE
      fr.chrg.cpMin = 0 'GET START OF TEXT
      fr.chrg.cpMax = -1 'END OF THE TEXT

      TextLength = Len(RTF.Text) 'GET LENGTH OF TEXT IN RTF

      Do 'LOOP PRINTING UNTIL EACH PAGE IS DONE
         NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr) 'PRINT PAGE
         If NextCharPosition >= TextLength Then Exit Do   'If done then exit
         fr.chrg.cpMin = NextCharPosition ' START POSTITION FOR NEXT PAGE
         Printer.NewPage 'GO TO NEXT PAGE
         Printer.Print Space(1)
         fr.hdc = Printer.hdc
         fr.hdcTarget = Printer.hdc
      Loop

      Printer.EndDoc 'COMMIT THE PRINTING

      r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0)) 'ALLOW RTF TO CLEAR UP MEMORY
   End Sub
