VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Map"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5640
      Top             =   2760
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6600
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Txt3 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Txt2 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Txt1 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Neuk"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   5
      Top             =   720
      Width           =   6255
   End
   Begin VB.ComboBox CboFonts 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Txtcopy 
      Height          =   375
      HideSelection   =   0   'False
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   6480
      Max             =   100
      Min             =   1
      TabIndex        =   20
      Top             =   2040
      Value           =   50
      Width           =   135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "&Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Char&acters to copy:"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Dec:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Bin:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   3465
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'******************************************************************************
'*Character Map recreation -17/jul/2000
'******************************************************************************
'By oigres P Email:oigres@postmaster.co.uk
Private Type POINTAPI  '  8 Bytes
    X As Long
    Y As Long
End Type
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As _
        Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As _
        Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
        ByVal Y As Long, lpPoint As POINTAPI)
Private Declare Function CreateRectRgnIndirect& Lib "gdi32" (lprect As RECT)
Private Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As _
        Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Private Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal _
        Y As Long)
Private Declare Function Rectangle& Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Const HORZRES = 8
Private Const VERTRES = 10
Const SRCCOPY = &HCC0020

Dim asciiList() ' list of character descriptions
Dim sizeX, sizeY, previousX, previousY
Dim mouseDown As Boolean, mouseVisible As Boolean

Private Sub CboFonts_Click()

    drawSquare CboFonts.List(CboFonts.ListIndex)
    Picture2.Font = CboFonts.List(CboFonts.ListIndex)
    Picture2.FontSize = 18
    Txtcopy.Font = CboFonts.List(CboFonts.ListIndex)

    'reselect last square
    drawfocusColour previousX, previousY


End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    'copy to clipboard
    Clipboard.Clear
    Clipboard.SetText Txtcopy.Text, vbCFText
    Picture1.SetFocus
End Sub

Private Sub cmdSelect_Click()
    '
    inserttext
    Picture1.SetFocus
End Sub
Sub inserttext()
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety, s
    s = selectedsquare
    Y1 = s \ 32
    X1 = s Mod 32
    char$ = Chr$((Y1 * 32) + (X1 + 1) + 30) '1)
    Txtcopy.SelText = char$

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ''MsgBox "keydon " & KeyCode & ":" & ActiveControl
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        MsgBox "alt+ A=frm key"
        'Txtcopy.SelText =
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If
    If KeyCode = Asc("F") And (Shift And vbAltMask) Then
        'MsgBox "alt+ A"
        'Txtcopy.SelText =
        CboFonts.SetFocus
    End If
    If KeyCode = Asc("S") And (Shift And vbAltMask) Then
        MsgBox "alt+ S"
        'Txtcopy.SelText =
        'CboFonts.SetFocus
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        ''MsgBox "alt+ A=frm key"
        'Txtcopy.SelText =
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
        Txtcopy.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim X, Y
    Dim index
    '' Form1.ScaleWidth = 32 * 7


    sizeX = (Picture1.ScaleWidth \ 32) ' + 1 '  32*7=224
    sizeY = (Picture1.ScaleHeight \ 7) ''' + 1 '  32*7=224
    '''MsgBox sizeX & ":" & sizeY
    createAsciiList
    '    Form1.ForeColor = vbBlack
    '    Form1.Picture = LoadPicture()

    ''Form1.Refresh

    ''Form1.AutoRedraw = True
    index = 32
    drawSquare "Times New Roman"


    CboFonts.Visible = False
    FillListWithFonts CboFonts 'List1
    CboFonts.ListIndex = 0
    CboFonts.Visible = True

    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False

    'previousX = 0 'start off with the first square
    'previousY = 0
    'starts off with first square selected
    '''drawfocusColour 0, 0
    '''updateLabel 0, 0
    Picture1_MouseDown 0&, 0&, 0, 0
    Picture1_MouseUp 0&, 0&, 0, 0
    Form1.Show
    Picture1.SetFocus
    cmdCopy.Enabled = False

    selectedsquare = 1
End Sub

'''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub updateLabel(X, Y)
    Dim key, k$
    'give keystroke and alt information

    key = (Y * 32) + (X + 1) ' + 31
    k$ = "Keystroke: "
    'MsgBox key
    Select Case key
    Case 1
        Label4.Caption = k$ & "Spacebar"
    Case 2 To 95 '
        If key = 7 Then 'need && to show in label for ampersand
            Label4.Caption = k$ & "&&" 'Chr$(key + 31)
        Else
            Label4.Caption = k$ & Chr$(key + 31)
        End If

    Case 96 To 97
        Label4.Caption = k$ & "Ctrl+" & (key - 95)
    Case 98 To 224
        Label4.Caption = k$ & "Alt+0" & key + 31
    End Select
    'hex / bin text
    Txt1.Text = Hex(key + 31)
    Txt2.Text = Bin(key + 31, 8)
    Txt3.Text = key + 31
    
    Label1.Caption = "Col: " & X & " Line: " & Y & " Square:" & (Y * 32) + (X + 1) & " Ascii: " & key + 31 ' * (y1 + 1)
    'Debug.Print key
    'asciilist array starts at 0 index
    Select Case key
    Case 1 To 98
        Label7.Caption = asciiList(key - 1)
    Case 99 To 129
        Label7.Caption = asciiList(key - 1)
    Case 130 To 224
        Label7.Caption = asciiList(key - 1)

    End Select
End Sub
Sub createAsciiList()
    ReDim asciiList(250)

    Dim a$, index As Long
    Open App.Path & "\asciiquoteds.txt" For Input As 1
    Do While Not (EOF(1))
        Input #1, a$
        asciiList(index) = a$
        index = index + 1
    Loop


    Close 1
    'For index = 0 To 255 - 31
    'MsgBox asciiList(index)
    'Next index
End Sub

Private Sub Picture1_DblClick()
    inserttext
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyDown
        If selectedsquare + 32 < 225 Then
            selectedsquare = selectedsquare + 32
        End If
    Case vbKeyUp
        If selectedsquare - 32 > 0 Then
            selectedsquare = selectedsquare - 32
        End If
    Case vbKeyRight
        If selectedsquare + 1 < 225 Then
            selectedsquare = selectedsquare + 1
        End If
    Case vbKeyLeft
        If selectedsquare - 1 > 0 Then
            selectedsquare = selectedsquare - 1
        End If
    Case Else
        Exit Sub
    End Select
    drawselected (selectedsquare - 1)
    updateLabel (selectedsquare - 1) Mod 32, (selectedsquare - 1) \ 32

End Sub
'/******************************************************************************
Sub drawselected(s As Long)
    '/******************************************************************************
    Dim X1, Y1, char$, lprect As RECT, offsetx, offsety
    Y1 = s \ 32
    X1 = s Mod 32
    'erase previous ?
    Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
    Picture1.CurrentX = (previousX * sizeX) + 3
    Picture1.CurrentY = (previousY * sizeY)

    Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
    previousX = X1
    previousY = Y1

    char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = False: Picture3.Visible = False
    offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
    offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
    Picture2.left = (X1 * sizeX - 5) + 10
    Picture2.top = (Y1 * sizeY - 5) + 35
    Picture3.left = Picture2.left + 5
    Picture3.top = Picture2.top + 5
    Picture2.CurrentX = offsetx
    Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
    Picture2.Picture = LoadPicture()
    Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
    Picture2.Visible = True: Picture3.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseDown
    '*
    '* Description:
    '*
    '* Date Created:  7/17/00
    '*
    '* Created By: oigres P
    '*
    '* Modified: 7/19/00
    '*
    '*******************************************************************************
    Dim X1, Y1, ret, lprect As RECT, offsetx, offsety, char$
    X1 = X \ sizeX
    Y1 = Y \ sizeY
    If Button = vbRightButton Then
        Exit Sub
    End If
    'if in square of picture
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
        ''If x1 <> previousX And y <> previousY Then
        'erase previous focus rectangle
        ''MsgBox IsEmpty(previousX)
        If Not (IsEmpty(previousX) And IsEmpty(previousY)) Then
            lprect.left = X1 * sizeX + 1
            lprect.top = Y1 * sizeY + 1
            lprect.right = X1 * sizeX + (sizeX - 1) + 1 '- 1
            lprect.bottom = Y1 * sizeY + (sizeY - 1) + 1
            ''DrawFocusRect Picture1.hdc, lprect

            Picture1.Line (previousX * sizeX, previousY * sizeY)-(previousX * sizeX + (sizeX), previousY * sizeY + (sizeY)), vbBlack, BF
            Picture1.Line (previousX * sizeX + 1, previousY * sizeY + 1)-(previousX * sizeX + (sizeX - 1), previousY * sizeY + (sizeY - 1)), vbWhite, BF
            '''''''''
            ''Picture1.CurrentX = (previousX * sizeX) + 3
            ''Picture1.CurrentY = (previousY * sizeY)
            char$ = Chr$((previousY * 32) + (previousX + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (previousX * sizeX) + offsetx
            Picture1.CurrentY = (previousY * sizeY) + offsety
            Picture1.Print char$;
            '''Picture1.Print Chr$((previousY * 32) + (previousX + 1) + 31);
        End If
        Picture2.Visible = False
        Picture3.Visible = False
        Picture2.left = (X1 * sizeX - 5) + 10
        Picture2.top = (Y1 * sizeY - 5) + 35
        Picture3.left = Picture2.left + 5
        Picture3.top = Picture2.top + 5
        Picture2.Visible = True
        Picture3.Visible = True
        selectedsquare = (Y1 * 32) + (X1 + 1)

        previousX = X1
        previousY = Y1
    End If ' in square

    'draw focus rectangle

    Call updateLabel(X1, Y1)
    'hide cursor
    If mouseDown = False Then
        ret = ShowCursor(False)
        'showcursor shows cursor if the return count >=0
        'force it to hide
        While ret >= 0
            ret = ShowCursor(False)
        Wend
        '    If ret >= 0 Then
        '
        '    End If
        mouseVisible = False
        '' Label5.Caption = "showcursor times= " & ret
    End If
    mouseVisible = False
    ''Form1.MousePointer = 15
    Picture2.Visible = True
    Picture3.Visible = True
    mouseDown = True
End Sub

'/******************************************************************************
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseMove
    '*
    '* Description:
    '*
    '* Date Created:  7/21/00
    '*
    '* Created By:
    '*
    '* Modified:
    '*
    '*******************************************************************************

    Dim X1, Y1, ret, char$, key
    Dim offsetx, offsety
    Static lastx
    Static lasty

    If mouseDown = True Then

        X1 = X \ sizeX
        Y1 = Y \ sizeY
        If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then


            If mouseVisible = True Then
                makeCursorInvisible
            End If
            If lastx = X1 And lasty = Y1 Then Exit Sub
            lastx = X1: lasty = Y1
            key = (Y1 * 32) + (X1 + 1)

            Picture2.Visible = False
            Picture3.Visible = False
            Picture2.left = (X1 * sizeX - 5) + 10
            Picture2.top = (Y1 * sizeY - 5) + 35
            Picture3.left = Picture2.left + 5
            Picture3.top = Picture2.top + 5
            '            Picture2.Visible = True
            '            Picture3.Visible = True
            char$ = Chr$((Y1 * 32) + (X1 + 1) + 31)
            If Picture2.Tag = char$ Then
            Else

                '        Picture1.Picture = LoadPicture()
                '        Picture1.CurrentX = 0: Picture1.CurrentY = 0
                '        Picture1.Print Chr$((y1 * 32) + (x1 + 1) + 31)
                previousX = X1
                previousY = Y1
                ''Picture2.Visible = False
                Picture2.Tag = char$


                offsetx = (Picture2.ScaleWidth - Picture2.TextWidth(char$)) \ 2
                offsety = (Picture2.ScaleHeight - Picture2.TextHeight(char$)) \ 2
                Picture2.CurrentX = offsetx
                Picture2.CurrentY = offsety '    Chr$((y1 * 32) + (x1 + 1) + 31)
                Picture2.Picture = LoadPicture()
                Picture2.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
                Picture2.Visible = True
                Picture3.Visible = True
            End If 'if tag

            Call updateLabel(X1, Y1)
            previousX = X1
            previousY = Y1
        Else 'not in square
            'showcursor ?
            makeCursorVisible


            Exit Sub
        End If ' x1 >= 0 And x1 <= 31 And y1 >= 0 And y1 <= 6
        ''ShowCursor False
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*******************************************************************************
    '* Name:  Picture1_MouseUp
    '*
    '* Description:
    '*
    '* Date Created:  7/21/00
    '*
    '* Created By:
    '*
    '* Modified:
    '*
    '*******************************************************************************
    Dim ret, X1, Y1, lprect As RECT
    X1 = X \ sizeX
    Y1 = Y \ sizeY
    If mouseVisible = False Then
        ret = ShowCursor(True)
        While ret < 0
            ret = ShowCursor(True)
        Wend
        mouseVisible = True
    End If
    drawfocusColour previousX, previousY
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then

    Else
        If mouseDown = True Then
            Picture2.Visible = False
            Picture3.Visible = False
            'draw red focus rectangle
            drawfocusColour previousX, previousY
        End If

    End If
    Picture2.Visible = False
    Picture3.Visible = False
    mouseDown = False
End Sub

'/******************************************************************************
Sub makeCursorInvisible()
    '/******************************************************************************
    Dim ret
    ret = ShowCursor(False)
    While ret >= 0
        ret = ShowCursor(False)
    Wend
    mouseVisible = False
End Sub
'/******************************************************************************
Sub makeCursorVisible()
    '/******************************************************************************
    Dim ret
    ret = ShowCursor(True)
    While ret < 0
        ret = ShowCursor(True)
    Wend
    mouseVisible = True
End Sub

'/******************************************************************************
Sub drawSquare(f As String)
    '/******************************************************************************
    Dim X As Long, Y As Long, char$, lpPT As POINTAPI
    Dim offsetx, offsety
    Picture1.Visible = False
    Picture1.FontName = f
    Picture1.FontSize = 8
    Picture1.Picture = LoadPicture()
    For X = 0 To 31 '32
        For Y = 0 To 6 '7
            ''Picture1.Line (x * sizex, y * sizey)-(x * sizex + (sizey - 1), y * sizex + (sizey - 1)), vbBlack, B
            char$ = Chr$((Y * 32) + (X + 1) + 31)
            offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
            offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
            Picture1.CurrentX = (X * sizeX) + offsetx
            Picture1.CurrentY = (Y * sizeY) + offsety
            Picture1.Print char$;

        Next Y
    Next X
    For X = 0 To 7
        MoveToEx Picture1.hdc, 0, X * sizeY, lpPT
        LineTo Picture1.hdc, sizeX * 32, X * sizeY
    Next X
    For X = 0 To 32
        MoveToEx Picture1.hdc, X * sizeX, 0, lpPT
        LineTo Picture1.hdc, X * sizeX, sizeY * 7 + 1 'Picture1.ScaleHeight - 1
    Next X
    Picture1.Visible = True
End Sub


Private Sub Timer1_Timer()
    Dim cp As POINTAPI, hr As Long, vr As Long, ret As Long
    GetCursorPos cp
    ''Label1.Caption = cp.x & Space(6 - Len(CStr(cp.x))) & ":" & cp.y

    Dim dsDC As Long, lpPT As POINTAPI, dshwnd As Long, Percent
    Dim lengthx, lengthy, offsetx, offsety, blitareax, blitareay
    'get desktop device context
    dsDC = GetDC(0&)
    'get screen width, height
    hr = GetDeviceCaps(dsDC, HORZRES)
    vr = GetDeviceCaps(dsDC, VERTRES)

    dshwnd = GetDesktopWindow()
    '      vscroll1=1..100 so 1/100=.1; 100/100=1;New Resolution
    Percent = VScroll1.value / 100
    lengthx = (Picture4.ScaleWidth - 0) * Percent
    lengthy = (Picture4.ScaleHeight - 0) * Percent
    'center image about mouse
    offsetx = lengthx \ 2
    offsety = lengthy \ 2
    blitareax = Picture4.ScaleWidth - 0 'actual area to blit to
    blitareay = Picture4.ScaleHeight - 0
    'Debug.Print lengthx; lengthy; Percent; offsetx; offsety
    'stop copying the screen off the edges <0 and  >horzres
    If cp.X - offsetx >= 0 And cp.X + offsetx < hr Then '800=screen width
        If cp.Y - offsety >= 0 And cp.Y + offsety < vr Then '600= screen height

            '                dest hdc ,destx,desty,width,height, sourceDC, source x,sourcey,sourcewidth,sourceheight,raster operation
            ret = StretchBlt(Picture4.hdc, 0, 0, blitareax, blitareay, dsDC, cp.X - offsetx, cp.Y - offsety, lengthx, lengthy, SRCCOPY)
        End If
    End If
    'Form1.Line (0, 0)-(Form1.ScaleWidth - VScroll1.Width, Form1.ScaleHeight - Label1.Height)
    'Form1.Line (Form1.ScaleWidth - VScroll1.Width, 0)-(0, Form1.ScaleHeight - Label1.Height)
    ReleaseDC dshwnd, dsDC 'previous bug not releasing memory
    Label5.Caption = Format(100 / VScroll1.value, "FIXED") & ":" & cp.X & ":" & cp.Y
End Sub

'/******************************************************************************
Private Sub Txtcopy_Change()
    '/******************************************************************************
    If Txtcopy.Text = "" Then
        cmdCopy.Enabled = False
    Else

        cmdCopy.Enabled = True
    End If
End Sub
'/******************************************************************************
Sub drawfocusColour(X, Y)
    Dim lprect As RECT, offsetx, offsety, char$
    Picture1.Line (X * sizeX + 1, Y * sizeY + 1)-(X * sizeX + (sizeX - 1), _
            Y * sizeY + (sizeY - 1)), vbHighlight, BF
    ''Picture1.FillColor = vbHighlight
    'Rectangle Picture1.hdc, x * sizeX + 1, y * sizeY + 1, x * sizeX + (sizeX), y * sizeY + (sizeY)
    ''Picture1.FillColor = vbWhite
    'Picture1.CurrentX = (x * sizeX) + 3
    'Picture1.CurrentY = (y * sizeY)
    '
    char$ = Chr$((Y * 32) + (X + 1) + 31)
    offsetx = (sizeX - Picture1.TextWidth(char$)) \ 2
    offsety = (sizeY - Picture1.TextHeight(char$)) \ 2
    Picture1.CurrentX = (X * sizeX) + offsetx
    Picture1.CurrentY = (Y * sizeY) + offsety

    '
    Picture1.ForeColor = vbWhite
    'Picture1.Print Chr$((y * 32) + (x + 1) + 31);
    Picture1.Print char$;
    Picture1.ForeColor = vbBlack
    ''previousX = x
    ''previousY = y
    lprect.left = X * sizeX + 1
    lprect.top = Y * sizeY + 1
    lprect.right = X * sizeX + (sizeX - 1) + 1 '- 1
    lprect.bottom = Y * sizeY + (sizeY - 1) + 1  '- 1

    DrawFocusRect Picture1.hdc, lprect
End Sub

'/******************************************************************************
Private Sub Txtcopy_GotFocus()
    '/******************************************************************************
    '    ''Debug.Print "1   tgoptfcs"

    '        'MsgBox "gfocus"
    '        Txtcopy.SelStart = 0
    '        Txtcopy.SelLength = Len(Txtcopy.Text)
    '    End If

End Sub

Private Sub Txtcopy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc("A") And (Shift And vbAltMask) Then
        MsgBox "tkdwn"
        Txtcopy.SelStart = 0
        Txtcopy.SelLength = Len(Txtcopy.Text)
    End If
End Sub

Private Sub Txtcopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '

End Sub

