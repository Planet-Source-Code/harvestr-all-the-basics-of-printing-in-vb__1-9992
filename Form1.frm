VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   420
      Left            =   1350
      ScaleHeight     =   360
      ScaleWidth      =   1950
      TabIndex        =   14
      Top             =   4290
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1650
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   12
      Top             =   390
      Width           =   1350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3405
      TabIndex        =   1
      Top             =   4305
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   4305
      Width           =   1185
   End
   Begin VB.Label Label11 
      Caption         =   "- print pictures"
      Height          =   300
      Left            =   750
      TabIndex        =   13
      Top             =   3660
      Width           =   3630
   End
   Begin VB.Line Line6 
      X1              =   180
      X2              =   1365
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label10 
      Caption         =   "- retrieve user's printer name"
      Height          =   300
      Left            =   750
      TabIndex        =   11
      Top             =   3420
      Width           =   3630
   End
   Begin VB.Label Label9 
      Caption         =   "- draw tables"
      Height          =   300
      Left            =   750
      TabIndex        =   10
      Top             =   3180
      Width           =   3630
   End
   Begin VB.Label Label8 
      Caption         =   "- use different fonts, font sizes and attributes"
      Height          =   300
      Left            =   750
      TabIndex        =   9
      Top             =   2955
      Width           =   3630
   End
   Begin VB.Label Label7 
      Caption         =   "- define the printable area"
      Height          =   300
      Left            =   750
      TabIndex        =   8
      Top             =   2715
      Width           =   3630
   End
   Begin VB.Label Label6 
      Caption         =   "- center text, align left, align right"
      Height          =   300
      Left            =   750
      TabIndex        =   7
      Top             =   2475
      Width           =   3630
   End
   Begin VB.Label Label5 
      Caption         =   "In this code sample, you'll find how to:"
      Height          =   270
      Left            =   285
      TabIndex        =   6
      Top             =   2195
      Width           =   4035
   End
   Begin VB.Label Label4 
      Caption         =   "Click the Print button to see what is printed, then you will be able to follow the comments in the code more easily."
      Height          =   450
      Left            =   300
      TabIndex        =   5
      Top             =   1680
      Width           =   4035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INSTRUCTIONS"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1365
      TabIndex        =   4
      Top             =   1320
      Width           =   1965
   End
   Begin VB.Line Line5 
      X1              =   173
      X2              =   4500
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line4 
      X1              =   4500
      X2              =   4500
      Y1              =   1455
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   3315
      X2              =   4530
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   180
      Y1              =   1455
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   1440
      Y2              =   1470
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "http://www.harvestr.org"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1073
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The Basics of Printing in VB6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   2
      Top             =   120
      Width           =   3660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'To better view this code and its comments, be sure you're at least in
'800*600 screen resolution, and expand this wondow to its max.

'All the variables defined in here are used and explained later in the code
Dim Answer As String, HorizontalMargin, VerticalMargin As Single
Dim MyCenteredText As String, MyCenteredTextWidth As Single
Dim MyLeftText As String, MyLeftTextWidth As Single
Dim MyRightText As String, MyRightTextWidth As Single
Dim txtGrid11, txtGrid12, txtGrid13, txtGrid14, txtGrid21, txtGrid22 As String
Dim txtGrid23, txtGrid24, txtGrid31, txtGrid32, txtGrid33, txtGrid34 As String
Dim MyGridTitle As String, MyGridTitleWidth As Single
Dim Row1Col1Left, Row1Col2Left, Row1Col3Left, Row1Col4Left As Single
Dim Row2Col1Left, Row2Col2Left, Row2Col3Left, Row2Col4Left As Single
Dim Row3Col1Left, Row3Col2Left, Row3Col3Left, Row3Col4Left As Single
Dim Row1Top, Row2Top, Row3Top, ImageLeft, ImageTop As Single

'We send a message for the User to confirm he really wants to print the doc
'we will also use this messagebox to display his Printer name
Answer = MsgBox("confirm printing on " & Printer.DeviceName, vbYesNo)
If Answer = vbNo Then Exit Sub

'We decide to measure in centimeters
Printer.ScaleMode = vbCentimeters

'We use the A4 format paper (21 * 29.7 centimeters = 8.5 * 11 inches)
'We check the physical borders of the Printer
HorizontalMargin = (21 - Printer.ScaleWidth) / 2
VerticalMargin = (29.7 - Printer.ScaleHeight) / 2

'It is highly recommended that you set a margin of 1 cm horizontally and
'1.5 cm vertically in your printable document, in addition to the physical
'margin, because the previous check isn't well working
HorizontalMargin = 1 + HorizontalMargin
VerticalMargin = 1.5 + VerticalMargin

'initialize the printer
Printer.Print "";

'We'll now draw a red line all aroound our printable area. We use the paper size
'(21*29.7) and we add the margins to the starting point (upper left). We'll
'substract these margins to the ending point (lower right).
'       NOTE: the syntax is: Printer.Line (X1,Y1)-(X2,Y2),color,flag
'             where flag can be: nothing (draw a line),
'                                B       (draw a box) or
'                                BF      (draw a filled box)
Printer.Line (HorizontalMargin, VerticalMargin)-(21 - HorizontalMargin, 29.7 - VerticalMargin), RGB(255, 0, 0), B

'Now, we'll center a line of text in the page. To get the correct text measures,
'we must first define the font and attributes of the text
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.FontBold = True          'we want bold
Printer.FontItalic = False       'no italic
Printer.FontUnderline = False    'no underline
Printer.FontStrikethru = False   'no strike
Printer.ForeColor = RGB(0, 0, 0) 'color black

'We put the text in a variable (easier to handle) and get the text width
MyCenteredText = "This is centered text in Arial 12 Bold"
MyCenteredTextWidth = Printer.TextWidth(MyCenteredText)

'To set a starting position, we use Printer.CurrentX and Printer.CurrentY
'functions. To know where the text is to be located horizontally, we will use
'a very simple formula:
'            (Page Width - Text Width) / 2
'For height, we will put it 0.5 cm under the top of our printable area
Printer.CurrentX = (21 - MyCenteredTextWidth) / 2
Printer.CurrentY = VerticalMargin + 0.5
Printer.Print MyCenteredText

'Now, we'll align text to the left, and on the same line, we will align another
'one on the right. We will use different fonts and attributes than for the
'centered text above
Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.FontBold = False         'no bold
Printer.FontItalic = True        'we use italic
Printer.FontUnderline = False    'no underline
Printer.FontStrikethru = False   'no strike
Printer.ForeColor = RGB(0, 0, 0) 'color black

'This time, we don't need the text width, but it could be useful to verify
'the text is not too long, for the cases we work with user's text
MyLeftText = "Courier New on the Left"
MyLeftTextWidth = Printer.TextWidth(MyLeftText)
If MyLeftTextWidth > 21 - (HorizontalMargin * 2) Then Exit Sub

'we just set the starting location on the horizontal margin. We will set the
'height to 2 cm + vertical margin
'       NOTE: by adding a semi-colon (;) after the text, we indicate to VB
'             that we want the text coming immediately after this one to be
'             printed on the same line
Printer.CurrentX = HorizontalMargin
Printer.CurrentY = VerticalMargin + 2
Printer.Print MyLeftText;

'And we do the same for the text aligned on the right. As we keep the same
'attributes, we don't need to specify the font. To set the starting location,
'we will use this formula:
'            Page Width - Horizontal Margin - Text Width
'We don't have to specify the vertical position, it's the same than the left
'aligned text, because of the semi-colon
MyRightText = "Courier New on the Right"
MyRightTextWidth = Printer.TextWidth(MyRightText)
If MyRightTextWidth > 21 - (HorizontalMargin) Then Exit Sub
Printer.CurrentX = 21 - (HorizontalMargin) - MyRightTextWidth
Printer.Print MyRightText

'Now, we will create a grid and print it, with the lines. We will use
'3 rows and 4 colums.

'We define 12 text variables to fill our grid. In your real program, you
'can use text and values from textBox, DataBase fields, ...
txtGrid11 = "row1 & col1"
txtGrid12 = "row1 & col2"
txtGrid13 = "row1 & col3"
txtGrid14 = "row1 & col4"
txtGrid21 = "row2 & col1"
txtGrid22 = "row2 & col2"
txtGrid23 = "row2 & col3"
txtGrid24 = "row2 & col4"
txtGrid31 = "row3 & col1"
txtGrid32 = "row3 & col2"
txtGrid33 = "row3 & col3"
txtGrid34 = "row3 & col4"

'In this example, I have set the columns width as bellow. Of course, in real
'program, you will choose font, fontsize and column sizes to have them to
'fit your needs:
'                   Column 1 Width = 2
'                   Column 2 Width = 6
'                   Column 3 Width = 5
'                   Column 4 Width = 2

'       NOTE: If we want our text not overwriting the lines, then we will add
'             0.1 cm before it, and substract 0.1 cm after. (see above)

'We will first build the grid.
'We could have it centered horizontally, but let's do it simple, and align it to
'1 cm from the left margin. We will start it vertically at 4 cm from the vertical
'up margin.

'We'll draw the lines in Green [= RGB(0, 255, 0)]
Printer.ForeColor = RGB(0, 255, 0)

' 1) Upper border of the grid
Printer.Line (1 + HorizontalMargin, 4 + VerticalMargin)-(16 + HorizontalMargin, 4 + VerticalMargin)
' 2) Lower border (we will use 1 cm for each row height)
Printer.Line (1 + HorizontalMargin, 7 + VerticalMargin)-(16 + HorizontalMargin, 7 + VerticalMargin)
' 3) Left and Right borders (connecting to the two previous lines)
Printer.Line (1 + HorizontalMargin, 4 + VerticalMargin)-(1 + HorizontalMargin, 7 + VerticalMargin)
Printer.Line (16 + HorizontalMargin, 4 + VerticalMargin)-(16 + HorizontalMargin, 7 + VerticalMargin)
' 4) Other horizontal lines (every centimeter)
Printer.Line (1 + HorizontalMargin, 5 + VerticalMargin)-(16 + HorizontalMargin, 5 + VerticalMargin)
Printer.Line (1 + HorizontalMargin, 6 + VerticalMargin)-(16 + HorizontalMargin, 6 + VerticalMargin)
' 5) Other vertical lines (according to the columns size we choose earlier)
Printer.Line (3 + HorizontalMargin, 4 + VerticalMargin)-(3 + HorizontalMargin, 7 + VerticalMargin)
Printer.Line (9 + HorizontalMargin, 4 + VerticalMargin)-(9 + HorizontalMargin, 7 + VerticalMargin)
Printer.Line (14 + HorizontalMargin, 4 + VerticalMargin)-(14 + HorizontalMargin, 7 + VerticalMargin)

'we print text in the grid in black
Printer.ForeColor = RGB(0, 0, 0)

'Let's put a little centered title. But we center it to the grid, not to the page.
'we use almost the same formula, but we don't base it on the Page Width, but on
'the grid width (wich is 15 cm):
'           Horizontal Margin + 1 + ((Grid Width - Text Width)/2)
'       NOTE : We decided our grid will start 1cm from the left margin, so
'              we add 1 to the left margin
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.FontBold = False         'no bold
Printer.FontItalic = False       'no italic
Printer.FontUnderline = False    'no underline
Printer.FontStrikethru = False   'no strike
MyGridTitle = "And this is a little table"
MyGridTitleWidth = Printer.TextWidth(MyGridTitle)
Printer.CurrentX = HorizontalMargin + 1 + ((15 - MyGridTitleWidth) / 2)

'Regarding the vertical position, we want our title to be printed BEFORE the
'grid, and not after. So, we will set the position to be at the upper border
'of the grid, and we will substract the text height (and 0.1 cm too, so our
'line won't be messed if we have some characters like "g" or "{"
Printer.CurrentY = 4 + VerticalMargin - (Printer.TextHeight(MyGridTitle) + 0.1)
Printer.Print MyGridTitle

'And finally, we write the content of each cell of our grid. We will also
'center them vertically, to have a nicer effect. To have it easier to manage,
'we'll use some variables to hold our locations.

' 1) first row
'    We set the horizontal position, wich is the left grid border + 0.1 cm
Row1Col1Left = 1 + HorizontalMargin + 0.1
'    then, we set the left position of all the text in row 1
Row1Col2Left = 3 + HorizontalMargin + 0.1
Row1Col3Left = 9 + HorizontalMargin + 0.1
Row1Col4Left = 14 + HorizontalMargin + 0.1
'    now, we will center the text vertically. we already know there is
'    1 cm between each line in the grid, so each cell height is 1 cm.
'    we'll use the formula:
'     Top Grid Line + Vert. Margin + ((Cell Height - Text Height) / 2)
Row1Top = 4 + VerticalMargin + ((1 - Printer.TextHeight(txtGrid11)) / 2)
'
' 2) other rows Top position
Row2Top = 5 + VerticalMargin + ((1 - Printer.TextHeight(txtGrid11)) / 2)
Row3Top = 6 + VerticalMargin + ((1 - Printer.TextHeight(txtGrid11)) / 2)
'
' 3) second row Left position
Row2Col1Left = 1 + HorizontalMargin + 0.1 '(column 1)
Row2Col2Left = 3 + HorizontalMargin + 0.1 '(column 2)
Row2Col3Left = 9 + HorizontalMargin + 0.1 '(column 3)
Row2Col4Left = 14 + HorizontalMargin + 0.1 '(column 4)
'
' 4) third row Left position
Row3Col1Left = 1 + HorizontalMargin + 0.1 '(column 1)
Row3Col2Left = 3 + HorizontalMargin + 0.1 '(column 2)
Row3Col3Left = 9 + HorizontalMargin + 0.1 '(column 3)
Row3Col4Left = 14 + HorizontalMargin + 0.1 '(column 4)
'
' 4) we print the first row (do not forget the semi-colons ";")
Printer.CurrentY = Row1Top
Printer.CurrentX = Row1Col1Left
Printer.Print txtGrid11;
Printer.CurrentX = Row1Col2Left
Printer.Print txtGrid12;
Printer.CurrentX = Row1Col3Left
Printer.Print txtGrid13;
Printer.CurrentX = Row1Col4Left
Printer.Print txtGrid14
'
' 5) we print the second row
Printer.CurrentY = Row2Top
Printer.CurrentX = Row2Col1Left
Printer.Print txtGrid21;
Printer.CurrentX = Row2Col2Left
Printer.Print txtGrid22;
Printer.CurrentX = Row2Col3Left
Printer.Print txtGrid23;
Printer.CurrentX = Row2Col4Left
Printer.Print txtGrid24
'
' 6) we print the last row (3rd)
Printer.CurrentY = Row3Top
Printer.CurrentX = Row3Col1Left
Printer.Print txtGrid31;
Printer.CurrentX = Row3Col2Left
Printer.Print txtGrid32;
Printer.CurrentX = Row3Col3Left
Printer.Print txtGrid33;
Printer.CurrentX = Row3Col4Left
Printer.Print txtGrid34

'IMAGE PRINTING SHORT OVERVIEW
'Image to be printed must be in a picture box. We use the hidden picture box
'on the Form to load a picture in , then we retrieve its size, applies this size
'to the hidden picture box, then we print it.

'Loading the image in the hidden picture box, then we auto-resize it so its
'size will exactly fit the image we loaded (always in centimeters)
Picture2.ScaleMode = vbCentimeters
Picture2.Picture = LoadPicture(App.Path & "\dark_shadow.jpg")
Picture2.AutoSize = True
Picture2.Refresh
Picture2.AutoSize = False

'we center the picture with the formula: (21 - Picture Width) / 2
ImageLeft = (21 - Picture2.ScaleWidth) / 2

'We'll print it at 12 centimeters from the top margin
ImageTop = 12

'We reset the Printer measure to centimeters, and we print the picture at the
'location we defined
Printer.ScaleMode = vbCentimeters
Printer.PaintPicture Picture2.Picture, ImageLeft, ImageTop

'We say to the printer that we're finished and so it can work. If we don't
'add this line at the end of the printing SUB, then the Printer would only
'start when we close the program.
Printer.EndDoc
End Sub

Private Sub Command2_Click()
End
End Sub

