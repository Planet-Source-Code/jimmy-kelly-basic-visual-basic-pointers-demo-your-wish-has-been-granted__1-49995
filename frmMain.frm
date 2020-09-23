VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB Pointers"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You may be asking.. Why use pointers?
'well because you can directly access so many things including the
'bitmap bits of a picture in THE MEMORY! This allows for super fast
'bitmap manipulation among hundreds of other things! Enjoy and prove
'those C/C++ coders wrong about VB!

'Just a word of warning.
'KNOW WHAT YOU ARE DOING!
'Directly manipulating memory is dangerous and can cause your VB or computer
'to crash. If you have read and C/C++ books they also warn about this when
'going into pointers.

'James

Private Sub Form_Load()

Me.Show

Dim x As Long, xptr As Long

MsgBox "Handling variables manually"

x = Var(2)
SetVar x, 65

MsgBox "X is located at " & x & " and reads " & GetVar(x, 2)

KillVar x

MsgBox "Creating pointers to existant variables"

xptr = VarPtr(x) 'VarPtr is a undocumented VB function
SetVar xptr, 65 'along with ObjPtr and StrPtr. Read more
'@ http://www.undergroundnews.com/boards/ubb-get_topic-f-11-t-000030.html
'and http://www.codeproject.com/vbscript/how_to_do_pointers_in_visual_basic.asp

MsgBox "X is located at " & xptr & " and reads " & GetVar(xptr, Len(Str(x)))

MsgBox "Creating pointers to bitmaps"

xptr = VarPtr(picMain.Picture)

MsgBox "picMain's picture is located at " & xptr

'I would do some bitmap manipulation in this demo but I have abosolutely
'no idea how to get the size of a picture in bytes so...

Unload Me

End Sub
