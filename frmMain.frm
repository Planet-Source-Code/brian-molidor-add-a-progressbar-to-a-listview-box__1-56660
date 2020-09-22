VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Progressbar in ListView Box"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   26
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "index"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Progressbar"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.CommandButton btnRunItem 
      Caption         =   "&Run Selected Item"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton btnAddItem 
      Caption         =   "&Add Item"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Brian Molidor
'Pulsewave@aol.com
'
'Add a Progress bar to a listview box
'like napster, kaza, limewire, bearshare
'and most other p2p network clients

Const mForeColor = &H0& 'progressbar forecolor
Const mBackColor = &HFF0000     'progressbar backcolor

Private Sub btnAddItem_Click()
    p = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add p, "", "" 'the first column is hidden
    ListView1.ListItems.Item(p).ListSubItems.Add 1, "", CStr(p) 'index column
    ListView1.ListItems.Item(p).ListSubItems.Add 2, "", , 1 'progressbar column
End Sub

Private Sub btnRunItem_Click()
    'get the selected item index
    r = ListView1.SelectedItem.Index
    'loop through every image in imagelist
    'and change it in the listview
    For i = 1 To ImageList1.ListImages.Count
        'change progressbar image
        ListView1.ListItems(r).ListSubItems(2).ReportIcon = i
        'refresh list
        ListView1.Refresh
        'pause so you can see it
        'normaly you'll be preforming
        'tasks so you wont have to use this
        Call Pause(0.1)
    Next i
'since this was just and example i didn't
'acutaly put any usefull code in here
'for you to use with a normal application
'so here's how i would use this:
'let's say your dling a file
'and you have 834k out of 2983k (what image index is that? who knows... but we can figure it out like so)
'
'so you would get the percent of 834/2983
'p = Round((834 / 2983) * 100)
'ok so you have 28% complete
'now you need to get 28% of
'imagelist1 total image count
'myImgIndx = Round((ImageList1.ListImages.Count * p) / 100)
'
'p = Round((834 / 2983) * 100)
'myImgIndx = Cint(Round((ImageList1.ListImages.Count * p) / 100))
'ListView1.ListItems(1).ListSubItems(2).ReportIcon = myImgIndx
'
'so now you would set your reporticon to myImgIndx
'easy as that.
End Sub

Private Sub Form_Load()
    'set picturebox's colors
    Picture1.ForeColor = &H0&
    Picture1.BackColor = &H8000000F
    Picture1.AutoRedraw = True
    
    'set scalemode to pixels
    frmMain.ScaleMode = vbpixel
    Picture1.ScaleMode = vbpixel
    
    'draw the progressbar's
    'you can change the 133 to any width
    'that you would like the progressbar to be
    'normaly it's the same measurement as the
    'column header..(set scalemode to pixels)
    Call reDrawImages(133)
    
    'add an item to the listview
    Call btnAddItem_Click
End Sub

Private Sub reDrawImages(mWidth As Integer)
    Picture1.Width = mWidth
    For i = 1 To Picture1.ScaleWidth
        Picture1.Cls
        'draw the progress line
        Picture1.Line (0, 0)-((i / Picture1.ScaleWidth) * Picture1.ScaleWidth, Picture1.ScaleHeight), mBackColor Xor mForeColor, BF
        'set text position in picturebox
        Picture1.CurrentX = (Picture1.ScaleWidth / 2) - (Picture1.TextWidth(CInt((Picture1.TextWidth(i / Picture1.Width) * 100))) / 2)
        Picture1.CurrentY = (Picture1.ScaleHeight / 2) - (Picture1.TextHeight("1") / 2)
        'print percent
        Picture1.Print CStr(Round(((i / Picture1.ScaleWidth) * 100), 0) & "%")

        'draw black lines around it
        'so it has a border
        Picture1.Line (0, 0)-(Picture1.ScaleWidth, 0), mForeColor, BF
        Picture1.Line (0, Picture1.ScaleHeight - 1)-(Picture1.ScaleWidth, Picture1.ScaleHeight - 1), mForeColor, BF
        Picture1.Line (Picture1.ScaleWidth - 1, 0)-(Picture1.ScaleWidth - 1, Picture1.ScaleHeight), mForeColor, BF
        Picture1.Line (0, 0)-(0, Picture1.ScaleHeight)
    
        'refresh the picture box
        Picture1.Refresh

        'add image to imagelist
        ImageList1.ListImages.Add i, "", Picture1.Image
    Next i
End Sub

Private Sub Pause(interval)
    current = Timer
        Do While Timer - current < Val(interval)
        DoEvents
    Loop
End Sub
