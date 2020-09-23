VERSION 5.00
Begin VB.Form MailLst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mail List"
   ClientHeight    =   3552
   ClientLeft      =   2076
   ClientTop       =   3276
   ClientWidth     =   6624
   Icon            =   "MAILLST.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3552
   ScaleWidth      =   6624
   Begin VB.PictureBox Tools 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   612
      ScaleWidth      =   6624
      TabIndex        =   2
      Top             =   2940
      Width           =   6624
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   6660
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   6660
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image PrtImage 
         Height          =   384
         Left            =   1200
         Picture         =   "MAILLST.frx":030A
         Top             =   60
         Width           =   384
      End
      Begin VB.Image Trash 
         Height          =   384
         Left            =   300
         Picture         =   "MAILLST.frx":0614
         Top             =   60
         Width           =   384
      End
   End
   Begin VB.ListBox MList 
      Height          =   2184
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Headings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listbox Headings"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   1680
   End
   Begin VB.Image NewMail 
      Height          =   384
      Left            =   5880
      Picture         =   "MAILLST.frx":091E
      Top             =   2820
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image nonew 
      Height          =   384
      Left            =   5280
      Picture         =   "MAILLST.frx":0C28
      Top             =   2880
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "MailLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_TemplateDerived = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module variable to hold MouseDown position information.
Dim ListX, ListY

Private Sub Form_Load()
    ' Resize the form.
    Height = 3945
    Call Tools_Resize

     ' Set list box headings.
     a$ = Mid$(Format$("From", "!" + String$(25, "@")), 1, 25)
     b$ = Mid$(Format$("Subject", "!" + String$(35, "@")), 1, 35)
     c$ = "Date"
     Headings = a$ + b$ + c$
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' If the user is closing the application, let this form unload.
    If UnloadMode = 4 Then
        ' Unloading is permitted.
    Else
        ' If the user is still logged on, minimize the form rather than closing it.
        If VBMail.MapiMess.SessionID <> 0 Then
            Me.WindowState = 1
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Resize()
    ' If the form isn't minimized, resize the list box to fit the form.
    If WindowState <> 1 Then

        If VBMail.DispTools.Checked Then
            xHeight% = Tools.Height
        Else
            xHeight% = 0
        End If

        ' Check for the minimum form height.
        If Height < 2500 - xHeight% Then
            Height = 2500
            Exit Sub
        End If

        MList.Width = ScaleWidth - MList.Left - 90
        MList.Height = ScaleHeight - 90 - MList.Top - xHeight%
    End If
End Sub

Private Sub MList_Click()
' Set the message index and enable the
' Previous and Next buttons as needed.
    Select Case MList.ListIndex
        Case 0
            VBMail.Previous.Enabled = False
        Case MList.ListCount - 1
            VBMail.Next.Enabled = False
        Case Else
            VBMail.Previous.Enabled = True
            VBMail.Next.Enabled = True
    End Select
    VBMail.MapiMess.MsgIndex = MList.ListIndex
End Sub

Private Sub MList_DBLClick()
' Check to see if the message is currently viewed,
' and if it isn't, load it into a new form.
    If Not MailLst.MList.ItemData(MailLst.MList.ListIndex) Then
       Dim Msg As New MsgView
       Call LoadMessage(MailLst.MList.ListIndex, Msg)
       MailLst.MList.ItemData(MailLst.MList.ListIndex) = True
    Else
        ' Search through the active windows to
        ' find the window with the correct message to view.
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = MailLst.MList.ListIndex Then
                    Forms(i).Show
                    Exit Sub
                End If
            End If
        Next i
     End If
End Sub

Private Sub MList_KeyPress(KeyAscii As Integer)
    ' If the user presses ENTER, process the action as a DblClick event.
    If KeyAscii = 13 Then
        Call MList_DBLClick
    End If
End Sub

Private Sub MList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Save the X and Y positions to determine the start of the drag-and-drop action.
    ListX = X
    ListY = Y
End Sub

Private Sub MList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If the mouse button is down and the X,Y position has changed, start dragging.
    If Button = 1 And ((X <> ListX) Or (Y <> ListY)) Then
        MList.Drag 1
    End If
End Sub

Private Sub PrtImage_DragDrop(Source As Control, X As Single, Y As Single)
    ' Same as File.PrintMessage on the VBMAIL File menu.
    Call PrintMail
End Sub

Private Sub Tools_Resize()
    ' Adjust the width of the lines on the top of the toolbar.
    Line1(0).X2 = Tools.Width
    Line1(1).X2 = Tools.Width
    Tools.Refresh
End Sub

Private Sub Trash_DragDrop(Source As Control, X As Single, Y As Single)
    ' Delete a message (Delete Button or Edit.Delete).
   Call DeleteMessage
End Sub

