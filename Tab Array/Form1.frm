VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   4500
      Width           =   6165
   End
   Begin VB.Frame Frame1 
      Caption         =   "Starter Controls"
      Height          =   2175
      Left            =   7200
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   1425
      Begin VB.Frame frmArray 
         Caption         =   "Frame1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1140
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ListBox IListArray 
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4365
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   7699
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      OLEDropMode     =   1
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
   End
   Begin VB.Menu MnuStuff 
      Caption         =   "Stuff to do"
      Begin VB.Menu MnuLoad 
         Caption         =   "Load Array"
      End
      Begin VB.Menu MnuTag 
         Caption         =   "Add Text IList Array"
      End
      Begin VB.Menu MnuAdd 
         Caption         =   "Add Tab"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' I only did this cause i had a bitch of a time trying to load
' controls into specific tabs dynamicly.
' this code is quite simplistic, but took me research and trial and error
' ... so maybe this will save some one some time :-)

' For more SStab info http://msdn.microsoft.com/library/en-us/mstab98/html/vbobjMSTabControl.asp

Private Sub Form_Resize()

On Error GoTo ErrHandler

SSTab1.Height = Form1.Height - 1100
SSTab1.Width = Form1.Width - 330
SSTab1.Left = 90
HScroll1.Width = Form1.Width - 360
HScroll1.Left = 120
HScroll1.Top = Form1.Height - 980

    For k = 0 To SSTab1.Tabs

        IListArray(k).Width = SSTab1.Width - 575
        IListArray(k).Height = SSTab1.Height - 920
        frmArray(k).Width = SSTab1.Width - 300
        frmArray(k).Height = SSTab1.Height - 650

    Next

ErrHandler:
End Sub

Private Sub HScroll1_Change()

' not sure why i put this here, cept i never used scroll before and was curious

HScroll1.Enabled = True ' enables control
HScroll1.Min = 0 ' sets min val
HScroll1.Max = SSTab1.Tabs - 1 ' sets max val, - 1 cause tabs index starts @ 0

    
    If HScroll1.Value >= SSTab1.Tabs Then GoTo End1 ' makes sure scroller count doesnt exceed tab count

    SSTab1.Tab = HScroll1.Value ' keeps focused tab and scroll value in sync


End1:
End Sub

Private Sub MnuLoad_Click()
' the good stuff ~~ setting tab contianer's and loading arrays into them
On Error GoTo ErrHandler

SSTab1.TabsPerRow = 9
SSTab1.Tabs = 9 ' sets amount tabs
SSTab1.Tab = 0 'brings 1st tab to focus
' starts loop
    For i = 1 To SSTab1.Tabs ' starts @ 1 cause we already have control on form - sets loop to end at tab amount
       Load IListArray(i) ' loads array 1
       Load frmArray(i)   ' loads array 2
         With frmArray(i)
           SSTab1.Tab = (i) - 1 ' brings current tab to focus
             Set .Container = Me.SSTab1 'so we can drop control here
             frmArray(i).Caption = "Frame" & " " & (i) - 1 'this and next 5 lines are control design / placement
             frmArray(i).Visible = True
             frmArray(i).Width = 5895
             frmArray(i).Top = 480
             frmArray(i).Left = 150
             frmArray(i).Height = 3675
             
            With IListArray(i) ' same as above just for array #2
                SSTab1.Tab = (i) - 1
                Set .Container = Me.SSTab1
                IListArray(i).Left = 270
                IListArray(i).Top = 690
                IListArray(i).Width = 5685
                IListArray(i).Height = 3375
                IListArray(i).Visible = True
             End With ' ends loop 1
         End With ' ends loop 2
     Next ' keeps loop moving along
     
 SSTab1.Tab = 0
 
ErrHandler:
End Sub

Private Sub MnuTag_Click()

On Error GoTo ErrHandler
'This is just for my own debug purposes

    For j = 0 To SSTab1.Tabs
        SSTab1.Tag = IListArray(j).Name & IListArray(j).Index
        IListArray(j).AddItem IListArray(j).Name & " " & IListArray(j).Index - 1 ' The - 1 is cheating !  if i didnt do, minus 1 then all the IListArray Text would be off by 1. because, tab index starts at zero.
        'only did it for visual asthetics
    Next


ErrHandler:
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

HScroll1.Value = SSTab1.Tab ' keeps focused tab and scrol value in sync

End Sub

