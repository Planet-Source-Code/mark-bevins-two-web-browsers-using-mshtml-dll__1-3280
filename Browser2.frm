VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Browser Form"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5880
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4125
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wbr 
      Height          =   1560
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   375
      Width           =   10920
      ExtentX         =   19262
      ExtentY         =   2752
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser wbr 
      Height          =   4650
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   2265
      Width           =   9315
      ExtentX         =   16431
      ExtentY         =   8202
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuFileReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Browser.frm (Vb6) Mar 1999   Written by Mark Bevins Orion Studios  markb@orionstudios.com
' main form for Project "TwoWbr"
' Requires Project/References entry for
'   "Microsoft HTML Object Library" (mshtml.dll)
'=================================================================================
' Module-level Variables
Private MARGINx2 As Long
Private mVertUsedArea As Long
Private WithEvents mFidelity As MSHTML.HTMLAnchorElement
Attribute mFidelity.VB_VarHelpID = -1
' Module-level Constants
Private Const MARGIN = 15
Private Const WBR_START = 0
Private Const WBR_LINK = 1
Private Const IMG_SRC = "http://www.orionstudios.net/fidelity.gif" ' links to IMG_LINK
Private Const IMG_NAME = "fidelity.gif" ' see wbr_DocumentComplete
Private Const IMG_TITLE = " Click for Fidelity! "
Private Const IMG_LINK = "http://www.orionstudios.com" '<--- Link to go to
Private Const BLANK_PAGE = "about:<HTML><BODY BGCOLOR=gainsboro SCROLL=NO></BODY></HTML>"
Private Const START_HTML _
        = "<HTML>" _
        & "<BODY BGCOLOR=lightskyblue SCROLL=NO>" _
        & "<CENTER>" _
        & "<A ID=idFidelity HREF=" & IMG_LINK & ">" _
        & "<IMG SRC=" & IMG_SRC & " TITLE='" & IMG_TITLE & "'>" _
        & "</A>" _
        & "</CENTER>" _
        & "</BODY>" _
        & "</HTML>"

Private Sub Form_Load()

    With Me
        .Caption = App.FileDescription
        .Move 1200, 0, 9600, 8400
    End With
    MARGINx2 = MARGIN * 2
    mVertUsedArea = MARGINx2 + sta.Height
    With wbr(WBR_LINK)
        .Visible = False
        .Navigate BLANK_PAGE
    End With
    wbr(WBR_START).Navigate BLANK_PAGE
    
End Sub

Private Sub mnuFile_Click()

    mnuFileStart = (mFidelity Is Nothing)
    
End Sub

Private Sub mnuFileStart_Click()

    With wbr(WBR_LINK)
        .Visible = False
        .Navigate BLANK_PAGE
    End With
    With wbr(WBR_START)
        .Navigate "about:" & START_HTML
        .Visible = True
    End With
    
End Sub


Private Sub mnuFileReset_Click()

    With wbr(WBR_LINK)
        .Visible = False
        .Navigate BLANK_PAGE
    End With
    With wbr(WBR_START)
        .Navigate BLANK_PAGE
        .Visible = True
    End With
    
End Sub

Private Sub mnuFileClose_Click()

    Unload Me
    
End Sub

Private Sub wbr_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)
    
    If pDisp = wbr(index).Object Then
        If InStr(1, URL, IMG_NAME, vbTextCompare) Then
            Set mFidelity = wbr(index).Document.All.idFidelity
        Else
            Set mFidelity = Nothing
        End If
    End If

End Sub

Private Sub wbr_StatusTextChange(index As Integer, ByVal Text As String)

    sta.SimpleText = Text
    
End Sub

Private Function mFidelity_onclick() As Boolean
'
' When IMG_SRC is clicked, use a second WebBrowser to navigate to the link
'
    mFidelity_onclick = False   ' cancel default behaviour
    wbr(WBR_START).Visible = False  ' hide WebBrowser containing image
    wbr(WBR_LINK).Visible = True
    
    DoEvents
    wbr(WBR_LINK).Navigate IMG_LINK ' use a second WebBrowser to display linked page
    
End Function

