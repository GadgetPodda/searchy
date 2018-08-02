VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form searchy 
   Caption         =   "Searchy 0.1 A"
   ClientHeight    =   8148
   ClientLeft      =   108
   ClientTop       =   732
   ClientWidth     =   11016
   Icon            =   "searchy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8148
   ScaleWidth      =   11016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabs 
      Height          =   11052
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   15132
      _ExtentX        =   26691
      _ExtentY        =   19495
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "Google"
      TabPicture(0)   =   "searchy.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "wb1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Yahoo!"
      TabPicture(1)   =   "searchy.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wb2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Bing"
      TabPicture(2)   =   "searchy.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "wb3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ask!"
      TabPicture(3)   =   "searchy.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "wb4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Zapmeta"
      TabPicture(4)   =   "searchy.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "wb5"
      Tab(4).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser wb5 
         Height          =   10692
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   15012
         ExtentX         =   26479
         ExtentY         =   18860
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser wb4 
         Height          =   10692
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   15012
         ExtentX         =   26479
         ExtentY         =   18860
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser wb3 
         Height          =   10692
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   15012
         ExtentX         =   26479
         ExtentY         =   18860
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser wb2 
         Height          =   10692
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   15012
         ExtentX         =   26479
         ExtentY         =   18860
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser wb1 
         Height          =   10692
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   15012
         ExtentX         =   26479
         ExtentY         =   18860
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.CommandButton SearchBtn 
      Caption         =   "Search in All"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   12960
      TabIndex        =   1
      Top             =   120
      Width           =   2052
   End
   Begin VB.TextBox InputSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Text            =   "Text2Search"
      Top             =   120
      Width           =   12612
   End
   Begin VB.Menu menu_help 
      Caption         =   "Help"
      Index           =   1
      NegotiatePosition=   2  'Middle
      Begin VB.Menu menu_doc 
         Caption         =   "FAQ"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu help_about_program 
         Caption         =   "About Program"
         Index           =   3
         Shortcut        =   {F2}
      End
      Begin VB.Menu menu_about_author 
         Caption         =   "About Author"
         Index           =   4
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "searchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
wb1.Silent = True
wb2.Silent = True
wb3.Silent = True
wb4.Silent = True
wb5.Silent = True
End Sub

Private Sub help_about_program_Click(Index As Integer)
MsgBox "This is a Simple Program to Search Something Online with 5 Different Search Engines at One Click. Every Search Engine is Loaded in Seperate Tabs, So You can use your Favourite One. Also, This Program can be used to Compare Listings of your Website in Defferent Search Engines. Current Version is 0.1a. Created with VB6. Licensed Under MIT License.", vbOKOnly + vbInformation, "About Searchy"
End Sub

Private Sub menu_about_author_Click(Index As Integer)
MsgBox "The First Author of Searchy is GadgetPodda. Contact Email and Paypal Donation Email : gadgetpodda2005@gmail.com GitHub Username : GadgetPodda", vbOKOnly + vbInformation, "About Author"
End Sub

Private Sub menu_doc_Click(Index As Integer)
faq.Visible = True
End Sub

Private Sub SearchBtn_Click()
wb1.Navigate "http://www.google.com/search?q=" + InputSearch.Text
wb2.Navigate "http://search.yahoo.com/search?ei=UTF-8&p=" + InputSearch.Text
wb3.Navigate "https://www.bing.com/search?q=" + InputSearch.Text
wb4.Navigate "http://www.ask.com/web?q=" + InputSearch.Text
wb5.Navigate "https://www.zapmeta.com/?q=" + InputSearch.Text
End Sub

