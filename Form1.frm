VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Collector"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList IList2 
      Left            =   2835
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":07DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1644
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Scan All Pages"
      Height          =   240
      Left            =   6915
      TabIndex        =   14
      Top             =   855
      Width           =   1590
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   90
      TabIndex        =   11
      Top             =   6105
      Width           =   8250
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5970
      Top             =   2010
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   10
   End
   Begin RichTextLib.RichTextBox Rich1 
      Height          =   1185
      Left            =   6165
      TabIndex        =   10
      Top             =   2970
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   2090
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":19DE
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   90
      TabIndex        =   9
      Top             =   4830
      Width           =   8250
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   7185
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2566
            Text            =   "URLs Prosessed: 0"
            TextSave        =   "URLs Prosessed: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Mails Collected: 0"
            TextSave        =   "Mails Collected: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Findrate: 0%"
            TextSave        =   "Findrate: 0%"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Collect emails from selected url"
      Height          =   240
      Left            =   105
      TabIndex        =   7
      Top             =   4365
      Width           =   8220
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Text            =   "http://search.yahoo.com/search?p=vb%2Bdownload%2B99"
      Top             =   810
      Width           =   6660
   End
   Begin VB.CommandButton Command3 
      Height          =   420
      Left            =   1230
      Picture         =   "Form1.frx":1A69
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Height          =   420
      Left            =   660
      Picture         =   "Form1.frx":1DF3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   420
      Left            =   90
      Picture         =   "Form1.frx":217D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   75
      Width           =   480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5175
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   12
      ImageHeight     =   15
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2507
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4663
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lw1 
      Height          =   2925
      Left            =   105
      TabIndex        =   0
      Top             =   1410
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   5159
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Urls prosessed:"
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   5880
      Width           =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Urls to prosess:"
      Height          =   225
      Left            =   165
      TabIndex        =   12
      Top             =   4620
      Width           =   2730
   End
   Begin VB.Label Label2 
      Caption         =   "Sites that needs manual check to see that the mails are not SPAM BAIT - Click on the link to view page:"
      Height          =   345
      Left            =   150
      TabIndex        =   6
      Top             =   1200
      Width           =   8205
   End
   Begin VB.Label Label1 
      Caption         =   "Startpoint/URL:"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   585
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbx As DAO.Database
Dim tdx As DAO.TableDef
Dim fdx As DAO.Field
Dim rsx As DAO.Recordset
Dim addit As Boolean
Dim stopit As Boolean
Dim continue As Boolean

Private Sub Command1_Click()
Command2.Enabled = True
Command1.Picture = IList2.ListImages(2).Picture
Command2.Picture = IList2.ListImages(5).Picture
Command3.Picture = IList2.ListImages(3).Picture
On Error GoTo errmess
Dim itmX As ListItem
  Dim clmX As ColumnHeader
    Set clmX = Form1.Lw1.ColumnHeaders.Add(, , "Mails found")

  Set clmX = Form1.Lw1.ColumnHeaders.Add(, , "Url", 10000)
Dim regEx
Set regEx = New RegExp

regEx.IgnoreCase = True
    regEx.Global = True


'open database
Set dbx = OpenDatabase(App.Path & "\data.mdb")
List1.AddItem Text1.Text
stopit = False
continue = True
Do 'start Loop While List1.ListCount > 0 And stopit = False
addit = True
'check already prosessed list for link
For x = 0 To List2.ListCount - 1
If List1.List(0) = x Then
addit = False
Exit For
End If
Next
'get the webpage
Rich1.Text = Inet1.OpenURL(List1.List(0))
'wait until page is loaded
Do
DoEvents
Loop While Inet1.StillExecuting = True
    teller2 = teller2 + 1
    StatusBar1.Panels(1).Text = "URLs Prosessed: " & teller2
'add link to already prosessed list
List2.AddItem List1.List(0)
'page is downloaded - remove it from list
List1.RemoveItem (0)
'Get links from webpage
regEx.Pattern = "(http|https|ftp)\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&%\$#\=~])*"
Set Matches = regEx.Execute(Rich1.Text)
'This is for stop filling list if we have 100 links to prosess
If List1.ListCount < 100 And addit = True Then
For Each Match In Matches
Rich1.SelStart = (Match.FirstIndex)
Rich1.SelLength = Len(Match.Value)
res1 = InStr(1, Rich1.SelText, ".htm", vbTextCompare)
res2 = InStr(1, Rich1.SelText, "srd.yahoo.com", vbTextCompare)
MyStr = Right(Rich1.SelText, 1)
If MyStr = "/" Then
res1 = 1
End If
'selecting checkbox means not just scanning static pages but everything
If Check1.Value = 1 Then
res1 = 1
End If
'yahoo adds their link to page links - remove it
If res2 > 0 Then
linkx = Split(Rich1.SelText, "*", -1, vbTextCompare)
link = linkx(1)
End If

If res1 > 0 Then
'add links to list to prosess
If res2 > 0 Then
List1.AddItem link
Else
List1.AddItem Rich1.SelText
End If
End If

Next 'end For Each Match In Matches
End If 'end If List1.ListCount < 100 And addit = True Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'start getting emails
regEx.Pattern = "[\w_-]+(\.[\w_-]+)*@[\w_-]+(\.[\w_-]+)*\.[a-z]{2,3}"
Set Matches = regEx.Execute(Rich1.Text)
mailcount = 0
For Each Match In Matches
mailcount = mailcount + 1
Next
 
If mailcount > 10 Then
  Set itmX = Form1.Lw1.ListItems.Add(, Form1.Lw1.ListItems.Count + 1, mailcount)
  itmX.SubItems(1) = Inet1.URL
Else
For Each Match In Matches
Rich1.SelStart = (Match.FirstIndex)
Rich1.SelLength = Len(Match.Value)
'check if you already got that mail
Set rsx = dbx.OpenRecordset("SELECT * FROM Mails WHERE mail Like " & "'" & Rich1.SelText & "'")

If rsx.RecordCount > 0 Then
Else
rsx.AddNew
rsx("mail") = Rich1.SelText
rsx.Update
teller1 = teller1 + 1
StatusBar1.Panels(2).Text = "Mails Collected: " & teller1
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'avoid getting the prossesed links list to large
If List2.ListCount > 2000 Then
For i = 0 To 500
List2.RemoveItem (0)
Next
End If

Next 'For Each Match In Matches
End If ' If mailcount > 10 Then
'set the pause loop
Do
DoEvents
Loop While continue = False

findrate = (teller1 / teller2) * 100
findratex = Format(findrate, "##0.00")
StatusBar1.Panels(3).Text = "Findrate: " & findratex & "%"
Loop While List1.ListCount > 0 And stopit = False
    Exit Sub
'remove bad url if error
errmess:
    If List1.ListCount > 0 Then
    List1.RemoveItem (0)
    End If
    Resume Next
End Sub

Private Sub Command2_Click()
If continue = False Then
continue = True
Command1.Picture = IList2.ListImages(2).Picture
Command2.Picture = IList2.ListImages(5).Picture
Command3.Picture = IList2.ListImages(3).Picture

Else
continue = False
Command1.Picture = IList2.ListImages(2).Picture
Command2.Picture = IList2.ListImages(6).Picture
Command3.Picture = IList2.ListImages(3).Picture
End If
End Sub

Private Sub Command3_Click()
stopit = True
Command2.Enabled = False
Command1.Picture = IList2.ListImages(1).Picture
Command2.Picture = IList2.ListImages(5).Picture
Command3.Picture = IList2.ListImages(4).Picture
End Sub



Private Sub Form_Load()
Command2.Enabled = False
Command1.Picture = IList2.ListImages(1).Picture
Command2.Picture = IList2.ListImages(5).Picture
Command3.Picture = IList2.ListImages(4).Picture
End Sub

Private Sub Lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Call ListViewFuncs.SortListView(Form1.Lw1, ColumnHeader.Index)
  
End Sub
