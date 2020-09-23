VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text Import Wizard"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "import.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      Picture         =   "import.frx":0442
      ScaleHeight     =   6375
      ScaleWidth      =   6015
      TabIndex        =   12
      Top             =   0
      Width           =   6015
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close Wizard"
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   4800
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   720
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Import succesfully!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      Picture         =   "import.frx":2968
      ScaleHeight     =   6375
      ScaleWidth      =   6015
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command2 
         Caption         =   "Finish"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "< Back"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "...."
         Height          =   255
         Left            =   5400
         TabIndex        =   6
         Top             =   2280
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "import.frx":4E8E
         Left            =   3120
         List            =   "import.frx":4EA1
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Choose file to import and choose a seperator"
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "File to Import :"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      Picture         =   "import.frx":4EB6
      ScaleHeight     =   6375
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Next >"
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "import.frx":73DC
         Top             =   720
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Visible = False
Picture2.Visible = True
End Sub
Function Does_Excist(TheFileName) As Boolean
If Dir(TheFileName) <> "" Then Does_Excist = True Else Does_Excist = False
End Function
Private Sub Command2_Click()
List1.Clear
Dim seperator

'FIRST CHECK WHAT USER HAS CHOSEN
If Combo1.Text = "Choose" Or Combo1.Text = "" Then
MsgBox "Select a seperator used in the textfile", vbCritical, "Import error!"
Exit Sub
ElseIf Combo1.Text = "#" Then
seperator = Chr$(35)
ElseIf Combo1.Text = "*" Then
seperator = Chr$(42)
ElseIf Combo1.Text = "!" Then
seperator = Chr$(33)
ElseIf Combo1.Text = "^" Then
seperator = Chr$(94)
Else
seperator = Chr$(9)
End If

Dim FileManually As String 'File to import
FileManually = Text2.Text

'CHECK FOR EXCISTING FILENAME
If Not Does_Excist(FileManually) Or FileManually = "" Then
MsgBox "The file " & FileManually & " cannot be opened!", vbCritical, "Importfout!"
Exit Sub

Else
On Error GoTo FileError
Dim strline As String
Dim strToSave
Dim strToSave2
Dim strToSave3
Dim strToSave4
Dim strToSave5

Dim CountStrParts
Dim CountStrParts2
Dim CountStrParts3
Dim CountStrParts4
Dim CountStrParts5
Dim CHARTOREPLACE
CHARTOREPLACE = 0
Dim x As Long 'Counter for number of records updated
x = 0
Dim y As Long 'Counter for number of records added new
y = 0


Set db = OpenDatabase(App.Path & "\" & "import.mdb")

Set rs = db.OpenRecordset("Import", dbOpenTable)
    Open FileManually For Input As #1
Do Until EOF(1)
Line Input #1, strline

'REPLACE A DOT FOR A COMMA
Do
    CHARTOREPLACE = InStr(strline, Chr$(46))
        If CHARTOREPLACE > 0 Then Mid(strline, CHARTOREPLACE) = Chr$(44)
Loop Until CHARTOREPLACE = 0
'END REPLACE

'HERE I BEGIN WITH CUTTING THE STRINT INTO 5 PARTS
'FIRST PART OF STRING
CountStrParts = InStr(1, strline, seperator) 'Search for first seperator
strToSave = Mid(strline, 1, CountStrParts - 1) 'Cut the string
'REMOVE SPACES
strToSave = Trim(strToSave) 'THIS IS THE VARIABLE I WILL USE TO CHECK IF THE RECORD EXCISTS OR NOT
'END REMOVE SPACES
'END FIRST PART

'SECOND PART
CountStrParts2 = InStr(CountStrParts + 1, strline, seperator) 'Search for second seperator
strToSave2 = Mid(strline, CountStrParts + 1, CountStrParts2 - CountStrParts - 1) 'Cut the string
'REMOVE SPACES
strToSave2 = Trim(strToSave2)
'END REMOVE SPACES
'END SECOUND PART

'THIRD PART
CountStrParts3 = InStr(CountStrParts2 + 1, strline, seperator) 'Search for third seperator
strToSave3 = Mid(strline, CountStrParts2 + 1, CountStrParts3 - CountStrParts2 - 1) 'Cut the string
'REMOVE SPACES
strToSave3 = Trim(strToSave3)
'END REMOVE SPACES
'END THIRD PART

'FOURTH PART
CountStrParts4 = InStr(CountStrParts3 + 1, strline, seperator) 'Search for fourth seperator
strToSave4 = Mid(strline, CountStrParts3 + 1, CountStrParts4 - CountStrParts3 - 1) 'Cut the string
'REMOVE SPACES
strToSave4 = Trim(strToSave4)
'END REMOVE SPACES
'END FOURTH PART

'FIFTH PART
CountStrParts5 = Len(strline) 'Search for end of string
strToSave5 = Mid(strline, CountStrParts4 + 1, CountStrParts5 - CountStrParts4) 'Cut the string
'REMOVE SPACES
strToSave5 = Trim(strToSave5)
'END REMOVE SPACES
'END FIFTH PART

'IF EXCISTS THEN UPDATE RECORD, ELSE ADD NEW
'QUERY TO CHECK IF CODE EXCIST
Dim mysql As String
mysql = "SELECT Code, Product, Price, Totalprice, Description FROM Import WHERE Code = '" & strToSave & "'"
Set rs2 = db.OpenRecordset(mysql)
Do Until rs2.EOF
List1.AddItem rs2.Fields!Code 'If code found add to listbox to get recordcount
rs2.MoveNext
Loop

If List1.ListCount = 0 Then  'Code doesnt excist, so add new
rs.AddNew
rs.Fields!Code = strToSave 'Save to field Code
rs.Fields!Product = strToSave2 'Save to field Produkt
rs.Fields!Price = strToSave3 'Save to field Inkoopprijs
rs.Fields!Totalprice = strToSave4 'Save to field Verkoopprijs
rs.Fields!Description = strToSave5 'Save to field Produktomschrijving
rs.Update
y = y + 1 'counter for number of records added new
'Else code does excist so update record
Else
Set rs2 = db.OpenRecordset(mysql)
rs2.Edit
rs2.Fields!Code = strToSave
rs2.Fields!Product = strToSave2
rs2.Fields!Price = strToSave3
rs2.Fields!Totalprice = strToSave4
rs2.Fields!Description = strToSave5
rs2.Update
x = x + 1 'Counter for number of records updated
List1.Clear 'Clear listbox
End If

Loop 'LOOP TO NEXT IMPORTLINE
'END IMPORTING INTO DATABASE

Dim xy As Long 'Number of records that have been added new/updated
xy = x + y

Picture3.Visible = True
Picture2.Visible = False
Label4.Caption = x & " records have been updated."
Label5.Caption = y & " new records have been added."
Label6.Caption = xy & " records have been checked."

Close #1 'Close the importfile
Exit Sub
End If

FileError:
Close #1
MsgBox "The importfile does not have the right format.", vbCritical, "Import error!"
End Sub

Private Sub Command4_Click()
 CommonDialog1.DialogTitle = "Textfile to import"
    CommonDialog1.Filter = "Textfiles(*.txt)|*.txt|"
CommonDialog1.ShowOpen
    Text2.Text = CommonDialog1.FileName
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Set db = OpenDatabase(App.Path & "\" & "import.mdb")

End Sub

Private Sub Form_Load()
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False

End Sub

