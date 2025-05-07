VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "ex8.frx":0000
      Height          =   5655
      Left            =   9000
      TabIndex        =   31
      Top             =   960
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9975
      _Version        =   393216
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   1455
      Left            =   11880
      Top             =   7800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2566
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   "dbms"
      RecordSource    =   "select * from lib_manage"
      UserName        =   "system"
      Password        =   "tiger"
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton updatedata 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   2400
      TabIndex        =   29
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton deletedata 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   4560
      TabIndex        =   28
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton viewdata 
      Caption         =   "VIEW"
      Height          =   495
      Left            =   6840
      TabIndex        =   27
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cleardata 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   480
      TabIndex        =   26
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton reportdata 
      Caption         =   "REPORT"
      Height          =   495
      Left            =   2400
      TabIndex        =   25
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton insertdata 
      BackColor       =   &H00C0C0C0&
      Caption         =   "INSERT"
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox bookname 
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox customerid 
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox customername 
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox penalty 
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox submissiondate 
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox returndate 
      Height          =   375
      Left            =   6600
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox issuedate 
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox publication 
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox edition 
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox authorid 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox authorname 
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox bookid 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "LIBRARY MANAGEMENT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   30
      Top             =   150
      Width           =   5295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "CUSTOMER ID"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "BOOK NAME"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "AUTHOR NAME"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "AUTHOR ID"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "EDITION"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "PUBLICATION"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "PENALTY"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "CUSTOMER NAME"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "ISSUE DATE"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "RETURN DATE"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "SUBMISSION DATE"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "BOOK ID"
      DataSource      =   "MSRDC1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As Connection
Public rs As Recordset
Public wrk As Workspace

Private Sub dbms()
Set wrk = CreateWorkspace("wrk", System, tiger, dbUseODBC)
Set con = wrk.OpenConnection("con", dbDriverNoPrompt, False, "odbc;uid=system;pwd=tiger;dsn=dbms")
MsgBox ("PROJECT CONNECTED SUCESSFULLY")
End Sub

Private Sub insertdata_Click()
dbms
MsgBox "insert into lib_manage values(" & Val(bookid.Text) & ",'" & Trim(bookname.Text) & "','" & Trim(authorname.Text) & "'," & Val(authorid.Text) & ",'" & Trim(edition.Text) & "','" & Trim(publication.Text) & "','" & Trim(issuedate.Text) & "','" & Trim(returndate.Text) & "','" & Trim(submissiondate.Text) & "'," & Val(penalty.Text) & ",'" & Trim(customername.Text) & "'," & Val(customerid.Text) & ")"
con.Execute ("insert into lib_manage values (" & Val(bookid.Text) & ",'" & Trim(bookname.Text) & "','" & Trim(authorname.Text) & "'," & Val(authorid.Text) & ",'" & Trim(edition.Text) & "','" & Trim(publication.Text) & "','" & Trim(issuedate.Text) & "','" & Trim(returndate.Text) & "','" & Trim(submissiondate.Text) & "'," & Val(penalty.Text) & ",'" & Trim(customername.Text) & "'," & Val(customerid.Text) & ")")
MsgBox "record inserted"
End Sub
Private Sub penalty_GotFocus()
n = DateDiff("d", returndate.Text, submissiondate.Text)
Dim p As Integer
p = 0
For i = 1 To n Step 1
p = p + 5
Next i
penalty.Text = p
End Sub

Private Sub updatedata_Click()
dbms
con.Execute ("update lib_manage set customer_name='" & Trim(customername.Text) & "' where customer_id=" & Val(customerid.Text) & "")
MsgBox ("record updated")
End Sub

Private Sub deletedata_Click()
con.Execute ("delete from lib_manage where customer_id=" & Val(customerid.Text) & "")
MsgBox ("record deleted")
End Sub

Private Sub viewdata_Click()
dbms
Set rs = con.OpenRecordset("select*from lib_manage where customer_id='" & Trim(customerid.Text) & "'")
bookid.Text = rs.Fields(0)
bookname.Text = rs.Fields(1)
authorname.Text = rs.Fields(2)
authorid.Text = rs.Fields(3)
edition.Text = rs.Fields(4)
publication.Text = rs.Fields(5)
issuedate.Text = rs.Fields(6)
returndate.Text = rs.Fields(7)
submissiondate.Text = rs.Fields(8)
penalty.Text = rs.Fields(9)
customername.Text = rs.Fields(10)
customerid.Text = rs.Fields(11)
End Sub

Private Sub cleardata_click()
bookid.Text = ""
bookname.Text = ""
authorname.Text = ""
authorid.Text = ""
edition.Text = ""
publication.Text = ""
issuedate.Text = ""
returndate.Text = ""
submissiondate.Text = ""
penalty.Text = ""
customername.Text = ""
customerid.Text = ""
End Sub

Private Sub reportdata_Click()
DataReport1.Show
End Sub




