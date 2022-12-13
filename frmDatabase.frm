VERSION 5.00
Begin VB.Form frmDatabase 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstThird 
      Height          =   3375
      Left            =   3840
      TabIndex        =   25
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtThird 
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Text            =   "Third"
      Top             =   1440
      Width           =   7575
   End
   Begin VB.ListBox lstSecondary 
      Height          =   3375
      Left            =   2040
      TabIndex        =   21
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtSecondary 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Text            =   "Secondary"
      Top             =   840
      Width           =   7575
   End
   Begin VB.ListBox lstCategories 
      Height          =   3375
      Left            =   7560
      TabIndex        =   14
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddCategory 
      Caption         =   "Add category"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtAddCategory 
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   3720
      Width           =   3495
   End
   Begin VB.ListBox lstCategory 
      Height          =   3375
      Left            =   5640
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtCategory 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Text            =   "Category"
      Top             =   2040
      Width           =   7575
   End
   Begin VB.ListBox lstID 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDeleteRecords 
      Caption         =   "DELETE RECORD"
      Height          =   615
      Left            =   6960
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtPrimary 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "Primary"
      Top             =   240
      Width           =   7575
   End
   Begin VB.ListBox lstPrimary 
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdListRecords 
      Caption         =   "List all records"
      Height          =   615
      Left            =   6960
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddRecord 
      Caption         =   "Add record"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreateDB 
      Caption         =   "Create Database"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Third field"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Third field"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Secondary field"
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Secondary field"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Categories"
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Field Category"
      Height          =   255
      Left            =   5640
      TabIndex        =   17
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Primary Field"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Record ID"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Primary field"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnuMain 
      Caption         =   "mnuMain"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete record"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelected 
         Caption         =   "Select category"
      End
      Begin VB.Menu mnuDeleteCategory 
         Caption         =   "Delete Category"
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDBpath As String
Dim Password As String


Private Function FileExists(ByVal sFileName As String) As Boolean
Dim intReturn As Integer

    On Error GoTo FileExists_Error
    intReturn = GetAttr(sFileName)
    FileExists = True
    
Exit Function
FileExists_Error:
    FileExists = False
End Function

Sub DeleteFromDB(ByVal theID As String)
On Error GoTo ErrorHandler


Dim daoRecordSet As DAO.Recordset
Dim dbDatabase As DAO.Database


Dim sqlStr As String



        
        Set dbDatabase = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        'dbDatabase.
        
   
   sqlStr = "DELETE * FROM Table1 WHERE ID = " & theID & ""
   
   
   dbDatabase.Execute (sqlStr)
   

   dbDatabase.Close
   
   MsgBox "Record deleted X2!"
   
   
ErrorHandler:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in CreateDB", vbInformation
End If
   
End Sub
Sub CreateDB()
On Error GoTo ErrorHandler

        
        Dim tdefMDB As TableDef, txtFieldone As Field
        Dim PrimaryOne As Field, memoFieldone As Field, dbDatabase As Database
        Dim txtFieldcategory As Field
        Dim TextField2 As Field
        Dim TextField3 As Field
        
        
        '~~> MDB to be created. Change this to relevant path and filename
        'sDBpath = "C:\my-source\MyDatabase.mdb"
        
        
        'Set dbDatabase = CreateDatabase(sDBpath, dbLangGeneral, 0)
        
        'original working code without password.
              
        
        
        Set dbDatabase = CreateDatabase(sDBpath, dbLangGeneral & _
            ";pwd=" & Password, 0)
                
          
        '~~> Create new TableDef (I am creating a table Table1)
        ' first table create
        
        Set tdefMDB = dbDatabase.CreateTableDef("Table1")
     
        'Assign fields
        
        Set PrimaryOne = tdefMDB.CreateField("ID", dbLong, dbAutoIncrField)
        Set txtFieldone = tdefMDB.CreateField("txtField1", dbText, 255)
        Set TextField2 = tdefMDB.CreateField("txtField2", dbText, 255)
        Set TextField3 = tdefMDB.CreateField("txtField3", dbText, 255)
        Set txtFieldcategory = tdefMDB.CreateField("Category", dbText, 255)
       
        PrimaryOne.Attributes = dbAutoIncrField
              
       'dateFieldone.Attributes =
       'Fields.Append fld
               
        '~~> Assign the field objects to the TableDef
        tdefMDB.Fields.Append PrimaryOne
        tdefMDB.Fields.Append txtFieldone
        tdefMDB.Fields.Append TextField2
        tdefMDB.Fields.Append TextField3
        tdefMDB.Fields.Append txtFieldcategory

     
        '~~> Save TableDef definition by appending it to TableDefs collection.
        dbDatabase.TableDefs.Append tdefMDB
        
        '-------
        'Create New table category
        Set tdefMDB = dbDatabase.CreateTableDef("Category")
     
        'Assign fields
        
        Set PrimaryOne = tdefMDB.CreateField("ID", dbLong, dbAutoIncrField)
        Set txtFieldcategory = tdefMDB.CreateField("Category", dbText, 255)
       
        PrimaryOne.Attributes = dbAutoIncrField
              
       'dateFieldone.Attributes =
       'Fields.Append fld
               
        '~~> Assign the field objects to the TableDef
        tdefMDB.Fields.Append PrimaryOne
         tdefMDB.Fields.Append txtFieldcategory

     
        '~~> Save TableDef definition by appending it to TableDefs collection.
        dbDatabase.TableDefs.Append tdefMDB
        
        
        
        
        
        '~~> Inform user.
        MsgBox "New .MDB Created - '" & sDBpath & "'", vbInformation
        
ErrorHandler:
If Err.Number = 3204 Then
' if Database already exists then create new one.


Dim answer As Integer
answer = MsgBox("The database already exists, do you wish to overwrite and create new one?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

    If answer = vbYes Then
        Kill sDBpath
        'Delete database
        
        Call CreateDB
        'Create new one.
    End If
    
    Exit Sub

End If


If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in CreateDB", vbInformation
End If



End Sub



Private Sub cmdAddCategory_Click()


On Error GoTo ErrorHandler


Dim TextCategory As String

'Text field contents

Dim dbDatabase As DAO.Database
Dim daoRecordSet As DAO.Recordset

Dim db As DAO.Database
Dim rs As DAO.Recordset


TextCategory = txtAddCategory.Text

        'sDBpath = "app.path & \MyDatabase.mdb"
        Set db = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        Set rs = db.OpenRecordset("Category")
       
'rs.MoveFirst

'search through records to check if exists.
'if exists then abort.

Do While Not rs.EOF
  Debug.Print rs!category
   
   If rs!category = TextCategory Then
   
        MsgBox TextCategory & " already exists!"
        'if found - then abort
        
        db.Close
        
    Exit Sub
   End If
   
  rs.MoveNext
Loop
   
       
   db.Close
   
     
  
   'sDBpath = assigned on form load
   'password = assigned on form load
   
    Set dbDatabase = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        'dbDatabase.
        
   Set daoRecordSet = dbDatabase.OpenRecordset("Category")
   
   daoRecordSet.AddNew
   daoRecordSet!category = TextCategory
     
    
   daoRecordSet.Update
   dbDatabase.Close
   
   MsgBox "Record added!"
   
   
   
ErrorHandler:

If Err.Number = 3024 Then
' if Database doesn't exists then create new one.


Dim answer As Integer
answer = MsgBox("The database doesn't exist, do you wish to create new one?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

    If answer = vbYes Then
       
    
    If FileExists(sDBpath) Then
        'MsgBox "File Exist."
        Kill sDBpath
    Else
        'MsgBox "File Doesn't Exist."
    End If

            
        'Delete database
        
        Call CreateDB
        'Create new one.
    End If
    
Exit Sub

End If



If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in AddCategory", vbInformation
End If







End Sub

Private Sub cmdAddRecord_Click()
On Error GoTo ErrorHandler

'add new record to database

Dim TextCon As String
Dim TextTwo As String
Dim TextThree As String

Dim TextCategory As String

'Text field contents

Dim daoRecordSet As DAO.Recordset
Dim dbDatabase As DAO.Database


TextCon = txtPrimary.Text

Dim db As DAO.Database
Dim rs As DAO.Recordset

        
        'sDBpath = "app.path & \MyDatabase.mdb"
        Set db = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        Set rs = db.OpenRecordset("Table1")
       
'rs.MoveFirst

'search through records to check if exists.
'if exists then abort.

Do While Not rs.EOF
  Debug.Print rs!txtField1
   
   If rs!txtField1 = TextCon Then
   
        MsgBox TextCon & " already exists!"
        'if found - then abort
        
        db.Close
        
    Exit Sub
   End If
   
  rs.MoveNext
Loop
   
       
   db.Close
   


TextTwo = txtSecondary.Text
TextCategory = txtCategory.Text
TextThree = txtThird.Text




        'sDBpath = "C:\my-source\MyDatabase.mdb"
        Set dbDatabase = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        'dbDatabase.
     
   
   Set daoRecordSet = dbDatabase.OpenRecordset("Table1")
 
   daoRecordSet.AddNew
   'daorecordset
   daoRecordSet!txtField1 = TextCon
   daoRecordSet!txtField2 = TextTwo
   daoRecordSet!txtField3 = TextThree
   daoRecordSet!category = TextCategory
   'MsgBox TextCategory
   Debug.Print TextCategory
    

 
   daoRecordSet.Update
   dbDatabase.Close
   
   MsgBox "Record added!"
   
ErrorHandler:

If Err.Number = 3024 Then
' if Database doesn't exists then create new one.


Dim answer As Integer
answer = MsgBox("The database doesn't exist, do you wish to create new one?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

    If answer = vbYes Then
       
    
    If FileExists(sDBpath) Then
        'MsgBox "File Exist."
        Kill sDBpath
    Else
        'MsgBox "File Doesn't Exist."
    End If

            
        'Delete database
        
        Call CreateDB
        'Create new one.
    End If
    
Exit Sub

End If



If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in AddRecord", vbInformation
End If




End Sub


Private Sub cmdCreateDB_Click()
Call CreateDB
End Sub

Private Sub cmdDeleteRecords_Click()
On Error GoTo ErrorHandler

'delete record from database

Dim TextCon As String
'text Fileds contents

Dim daoRecordSet As DAO.Recordset
Dim dbDatabase As DAO.Database

Dim sqlStr As String


TextCon = txtPrimary.Text


        'sDBpath = "C:\my-source\MyDatabase.mdb"
        Set dbDatabase = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        'dbDatabase.
        
   
   sqlStr = "DELETE * FROM Table1 WHERE txtField1 = '" & TextCon & "'"
   
   
   dbDatabase.Execute (sqlStr)
      

   dbDatabase.Close
   
   MsgBox "Record deleted!"
   
ErrorHandler:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in DeleteRecord", vbInformation
End If
End Sub

Private Sub cmdListRecords_Click()
On Error GoTo ErrorHandler
On Error Resume Next



'list database records

lstCategories.Clear
lstCategory.Clear

lstPrimary.Clear
lstID.Clear
lstSecondary.Clear
lstThird.Clear





Dim db As DAO.Database
Dim rs As DAO.Recordset

Dim sqlStr As String

On Error GoTo ErrorHandler



sqlStr = "SELECT * FROM Category as c WHERE 1"

        'sDBpath = "C:\my-source\MyDatabase.mdb"
        Set db = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
            
        Set rs = db.OpenRecordset(sqlStr)
              

        'sDBpath = "C:\my-source\MyDatabase.mdb"
        'Set db = OpenDatabase(sDBpath)
        'Set rs = db.OpenRecordset(sqlStr)
              

rs.MoveFirst


              'MsgBox rs.RecordCount

Do While Not rs.EOF
  'Debug.Print (" txtPrimary: " & rs!txtField1)
  'lstPrimary.AddItem rs!txtField1
  'lstID.AddItem rs!ID
  lstCategories.AddItem rs!category
  rs.MoveNext
Loop


On Error GoTo ErrorHandlerTwo

sqlStr = "SELECT * FROM Table1 as c WHERE 1"

Set db = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
Set rs = db.OpenRecordset(sqlStr)


rs.MoveFirst

Do While Not rs.EOF
  'Debug.Print (" txtPrimary: " & rs!txtField1)
  lstPrimary.AddItem rs!txtField1
  lstSecondary.AddItem rs!txtField2
  lstThird.AddItem rs!txtField3
  
  lstID.AddItem rs!ID
  lstCategory.AddItem rs!category
  rs.MoveNext
Loop




db.Close

MsgBox ("End of list")


Dim answer As Integer

ErrorHandler:

If Err.Number = 3021 Then
MsgBox "no existing category records"
    Exit Sub
    
End If



If Err.Number = 3024 Then
' if Database doesn't exists then create new one.


answer = MsgBox("The database doesn't exist, do you wish to create new one?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

    If answer = vbYes Then
       
    
    If FileExists(sDBpath) Then
        'MsgBox "File Exist."
        Kill sDBpath
    Else
        'MsgBox "File Doesn't Exist."
    End If

            
        'Delete database
        
        Call CreateDB
        'Create new one.
    End If
    
Exit Sub

End If



If Err.Number <> 0 Then
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in ListRecord", vbInformation

End If



ErrorHandlerTwo:

If Err.Number = 3021 Then
MsgBox "no existing table records"
    Exit Sub
    
End If



If Err.Number = 3024 Then
' if Database doesn't exists then create new one.


answer = MsgBox("The database doesn't exist, do you wish to create new one?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

    If answer = vbYes Then
       
    
    If FileExists(sDBpath) Then
        'MsgBox "File Exist."
        Kill sDBpath
    Else
        'MsgBox "File Doesn't Exist."
    End If

            
        'Delete database
        
        Call CreateDB
        'Create new one.
    End If
    
Exit Sub

End If



If Err.Number <> 0 Then
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in ListRecord", vbInformation

End If
End Sub

Private Sub Form_Load()

sDBpath = App.Path & "\911Database.mdb"
Password = "foobar666"

End Sub

Private Sub lstPrimary_Click()
Dim i As Integer

For i = 0 To lstPrimary.ListCount - 1
If lstPrimary.Selected(i) = True Then

PopupMenu mnuMain

End If


Next

End Sub

Private Sub lstCategories_Click()
Dim i As Integer

For i = 0 To lstCategories.ListCount - 1
If lstCategories.Selected(i) = True Then

PopupMenu mnuSelect

End If


Next
End Sub

Private Sub mnuDelete_Click()
On Error GoTo ErrorHandler

Dim X As Integer

For i = 0 To lstPrimary.ListCount - 1
If lstPrimary.Selected(i) = True Then

X = lstID.List(i)
'MsgBox X

Call DeleteFromDB(X)




End If


Next

ErrorHandler:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in CreateDB", vbInformation
End If

End Sub

Private Sub mnuDeleteCategory_Click()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs As DAO.Recordset

Dim sqlStr As String
Dim X As String


For i = 0 To lstCategories.ListCount - 1
If lstCategories.Selected(i) = True Then

X = lstCategories.List(i)
MsgBox X
'Debug.Print X


        'sDBpath = "C:\my-source\MyDatabase.mdb"
        Set dbDatabase = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
        'dbDatabase.
        
   
        sqlStr = "DELETE * FROM Category WHERE Category = '" & X & "'"
   
   
        dbDatabase.Execute (sqlStr)
        dbDatabase.Close
   
        MsgBox "Record deleted!"


End If
Next i


ErrorHandler:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in mnuDeletCategory", vbInformation
End If



End Sub

Private Sub mnuSelected_Click()
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs As DAO.Recordset

Dim sqlStr As String
Dim X As String


For i = 0 To lstCategories.ListCount - 1
If lstCategories.Selected(i) = True Then

X = lstCategories.List(i)
'MsgBox X
Debug.Print X



sqlStr = "SELECT * FROM Table1 as c WHERE Category='" & X & "'"

Set db = OpenDatabase(sDBpath, 0, 0, ";pwd=" & Password)
Set rs = db.OpenRecordset(sqlStr)

If rs.RecordCount > 0 Then

lstPrimary.Clear
lstID.Clear
lstCategory.Clear
'lstCategories.Clear
lstSecondary.Clear
lstThird.Clear



End If


rs.MoveFirst

Do While Not rs.EOF
  'Debug.Print (" txtPrimary: " & rs!txtField1)
  lstPrimary.AddItem rs!txtField1
  lstSecondary.AddItem rs!txtField2
  lstThird.AddItem rs!txtField3
  
  lstID.AddItem rs!ID
  lstCategory.AddItem rs!category
  rs.MoveNext
Loop




db.Close

MsgBox ("End of list")





End If


Next

ErrorHandler:
If Err.Number <> 0 Then
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in mnuSelected", vbInformation
End If

End Sub

