VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Add / Delete Field from table"
      Height          =   4095
      Left            =   6840
      TabIndex        =   23
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command9 
         Caption         =   "List Tables"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ListBox List2 
         Height          =   2400
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Delete Field"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add Field"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Enter Table Name"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Field Name"
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add / Delete Table"
      Height          =   1575
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   6735
      Begin VB.CommandButton Command8 
         Caption         =   "Delete Table"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3120
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Table"
         Height          =   375
         Left            =   5520
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Only Enter Fields if your adding a table"
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Enter Fields to Add"
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Enter Table Name"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "List / Edit / Delete Record"
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "List"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Name"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Record"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Database Tutorial v1.0
'Dustin Davis
'VB Live
'http://www.vblive.com
'
'This tutorial is to show you how to talk to databases with pure code
'no data object. This comes in handy for ASP applications or DHTML projects
'
'You will need a reference to DAO 3.6 (only if you use access 2000)
'and if you dont have it, get the service pack 3 from microsoft.com
'If you use access 98 or lower, you can use DAO 3.5
'
'Questions or comments, please send em to cheater@smallwww.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Add()
'This function will show you how to add a record to the database which
'i have created already with access 2000

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
'this opens a table inside the database
Set RS = DB.OpenRecordset("email", dbOpenTable)

'Tells the database we want to add a new record to the recordset
RS.AddNew

'Put the data in the proper fields
'RS is your recordset and ("field_name") points to the field you want to
'set the data for
RS("Name") = Text1.Text
RS("E-Mail") = Text2.Text

'Update the database. If you dont, the database will add it, but
'it wont be visible
RS.Update

'close the database
DB.Close

End Function


Private Sub Command1_Click()
'Add to the DB
Add
'clear the text boxes
Text1.Text = ""
Text2.Text = ""
'relist everything
List
End Sub

Private Sub Command2_Click()
'call the list function
List
End Sub

Private Sub Command3_Click()
'call the edit function
Edit
'clear the text boxes
Text3.Text = ""
Text4.Text = ""
'relist everything
List
End Sub

Private Sub Command4_Click()
'call the delete function
Delete
'relist everything
List
End Sub

Private Sub Edit()
'This function will show how to read from the fields

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace
Dim i As Long

If List1.SelCount = 1 Then

    'we will use this to move to that record
    i = List1.ListIndex
    
    'This sets a workspace for the database
    Set WS = DBEngine.Workspaces(0)
    'this opens the database
    Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    'this opens a table inside the database
    Set RS = DB.OpenRecordset("email", dbOpenTable)
    
    'Move to the record. NOTE: if it is off and not changing
    'the correct record, try adding one to i
    RS.Move i
    
    'Tell the Database you want to change the info
    'Or in other words, opens edit mode
    RS.Edit
    
    'change the info
    RS("Name") = Text3.Text
    RS("E-Mail") = Text4.Text
    
    'Update the DB
    RS.Update
    
    'close the database
    DB.Close
Else
    MsgBox "Select a name from the list box"
    Exit Sub
End If
End Sub


Private Function List()
'This function will show you how to list all of the records
'in a table

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace
Dim Max As Long

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
'this opens a table inside the database
Set RS = DB.OpenRecordset("email", dbOpenTable)

'Get how manby records are in the table
Max = RS.RecordCount

'Move to the begining of the file, or you can do
'RS.MoveFirst or RS.Move 1, but i prefer this
RS.Move BOF

'clear the list
List1.Clear

'do the loop
For i = 1 To Max
    'Add the data from the fields to the listbox. Notice i used
    'two different methods. One is kind of a shortcut RS!Name
    'is easy and simple. But, if you want to put in a name
    'that has a Dash '-' or use a varable like;
    'dim FieldName as String
    'FieldName = "E-Mail"
    'rs(FieldName)
    'You can not do this with the RS!FieldName Method. It dont
    'work.
    List1.AddItem RS!Name & "," & RS("E-Mail")
    
    'move to next record. You can do different things such as
    'RS.MoveNext or RS.Move i, i will use
    RS.MoveNext
Next i

End Function

Public Function Delete()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim RS As Recordset
Dim WS As Workspace
Dim i As Long

If List1.SelCount = 1 Then
    i = List1.ListIndex
    
    'This sets a workspace for the database
    Set WS = DBEngine.Workspaces(0)
    'this opens the database
    Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    'this opens a table inside the database
    Set RS = DB.OpenRecordset("email", dbOpenTable)
    
    'Move to the record
    RS.Move i
    
    'Tell the Database you want to delete the record
    RS.Delete
    'Simple isnt it?!
           
    'close the database
    DB.Close
Else
    MsgBox "Select a name from the list box"
    Exit Function
End If

End Function

Private Sub Command5_Click()
'Call add table function
Add_Table
End Sub

Public Function Add_Table()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD1 As Field
Dim FD2 As Field
Dim FD3 As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    
'Set the table info
Set TD = DB.CreateTableDef(Text5.Text)
           
'create new fields and bind it to the table.
'For the dbText, you can use dbInteger or whatever else
'you wish to set the field type to. I would stick with those
'2 though.
Set FD1 = TD.CreateField(Text6.Text, dbText)
Set FD2 = TD.CreateField(Text7.Text, dbText)
Set FD3 = TD.CreateField(Text8.Text, dbText)

'bind the Fields to the table
TD.Fields.Append FD1
TD.Fields.Append FD2
TD.Fields.Append FD3

'Now bind the table to the database
DB.TableDefs.Append TD

'close the database
DB.Close

End Function

Private Sub Command6_Click()
'Call add field function
Add_Field
End Sub

Public Function Add_Field()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    
'Set the table to open
Set TD = DB.TableDefs(Text10.Text)

'Once again, you can use dbText, or dbInterger
'or whatever else you wish to set the field type
Set FD = TD.CreateField(Text9.Text, dbText)

'Bind field to the table
TD.Fields.Append FD

'close the database
DB.Close

End Function

Public Function Delete_Field()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    
'Set the table to open
Set TD = DB.TableDefs(Text10.Text)

'Erase the field
TD.Fields.Delete Text9.Text

'close the database
DB.Close

End Function

Private Sub Command7_Click()
'call the delete field function
Delete_Field
End Sub

Private Sub Command8_Click()
'call the delete table function
Delete_Table
End Sub

Public Function Delete_Table()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    
'Set the table to open
DB.TableDefs.Delete Text5.Text

'close the database
DB.Close

End Function

Private Sub Command9_Click()
'call the list tables function
list_Tables
End Sub

Public Function list_Tables()
'This function will show how to list tables in a database

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim Temp As String
Dim Max As Long


'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(App.Path & "\dbtut.mdb")
    
'find the number of tables in the database
Max = DB.TableDefs.Count

For i = 0 To Max - 6
'Take away six from max because the last 5
'are not for you to mess with

'select the table
Set TD = DB.TableDefs(i)

'List the tables in the listbox
List2.AddItem TD.Name

Next i

'close the database
DB.Close

End Function
