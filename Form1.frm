VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid dgFileContents 
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Data"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim csvLine As String
    Dim csvFileName As String
    Dim csvFile As TextStream
    Dim fs As Scripting.FileSystemObject
    Dim csvFields() As String
    Dim numberOfLines As Integer
    Dim i As Integer
    
    csvFileName = "C:\Documents and Settings\Administrator\Desktop\ANPR_archivio_comuni.csv"
    
    Set fs = New Scripting.FileSystemObject
    Set csvFile = fs.OpenTextFile(csvFileName, ForReading)
    csvLine = csvFile.ReadLine   ' headings
    
    ' *********************  add the columns to the grid
    ' remove any that are there first
    dgFileContents.Cols = 0
    
    ' split the line on the comma and remove the quotes as we go to write the headers
    csvFields = Split(csvLine, ",")
    dgFileContents.Cols = UBound(csvFields) + 1
    dgFileContents.Rows = 1
    For i = 0 To UBound(csvFields)
        dgFileContents.TextMatrix(0, i) = Replace(csvFields(i), """", "")
    Next i
    
    numberOfLines = 1
    Label1.Caption = "Loading data from file"
    Me.MousePointer = vbHourglass
    DoEvents
    
    Do While csvFile.AtEndOfStream = False
        csvLine = csvFile.ReadLine
        csvFields = Split(csvLine, ",")
        
        For i = 0 To UBound(csvFields)
            Label1.Caption = "Loading data from file - on row " & CStr(numberOfLines)
            dgFileContents.Rows = numberOfLines + 1
            dgFileContents.TextMatrix(numberOfLines, i) = Replace(csvFields(i), """", "")
            DoEvents
        Next i
        
        numberOfLines = numberOfLines + 1
    Loop
  
    csvFile.Close
    Label1.Caption = ""
    Me.MousePointer = vbDefault
    DoEvents
    
End Sub


