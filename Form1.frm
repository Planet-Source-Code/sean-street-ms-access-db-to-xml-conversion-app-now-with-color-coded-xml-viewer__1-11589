VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MS Access to XML Conversion"
   ClientHeight    =   7815
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   7305
      Left            =   3990
      TabIndex        =   1
      Top             =   330
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   12885
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMain 
      Height          =   7305
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   12885
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   6
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cdcMain 
      Left            =   30
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Label1"
      Height          =   195
      Left            =   3210
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load Database"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile XML"
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbConnect As ADODB.Connection
Dim catDB As New ADOX.Catalog
Dim strDBLocation As String
Dim strDBName As String
Dim strXMLLocation As String
Dim intTableCount As Integer
Dim intFieldCount As Integer
Dim intCounter1 As Integer
Dim intCounter2 As Integer
Dim rstTable As ADODB.Recordset
Private Function GetType(intType As Integer) As String
    Select Case intType
        Case 2
            GetType = "Integer"
        Case 3
            GetType = "Long Ineger"
        Case 4
            GetType = "Single"
        Case 5
            GetType = "Double"
        Case 6
            GetType = "Currency"
        Case 7
            GetType = "Date/Time"
        Case 11
            GetType = "Yes/No"
        Case 17
            GetType = "Byte"
        Case 72
            GetType = "Replication ID"
        Case 202
            GetType = "Text"
        Case 203
            GetType = "Memo"
        Case 205
            GetType = "OLE Object"
        Case Else
            GetType = "Unknown"
    
    End Select
    
End Function

Private Sub CreateXML()
    Dim intTableCount As Integer

    If Len(Dir(strXMLLocation)) > 1 Then Kill strXMLLocation
    
    Open strXMLLocation For Output As #1

    Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    Print #1, "<" & UCase(strDBName) & ">"

    intTableCount = 0
    intTableCount = catDB.Tables.Count

    For intCounter1 = 0 To intTableCount - 1
        With catDB
            If (.Tables(intCounter1).Type = "TABLE") Then
            
                intTableCount = intTableCount + 1
                
                If (trvMain.Nodes("T" & Trim(intTableCount)).Checked) Then
                
                    Set rstTable = New ADODB.Recordset
                    rstTable.Open "SELECT * FROM " & Trim(.Tables(intCounter1).Name), dbConnect, adOpenForwardOnly
                    
                    Print #1, "     <" & UCase(Trim(.Tables(intCounter1).Name)) & ">"
                    
                    If (Not rstTable.EOF) And (Not rstTable.BOF) Then
                        
                        rstTable.MoveFirst
                        
                        Do While Not rstTable.EOF
                            
                            Print #1, "          <RECORD>"
                        
                            intFieldCount = rstTable.Fields.Count
                        
                            For intCounter2 = 0 To intFieldCount - 1
                                If trvMain.Nodes("T" & Trim(intTableCount) & "F" & Trim(intCounter2)).Checked Then
                                    
                                    
                                    With rstTable.Fields(intCounter2)
                                        If (.Type = 202) Or (.Type = 203) Then
                                            Print #1, "               <" & UCase(Trim(.Name)) & " Type= '" & Trim(GetType(.Type)) & "'>"
                                            Print #1, "               " & .Value
                                            Print #1, "               </" & UCase(Trim(.Name)) & ">"
                                        ElseIf Trim(GetType(.Type)) = "Unknown" Then
                                            Print #1, "               <" & UCase(Trim(.Name)) & " Type= 'Unknown'> </" & UCase(Trim(.Name)) & ">"
                                        ElseIf .Type <> 205 Then
                                            Print #1, "               <" & UCase(Trim(.Name)) & " Type= '" & Trim(GetType(.Type)) & "'> " & .Value & _
                                            " </" & UCase(Trim(.Name)) & ">"
                                        End If
                                    End With
                                    
                                End If
                            Next intCounter2
                            Print #1, "          </RECORD>"
                            rstTable.MoveNext
                        Loop
                    End If
                    
                    Print #1, "     </" & UCase(Trim(.Tables(intCounter1).Name)) & ">"
                End If
            End If
        End With
        Set rstTable = Nothing
    Next intCounter1

    Print #1, "</" & UCase(strDBName) & ">"
    Close
End Sub

Private Sub Form_Load()
    mnuCompile.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstTable = Nothing
    Set catDB = Nothing
    Set dbConnect = Nothing
End Sub

Private Sub mnuCompile_Click()
    CreateXML
    LoadXMLFile
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLoad_Click()
    On Error GoTo ErrHandle
    Dim strXMLName As String
    
    With cdcMain
        .Flags = cdlOFNFileMustExist & cdlOFNHideReadOnly
        .CancelError = True
        .Filter = "MS Access DB Files|*.MDB|"
        .ShowOpen
        strDBLocation = .FileName
        strDBName = .FileTitle
    End With
    
    strXMLName = InputBox("What would you like to name the XML file?", "XML Name")
    
    If Trim(Len(strXMLName)) = 0 Then Exit Sub
    
    If UCase(Right(strXMLName, 4)) <> ".XML" Then strXMLName = strXMLName & ".xml"
    
    strXMLLocation = Replace(strDBLocation, strDBName, strXMLName)
    
    strDBName = Replace(UCase(strDBName), ".MDB", "")
    
    Populate_TreeView
    
    mnuCompile.Enabled = True
    
    Exit Sub
ErrHandle:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        Unload Me
    End If
End Sub

Private Sub Populate_TreeView()
    Dim intTableCount As Integer
    Dim nodNode As Node

    Set dbConnect = New ADODB.Connection
    dbConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0"
    dbConnect.Open strDBLocation
    Set catDB.ActiveConnection = dbConnect
    
    trvMain.Nodes.Clear
    
    intTableCount = 0
    intTableCount = catDB.Tables.Count

    For intCounter1 = 0 To intTableCount - 1
        With catDB
            If .Tables(intCounter1).Type = "TABLE" Then
                
                intTableCount = intTableCount + 1
                
                Set nodNode = trvMain.Nodes.Add(, , "T" & Trim(intTableCount), Trim(.Tables(intCounter1).Name))
                trvMain.Nodes("T" & Trim(intTableCount)).Checked = True
                                
                Set rstTable = New ADODB.Recordset
                rstTable.Open "SELECT * FROM " & Trim(.Tables(intCounter1).Name), dbConnect
                
                intFieldCount = rstTable.Fields.Count
                
                For intCounter2 = 0 To intFieldCount - 1
                    Set nodNode = trvMain.Nodes.Add("T" & Trim(intTableCount), tvwChild, "T" & Trim(intTableCount) & "F" & Trim(intCounter2), Trim(rstTable.Fields(intCounter2).Name))
                    trvMain.Nodes.Item("T" & Trim(intTableCount) & "F" & Trim(intCounter2)).Checked = True
                Next intCounter2
                
            End If
        End With
    Next intCounter1
    
End Sub

Private Sub LoadXMLFile()
    Dim lngLastPos As Long
    Dim lngLength As Long

    rtfMain.Text = ""
    rtfMain.SelColor = vbBlack
    rtfMain.SelBold = True
    
    rtfMain.LoadFile strXMLLocation, 1
   
   rtfMain.Span ("<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>")
   lngLastPos = rtfMain.SelLength - 1
   rtfMain.SelColor = vbBlue
   rtfMain.SelBold = False
   
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find("<", lngLastPos + 1)
        If lngLastPos = -1 Then Exit Do
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
        
        lngLength = (rtfMain.Find(">", lngLastPos)) - lngLastPos
        rtfMain.SelStart = lngLastPos + 1
        rtfMain.SelLength = lngLength
        rtfMain.SelColor = &HC0&
        rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find(">", lngLastPos + 1)
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find("</", lngLastPos + 1)
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find("=", lngLastPos + 1)
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find("'", lngLastPos + 1)
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
        
        If lngLastPos = -1 Then Exit Do
        lngLength = lngLastPos
        lngLastPos = rtfMain.Find("'", lngLastPos + 1)
        
        rtfMain.SelStart = lngLength + 1
        rtfMain.SelLength = (lngLastPos - lngLength) - 1
        rtfMain.SelColor = &H8000&
        rtfMain.SelBold = True
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = rtfMain.Find("'", lngLastPos + 1)
        rtfMain.SelColor = vbBlue
        rtfMain.SelBold = False
    Loop
    
    rtfMain.SelStart = 0
    rtfMain.SetFocus
End Sub
















