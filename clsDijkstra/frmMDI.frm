VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "DijkstraSPaths 1.1"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9510
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2580
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   9510
      TabIndex        =   13
      Top             =   5940
      Width           =   9510
      Begin VB.Frame fraSolution 
         Caption         =   "Solution"
         Height          =   825
         Left            =   6795
         TabIndex        =   23
         Top             =   60
         Width           =   2640
         Begin VB.Label lblExplored 
            Caption         =   "Graph Explored = 0 %"
            Height          =   255
            Left            =   135
            TabIndex        =   29
            Top             =   525
            Width           =   2355
         End
         Begin VB.Label lblCost 
            Caption         =   "Cost = 0"
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   255
            Width           =   2355
         End
      End
      Begin VB.Frame fraCoords 
         Caption         =   "Coordinates"
         Height          =   825
         Left            =   2505
         TabIndex        =   21
         Top             =   60
         Width           =   1620
         Begin VB.Label lblY 
            Caption         =   "Y = 0"
            Height          =   225
            Left            =   135
            TabIndex        =   26
            Top             =   525
            Width           =   1290
         End
         Begin VB.Label lblX 
            Caption         =   "X = 0"
            Height          =   225
            Left            =   135
            TabIndex        =   22
            Top             =   255
            Width           =   1290
         End
      End
      Begin VB.Frame fraTime 
         Caption         =   "Processing Time"
         Height          =   825
         Left            =   4185
         TabIndex        =   19
         Top             =   60
         Width           =   2550
         Begin VB.Label lblDrawTime 
            Caption         =   "Drawing time = 0 sec."
            Height          =   195
            Left            =   135
            TabIndex        =   28
            Top             =   525
            Width           =   2265
         End
         Begin VB.Label lblDijTime 
            Caption         =   "Dijksta time = 0 sec."
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   255
            Width           =   2265
         End
      End
      Begin VB.Frame fraNodesEdges 
         Caption         =   "Graph Properties"
         Height          =   825
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   2385
         Begin VB.Label lblConnections 
            Caption         =   "0 Connections"
            Height          =   240
            Left            =   135
            TabIndex        =   27
            Top             =   525
            Width           =   2115
         End
         Begin VB.Label lblNodes 
            Caption         =   "0 Nodes"
            Height          =   240
            Left            =   135
            TabIndex        =   15
            Top             =   255
            Width           =   2085
         End
      End
   End
   Begin VB.PictureBox pctControls 
      Align           =   3  'Align Left
      Height          =   5940
      Left            =   0
      ScaleHeight     =   5880
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2475
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   1515
         Left            =   60
         TabIndex        =   30
         Top             =   3900
         Width           =   2325
         Begin VB.TextBox txtMDI 
            Height          =   300
            Index           =   3
            Left            =   1050
            TabIndex        =   9
            Text            =   "1"
            ToolTipText     =   "Run algorithm X times in a row."
            Top             =   300
            Width           =   1005
         End
         Begin VB.CheckBox cbTileDiag 
            Caption         =   "Allow diagonal edges"
            Enabled         =   0   'False
            Height          =   240
            Left            =   210
            TabIndex        =   11
            ToolTipText     =   "Allow diagonal edges for tile instances"
            Top             =   1140
            Width           =   1845
         End
         Begin VB.CheckBox cbPruneGraph 
            Caption         =   "Prune imported graph"
            Height          =   315
            Left            =   210
            TabIndex        =   10
            ToolTipText     =   "Prune sub-trees from imported graph"
            Top             =   780
            Width           =   2025
         End
         Begin VB.Label lblBenchmark 
            Caption         =   "Benchmark Iterations:"
            Height          =   405
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame fraAPSP 
         Caption         =   "All-Pairs Shortest Paths"
         Enabled         =   0   'False
         Height          =   1605
         Left            =   60
         TabIndex        =   25
         Top             =   2190
         Width           =   2325
         Begin VB.CheckBox cbShowProgress 
            Caption         =   "Show Progress Bar"
            Height          =   255
            Left            =   105
            TabIndex        =   8
            ToolTipText     =   "If disabled the app will stop responding until the calculations are finished."
            Top             =   1230
            Value           =   1  'Checked
            Width           =   2100
         End
         Begin VB.CommandButton cmdAPSP 
            Caption         =   "Save to file"
            Height          =   405
            Index           =   1
            Left            =   105
            TabIndex        =   7
            ToolTipText     =   "Save the All-Pairs Shortest paths to a binary file. Lots of disk space might be required."
            Top             =   735
            Width           =   2115
         End
         Begin VB.CommandButton cmdAPSP 
            Caption         =   "Calculate"
            Height          =   405
            Index           =   0
            Left            =   105
            TabIndex        =   6
            ToolTipText     =   "Calculates the All-Pairs Shortest Paths (they are not kept in RAM, though)."
            Top             =   255
            Width           =   2115
         End
      End
      Begin VB.Frame fraSP 
         Caption         =   "Shortest Paths"
         Enabled         =   0   'False
         Height          =   1935
         Left            =   60
         TabIndex        =   16
         Top             =   135
         Width           =   2325
         Begin VB.CommandButton cmdReverse 
            Caption         =   "R"
            Height          =   405
            Left            =   1860
            TabIndex        =   3
            ToolTipText     =   "Reverse Source/Destination"
            Top             =   465
            Width           =   330
         End
         Begin VB.CheckBox cbSPtrees 
            Caption         =   "View Shortest Path Trees"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Could also be thought as a service area around the Source."
            Top             =   1545
            Width           =   2175
         End
         Begin VB.TextBox txtMDI 
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   2
            Text            =   "1"
            ToolTipText     =   "Enter 0 for unspecified destination"
            Top             =   705
            Width           =   885
         End
         Begin VB.TextBox txtMDI 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   1
            Text            =   "1"
            Top             =   345
            Width           =   885
         End
         Begin VB.CommandButton cmdCalcSP 
            Caption         =   "Calculate"
            Height          =   405
            Left            =   105
            TabIndex        =   4
            ToolTipText     =   "Calculate the single-source shortest paths."
            Top             =   1080
            Width           =   2115
         End
         Begin VB.Label lblSPdestination 
            Caption         =   "Destination:"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   780
            Width           =   945
         End
         Begin VB.Label lblSPsource 
            Caption         =   "Source:"
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   390
            Width           =   555
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate Shortest Path"
         Height          =   375
         Left            =   3540
         TabIndex        =   12
         Top             =   90
         Width           =   1875
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mFileImport 
         Caption         =   "&Import..."
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mFileRandom 
         Caption         =   "&Random Graph..."
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mWindow 
      Caption         =   "&Window"
      Begin VB.Menu mWindowHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mWindowVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mWindowCloseAll 
         Caption         =   "C&lose All Windows"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MDI FORM
'
'Not a lot worth mentioning here...
'Management of the MDIchildren, textbox validation, command issuing etc. etc.

Option Explicit

Private blUpdate_cbDaig As Boolean
Private ActiveChild     As Long
Private MDIchildren()   As frmNewInstance

Public Property Get GetActiveChild() As Long
    GetActiveChild = ActiveChild
End Property

Private Sub MDIForm_Load()
    mAbout_Click '< rem it out if it annoys you
    
    CommonDialog1.InitDir = App.Path
    blUpdate = cbShowProgress.Value
    blPruneGraph = cbPruneGraph.Value
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    mWindowCloseAll_Click
End Sub

Private Sub mFileNew_Click()
Dim i As Long
    If (Not MDIchildren) = -1 Then
        ReDim MDIchildren(1)
    Else
        ReDim Preserve MDIchildren(UBound(MDIchildren) + 1)
    End If
    i = UBound(MDIchildren)
    Set MDIchildren(i) = New frmNewInstance
    MDIchildren(i).ID = i
    MDIchildren(i).Show
End Sub

Private Sub mFileImport_Click()
    With CommonDialog1
        .FileName = vbNullString
        .Filter = "ASCII Graph (*.ga)|*.ga|Tile Map (*.at))|*.at"
        .Filter = .Filter & "|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .ShowOpen
        If LenB(.FileName) <> 0 Then
            If Dir(.FileName) <> vbNullString Then
                .InitDir = .FileName
                Select Case Trim$(LCase$(Right$(.FileName, 3)))
                Case ".at" ' Tile Map
                    MDIchildren(ActiveChild).ImportBinaryA (Trim$(.FileName))
                Case ".ga" ' ASCII Graph
                    MDIchildren(ActiveChild).ImportGraphData (Trim$(.FileName))
                Case Else  ' Image -> Tile Map
                    frmPic2at.Show
                    .FileName = frmPic2at.mConvert(Trim$(.FileName))
                    If LenB(.FileName) <> 0 Then
                        If Dir(.FileName) <> vbNullString Then
                            MDIchildren(ActiveChild).ImportBinaryA (Trim$(.FileName))
                        End If
                    End If
                End Select
            End If
        End If
    End With
End Sub

Private Sub mFileRandom_Click()
Dim iFile         As Integer
Dim i             As Long
Dim lNodes        As Long
Dim lEdgesPerNode As Long
Dim BenchStep     As Long

    On Error GoTo CreationError
    
    lNodes = Int(InputBox("Enter number of nodes", , "100"))
    lEdgesPerNode = Int(InputBox("Enter number of edges per node, or 0 for benchmark", , "10"))
    
    If lNodes > 0 Then
        If lEdgesPerNode > 0 Then
            If lEdgesPerNode <= lNodes Then
                MDIchildren(ActiveChild).RandomGraph lNodes, lEdgesPerNode
            End If
        Else
            'This Benchmark creates a random graph with an increasing number of edges
            'and repeatedly runs All-Pairs Shortest Paths. The elapsed running times
            'are saved in a tab-telimeted file.
            BenchStep = Int(InputBox("Enter benchmark stepping (1 to " & lNodes & ")", , "1"))
            If BenchStep > 0 Then
                If BenchStep <= lNodes Then
                    With CommonDialog1
                        .FileName = vbNullString
                        .Filter = "Text (*.txt)|*.txt"
                        .ShowSave
                        If LenB(.FileName) Then
                            iFile = FreeFile
                            Open .FileName For Output As iFile
                            Print #iFile, "Total Nodes=" & lNodes
                            
                            blBenchMark = True
                            For i = 1 To lNodes Step BenchStep
                                MDIchildren(ActiveChild).RandomGraph lNodes, i, True
                                SetRTime
                                If MDIchildren(ActiveChild).CalculateShortestPaths(, , 1) Then
                                    Exit For
                                End If
                                Print #iFile, i * lNodes & vbTab & GetRTime
                                DoEvents
                            Next i
                            Unload frmProgress ' Update #1 (unload only once, for the user to be able to press Stop)
                            blBenchMark = False
                            
                            Close #iFile
                        End If
                    End With
                End If
            End If
        End If
    End If
    Exit Sub
CreationError:
MsgBox "Errors encountered while creating random graph"
Me.MousePointer = vbDefault
End Sub

Public Sub GetChildFocus(ByVal i As Long)
    ActiveChild = i
    mFileImport.Enabled = True
    mFileRandom.Enabled = True
    With MDIchildren(ActiveChild)
        lblNodes.Caption = FormatNumber(.GNodes, 0) & " Nodes"
        lblConnections.Caption = FormatNumber(.GConnections, 0) & " Connections"
        fraSP.Enabled = .IsInit
        fraAPSP.Enabled = .IsInit

        If .IsTile Then
            cbTileDiag.Enabled = True
            If cbTileDiag.Value = 1 Then
                If .IsTile = 1 Then
                    blUpdate_cbDaig = True
                    cbTileDiag.Value = 0
                End If
            Else
                If .IsTile = 2 Then
                    blUpdate_cbDaig = True
                    cbTileDiag.Value = 1
                End If
            End If
        Else
            cbTileDiag.Enabled = False
        End If
        
    End With
End Sub

Public Sub DeleteChild(ByVal i As Long)
Dim UB As Long
    
    If i > 0 Then
        If i <= UBound(MDIchildren) Then
            UB = UBound(MDIchildren)
            mFileImport.Enabled = False
            mFileRandom.Enabled = False
            fraSP.Enabled = False
            fraAPSP.Enabled = False
            cbTileDiag.Enabled = False

            Set MDIchildren(i) = MDIchildren(UB)
            MDIchildren(i).ID = i
            Set MDIchildren(UB) = Nothing
            ReDim Preserve MDIchildren(UB - 1)
        End If
    End If
End Sub

Private Sub txtMDI_KeyPress(Index As Integer, _
                            KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) Then
        If Not (KeyAscii = 3 Or KeyAscii = 22 Or KeyAscii = 8) Then
             If KeyAscii = 13 Then cmdCalcSP_Click
                KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtMDI_Change(Index As Integer)
Dim lTemp As Long
    On Error Resume Next
    
    lTemp = Int(Val(txtMDI(Index).Text))
    If lTemp > MDIchildren(ActiveChild).GNodes Then
        If Index < 3 Then lTemp = 1
    End If
    If lTemp = 0 Then
        If Index <> SpDestination Then lTemp = 1
    End If
    txtMDI(Index).Text = lTemp
    MDIchildren(ActiveChild).ChangeText Index, lTemp
End Sub

Private Sub cmdReverse_Click()
Dim strTemp As String

    strTemp = txtMDI(SpSource).Text
    txtMDI(SpSource).Text = txtMDI(SpDestination).Text
    txtMDI(SpDestination).Text = strTemp
    cmdCalcSP_Click
End Sub

Private Sub cmdCalcSP_Click()
    MDIchildren(ActiveChild).CalculateShortestPaths Int(txtMDI(SpSource)), Int(txtMDI(SpDestination))
End Sub

Private Sub cmdAPSP_Click(Index As Integer)
    MDIchildren(ActiveChild).CalculateShortestPaths , , 1 + Index
End Sub

Private Sub cbSPtrees_Click()
    blSPTrees = cbSPtrees.Value
    cmdCalcSP_Click
End Sub

Private Sub cbShowProgress_Click()
    blUpdate = cbShowProgress.Value
End Sub

Private Sub cbTileDiag_Click()
'//Update 1.1 (Now recalculates the shortest path after change)
    blAllowTileDiag = cbTileDiag.Value
    If Not blUpdate_cbDaig Then
        If MDIchildren(ActiveChild).IsTile Then
            txtMDI(SpSource).Tag = txtMDI(SpSource).Text
            txtMDI(SpDestination).Tag = txtMDI(SpDestination).Text
            
            MDIchildren(ActiveChild).ImportBinaryA
            
            txtMDI(SpSource).Text = txtMDI(SpSource).Tag
            txtMDI(SpDestination).Text = txtMDI(SpDestination).Tag
            
            cmdCalcSP_Click
        End If
    End If
    blUpdate_cbDaig = False
End Sub

Private Sub cbPruneGraph_Click()
    blPruneGraph = cbPruneGraph.Value
    On Error GoTo ItIsNothing
    With MDIchildren(ActiveChild)
        If .IsInit Then
            If .IsTile Then
                .ImportBinaryA
            Else
                .ImportGraphData (.FileName)
            End If
        End If
    End With
ItIsNothing:
End Sub

Private Sub mWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mWindowHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mWindowVertically_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mWindowCloseAll_Click()
Dim i As Long
Dim UB As Long
    On Error GoTo NoArray
        UB = UBound(MDIchildren)
        For i = UB To 1 Step -1
           Unload MDIchildren(i)
        Next i
NoArray:
End Sub

Private Sub mAbout_Click()
    MsgBox "Created by Stavros Sirigos." & vbLf & "<ssirig@uth.gr>" & vbLf & vbLf _
           & "Quick instructions: (can also be accessed through Help->About)" & vbLf & vbLf _
           & "File->New instance" & vbLf _
           & "File->Import.. or File->Random.." & vbLf _
           & "Left mouse button: Set Destination" & vbLf _
           & "Right mouse button: Set Source" & vbLf _
           & "Middle mouse button + drag: Zoom to selected nodes" & vbLf _
           & "Middle mouse button: Zoom out" & vbLf & vbLf _
           & ">>Do not miss the new Import Image function!! :)<<" & vbLf & vbLf _
           & "No graph editing capabilities yet, maybe in the next version!" & vbLf & vbLf _
           & "Have fun!"
End Sub

Private Sub mFileExit_Click()
    Unload Me
End Sub
