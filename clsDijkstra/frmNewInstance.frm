VERSION 5.00
Begin VB.Form frmNewInstance 
   BackColor       =   &H00FFFFFF&
   Caption         =   "New Instance"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2790
   Icon            =   "frmNewInstance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   2790
   Begin VB.Timer ResizeTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   90
      Top             =   2220
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      DrawStyle       =   2  'Dot
      Height          =   2010
      Left            =   555
      ScaleHeight     =   1950
      ScaleWidth      =   2130
      TabIndex        =   0
      Top             =   630
      Width           =   2190
   End
   Begin VB.Timer MouseTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   1785
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   1710
      Left            =   45
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   45
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "frmNewInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Few comments here, due to lack of time resources :) .
'It will probably be commented in the next version!

Option Explicit

Private Type ZoomRect
    Xmin As Single
    Xmax As Single
    Ymin As Single
    Ymax As Single
End Type

Private Zoom          As ZoomRect

Private MyID          As Long                                   'ID that corresponds to MDIchildren() index
Private MyFile        As String                                 'The last file imported

Private blInit        As Boolean                                '=True if a graph has been initialized
Private blCoord       As Boolean                                '=True if node coordinates exist
Private blIsRunning   As Boolean                                '=True if algorithm is running
Private blZoomStart   As Boolean

Private iLastButton   As Integer                                'Last mouse button clicked
Private lTileInstance As Long                                   '0: no tile instance, 1:tile instance
                                                                '2: tile instance with diagonals

Private lTextOnMDI(2) As Long                                   'Holds the frmMDI textbox values

Private Dijkstra      As clsDijkstra
Private TheGraph      As clsGraph
Private SimpleDraw    As clsSimpleDraw

Public Property Get ID() As Long
    ID = MyID
End Property

Public Property Let ID(ByVal ChildID As Long)
    MyID = ChildID
End Property

Public Property Get IsInit() As Boolean
    IsInit = blInit
End Property

Public Property Get IsTile() As Long
    IsTile = lTileInstance
End Property

Public Property Get FileName() As String
    FileName = MyFile
End Property

Public Property Get GNodes() As Long
    If blInit Then
        GNodes = TheGraph.GNodes
    End If
End Property

Public Property Get GEdges() As Long
    If blInit Then
        GEdges = TheGraph.GEdges
    End If
End Property

Public Property Get GConnections() As Long
    If blInit Then
        GConnections = TheGraph.GConnections
    End If
End Property

Private Sub Form_Load()
    Icon = frmMDI.Icon
    Picture1.BorderStyle = 0
    Picture2.BorderStyle = 0
    Picture1.Top = 60
    Picture1.Left = 60
    lTextOnMDI(SpSource) = 1
    lTextOnMDI(SpDestination) = 1
    If WindowState = 0 Then
        If GlobalFormWidth Then
            Me.Width = GlobalFormWidth
        End If
        If GlobalFormHeight Then
            Me.Height = GlobalFormHeight
        End If
    End If
End Sub

Private Sub Form_Resize()
    If blInit Then
        If WindowState = 0 Then
            GlobalFormWidth = Me.Width
            GlobalFormHeight = Me.Height
        End If
    End If
    Picture1.Width = Me.Width - 270
    Picture2.Width = Me.Width - 270
    If Me.Height - 570 > 0 Then
        Picture1.Height = Me.Height - 570
        Picture2.Height = Me.Height - 570
    End If

    ResizeTimer.Enabled = False
    If blInit Then
        ResizeTimer.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Dijkstra = Nothing
    Set SimpleDraw = Nothing
    Set TheGraph = Nothing
    frmMDI.DeleteChild (MyID)
End Sub

Private Sub ResizeTimer_Timer()
    '// We update the resized map only after the timer interval.
    '   Usefull for slow drawing, large graphs
    SimpleDraw.ResizeMap
    SetRTime
    SimpleDraw.DrawNet &HC0C0C0
    RedrawShortestPath
    ResizeTimer.Enabled = False
    frmMDI.lblDrawTime.Caption = "Drawing time = " & Round(GetRTime, 3) & " sec."
End Sub

Public Sub GetFocus()
    With frmMDI
        .GetChildFocus (MyID)
        .txtMDI(SpSource) = lTextOnMDI(SpSource)
        .txtMDI(SpDestination) = lTextOnMDI(SpDestination)
        .cbTileDiag.Enabled = IIf(lTileInstance, True, False)
    End With
    If LenB(Trim$(MyFile)) Then
        Me.Caption = Mid$(MyFile, InStrRev(MyFile, "\") + 1)
    Else
        Me.Caption = "[Untitled - " & MyID & "]"
    End If
End Sub

Public Function ChangeText(txtID As Integer, _
                           Optional ByVal txtNew As Long) As Long
    If Not txtID = 3 Then
        ChangeText = lTextOnMDI(txtID)
        If txtNew Or txtID = SpDestination Then
            lTextOnMDI(txtID) = txtNew
        End If
    End If
End Function

Private Sub RedrawShortestPath()
    If blInit Then
        Dim Predecessors() As Long
        ReDim Predecessors(1 To Dijkstra.GNodes)
        Dijkstra.GetPredecessorsArray Predecessors
        
        If blSPTrees Then
            SimpleDraw.DrawShortestPath Predecessors
        Else
            SimpleDraw.DrawShortestPath Predecessors, lTextOnMDI(SpDestination)
        End If
    End If
End Sub

Private Sub Picture1_GotFocus()
    GetFocus
End Sub

Private Sub Picture1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    Dim dTemp As Double
    Dim i     As Long
    Dim j     As Long

    If blCoord Then
        blIsRunning = True
        dTemp = dInf
        With TheGraph
            For i = 1 To .GNodes
                If Abs(.NodeXcoord(i) - x) + Abs(.NodeYcoord(i) - y) < dTemp Then
                    dTemp = Abs(.NodeXcoord(i) - x) + Abs(.NodeYcoord(i) - y)
                    j = i
                End If
            Next i

            If Button = 1 Then 'Set destination and calculate shortest paths
                frmMDI.txtMDI(SpDestination) = j
                CalculateShortestPaths Int(frmMDI.txtMDI(SpSource)), j
            ElseIf Button = 2 Then  'Set source and calculate shortest paths
                frmMDI.txtMDI(SpSource) = j
                CalculateShortestPaths j, Int(frmMDI.txtMDI(SpDestination))
            ElseIf Button = 4 Then   'Start zooming rectangle
                Zoom.Xmin = x
                Zoom.Ymax = y
            End If
        End With
        blIsRunning = False
    End If
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
    
    frmMDI.lblX.Caption = "X = " & Round(x, 3)
    frmMDI.lblY.Caption = "Y = " & Round(y, 3)
    MouseTimer.Enabled = False
    If blCoord Then
        If frmMDI.GetActiveChild = MyID Then
            If Button = 1 Or Button = 2 Then
                blMouseMoveCalc = True
                If Not blIsRunning Then
                    '// The concept here is that we want to reduce an excessive number of calls
                    '   to the Dijkstra algorithm / Drawing routine.
                    '   The mouse pointer must have been stablished over a given X,Y
                    '   for a sufficient time (MouseTimer.Interval) for the call to be made.
                    Picture1.CurrentX = x
                    Picture1.CurrentY = y
                    iLastButton = Button
                    MouseTimer.Enabled = True
                End If
            Else
                blMouseMoveCalc = False
                If Button = 4 Then ' draw zoom rectangle
                    If Not blZoomStart Then Picture1.Picture = Picture1.Image
                    Picture1.Cls
                    Picture1.DrawMode = 6
                    Picture1.Line (Zoom.Xmin, Zoom.Ymax)-(x, y), vbBlack, B
                    Picture1.DrawMode = 13
                    blZoomStart = True
                End If
            End If
        End If
    End If
    
End Sub

Private Sub MouseTimer_Timer()
    Picture1_MouseDown iLastButton, 0, Picture1.CurrentX, Picture1.CurrentY
    MouseTimer.Enabled = False
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 4 Then
        If blCoord Then
            With Zoom
                .Xmax = x
                .Ymin = y
                If .Xmax > .Xmin And .Ymax > .Ymin Then
                    SimpleDraw.ZoomMap .Xmin, .Ymin, .Xmax, .Ymax
                    Form_Resize
                Else
                    SimpleDraw.InitializeMap Picture1, Picture2, TheGraph
                    Form_Resize
                End If
            End With
        End If
        blZoomStart = False
    End If
     
End Sub

'++----------------------------------------------------------------------------
'//   CALCULATE SHORTEST PATHS
'++----------------------------------------------------------------------------
Public Function CalculateShortestPaths(Optional Source As Long, _
                                       Optional Destination As Long, _
                                       Optional lAPSP As Long) As Double
    If blInit Then
        CalculateShortestPaths = CalcShortestPaths(Dijkstra, SimpleDraw, TheGraph, Source, Destination, lAPSP, blCoord)
    End If
End Function

'++----------------------------------------------------------------------------
'//   IMPORT .ga GRAPH
'++----------------------------------------------------------------------------
Public Function ImportGraphData(strFileName As String) As Long
Dim i As Long
    
    If Dir(strFileName) = vbNullString Then Exit Function
    
    Set SimpleDraw = Nothing
    Set TheGraph = New clsGraph
    Set Dijkstra = New clsDijkstra

    blInit = False
    blCoord = False

    i = ImportGraph(strFileName, TheGraph)
    Dijkstra.InitList TheGraph, blPruneGraph

    '//Initialization (currently is a mess...)

    If i Then  'Update #1: fixed ( was if i = 0 Then )
        blCoord = True
        Form_Resize
        Set SimpleDraw = New clsSimpleDraw

        With SimpleDraw
            .InitializeMap Picture1, Picture2, TheGraph
            .DrawNet &HC0C0C0
            blCoord = .MapInit
        End With
    End If
    blInit = True
    lTextOnMDI(SpSource) = 1
    lTextOnMDI(SpDestination) = 1
    MyFile = strFileName
    lTileInstance = 0
    frmMDI.cbTileDiag.Enabled = False
    GetFocus

End Function

'++----------------------------------------------------------------------------
'//   MAKE A RANDOM GRAPH
'++----------------------------------------------------------------------------
Public Function RandomGraph(ByVal lNodes As Long, _
                            ByVal lEdgesPerNode As Long, _
                            Optional NoCoords As Boolean = False)
Dim Permutations() As Long
Dim i As Long, j As Long, lTmp As Long
    
    frmMDI.MousePointer = vbHourglass
    DoEvents

    blInit = False
    blCoord = False
    Set SimpleDraw = Nothing
    Set Dijkstra = New clsDijkstra
    Set TheGraph = New clsGraph

    ReDim Permutations(1 To lNodes)
    For i = 1 To lNodes
        Permutations(i) = i
    Next i
        
    Randomize (Timer)
    
    TheGraph.RedimStep = lNodes * lEdgesPerNode
    For i = 1 To lNodes
        For j = 1 To lEdgesPerNode
            lTmp = Permutations(j)
            Permutations(j) = Permutations(Int(Rnd() * (lNodes - j)) + j)
            Permutations(lTmp) = lTmp
            
            TheGraph.NewEdge i, Permutations(j), CDbl(Rnd() * 1000), CDbl(Rnd() * 1000)
        Next j
        If Not NoCoords Then
            TheGraph.SetNodeCoordinates (i), Rnd() * 1000, Rnd() * 1000
        End If
    Next i
    TheGraph.TrimBuffers
    
    '//Initialization (currently is a mess...)
    
    frmMDI.MousePointer = vbDefault
    Form_Resize
    lTextOnMDI(SpSource) = 1
    lTextOnMDI(SpDestination) = 1
    Dijkstra.InitList TheGraph, blPruneGraph
    
    blCoord = Not NoCoords
    If blCoord Then
        Set SimpleDraw = New clsSimpleDraw
        With SimpleDraw
            .InitializeMap Picture1, Picture2, TheGraph
            .DrawNet &HC0C0C0
            blCoord = .MapInit
        End With
    End If
    
    blInit = True
    lTileInstance = 0
    frmMDI.cbTileDiag.Enabled = False
    GetFocus
    
End Function

'++----------------------------------------------------------------------------
'//   IMPORT A .at GRAPH (tile instance)
'++----------------------------------------------------------------------------
'//   - At this time it can import only square tilemaps (you can override it by
'//     assigning "borders" of nodes with an "impassable" value.
'++----------------------------------------------------------------------------
Public Function ImportBinaryA(Optional strFileName As String) As String

Dim iFileBin       As Integer
Dim lDim           As Long
Dim aiMap()        As Long
Dim DeleteNodes()  As Integer
Dim i              As Long
Dim j              As Long
Dim lTmp           As Long
Dim lTmp2          As Long
Dim dCostTF        As Double
Dim dCostFT        As Double
Dim sTmpX          As Single
Dim sTmpY          As Single
Dim ImpassableTile As Long

Dim AtPtTColor As AtPoint
Dim AtPtTGscale As AtPoint

    
    Const Sqr2        As Double = 1.00000000001 '1.4142135623731 '<Priority given to diagonals

    If strFileName = vbNullString Then
        strFileName = MyFile
    End If
    If Dir(strFileName) = vbNullString Then Exit Function

    blInit = False
    blCoord = False
    Set Dijkstra = Nothing
    Set SimpleDraw = Nothing
    Set TheGraph = New clsGraph

    frmMDI.MousePointer = vbHourglass
    DoEvents

    With TheGraph

        iFileBin = FreeFile
        Open strFileName For Binary Access Read As #iFileBin

        lDim = (LOF(iFileBin) - 16) \ 4
        ReDim DeleteNodes(1 To lDim) As Integer
        
        .RedimStep = lDim * IIf(blAllowTileDiag, 8, 4) 'If we omit this in big instances,'
                                                       'we're going to wait for a long time..
        
        lDim = Sqr(lDim) - 1
        ReDim aiMap(lDim, lDim) As Long

        Get #iFileBin, , AtPtTColor 'lStartEnd()
        Get #iFileBin, , AtPtTGscale
        Get #iFileBin, , aiMap()
        Close #iFileBin
        
        '//Update 1.1
        If AtPtTColor.End = -1 Then
           ImpassableTile = AtPtTColor.Start                    'For .at converted from images, retrieve Impassable color
        ElseIf AtPtTGscale.End = -1 Then
           ImpassableTile = AtPtTGscale.Start                   'For .at converted from images, retrieve Impassable color
        Else
           ImpassableTile = 10                                  'For original .at files (CodeId=31654)
        End If
        '//End Update

        For i = 0 To lDim ' we move at Y axis

            lTmp = i * (lDim + 1) + 1

            For j = 0 To lDim 'we move at X axis

                lTmp2 = lTmp + j
                dCostFT = aiMap(j, i)
                If dCostFT = ImpassableTile Then
                    DeleteNodes(lTmp2) = 1 ' mark node as deleted
                End If
                
                '// Import Edges
                
                'Verticals   .
                '            |
                
                If i < lDim Then
                    dCostTF = aiMap(j, i + 1)
                    .NewEdge lTmp2, lTmp2 + lDim + 1, dCostFT, dCostTF
                End If
                
                If j < lDim Then
                
                'Horizontals .-
                
                    dCostTF = aiMap(j + 1, i)
                    .NewEdge lTmp2, lTmp2 + 1, dCostFT, dCostTF

                    If blAllowTileDiag Then
                        
                'Diagonals   .
                '             \
                
                        If i < lDim Then
                            dCostTF = aiMap(j + 1, i + 1)
                            .NewEdge lTmp2, lTmp2 + 2 + lDim, dCostFT * Sqr2, dCostTF * Sqr2
                        End If
                        
                'Diagonals    /
                '            .
                        
                        If i Then
                            dCostTF = aiMap(j + 1, i - 1)
                            .NewEdge lTmp2, lTmp2 - lDim, dCostFT * Sqr2, dCostTF * Sqr2
                        End If
                    End If

                End If

                .SetNodeCoordinates lTmp2, Val(j), Val(lDim - i)

            Next j
        Next i
        .FixGraph DeleteNodes ' remove nodes marked as deleted
        
        '//Update #1
        '/ Trim the unused Bottom and Right Pixel rows (I can't figure out why they exist...)
        If AtPtTColor.End = -1 Or AtPtTGscale.End = -1 Then
            sTmpY = sInf '<<
            For i = 1 To .GNodes
                If .NodeXcoord(i) > sTmpX Then sTmpX = .NodeXcoord(i)
                If .NodeYcoord(i) < sTmpY Then sTmpY = .NodeYcoord(i)
            Next i
            ReDim DeleteNodes(1 To .GNodes)
            For i = 1 To .GNodes
                If .NodeXcoord(i) = sTmpX Then DeleteNodes(i) = 1
                If .NodeYcoord(i) = sTmpY Then DeleteNodes(i) = 1
            Next i
            .FixGraph DeleteNodes
        End If
        '//End Update
        
        '//Initialization (currently is a mess...)
        
        Form_Resize
        lTextOnMDI(SpSource) = 1
        lTextOnMDI(SpDestination) = 1
        Set Dijkstra = New clsDijkstra
        Dijkstra.InitList TheGraph, blPruneGraph
        Set SimpleDraw = New clsSimpleDraw

        With SimpleDraw
            
            '//Update #1
            If AtPtTColor.End = -1 Or AtPtTColor.Start = -1 Then
                .ImageType = TrueColor
            ElseIf AtPtTGscale.End = -1 Or AtPtTGscale.End = -1 Then
                .ImageType = GreyScale
            Else
                .ImageType = NormalAt ' < draws like before update
            End If
            '//End Update
            
            .InitializeMap Picture1, Picture2, TheGraph
            .DrawNet &HC0C0C0
            blCoord = .MapInit
        End With
        
        blInit = True
        MyFile = strFileName
        lTileInstance = 1 + IIf(blAllowTileDiag, 1, 0)
        frmMDI.cbTileDiag.Enabled = True
        GetFocus
    End With

    frmMDI.MousePointer = vbDefault

End Function
