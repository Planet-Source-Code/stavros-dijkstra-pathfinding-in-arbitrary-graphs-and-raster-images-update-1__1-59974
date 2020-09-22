VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimpleDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This is a simple support class that draws the graph and shortest paths
'
' It can be enhanced in many ways e.g.:
'   DMA for screen drawing
'   .EMF export
'   The way nodes are drawn (maybe circles?).
'   Some hard coded brushes/colors/properties should be removed from here and be defined in the UI.
'   Also the zooming / panning should be improved.
'   // Update #1 Almost done ->> Support for actually showing raster images (with zooming etc.) instead of showing a "plain" graph.

Option Explicit

Private Const sInf   As Single = 1E+38       '/"Infinity" (single)

'// PointAPI udt
Private Type POINTAPI
    x   As Long
    y   As Long
End Type

'//Update #1
Private lImageType    As Long
'//End Update

'// ViewPort coordinates
Private sXmin        As Single
Private sXmax        As Single
Private sYmin        As Single
Private sYmax        As Single

'// Coordinate --> Pixel multipliers
Private sPxlX        As Single
Private sPxlY        As Single

Private blInit       As Boolean             '/If=True The map is initialized

Private TheGraph     As clsGraph            '/Reference to the actual graph
Private TheMap       As PictureBox          '/The "Front" layer, where we draw the shortest paths.
Private TheBuffer    As PictureBox          '/The "Back" layer, where we draw the graph and update it
                                            ' only if the graph has changed or the picture is resized.

Private QuickX()     As Single              '/Local copy of the graph node's X coordinates
Private QuickY()     As Single              '/Local copy of the graph node's Y coordinates
Private DrawMe()     As Integer             '/=1 for nodes close to the ViewPort

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'++----------------------------------------------------------------------------
'//   Class
'++----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set TheMap = Nothing
    Set TheBuffer = Nothing
End Sub

'++----------------------------------------------------------------------------
'//   Public Properties
'++----------------------------------------------------------------------------
Public Property Get MapInit() As Boolean
    MapInit = blInit
End Property

'//Update #1
Public Property Get ImageType() As Long
    ImageType = lImageType
End Property
Public Property Let ImageType(lType As Long)
    lImageType = lType
End Property
'//End Update

'++----------------------------------------------------------------------------
'//   REFRESH GRAPH - Copies the graph image (nodes and edges) to the visible picturebox
'++----------------------------------------------------------------------------
Public Sub RefreshGraph()
    If blInit Then
        TheMap.Picture = TheBuffer.Image
    End If
End Sub


'++----------------------------------------------------------------------------
'//   INITALIZE MAP
'++----------------------------------------------------------------------------
'//   - It is called after the graph has been imported or altered.
'++----------------------------------------------------------------------------
Public Sub InitializeMap(ByRef SourceMap As PictureBox, _
                         ByRef BufferMap As PictureBox, _
                         ByRef SourceGraph As clsGraph)
Dim i     As Long
Dim sXdif As Single
Dim sYdif As Single

    Set TheMap = SourceMap
    Set TheBuffer = BufferMap
    Set TheGraph = SourceGraph
    
    TheMap.ScaleMode = 0
    TheBuffer.ScaleMode = 0
    
    sXmax = 0
    sYmax = 0
    sXmin = sInf
    sYmin = sInf

    With TheGraph
        ReDim QuickX(1 To .GNodes)
        ReDim QuickY(1 To .GNodes)
        ReDim DrawMe(1 To .GNodes)
        .GetCoords_X QuickX
        .GetCoords_Y QuickY
    End With
    
    For i = 1 To UBound(QuickX)
        DrawMe(i) = 1
        If QuickX(i) < sXmin Then sXmin = QuickX(i)
        If QuickX(i) > sXmax Then sXmax = QuickX(i)
        If QuickY(i) < sYmin Then sYmin = QuickY(i)
        If QuickY(i) > sYmax Then sYmax = QuickY(i)
    Next i

    sXdif = sXmax - sXmin
    sYdif = sYmax - sYmin
    sXmin = sXmin - 0.01 * sXdif
    sXmax = sXmax + 0.01 * sXdif
    sYmin = sYmin - 0.01 * sYdif
    sYmax = sYmax + 0.01 * sYdif

    For i = 1 To UBound(QuickX)
        QuickX(i) = QuickX(i) - sXmin
        QuickY(i) = sYmax - QuickY(i)
    Next i

    ResizeMap
    blInit = True

End Sub

'++----------------------------------------------------------------------------
'//   RESIZE MAP
'++----------------------------------------------------------------------------
'//   - It calculates the correct multipliers for drawing in pixel mode using
'//     the real node XY coordinates.
'++----------------------------------------------------------------------------
Public Sub ResizeMap()
Dim sTemp As Single

    sTemp = (sXmax - sXmin) / (sYmax - sYmin)

    If TheMap.Width / TheMap.Height > sTemp Then
        TheMap.Width = TheMap.Height * sTemp
        TheBuffer.Width = TheMap.Width
    Else
        TheMap.Height = TheMap.Width / sTemp
        TheBuffer.Height = TheMap.Height
    End If
    TheMap.Scale (sXmin, sYmax)-(sXmax, sYmin)

    sPxlX = TheBuffer.ScaleWidth / (sXmax - sXmin)
    sPxlY = TheBuffer.ScaleHeight / (sYmax - sYmin)

End Sub

'++----------------------------------------------------------------------------
'//   ZOOM MAP
'++----------------------------------------------------------------------------
'//   - Zooms to the nodes that lay inside the defined zoom rectangle
'++----------------------------------------------------------------------------
Public Sub ZoomMap(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
Dim i As Long

    For i = 1 To UBound(QuickX)
        If DrawMe(i) Then
            QuickX(i) = QuickX(i) + sXmin
            QuickY(i) = sYmax - QuickY(i)
            
            If QuickX(i) < 0.5 * X1 Or QuickX(i) > 2 * X2 Or _
               QuickY(i) < 0.5 * Y1 Or QuickY(i) > 2 * Y2 Then
               DrawMe(i) = 0
            End If
            
            QuickX(i) = QuickX(i) - X1
            QuickY(i) = Y2 - QuickY(i)
        End If
    Next i
    
    sXmin = X1
    sXmax = X2
    sYmin = Y1
    sYmax = Y2

    ResizeMap
End Sub

'++----------------------------------------------------------------------------
'//   DRAWING OF THE GRAPH EDGES AND NODES
'++----------------------------------------------------------------------------
Public Sub DrawNet(Optional ByVal TheColor As Long = vbBlack, _
                   Optional ByVal l_hDC As Long, _
                   Optional ByVal Xmulti As Single, _
                   Optional ByVal Ymulti As Single)

Dim i           As Long
Dim lPen        As Long
Dim lGNodes     As Long
Dim lGEdges     As Long
Dim lLastColor  As Long 'Update #1 (used to minimize new pen creation)

    Select Case lImageType
    Case NormalAt
        TheBuffer.BackColor = vbWhite
    Case Else
        TheBuffer.BackColor = modPublic.ImpassableColor
    End Select
    
    If Not CheckDefaults(TheBuffer.hDC, l_hDC, Xmulti, Ymulti) Then

        With TheGraph
            lGEdges = .GEdges
            lGNodes = .GNodes
            
            '//Update #1
            Select Case lImageType
            
            '/As in Original Version
            Case NormalAt
                lPen = CreatePen(0, 1, TheColor)
                DeleteObject SelectObject(l_hDC, lPen)
                For i = 1 To lGEdges
                    DrawLine l_hDC, .EdgeFrom(i), .EdgeTo(i), Xmulti, Ymulti
                Next i
                For i = 1 To lGNodes
                    DrawPixel l_hDC, i, Xmulti, Ymulti
                Next i
            
            '/In TrueColor mode (costs must have been pre-normalized)
            Case TrueColor
                For i = 1 To lGEdges
                    If .CostFromTo(i) <> lLastColor Then
                        lLastColor = .CostFromTo(i)
                        lPen = CreatePen(0, 1, lLastColor)
                        DeleteObject SelectObject(l_hDC, lPen)
                    End If
                    DrawLine l_hDC, .EdgeFrom(i), .EdgeTo(i), Xmulti, Ymulti
                Next i
            
            '/In GreyScale mode (costs must have been pre-normalized)
            Case GreyScale
                For i = 1 To lGEdges
                    If .CostFromTo(i) <> lLastColor Then
                        lLastColor = 256 - .CostFromTo(i)
                        lPen = CreatePen(0, 1, RGB(lLastColor, lLastColor, lLastColor))
                        DeleteObject SelectObject(l_hDC, lPen)
                    End If
                    DrawLine l_hDC, .EdgeFrom(i), .EdgeTo(i), Xmulti, Ymulti
                Next i
            End Select
            '//End Update
                
        End With
        
        RefreshGraph
        DeleteObject lPen
    End If

End Sub

'++----------------------------------------------------------------------------
'//   DRAWING OF THE SHORTEST PATHS / SHORTEST PATH TREES
'++----------------------------------------------------------------------------
Public Sub DrawShortestPath(ByRef Predecessors() As Long, _
                            Optional ByVal Destination As Long, _
                            Optional ByVal TheColor As Long = vbRed, _
                            Optional ByVal l_hDC As Long, _
                            Optional ByVal Xmulti As Single, _
                            Optional ByVal Ymulti As Single)
Dim lPen           As Long
Dim i              As Long
Dim lGNodes        As Long
Dim HasBeenDrawn() As Long

    If Not CheckDefaults(TheMap.hDC, l_hDC, Xmulti, Ymulti) Then

        lPen = CreatePen(0, IIf(blSPTrees, 2, 3), TheColor)
        DeleteObject SelectObject(l_hDC, lPen)

        If Not blSPTrees Then                                   '//Draw the Shortest Path:
            If Destination Then
                While Predecessors(Destination)
                    DrawLine l_hDC, Destination, Predecessors(Destination), Xmulti, Ymulti
                    Destination = Predecessors(Destination)
                Wend
            End If
        Else                                                    '//Draw the Shortest Path Tree:
            lGNodes = UBound(Predecessors)
            ReDim HasBeenDrawn(1 To lGNodes) As Long
            For i = 1 To lGNodes
                If HasBeenDrawn(i) = 0 Then                     '/Draw each edge only once
                    Destination = i
                    Do While Predecessors(Destination)
                        If HasBeenDrawn(Destination) Then
                            Exit Do
                        End If
                        HasBeenDrawn(Destination) = 1
                        DrawLine l_hDC, Destination, Predecessors(Destination), Xmulti, Ymulti
                        Destination = Predecessors(Destination)
                    Loop
                End If
            Next i
        End If
        
        DeleteObject lPen
    End If
End Sub

'++----------------------------------------------------------------------------
'//   assigning default (non-constant) values for Optional variables
'++----------------------------------------------------------------------------
Private Function CheckDefaults(ByVal Def_hDC As Long, _
                               ByRef l_hDC As Long, _
                               ByRef Xmulti As Single, _
                               ByRef Ymulti As Single) As Boolean

    If l_hDC = 0 Then l_hDC = Def_hDC
    If Xmulti = 0 Then Xmulti = sPxlX
    If Ymulti = 0 Then Ymulti = sPxlY
    CheckDefaults = Not blInit

End Function

'++----------------------------------------------------------------------------
'//   DRAW LINE
'++----------------------------------------------------------------------------
'//   - Draws line to the target DC (only if endpoint coordinates are different)
'++----------------------------------------------------------------------------
Private Sub DrawLine(ByVal l_hDC As Long, _
                     ByVal FromNode As Long, _
                     ByVal ToNode As Long, _
                     Optional ByVal Xmulti As Single = 1, _
                     Optional ByVal Ymulti As Single = 1)
Dim ptX As Long
Dim ptY As Long
Dim pt  As POINTAPI
    
    If DrawMe(FromNode) Then
        If DrawMe(ToNode) Then
            pt.x = QuickX(FromNode) * Xmulti
            pt.y = QuickY(FromNode) * Ymulti
            ptX = QuickX(ToNode) * Xmulti
            ptY = QuickY(ToNode) * Ymulti
            
            If pt.x <> ptX Or pt.y <> ptY Then
                MoveToEx l_hDC, pt.x, pt.y, pt
                LineTo l_hDC, ptX, ptY
            End If
        End If
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   DRAW PIXEL
'++----------------------------------------------------------------------------
'//   - Sets a pixel to the target DC (if it is deferent since the last call)
'++----------------------------------------------------------------------------
Private Sub DrawPixel(ByVal l_hDC As Long, _
                      ByVal Node As Long, _
                      Optional ByVal Xmulti As Single = 1, _
                      Optional ByVal Ymulti As Single = 1)
Dim pt1X As Long
Dim pt1Y As Long
Static pt2X As Long
Static pt2Y As Long

    If DrawMe(Node) Then
        pt1X = QuickX(Node) * Xmulti
        pt1Y = QuickY(Node) * Ymulti
        If pt1X <> pt2X Or pt1Y <> pt2Y Then
            pt2X = pt1X
            pt2Y = pt1Y
            SetPixel l_hDC, pt1X, pt1Y, vbBlue
        End If
    End If
            
End Sub
