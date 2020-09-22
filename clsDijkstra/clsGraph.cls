VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Graph Class
'
' Nothing special here, just a very simple and not very efficient graph class
' to get things started. If the project is going to evolve, this part should
' see a lot of changes.
'
Option Explicit

Private Enum BufferConst
    UPDT_BUF = 0 ' Update the Edges potential UBound
    INIT_BUF = 1 ' Initialize the Edges
    TRIM_BUF = 2 ' Trim the Edges arrays to the actual Edges UBound
End Enum

Private Const lInitStep As Long = 256   'Default buffer step
Private lRedimStep      As Long         'This is the number of Edges "slots" we buffer each time the buffer is full.

Private lGNodes         As Long         'Number of nodes in the graph (n)
Private lGEdges         As Long         'Number of edges in the graph (m) (0<= m <=n^2)
Private lGConnections   As Long         'Number of connections in the graph (c) (0<= c <=2*m)
Private blInitXY        As Boolean      '=True, If the graph nodes are associated with XY coordinates

Private Coords_X()      As Single       'Node X coordinates. Index refers to node ID
Private Coords_Y()      As Single       'Node Y coordinates. Index refers to node ID
Private Topology_From() As Long         'Edge Start, or "From", node. Index refers to edge ID
Private Topology_To()   As Long         'Edge End, or "To", node. Index refers to edge ID
Private Cost_FromTo()   As Double       'Cost of connection From -> To. Index refers to edge ID
Private Cost_ToFrom()   As Double       'Cost of connection To -> From. Index refers to edge ID

'++----------------------------------------------------------------------------
'//   Class
'++----------------------------------------------------------------------------

Public Sub Class_Initialize()
    RedimStep = lInitStep
    lGNodes = 0
    lGEdges = 0
    blInitXY = False
    BufferEdges INIT_BUF
End Sub

'++----------------------------------------------------------------------------
'//   Public Properties and simple Methods
'++----------------------------------------------------------------------------

'----------------------------------
'// Get graph properties one by one
'----------------------------------

Public Property Get CostFromTo(EdgeID As Long) As Double
    If EdgeID <= lGEdges Then
        CostFromTo = Cost_FromTo(EdgeID)
    End If
End Property

Public Property Get CostToFrom(EdgeID As Long) As Double
    If EdgeID <= lGEdges Then
        CostToFrom = Cost_ToFrom(EdgeID)
    End If
End Property

Public Property Get EdgeFrom(EdgeID As Long) As Long
    If EdgeID <= lGEdges Then
        EdgeFrom = Topology_From(EdgeID)
    End If
End Property

Public Property Get EdgeTo(EdgeID As Long) As Long
    If EdgeID <= lGEdges Then
        EdgeTo = Topology_To(EdgeID)
    End If
End Property

Public Property Get GConnections() As Long
    GConnections = lGConnections
End Property

Public Property Get GEdges() As Long
    GEdges = lGEdges
End Property

Public Property Get GNodes() As Long
    GNodes = lGNodes
End Property

Public Property Get NodeXcoord(NodeID As Long) As Single
    If NodeID <= lGNodes Then
        NodeXcoord = Coords_X(NodeID)
    End If
End Property

Public Property Get NodeYcoord(NodeID As Long) As Single
    If NodeID <= lGNodes Then
        NodeYcoord = Coords_Y(NodeID)
    End If
End Property

Public Property Get RedimStep() As Long
    RedimStep = lRedimStep
End Property

Public Property Let RedimStep(ByVal SetStep As Long)
    lRedimStep = IIf(SetStep, Abs(SetStep), lInitStep)
    If lGEdges = 0 Then
        BufferEdges INIT_BUF
    End If
End Property

'------------------------------------
'// Get complete arrays of properties
'------------------------------------

Public Sub GetCoords_X(ByRef TheArray() As Single)
    TheArray = Coords_X
End Sub

Public Sub GetCoords_Y(ByRef TheArray() As Single)
    TheArray = Coords_Y
End Sub

Public Sub GetCost_FromTo(ByRef TheArray() As Double)
    TheArray = Cost_FromTo
End Sub

Public Sub GetCost_ToFrom(ByRef TheArray() As Double)
    TheArray = Cost_ToFrom
End Sub

Public Sub GetTopology_From(ByRef TheArray() As Long)
    TheArray = Topology_From
End Sub

Public Sub GetTopology_To(ByRef TheArray() As Long)
    TheArray = Topology_To
End Sub



'++----------------------------------------------------------------------------
'//   BUFFER EDGES
'++----------------------------------------------------------------------------
'//   - It is used to handle how frequently the edge arrays are redimensioned.
'++----------------------------------------------------------------------------
Public Sub TrimBuffers()
    BufferEdges TRIM_BUF                                        '/Partial interface to BufferEdges
End Sub

Private Sub BufferEdges(Optional ByVal lOp As BufferConst = UPDT_BUF)
Dim i As Long
    
    If lOp = INIT_BUF Then                                      '/Initialize
        ReDim Topology_From(1 To lRedimStep)
        ReDim Topology_To(1 To lRedimStep)
        ReDim Cost_FromTo(1 To lRedimStep)
        ReDim Cost_ToFrom(1 To lRedimStep)
    Else
        If lOp = UPDT_BUF Then
            i = UBound(Topology_From) + lRedimStep              '/Update
        ElseIf lOp = TRIM_BUF Then
            i = lGEdges                                         '/Trim
        End If
        ReDim Preserve Topology_From(1 To i)
        ReDim Preserve Topology_To(1 To i)
        ReDim Preserve Cost_FromTo(1 To i)
        ReDim Preserve Cost_ToFrom(1 To i)
    End If

End Sub

'++----------------------------------------------------------------------------
'//   ADD A NEW EDGE
'++----------------------------------------------------------------------------
'//   - Adds a new edge to the graph, redimensioning the arrays if needed.
'//   - This routine also maintains an upper bound in the number of nodes.
'//   - Connection Cost = - 1 means impassable.
'++----------------------------------------------------------------------------
Public Sub NewEdge(ByVal FromNode As Long, _
                   ByVal ToNode As Long, _
                   ByVal CostFromTo As Double, _
                   ByVal CostToFrom As Double)
    
    lGEdges = lGEdges + 1
    If Not CostFromTo Then                                      '/Not impassable.
        lGConnections = lGConnections + 1                       ' increment connections.
    End If
    If Not CostToFrom Then                                      '/Not impassable.
        lGConnections = lGConnections + 1                       ' increment connections.
    End If

    If lGEdges Mod lRedimStep = 0 Then                          '/We must redimension
        BufferEdges
    End If
    
    Topology_From(lGEdges) = FromNode                           '/Set the edge properties
    Topology_To(lGEdges) = ToNode                               '
    Cost_FromTo(lGEdges) = CostFromTo                           '
    Cost_ToFrom(lGEdges) = CostToFrom                           '

    If lGNodes < FromNode Then                                  '/Update the number of nodes
        lGNodes = FromNode                                      '
    End If                                                      '
    If lGNodes < ToNode Then                                    '
        lGNodes = ToNode                                        '
    End If                                                      '

End Sub

'++----------------------------------------------------------------------------
'//   UPDATE AN EDGE
'++----------------------------------------------------------------------------
'//   - Updates an edge's properties, adjusting the connection count appropriately
'++----------------------------------------------------------------------------
Public Sub UpdateEdge(ByVal EdgeID As Long, _
                      Optional ByVal FromNode As Long, _
                      Optional ByVal ToNode As Long, _
                      Optional ByVal CostFromTo As Double = -2, _
                      Optional ByVal CostToFrom As Double = -2)
    
    If FromNode Then
        Topology_From(EdgeID) = FromNode                        '/Update "From" Node ID
    End If
    
    If ToNode Then
        Topology_To(EdgeID) = ToNode                            '/Update "To" Node ID
    End If
    
    If CostFromTo <> -2 Then                                    '/Update FromTo costs & connection count
        If (Not CostFromTo) And Cost_FromTo(EdgeID) = -1 Then
            lGConnections = lGConnections + 1                   'Previous=impassable, new=passable
        End If
        If CostFromTo = -1 And (Not Cost_FromTo(EdgeID)) Then
            lGConnections = lGConnections - 1                   'Previous=passable, new=impassable
        End If
        Cost_FromTo(EdgeID) = CostFromTo                        'New FromTo connection cost
    End If
        
    If CostToFrom <> -2 Then                                    '/Update ToFrom costs & connection count
        If (Not CostToFrom) And Cost_ToFrom(EdgeID) = -1 Then
            lGConnections = lGConnections + 1                   'Previous=impassable, new=passable
        End If
        If CostToFrom = -1 And (Not Cost_ToFrom(EdgeID)) Then
            lGConnections = lGConnections - 1                   'Previous=passable, new=impassable
        End If
        Cost_ToFrom(EdgeID) = CostToFrom                        'New ToFrom connection cost
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   DELETE AN EDGE
'++----------------------------------------------------------------------------
'//   - Swaps deleted edge with the last edge and adjust connection and edge count appropriately
'++----------------------------------------------------------------------------
Public Sub DeleteEdge(ByVal EdgeID As Long)

    If Not Cost_FromTo(EdgeID) Then                             '/The connection existed (was passable)
        lGConnections = lGConnections - 1
    End If
    If Not Cost_ToFrom(EdgeID) Then                             '/The connection existed (was passable)
        lGConnections = lGConnections - 1
    End If
    If EdgeID < lGEdges Then                                    '/Swap with last edge
        Topology_From(EdgeID) = Topology_From(lGEdges)          ' (Note: we are not storing edges
        Topology_To(EdgeID) = Topology_To(lGEdges)              '  in increasing ID order - in
        Cost_FromTo(EdgeID) = Cost_FromTo(lGEdges)              '  contrast with the nodes.)
        Cost_ToFrom(EdgeID) = Cost_ToFrom(lGEdges)              '
    End If                                                      '
    lGEdges = lGEdges - 1

End Sub

'++----------------------------------------------------------------------------
'//   ADD NODES / UPDATE NODE'S XY COORDINATES
'++----------------------------------------------------------------------------
'//   - It also ReDim Preserve the array of nodes, if we exceed the UBound.
'//     '(This shouldn't happen if we load the edges first. It is useful however,
'//       if we want to add new nodes later, e.g. interactively on screen).
'++----------------------------------------------------------------------------
Public Sub SetNodeCoordinates(ByVal NodeID As Long, _
                              ByVal Xcoord As Single, _
                              ByVal Ycoord As Single)
    If lGNodes < NodeID Then
        lGNodes = NodeID
    End If
    If Not blInitXY Then
        ReDim Coords_X(1 To lGNodes)                            '/Initialize the arrays
        ReDim Coords_Y(1 To lGNodes)
        blInitXY = True
    Else
        If UBound(Coords_X) < lGNodes Then
            ReDim Preserve Coords_X(1 To lGNodes)               '/New nodes have been added
            ReDim Preserve Coords_Y(1 To lGNodes)
        End If
    End If
    Coords_X(NodeID) = Xcoord
    Coords_Y(NodeID) = Ycoord
End Sub

'++----------------------------------------------------------------------------
'//   FIX THE GRAPH - (READ: BULK DELETE NODES AND ASSOCIATED EDGES AND FIX ID MESS :) )
'++----------------------------------------------------------------------------
'//   - NOTE: This routine is rather inefficient (it should use an adjacency list for speed up).
'//     In addition, it potentially messes up with the ID's initially imported.
'//
'//   - It finds and deletes all edges that have at least one of their endpoint nodes marked
'//     as deleted (<> 0). Nodes marked as deleted are deleted as well (duh!..).
'//     It then assigns new IDs to the remaining nodes in order for them to be in the new
'//     (1,...,lGNodes) interval and updates all edges with the new node IDs.
'//
'//     This means that the new node with ID=12345 is probably different than the
'//     node with the same ID before this routine was applied.
'++----------------------------------------------------------------------------
Public Sub FixGraph(ByRef DeleteNodes() As Integer)
Dim i             As Long
Dim j             As Long
Dim MapNewNodes() As Long

    ReDim MapNewNodes(1 To lGNodes) As Long
    ReDim Preserve DeleteNodes(1 To lGNodes)
    
    i = 1
    j = 0

    Do
        If DeleteNodes(Topology_From(i)) Or DeleteNodes(Topology_To(i)) Then
            DeleteEdge (i)
        Else
            i = i + 1
        End If
    Loop Until i > lGEdges
    BufferEdges TRIM_BUF

    j = 0
    For i = 1 To lGNodes
        If DeleteNodes(i) = 0 Then
            j = j + 1
            MapNewNodes(i) = j
            Coords_X(j) = Coords_X(i)
            Coords_Y(j) = Coords_Y(i)
        End If
    Next i

    lGNodes = j
    ReDim Preserve Coords_X(1 To lGNodes)
    ReDim Preserve Coords_Y(1 To lGNodes)

    For i = 1 To lGEdges
        UpdateEdge i, MapNewNodes(Topology_From(i)), MapNewNodes(Topology_To(i))
    Next i

End Sub
