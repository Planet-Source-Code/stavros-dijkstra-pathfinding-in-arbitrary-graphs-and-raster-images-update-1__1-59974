VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDijkstra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' STANDARD DIJKSTRA's ALGORITHM USING A PRIORITY QUEUE IMPLEMENTED AS A BINARY HEAP
'
' It solves the Single Source Shortest Path Problem (SSSPP) in O(m*log(n)) time,
' and can be applied to any arbitrary directed graph consisting of n nodes and m edges.
' Each edge may have up to two positive (or 0) weighted connections.
' If an edge is one-directed, the impassable connection should be assigned a cost of -1.
'
' The algorithm makes use of a single dimensional node adjacency list, which must be
' Initialized prior running any shortest paths. After the initialization, no other
' external references of the original graph are required.
'
' Useful for vehicle routing, location analysis, general AI path finding etc.
'
' Although it can also be applied in raster-based graphs (i.e. images) it is less
' efficient and much more resource-hungry than other specialized techniques
' (see LaVolp's pathfinder for a remarkable example [CodeId=50278]).
'
' Any comments or suggestions on improving this class would be greatly appreciated!
' Also the whole project could (with a lot of work) evolve into a general platform
' for experimentation with graph algorithms for both Vector and Raster inputs.
' If you are interested, drop me an e-mail Stavros Sirigos <ssirig@prd.uth.gr>

Option Explicit

Private Const dInf     As Double = 1E+308        '"Infinity".

Private blListInit     As Boolean                'List has been initialized.

Private lGNodes        As Long                   'Number of nodes (vertices) in the graph.
Private lGEdges        As Long                   'Number of edges (arcs) in the graph (max.=lGNodes^2).
Private lGConnections  As Long                   'Number of Connections in the graph (max.=2*lGEdges).

Private lHeapSize      As Long                   'Number of elements in the priority queue (max.=lGNodes).

Private ListPtr()      As Long                   'Holds each node's UBound position in the Adjacency List.
Private ListAdjNode()  As Long                   'For each node, it holds the ID's of the linked nodes.
Private ListAdjCost()  As Double                 'For each node, it holds the connection cost of the linked nodes.
Private dCost()        As Double                 'Cost of travelling to each node. Array Index refers to node ID.
Private lPredecessor() As Long                   'Keeps track of the node ID via which we rach each destination node. Array Index refers to destination's ID.
Private HeapID()       As Long                   'Node ID of each element in the priority queue. Array index refers to HeapPos()
Private HeapPos()      As Long                   'Position of each node in the priority queue.

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'++----------------------------------------------------------------------------
'//   Public Properties and simple Methods
'++----------------------------------------------------------------------------

'//Number of Nodes in the adjacency list.
Public Property Get GNodes() As Long
    GNodes = lGNodes
End Property

'//Number of Edges in the adjacency list.
Public Property Get GEdges() As Long
    GEdges = lGEdges
End Property

'//Number of Connections in the adjacency list.
Public Property Get GConnections() As Long
    GConnections = lGConnections
End Property

'//Get the precalculated Predecessor's ID on the path from Source to lDestination.
Public Property Get Predecessor(ByVal lDestination As Long) As Long
    Predecessor = lPredecessor(lDestination)
End Property

'//Get the precalculated Cost value from Source to lDestination.
Public Property Get Cost(ByVal lDestination As Long) As Double
    Cost = dCost(lDestination)
End Property

'//Get the array of Predecessors' IDs on the path from Source to all other nodes
' (index refers to Destination ID).
Public Sub GetPredecessorsArray(ByRef TheArray() As Long)
    TheArray = lPredecessor
End Sub

'//Get the array of Costs from Source to all other nodes
' (index refers to Destination ID).
Public Sub GetCostArray(ByRef TheArray() As Double)
    TheArray = dCost
End Sub

'//Get the ratio of the connections traversed during last shortest paths calculation
Public Function GraphExplored() As Double
Dim i         As Long 'Node counter
Dim lExplored As Long 'Counter of explored connections
    
    If blListInit And lGConnections Then                        '/Adjacency List is Initialised
                                                                ' and at least one connection exists.
      For i = 1 To lGNodes                                      '/For every node in the graph..
          If dCost(i) <> dInf Then                              '/If a path from the source exists...
            lExplored = lExplored + ListPtr(i) - ListPtr(i - 1) '/add nr. of links whose origin is i.
          End If
      Next i
      GraphExplored = lExplored / lGConnections                 'Ratio of explored VS total connections.
    End If
End Function


'++/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'//
'//   ADJACENCY LIST STRUCTURE (sort of a linked-list implementation)
'//
'++\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'++----------------------------------------------------------------------------
'//   ADJACENCY LIST INITIALIZATION
'++----------------------------------------------------------------------------
'//   -The adjacency list is a way of representing the network topology:
'//    For each node, it allows sequential access to all the nodes that are
'//    directly linked to it, retrieving the IDs and respective connection costs.
'//
'//   -This routine must be run prior calculating any shortest paths.
'//    It is also the only routine in this class that requires an external
'//    object reference (TheGraph). Instead of TheGraph, you could also pass only
'//    the Cost and Topology arrays, in order to easily include the class
'//    in an other application / library.
'++----------------------------------------------------------------------------

Public Function InitList(ByRef TheGraph As clsGraph, Optional blPrune As Boolean) As String
Dim i             As Long
Dim lF            As Long    'Temp variable, refers to TopoFrom()
Dim lT            As Long    'Temp variable, refers to TopoTo()
Dim lNodeLBound() As Long    'This is a temp pointer for node insertion.
Dim CostFromTo()  As Double  'A copy of TheGraph's respective array.
Dim CostToFrom()  As Double  'A copy of TheGraph's respective array.
Dim TopoFrom()    As Long    'A copy of TheGraph's respective array.
Dim TopoTo()      As Long    'A copy of TheGraph's respective array.
Dim DeleteNodes() As Integer 'Used only if blPrune=True.
Dim blHasDangles  As Boolean '=True if TheGraph contains nodes whose connectivity is < 2 (could be any positive number).
    
    With TheGraph
        lGNodes = .GNodes                                       '/Get the nr. of graph nodes.
        lGEdges = .GEdges                                       '/Get the nr. of graph edges.
        lGConnections = .GConnections                           '/Get the nr. of graph connections.
        
        ReDim CostFromTo(1 To lGEdges)                          '/Initialize the temp arrays to hold
        ReDim CostToFrom(1 To lGEdges)                          ' TheGraph Data.
        ReDim TopoFrom(1 To lGEdges)                            '
        ReDim TopoTo(1 To lGEdges)                              '
                    
        .GetCost_FromTo CostFromTo                              '/Copy each temp array from TheGraph
        .GetCost_ToFrom CostToFrom                              ' to access the data more efficiently.
        .GetTopology_From TopoFrom                              '
        .GetTopology_To TopoTo                                  '
    End With
                                                                '/Initialize the 'pointer' array to the Adjacency List.
    ReDim ListPtr(lGNodes)

    For i = 1 To lGEdges                                        '/Looping through TheGraph's edges, we
        If Not CostFromTo(i) Then                               ' count every nodes connectivity.
            ListPtr(TopoFrom(i)) = ListPtr(TopoFrom(i)) + 1     ' Edges Cost=-1 means the connection
        End If                                                  ' is impassable.
        If Not CostToFrom(i) Then                               '
            ListPtr(TopoTo(i)) = ListPtr(TopoTo(i)) + 1         '
        End If                                                  '
    Next i                                                      '
    
    If blPrune Then                                             '/This optional part is useful if we
        ReDim DeleteNodes(1 To lGNodes)                         ' want to mark the nodes whose
        For i = 1 To lGNodes                                    ' connectivity is less than a specified bound.
            If ListPtr(i) < 2 Then                              ' (Here, nodes with connectivity<2,
                blHasDangles = True                             ' i.e. belonging in sub-trees).
                DeleteNodes(i) = 1                              '
            End If                                              '
        Next i                                                  '
        If blHasDangles Then                                    ' We remove nodes marked as Deleted
            TheGraph.FixGraph DeleteNodes                       ' in a recursive manner.
            
            Erase CostFromTo                                    '//Update #1
            Erase CostToFrom                                    '/   If we do not erase
            Erase TopoTo                                        '/   the arrays, the recursion will
            Erase TopoFrom                                      '/   consume more and more resources
            Erase ListPtr                                       '/   before it is finished.
            Erase DeleteNodes                                   '//End update
            
            InitList TheGraph, blPrune                          '
            Exit Function                                       '
        End If                                                  '
    End If                                                      '
    
    ReDim ListAdjNode(1 To lGConnections)                       '/Initialization of the Adjacency List.
    ReDim ListAdjCost(1 To lGConnections)                       '
    ReDim lNodeLBound(1 To lGNodes) As Long                     '/Initialization of the temp pointer array.
    
    For i = 1 To lGNodes                                        '/We make the array of pointers incremetal
        ListPtr(i) = ListPtr(i) + ListPtr(i - 1)                ' and initialise each node's Lower Bound
        lNodeLBound(i) = ListPtr(i - 1)                         ' for insertion in the List.
    Next i

    For i = 1 To lGEdges                                        '/This is the construction of the actuall List.
        lF = TopoFrom(i)                                        ' For each connection found, the respective
        lT = TopoTo(i)                                          ' nodes' LBound is incremented by 1, and the
        If Not CostFromTo(i) Then                               ' adjacent node's ID and connection cost are
            lNodeLBound(lF) = lNodeLBound(lF) + 1               ' entered into that position.
            ListAdjNode(lNodeLBound(lF)) = lT                   '
            ListAdjCost(lNodeLBound(lF)) = CostFromTo(i)        '
        End If                                                  '
        If Not CostToFrom(i) Then                               '
            lNodeLBound(lT) = lNodeLBound(lT) + 1               '
            ListAdjNode(lNodeLBound(lT)) = lF                   '
            ListAdjCost(lNodeLBound(lT)) = CostToFrom(i)        '
        End If                                                  '
    Next i                                                      '

    ReDim dCost(1 To lGNodes)                                   '/Initialization of Dijkstra and
    ReDim lPredecessor(1 To lGNodes)                            ' Heap related arrays.
    ReDim HeapID(1 To lGNodes)                                  '
    ReDim HeapPos(1 To lGNodes)                                 '

    blListInit = True                                           'We are ready to start computing shortest paths!
    
End Function

'++/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'//   DIJKSTRA ALGORITHM USING A PRIORITY QUEUE IMPLEMENTED AS A BINARY HEAP
'//
'//   -Step 1: Initialization ~ O(lGNodes).
'//   -Step 2: FindMin, DeleteRoot ~ O(log(lGNodes)).
'//   -Step 3: Relax tentative costs ~ O(lLocalConnections*log(lGNodes)).
'//   -Step 4: Repeat Step 2, Step 3 until Heap is empty or Destination Reached.
'//
'//   Amortised worst case complexity ~ O(lGEdges*log(lGNodes)).
'++\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'++----------------------------------------------------------------------------
'//   DIJKSTRA INITIALIZATION ~ complexity O(lGNodes)
'++----------------------------------------------------------------------------
'//   -What does it do?
'//   -First we set the array of costs equal to "Infinity", and clear the
'//    contents of the rest of the arrays used. Then we initialize the Heap by
'//    inserting the Source node in the Root. Finally we call the main algorithm.
'++----------------------------------------------------------------------------
Public Function DijkstraHeap(Source As Long, _
                             Optional Destination As Long) As Double
Dim lNodes As Long
    
    DijkstraHeap = -1                                           'In case something goes wrong.
    If blListInit Then
        If Source > 0 And Source <= lGNodes Then
            If Destination >= 0 And Destination <= lGNodes Then
            
                For lNodes = 1 To lGNodes                       '/Set cost array = Infinity.
                    dCost(lNodes) = dInf                        '
                Next lNodes                                     '
                    dCost(Source) = 0                           '/Set the cost to the Source=0
                                                                '(this could also be optional).
                                                                
                lNodes = lGNodes * 4                            '/Size of each array in bytes.
                ZeroMemory lPredecessor(1), lNodes              '/Clear the arrays.
                ZeroMemory HeapID(1), lNodes                    '
                ZeroMemory HeapPos(1), lNodes                   '
                
                lHeapSize = 1                                   '/Initialization of the Heap
                HeapID(lHeapSize) = Source                      ' by inserting the Source
                HeapPos(Source) = lHeapSize                     ' as the root node.
            
                DijkstraHeap = DijkstraMain(Destination)        '/Call the main algorithm.
                
            End If
        End If
    End If
End Function

'++----------------------------------------------------------------------------
'//   MAIN ALGORITHM ~ complexity O(lGEdges*log(lGNodes))
'++----------------------------------------------------------------------------
'//   -What does it do?
'//   -First, it finds the unscanned node closest to the Source. This is the
'//    root of the Heap, HeapID(1). If this is the target destination, we stop,
'//    else we delete the Root:
'//    This is done by swapping the Root (which is now lLastScanned) with the
'//    lLastElement of the heap. Then, we perform a Shift Down on the root since
'//    it probably violates the heap property (ie. it's key [cost in our case]
'//    is probably greater than one of it's children).
'//    Finally, we try improving the costs found so far by considering the
'//    connections that come out of lLastScanned node.
'++----------------------------------------------------------------------------
Private Function DijkstraMain(ByVal Destination As Long) As Double
Dim lLastScanned As Long 'The (until now unscanned) node's ID which is closest to the Source.
Dim lLastElement As Long 'ID of the last element in the priority queue.
    
    Do                                                          '//Main Loop.
    
        lLastScanned = HeapID(1)                                '/lLastScanned node's ID = Root ID.
        If lLastScanned = Destination Then                      '/Check if we have reached our
            DijkstraMain = dCost(Destination)                   ' Destination.
            Exit Do                                             ' If Destination = 0
        End If                                                  ' we simply calculate all paths.

        lLastElement = HeapID(lHeapSize)                        '/We delete the Heap's (now scanned)
        HeapID(1) = lLastElement                                ' Root element, by swapping it with the
        HeapPos(lLastElement) = 1                               ' Heap's lLastElement, decreasing the
        lHeapSize = lHeapSize - 1                               ' HeapSize by 1
                                                                ' and shifting the new Root down to
        HeapShiftDown lLastElement                              ' an appropriate position.
                                                                
        DijkstraRelaxCosts lLastScanned                         '/We try improving the costs found so far.

    Loop While lHeapSize                                        '//While the Heap is not empty.
    
    DijkstraMain = dCost(lLastScanned)
End Function

'++----------------------------------------------------------------------------
'//   RELAX COSTS ~ complexity O((LastCon-FirstCon)*log(lGNodes))
'++----------------------------------------------------------------------------
'//   -What does it do?
'//   -For each node adjacent to lLastScannedNode (lAdjNode), we try relaxing
'//    it's so-far Best Known ("tentative") Cost (dCost(lAdjNode)).
'//    If that was an improvement, we update dCost(lAdjNode), insert
'//    lAdjNode in the Heap (if it wasn't already there) and Percolate lAdjNode
'//    to it's final Heap Position.
'//
'//    We could also use an array to keep track of which nodes have reached
'//    their optimal costs (i.e. Scanned), to prevent summing and comparing
'//    double Vars. However, I found that the overhead caused by this, actually
'//    slows the procedure down marginally, even for dense graphs.
'++----------------------------------------------------------------------------
Private Sub DijkstraRelaxCosts(ByVal lLastScanned As Long)
Dim i            As Long
Dim FirstCon     As Long   'LBound in the Adjacency List.
Dim LastCon      As Long   'Ubound in the Adjacency List.
Dim lAdjNode     As Long   'ID of each of lLastScanned node's adjacent nodes.
Dim dCostTemp    As Double 'Optimal cost from Source node to lLastScanned node.
Dim dLastScanned As Double 'Cost from Source node to lAdjNode node via lLastScanned node.
    
    dLastScanned = dCost(lLastScanned)                          '/These assignments should
    FirstCon = ListPtr(lLastScanned - 1) + 1                    ' save some overhead later on.
    LastCon = ListPtr(lLastScanned)                             '

    For i = FirstCon To LastCon                                 '/Check each of LastScannedNode's direct connections.

        lAdjNode = ListAdjNode(i)                               '/These assignments should
        dCostTemp = ListAdjCost(i) + dLastScanned               ' save some overhead later on.

        If dCostTemp < dCost(lAdjNode) Then                     '/If we have found a shorter path...
            dCost(lAdjNode) = dCostTemp                         ' Update the current best cost.
            lPredecessor(lAdjNode) = lLastScanned               ' Update the last node we have to reach
                                                                ' prior reaching lAdjNode (for drawing purposes).

            If HeapPos(lAdjNode) Then                           '/lAdjNode is in the Heap:
                HeapPercolate lAdjNode                          ' Increase lAdjNode's Priority.
                
            Else                                                '/lAdjNode is not in the Heap:
                lHeapSize = lHeapSize + 1                       ' Increase the lHeapSize by 1.
                HeapID(lHeapSize) = lAdjNode                    ' Insert lAdjNode as the last element.
                HeapPos(lAdjNode) = lHeapSize                   '
                HeapPercolate lAdjNode                          ' Increase AdjNode's Priority.
            End If
        End If
    Next i
End Sub

'++----------------------------------------------------------------------------
'//   PERCOLATE ~ complexity O(log(lHeapSize))
'++----------------------------------------------------------------------------
'//   -What does it do?
'//   -It maintains the Heap property, i.e. ensuring that the target Child's
'//    key (cost in our case) is greater than it's respective Parent's key.
'//
'//    If the child's key is smaller, we swap Parent and Child and continue
'//    the process until no further swaps are possible.
'//
'//    In a binary Heap, ith Child's Parent is located at i\2
'//
'//    Also known as: Decrease Key, Increase Priority, Bubble Up or Shift Up.
'++----------------------------------------------------------------------------
Private Sub HeapPercolate(ByVal lChildID As Long)
Dim lParentID  As Long   'Parent's ID at each iteration.
Dim lParentPos As Long   'Parent's position in the Heap at each iteration.
Dim lChildPos  As Long   'Child's position in the Heap at each iteration.
Dim dChildCost As Double 'Child's updated key (cost from Source).

    dChildCost = dCost(lChildID)
    lChildPos = HeapPos(lChildID)                                '/Current Child's position.
    
    Do Until lChildPos = 1                                       '//Until we reach the Root...
        lParentPos = lChildPos \ 2                               '/Parent's Current position.
        lParentID = HeapID(lParentPos)                           '/Parent's ID.

        If dCost(lParentID) > dChildCost Then                    '/Child is violating the Heap property.
            HeapID(lChildPos) = lParentID                        ' Update the HeapID at Child's position to hold Parent's ID.
            HeapPos(lParentID) = lChildPos                       ' Update Parent's position.
            lChildPos = lParentPos                               ' This is the Child's new position.
        Else                                                     '
            Exit Do                                              '/Child is NOT violating the Heap property.
        End If                                                   '
    Loop                                                         '//Try shifting the Child further UP the Heap.
    
    HeapID(lChildPos) = lChildID                                 '/Establish the Child in it's final position.
    HeapPos(lChildID) = lChildPos                                '
End Sub

'++----------------------------------------------------------------------------
'//   SHIFT DOWN ~ complexity O(log(lHeapSize))
'++----------------------------------------------------------------------------
'//   -What does it do?
'//   -It maintains the Heap property, i.e. ensuring that the target parent's
'//    key (cost in our case) is smaller than it's respective children keys.
'//
'//    If the parent's key is greater, we swap Parent with the Child that has the
'//    smallest key and continue the process until no further swaps are possible.
'//
'//    In a binary Heap, ith Parent's Left and Right Children are located
'//    at i*2 and i*2+1 respectively (here we do Left=i+i and Right=Left+1).
'//
'//    Also known as: Increase Key, Decrease Priority or Bubble Down.
'++----------------------------------------------------------------------------
Private Sub HeapShiftDown(ByVal lParentID As Long)
Dim lChildID    As Long    'Selected Child's ID at each iteration.
Dim lParentPos  As Long    'Parent's position in the Heap at each iteration.
Dim lLeftPos    As Long    'Left Child's position in the Heap at each iteration.
Dim lRightPos   As Long    'Right Child's position in the Heap at each iteration.
Dim dLeftCost   As Double  'Left Child's key (cost from Source).
Dim dParentCost As Double  'Parent's key (cost from Source).

    dParentCost = dCost(lParentID)
    lParentPos = HeapPos(lParentID)                             '/Current Parent's position.
    lLeftPos = lParentPos + lParentPos                          '/Current Left Child's position.

    Do Until lLeftPos > lHeapSize                               '//Until Left position is out of bounds...
        lRightPos = lLeftPos + 1                                '/Current Right Child's position.
        dLeftCost = dCost(HeapID(lLeftPos))                     '/Current Left Child's key.

        If lRightPos <= lHeapSize Then                          '//If Right position is inside bounds...
        
            If dLeftCost < dParentCost Then                     '/Parent is violating the Heap property.
                If dLeftCost > dCost(HeapID(lRightPos)) Then    ' We consider as "Left" Child the Child
                    lLeftPos = lRightPos                        ' with the minimum key.
                End If                                          '
            Else                                                '
                If dParentCost > dCost(HeapID(lRightPos)) Then  '/Parent is violating the Heap property.
                    lLeftPos = lRightPos                        ' We consider as "Left" Child the Child
                                                                ' with the minimum key.
                                                                
                Else                                            '/Parent is NOT violating the Heap property.
                    Exit Do                                     '
                End If
            End If
        Else                                                    '//If Right position is outside bounds...
            If dParentCost <= dLeftCost Then                    '/Parent is NOT violating the Heap property.
                Exit Do                                         '
            End If
        End If
                                                                '/Perform the swap.
        lChildID = HeapID(lLeftPos)                             'Child's ID (could actually be Left's or Right's)
        HeapID(lParentPos) = lChildID                           'Update the HeapID at Parent's position to hold Child's ID.
        HeapPos(lChildID) = lParentPos                          'Update Child's position.

        lParentPos = lLeftPos                                   'Current Parent's position.
        lLeftPos = lLeftPos + lLeftPos                          'Current Left Child's position.
    Loop                                                        '//Try shifting the Parent further DOWN the Heap.
    
    HeapID(lParentPos) = lParentID                              '/Establish the Parent in it's final position.
    HeapPos(lParentID) = lParentPos                             '
End Sub
