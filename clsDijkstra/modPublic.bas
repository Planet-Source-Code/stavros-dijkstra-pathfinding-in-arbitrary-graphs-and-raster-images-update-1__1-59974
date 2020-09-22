Attribute VB_Name = "modPublic"
Option Explicit

Public Enum MDItext
    SpSource = 0
    SpDestination = 1
End Enum

'// Update #1
Public Enum EnumImageType
    NormalAt = 0
    TrueColor = 1
    GreyScale = 2
End Enum

Public Type AtPoint ' < This structure was use by the original author of *.at files
    Start As Long
    End As Long
End Type

Public Const ImpassableColor As Long = vbBlack      '/Determines the color of "obstacles" when importing images.
                                                    ' Nodes that are imported as obstacles are removed from the graph.
'// End Update

Public Const dInf          As Double = 1E+308       '"Infinity"
Public Const sInf          As Single = 1E+38        '/"Infinity" (single)
Public Const lProgress_ms  As Long = 250            'How often should we update the progressbar (ms)

Public GlobalFormHeight    As Long
Public GlobalFormWidth     As Long
Public blMouseMoveCalc     As Boolean               '=True if we are calculating shortest paths via the mouse
Public blAllowTileDiag     As Boolean               'Only for tile instances - toggles diagonals.
Public blSPTrees           As Boolean               'Show Shortest path trees instead of a single path
Public blPruneGraph        As Boolean               'see clsDijkstra InitList
Public blBenchMark         As Boolean               'Udate #1: =True if we are running the benchmark

Public blStopProcess       As Boolean               'Stop calculating (used by frmProgress)
Public blUpdate            As Boolean               'update the progressbar or not
Public lElapsedPrevious    As Long                  'elapsed ms until last progressbar update

Private Const NON_RELATIVE As Boolean = False       'Refers to imported graphs (see bellow)
Private Const RELATIVE     As Boolean = True        'Refers to imported graphs (see bellow)

'++----------------------------------------------------------------------------
'//   CALCULATE SHORTEST PATHS
'++----------------------------------------------------------------------------
'//   - This procedure is called from a frmNewInstance().CalculateShortestPaths
'++----------------------------------------------------------------------------
Public Function CalcShortestPaths(ByRef Dijkstra As clsDijkstra, _
                                  ByRef SimpleDraw As clsSimpleDraw, _
                                  ByRef TheGraph As clsGraph, _
                                  Optional ByVal Source As Long, _
                                  Optional ByVal Destination As Long, _
                                  Optional ByVal lAPSP As Long, _
                                  Optional ByVal blCoord As Boolean) As Double
Dim strTmp         As String
Dim iFileBin       As Integer
Dim i              As Long
Dim j              As Long
Dim k              As Long
Dim lButton        As Long
Dim lGNodes        As Long
Dim lGEdges        As Long
Dim sStepConst     As Single
Dim dMB            As Double
Dim dTemp          As Double
Dim dDijTimer      As Double
Dim dDrwTimer      As Double

Dim Predecessors() As Long
Dim Costs()        As Double

    With Dijkstra

        lGNodes = .GNodes
        lGEdges = .GEdges ' Update #1: Only for the benchmark (used in the progressbar)

        Select Case lAPSP
        Case 0
        '--------------------------------
        'Simple Shortest Path calculation
        '--------------------------------
            
            If Source And Source <> Destination Then
                
                If Not blMouseMoveCalc Then
                    frmMDI.MousePointer = vbHourglass
                    DoEvents
                End If

                j = Int(frmMDI.txtMDI(3))
                
                ReDim Predecessors(1 To lGNodes) As Long

                SetRTime
                For i = 1 To j
                    dTemp = .DijkstraHeap(Source, Destination)
                Next i
                dDijTimer = GetRTime
              
                .GetPredecessorsArray Predecessors

                SetRTime
                
                If blCoord Then
                    SimpleDraw.RefreshGraph
                    SimpleDraw.DrawShortestPath Predecessors, Destination
                End If
                
                dDrwTimer = GetRTime
                
            End If

        Case 1
        '-----------------------------------
        'All-Pairs Shortest Path calculation
        '-----------------------------------

            j = Int(frmMDI.txtMDI(3))
            
            frmMDI.MousePointer = vbHourglass
            DoEvents
            
            If blUpdate Then StartProcess lGNodes, sStepConst
            
            SetRTime
            For i = 1 To j
                For k = 1 To lGNodes
                    If blUpdate Then
                        If UpdateProcess(k, sStepConst, lGNodes, lGEdges) Then
                            i = j
                            CalcShortestPaths = -1
                            Exit For
                        End If
                    End If
                    
                    .DijkstraHeap k
                Next k
            Next i
            dDijTimer = GetRTime
            
            If blUpdate Then StopProcess

        Case 2
        '-------------------------------------
        'Save All-Pairs Shortest Paths to file
        '-------------------------------------

            dMB = (.GNodes / 1024)
            dMB = dMB * dMB * 8
            lButton = vbYes
            If dMB > 64 Then
                lButton = MsgBox("That will take up approximately " & vbLf & _
                          FormatNumber(dMB, 3, vbTrue) & " MBytes of disk space! Continue?", vbYesNo)
            End If
            
            If lButton = vbYes Then
                ReDim Costs(1 To .GNodes) As Double
                With frmMDI.CommonDialog1
                    .FileName = vbNullString
                    .Filter = "Binary Cost Matrix (*.bcm)|*.bcm"
                    .ShowSave
                End With
                    
                    strTmp = frmMDI.CommonDialog1.FileName
                
                    If LenB(strTmp) Then
                    
                        '//Update #1
                        If Dir$(strTmp) <> vbNullString Then
                            lButton = MsgBox("File exists. Overwrite?", vbYesNo)
                            If lButton = vbYes Then
                                If GetAttr(strTmp) = vbNormal Or vbArchive Then
                                    Kill (strTmp)
                                Else
                                    lButton = vbNo
                                End If
                            End If
                        End If
                        '//End Update
                        
                        If lButton = vbYes Then
                            iFileBin = FreeFile
                            Open Trim$(strTmp) For Binary Access Write As iFileBin
                            
                            frmMDI.MousePointer = vbHourglass
                            DoEvents
                            
                            If blUpdate Then StartProcess lGNodes, sStepConst
                            
                            SetRTime
                            For k = 1 To .GNodes
                                If blUpdate Then
                                    If UpdateProcess(k, sStepConst, lGNodes, lGEdges) Then
                                        CalcShortestPaths = -1 'clicked stop
                                        Exit For
                                    End If
                                End If
                                
                                .DijkstraHeap k
                                .GetCostArray Costs
                                Put iFileBin, , Costs()
                            Next k
                            Close iFileBin
                            dDijTimer = GetRTime
                            
                            If blUpdate Then StopProcess
                        End If
                    End If
                End If

        End Select

        With frmMDI
            .lblDijTime.Caption = "Dijkstra time = " & Round(dDijTimer, 3) & " sec."
            .lblDrawTime.Caption = "Drawing time = " & Round(dDrwTimer, 3) & " sec."
            .lblCost.Caption = "Cost = " & IIf(dTemp = dInf, "N/A", Round(dTemp, 5))
        End With
        frmMDI.lblExplored.Caption = "Graph Explored = " & Round(100 * .GraphExplored, 3) & "%"
        frmMDI.MousePointer = vbDefault
    End With

End Function

'++----------------------------------------------------------------------------
'//   IMPORT GRAPH
'++----------------------------------------------------------------------------
'//   - This procedure is called from a frmNewInstance().ImportGraphData
'//   - Imports the *.ga ASCII graphs
'++----------------------------------------------------------------------------
Public Function ImportGraph(ByVal strFileName As String, _
                            ByRef TheGraph As clsGraph) As Long
Dim Tmp         As String
Dim blEuclidean As Boolean
Dim dFTmulti    As Double
Dim dTFmulti    As Double
Dim iFile       As Integer

    dFTmulti = 1
    dTFmulti = 1
    iFile = FreeFile

    frmMDI.MousePointer = vbHourglass
    DoEvents

    Open strFileName For Input As #iFile

    Do
        Input #iFile, Tmp
        Tmp = UCase$(Tmp)

        If InStr(Tmp, "[METRIC]=EUCLIDEAN") Then
            '//This is an optional line and must be before graph topology.
            '//If this line exists, all topology costs are overriden by Euclidean distances,
            '//based on the XY coordinates.
            blEuclidean = True
        End If
        If InStr(Tmp, "[COST_FT*]=") Then
            '//This is an optional line and must be before graph topology.
            '//If this line exists, all topology FromTo costs (except -1)
            '//are multiplied by that value.
            dFTmulti = Val(Mid$(Tmp, 12, Len(Tmp)))
        End If
        If InStr(Tmp, "[COST_TF*]=") Then
            '//This is an optional line and must be before graph topology.
            '//If this line exists, all topology ToFrom costs (except -1)
            '//are multiplied by that value.
            dTFmulti = Val(Mid$(Tmp, 12, Len(Tmp)))
        End If
        If InStr(Tmp, "[GRAPH]") Then
            '//Start parsing the NON_RELATIVE graph. This means that node ID's are for e.g.:
            '// From To
            '// 10   23
            '// 11   25
            '// 7    22
            '// ...  ...
            Input #iFile, Tmp ' < This is the topology header (we don't care what it contains)
            Tmp = UCase$(ImportTopo(iFile, NON_RELATIVE, TheGraph, dFTmulti, dTFmulti))
        End If
        If InStr(Tmp, "[RELATIVE_GRAPH]") Then
            '//Start parsing the RELATIVE graph. This means that node ID's are for e.g.:
            '// From To
            '// 10   23
            '// 1     2
            '// -4   -3
            '// ...  ...
            '// If we iterativelly add each number with the one bellow, we get the
            '// NON_RELATIVE equivelant. The only practical purpose for this representation is that
            '// it can save considerable disk space for large graphs.
            Input #iFile, Tmp ' < This is the topology header (we don't care what it contains)
            Tmp = UCase$(ImportTopo(iFile, RELATIVE, TheGraph, dFTmulti, dTFmulti))
        End If
        If InStr(Tmp, "[COORDS]") Then
            '//Start parsing the NON_RELATIVE coordinates. The concept is identical to NON_RELATIVE graph.
            Input #iFile, Tmp ' < This is the coordinate header (we don't care what it contains)
            If ImportCoords(iFile, NON_RELATIVE, TheGraph, blEuclidean, dFTmulti, dTFmulti) Then
                ImportGraph = 1
            End If
            Exit Do
        End If
        If InStr(Tmp, "[RELATIVE_COORDS]") Then
            '//Start parsing the RELATIVE coordinates. The concept is identical to RELATIVE graph.
            Input #iFile, Tmp ' < This is the coordinate header (we don't care what it contains)
            If ImportCoords(iFile, RELATIVE, TheGraph, blEuclidean, dFTmulti, dTFmulti) Then
                ImportGraph = 1
            End If
            Exit Do
        End If

    Loop Until EOF(iFile) ' < Update #1: Fixed, in case of missing coordinates
    Close #iFile

    frmMDI.MousePointer = vbDefault

End Function

Public Function ImportTopo(ByVal iFile As Integer, _
                           ByVal bRelative As Boolean, _
                           ByRef TheGraph As clsGraph, _
                           Optional ByVal dFTmulti As Double = 1, _
                           Optional ByVal dTFmulti As Double = 1) As String
Dim Tmp       As String
Dim strArr()  As String
Dim lLastFrom As Long
Dim lLastTo   As Long
Dim dCostFT   As Double
Dim dCostTF   As Double
    
    With TheGraph

        Do While Not EOF(iFile)

            Input #iFile, Tmp
            If InStr(1, Tmp, "[") Then
                ImportTopo = Tmp ' We are finished, and Tmp is the current line
                Exit Do
            Else
                strArr = Split(Tmp, vbTab)
                If bRelative Then
                    lLastFrom = lLastFrom + Int(strArr(0))
                    lLastTo = lLastTo + Int(strArr(1))
                Else
                    lLastFrom = Int(strArr(0))
                    lLastTo = Int(strArr(1))
                End If
                dCostFT = 0
                dCostTF = 0
                If UBound(strArr) > 1 Then ' The value exists
                    dCostFT = Val(strArr(2)) * dFTmulti
                End If
                If UBound(strArr) > 2 Then ' The value exists
                    dCostTF = Val(strArr(3)) * dTFmulti
                End If
            End If

            .NewEdge lLastFrom, lLastTo, dCostFT, dCostTF

        Loop
        .TrimBuffers
    End With
End Function

Public Function ImportCoords(ByVal iFile As Integer, _
                             ByVal bRelative As Boolean, _
                             ByRef TheGraph As clsGraph, _
                             Optional ByVal bEucl As Boolean, _
                             Optional ByVal dFTmulti As Double = 1, _
                             Optional ByVal dTFmulti As Double = 1) As Long
Dim Tmp   As String
Dim i     As Long
Dim j     As Long
Dim LastX As Single
Dim LastY As Single
Dim dDist As Double

    With TheGraph

        For i = 1 To .GNodes
            If EOF(iFile) Then Exit For

            Input #iFile, Tmp
            j = InStr(1, Tmp, vbTab)

            If bRelative Then
                LastX = LastX + Val(Mid$(Tmp, 1, j - 1))
                LastY = LastY + Val(Mid$(Tmp, j + 1, Len(Tmp) - j))
            Else
                LastX = Val(Mid$(Tmp, 1, j - 1))
                LastY = Val(Mid$(Tmp, j + 1, Len(Tmp) - j))
            End If

            .SetNodeCoordinates i, LastX, LastY
        Next i

        If bEucl Then ' Polulate edges with euclidean distances
            For i = 1 To .GEdges
                LastX = .NodeXcoord(.EdgeFrom(i))
                LastX = (LastX - .NodeXcoord(.EdgeTo(i)))
                LastY = .NodeYcoord(.EdgeFrom(i))
                LastY = (LastY - .NodeYcoord(.EdgeTo(i)))

                dDist = Sqr(LastX * LastX + LastY * LastY)

                .UpdateEdge i, , , dDist * dFTmulti, dDist * dTFmulti
            Next i
        End If
    End With
    
    ImportCoords = 1 'Update #1: all coordinates exist
    
End Function

'++----------------------------------------------------------------------------
'//   PROGRESS BAR CONTROL
'++----------------------------------------------------------------------------

Private Sub StartProcess(ByVal lGNodes As Long, _
                               sStepConst As Single)
    frmMDI.Enabled = False
    frmProgress.Show
    sStepConst = frmProgress.shpBack.Width / lGNodes
    lElapsedPrevious = 0
End Sub

Private Function UpdateProcess(ByVal k As Long, _
                               ByVal sStepConst As Single, _
                               ByVal lGNodes As Long, _
                               ByVal lGEdges As Long) As Boolean
    Dim lElapsed As Long
    
    lElapsed = GetRTime * 1000
    
    If lElapsed - lElapsedPrevious > lProgress_ms Then
        If blStopProcess Then
            UpdateProcess = True
        Else
            lElapsedPrevious = lElapsed
            
            If Not blBenchMark Then
                frmProgress.shpProgress.Width = sStepConst * k
                frmProgress.lblProgress.Caption = Round(100 * k / lGNodes) & "%"
            Else ' Update #1: (used for benchmark)
                frmProgress.shpProgress.Width = sStepConst * (lGEdges / lGNodes)
                frmProgress.lblProgress.Caption = Round(100 * lGEdges / (lGNodes * lGNodes)) & "%"
            End If
            
            frmProgress.Caption = frmProgress.lblProgress.Caption
            DoEvents
        End If
    End If
    
End Function

Private Sub StopProcess()
    blStopProcess = False
    frmMDI.Enabled = True
    If Not blBenchMark Then Unload frmProgress ' Update #1 (unload only once, for the user to be able to press Stop)
End Sub




Public Function ShowTip(Optional InSt As String) As String
Dim i As Long, StIn As String
    Randomize Timer
    Dim Arr() As String
    If InSt = vbNullString Then
        Arr = Split("%%%*ngbc|+j+n`j+rgijidy{+ggb|+jc+ffC@*~dr+do+'ledg+_JC_+bj|+d+ej|+,edo+~dr+|dE@*xnfb+lebee~y+hnmmj+ggb|+B+*{bgdd+nc+cb|+rjg{+,edo+'qg[@*ree~m+ni+d+onxd{{~x+xj|+xbc_+*ngbfX+*ed+nfdH@*{{J+nc+d+`hjI+%n`da+{bgdd+xbc+md+cl~den+'`D@*{bgdd+j+fj+B+*bC@*yn~{fdh+ynxjm+j+r~i+1{b_@*ynjg+`hji+nfdh+oej+nnmmdh+j+n}jc+1{b_@*;:!+nfb+rg{bg~f+oej+.;:+gbe~+bj|+1{b_@bj|+nxjng[@{bgdd+yncdej+x~A@*xng~y+HX[@*{dx+`hbgH+4^[H+y~dr+lebjncyn}d+`xby+rc\@*FJY+eb+xcj[+xnydcX+nc+ggJ+nydx+,edo+B+0ryyd|+,edO@%xedbjg~hgjh+njebfyn+d+nync+`hbgH@%xcj{+xnydcx+nc+lebjg~hgjh+{dx+d+nf+`hbgH", "@")
        i = Int(Rnd() * (UBound(Arr))) + 1
        StIn = StrReverse(Arr(i))
    Else
        StIn = StrReverse(InSt)
    End If
    For i = 1 To Len(StIn)
        ShowTip = ShowTip & Chr$(Asc(Mid$(StIn, i, 1)) Xor 11)
    Next i
End Function
