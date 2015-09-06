Attribute VB_Name = "Mod_C2_Line_Mover"
Option Explicit
Sub Run_Lines_Mover()
'Called from a button in Worksheet plan
'For each collumn marked with X in checkRow will call Cones
Dim plan As Worksheet
Set plan = ThisWorkbook.Sheets("Line_mover_tool")

'typed arguments to pass to Cones
Dim msgCell As Range

'locations in plan where to find arguments
Dim startCol As Integer: startCol = 4
Dim endCol As Integer: endCol = 104

Dim folderRow As Integer: folderRow = 2
Dim netINrow  As Integer: netINrow = 4
Dim transitINrow  As Integer: transitINrow = 5
Dim netOUTrow  As Integer: netOUTrow = 7
Dim transitOUTRow  As Integer: transitOUTRow = 8
Dim instructionsRow  As Integer: instructionsRow = 10

Dim checkRow As Integer: checkRow = 11
Dim messageRow As Integer: messageRow = 12

On Error GoTo NoParam
Load_Network_Parameters

On Error GoTo 0
Dim icol As Integer
Dim folder As String
For icol = startCol To endCol
    If plan.Cells(checkRow, icol) = "X" Then

        folder = plan.Cells(folderRow, icol)
        If folder <> "" And Right(folder, 1) <> "\" Then folder = folder & "\"
        
        Set msgCell = plan.Cells(messageRow, icol)
        msgCell.Value = "Start: " & Format(Now(), "MMM/DD/YYYY hh:mm:ss")
        Call RunSplitCol(folder & plan.Cells(netINrow, icol), _
                    folder & plan.Cells(transitINrow, icol), _
                    folder & plan.Cells(netOUTrow, icol), _
                    folder & plan.Cells(transitOUTRow, icol), _
                    folder & plan.Cells(instructionsRow, icol), _
                    msgCell)
    End If
Next icol

Exit Sub:
NoParam:
    MsgBox "Fail to Load Network Parameters"
Exit Sub
Other:
    
End Sub
Sub RunSplitCol(FileBaseNetIN As String, _
                FileRouteIN As String, _
                FileBaseNetOUT As String, _
                FileRouteOUT As String, _
                FileSPLITTXT As String, _
                msgCell As Range)

marker = 0
ResetNetWork
ResetRoutes
'On Error GoTo Error_Handler

If Dir(FileBaseNetIN) = "" Then msgCell = msgCell & Chr(10) & "File " & FileBaseNetIN & " not found": Exit Sub
If Dir(FileRouteIN) = "" Then msgCell = msgCell & Chr(10) & "File " & FileRouteIN & " not found": Exit Sub

msgCell.Value = msgCell.Value & Chr(10) & "Reading Network    " & Format(Now(), "hh:mm:ss") & " ..."
Call ReadM2NetworkFile(FileBaseNetIN, msgCell)
msgCell = msgCell & Chr(10) & "... DONE: " & Npoints & " nodes and " & NLinks & " links  " ' & Format(Now(), "hh:mm:ss")
OldNpoints = Npoints
OldNLinks = NLinks

msgCell.Value = msgCell.Value & Chr(10) & "Reading Routes    " & Format(Now(), "hh:mm:ss") & " ..."
Call ReadM2RoutesFile(FileRouteIN, msgCell, False, False)
msgCell = msgCell & Chr(10) & "... DONE: " & nRoutes & " transit lines  " '& Format(Now(), "hh:mm:ss")
OldNRoutes = nRoutes
CountRoutesinNetwork ' this is called to fill links and nodes positions (used to move routes later)


SelectAllnotConnector
msgCell.Value = msgCell.Value & Chr(10) & "Reading Moving Instructions and Processing    " & Format(Now(), "hh:mm:ss") & " ..."
Call Read_LineMover_file(FileSPLITTXT, msgCell)

Dim onlyChanges As Boolean
onlyChanges = True ' (Format(OnlyOutputChanges, ">") = "X")

msgCell.Value = msgCell.Value & Chr(10) & "Writting routes  " & Format(Now(), "hh:mm:ss") & " ..."
WriteEmmeRoutes FileRouteOUT, False, True, True, onlyChanges

msgCell.Value = msgCell.Value & Chr(10) & "Writting changed network  " & Format(Now(), "hh:mm:ss") & " ..."
WriteEmmeNetwork FileBaseNetOUT, onlyChanges
   
msgCell = msgCell & Chr(10) & "... DONE" & Format(Now(), "hh:mm:ss")
   
Exit Sub

Error_Handler:
    Select Case Err.number
        Case 55
            Close #1
            Resume
        Case 52
            msgCell = msgCell & Chr(10) & "Erro 52 - File Name not recognized "
        Case Else
            msgCell = msgCell & Chr(10) & " Error: " & Err.number
    End Select
    Exit Sub
End Sub
Sub Read_LineMover_file(fullfile, msgCell As Range)
Dim stringline As String
Open fullfile For Input As #1
Dim iroute As Integer
Dim ipoint As Long
Dim i As Integer
Dim N As Integer

Do While Not EOF(1)
Read_Line:
    Line Input #1, stringline
    N = CountWords(stringline)
    
    If N = 0 Or Splitword(1) = "c" Or Splitword(1) = "C" Then
        'If empty line or comment goto next line
        GoTo Read_Line
    End If
        
    'Checking if route, (first word)  is valid
    iroute = Get_Route(Splitword(1))
    If iroute = 0 Then
        If vbNo = MsgBox("Transit line " & Splitword(1) & " in file " & fullfile & " not found in network. Proceed ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
            msgCell = msgCell = "Transit line " & Splitword(i) & " not found. Aborted moving lines after command:" & Chr(13) & stringline & Chr(13)
            Exit Sub
        Else
            msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
            GoTo Read_Line
        End If
    End If
        
    'Check if command (the second word) is valid
    If Splitword(2) <> "X>" And Splitword(2) <> "X<" And Splitword(2) <> "x>" And Splitword(2) <> "x<" And Splitword(2) <> "+>" And Splitword(2) <> "+<" And Splitword(2) <> "-" Then
        'Unknown signal
        If vbNo = MsgBox("Operation" & Splitword(2) & " in file " & fullfile & " unknown. Proceed ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
            msgCell = msgCell = "Command " & Splitword(2) & " unknown. Aborted moving lines after command:" & Chr(13) & stringline & Chr(13)
            Exit Sub
        Else
            msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
            GoTo Read_Line
        End If
    End If
        
    'Check all words after 2, shall be network points
    For i = 3 To N
        'uses point list i-2, to let in right poistion (first point in list is word 3)
        PointList(i - 2) = PointNamed(Splitword(i))
        If PointList(i - 2) = 0 Then
            If vbNo = MsgBox("Point" & Splitword(i) & " in file " & fullfile & " not found in network. Move ahead ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "?(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
                msgCell = msgCell = "Point " & Splitword(i) & " not found. Aborted moving lines after command:" & Chr(13) & stringline & Chr(13)
                Exit Sub
            Else
                msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
                GoTo Read_Line
            End If
         End If
    Next i
        
    'If everything is write up to this point, start processing
    
    If Splitword(2) = "-" Then
        'Remove line
        If N > 2 Then
            If vbNo = MsgBox("Operation" & Splitword(2) & " in file " & fullfile & " has more parameters than requested. Proceed ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
                msgCell = msgCell = "More parameters than required. Aborted moving lines after command:" & Chr(13) & stringline & Chr(13)
                Exit Sub
            Else
                msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
            End If
        Else
            route(iroute).changed = True
            route(iroute).deleted = True
        End If
    ElseIf Left(Splitword(2), 1) = "x" Or Left(Splitword(2), 1) = "X" Then
        'Cut line
        If N > 3 Then
            If vbNo = MsgBox("Operation" & Splitword(2) & " in file " & fullfile & " has more parameters than requested. Proceed ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
                msgCell = msgCell = "More parameters than required. Aborted moving lines after command:" & Chr(13) & stringline & Chr(13)
                Exit Sub
            Else
                msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
            End If
        Else
            ipoint = PointNamed(Splitword(3))
            If Not DoesRoutepasshere(iroute, ipoint) <> 0 Then
                If vbNo = MsgBox("Transit line " & Splitword(1) & " doesn't goes thru point " & Splitword(3) & ", as suposed in file " & fullfile & "Proceed ignoring command:" & Chr(13) & stringline & Chr(13) & Chr(13) & "(Não, aborta a execução do programa)", vbYesNo, "OOPS!") Then
                    msgCell = msgCell = "Transit line " & Splitword(1) & " doesn't goes thru point " & Splitword(3) & Chr(13) & stringline & Chr(13)
                    Exit Sub
                Else
                    msgCell = msgCell = "Ignoring command:" & Chr(13) & stringline & Chr(13)
                    GoTo Read_Line
                End If
            ElseIf Right(Splitword(2), 1) = ">" Then
                route(iroute).changed = Cut_Route(iroute, 0, ipoint)
            ElseIf Right(Splitword(2), 1) = "<" Then
                route(iroute).changed = Cut_Route(iroute, ipoint, 0)
            Else
                MsgBox "Bug found.. Error found after checking input... Can't process line..." & stringline
            End If
        End If
    ElseIf Left(Splitword(2), 1) = "+" Then
        If Right(Splitword(2), 1) = ">" Then
            route(iroute).changed = Extend_Route_TO(iroute, N - 2)
        ElseIf Right(Splitword(2), 1) = "<" Then
            route(iroute).changed = Extend_Route_FROM(iroute, N - 2)
        Else
                MsgBox "Bug found.. Error found after checking input... Can't process line..." & stringline
        End If
    Else
                MsgBox "Bug found.. Error found after checking input... Can't process line..." & stringline
    End If
Loop
Close #1
End Sub
Sub DetourAllRoutesBetween(P1 As Long, P2 As Long, Optional NNodesinPointList As Long = 0)
Dim iroute As Integer
For i = 1 To point(P1).nRoutes
    iroute = point(P1).iroute(i)
    For j = 1 To point(P2).nRoutes
        jroute = point(P2).iroute(j)
        If iroute = jroute And Not route(iroute).checked Then
            route(iroute).changed = RoutePassesThru(iroute, P1, P2, NNodesinPointList)
            Debug.Print route(iroute).number & " " & point(P1).Name & " " & point(P2).Name
        End If
    Next j
Next i

End Sub
Function RoutePassesThru(iroute As Integer, ByVal onepoint As Long, ByVal otherpoint As Long, Optional NNodesinPointList As Long = 0) As Boolean
Dim are As Boolean
Dim netdist As Single
Dim guarda(300)
    first = 0
    For i = 1 To route(iroute).Npoints
        If route(iroute).ipoint(i) = onepoint Then
            first = i
            Exit For
        End If
        If route(iroute).ipoint(i) = otherpoint Then
            first = i
            memo = otherpoint
            otherpoint = onepoint
            onepoint = memo
            Exit For
        End If
    Next i
    If first = 0 Then MsgBox ("just move along"): Exit Function
    last = 0
    For i = first To route(iroute).Npoints
        If route(iroute).ipoint(i) = otherpoint Then
            last = i
            Exit For
        End If
    Next i
    If first = 0 Or last = 0 Then MsgBox ("just move along"): Exit Function
    nguarda = 0
    For i = 1 To first
        nguarda = nguarda + 1
        guarda(nguarda) = route(iroute).ipoint(i)
    Next i
    For i = 1 To NNodesinPointList
        nguarda = nguarda + 1
        guarda(nguarda) = PointList(i)
    Next i
    For i = last To route(iroute).Npoints
        nguarda = nguarda + 1
        guarda(nguarda) = route(iroute).ipoint(i)
    Next i
    route(iroute).Npoints = nguarda
    ReDim route(iroute).ipoint(nguarda)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i
    RouteCaminhoMínimo (iroute)
    RoutePassesThru = True

End Function

Function RouteWillFindaWay(iroute As Integer, ByVal onepoint As Long, ByVal otherpoint As Long, Optional middlepoint As Long = 0) As Boolean
Dim are As Boolean
Dim netdist As Single
Dim guarda(300)
    first = 0
    For i = 1 To route(iroute).Npoints
        If route(iroute).ipoint(i) = onepoint Then
            first = i
            Exit For
        End If
        If route(iroute).ipoint(i) = otherpoint Then
            first = i
            memo = otherpoint
            otherpoint = onepoint
            onepoint = memo
            Exit For
        End If
    Next i
    If first = 0 Then MsgBox ("just move along"): Exit Function
    last = 0
    For i = first To route(iroute).Npoints
        If route(iroute).ipoint(i) = otherpoint Then
            last = i
            Exit For
        End If
    Next i
    If first = 0 Or last = 0 Then MsgBox ("just move along"): Exit Function
    nguarda = 0
    For i = 1 To first
        nguarda = nguarda + 1
        guarda(nguarda) = route(iroute).ipoint(i)
    Next i
    If middlepoint <> 0 Then
        nguarda = nguarda + 1
        guarda(nguarda) = middlepoint
    End If
    For i = last To route(iroute).Npoints
        nguarda = nguarda + 1
        guarda(nguarda) = route(iroute).ipoint(i)
    Next i
    route(iroute).Npoints = nguarda
    ReDim route(iroute).ipoint(nguarda)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i
    RouteCaminhoMínimo (iroute)
    RouteWillFindaWay = True

End Function
Function EsticaAteTerminal(iroute As Integer, terminal As Long) As Boolean
Dim are As Boolean
Dim netdist As Single
Dim guarda(300)
    For i = 1 To route(iroute).Npoints
        point(route(iroute).ipoint(i)).Distance = 100000
    Next i
    nguarda = 0
    selectedin = 1
    N = NLinksInShortestPathDijkstraTO(terminal, route(iroute).ipoint(1), are, netdist)
    For i = 2 To route(iroute).Npoints
        If point(route(iroute).ipoint(i)).Distance < point(route(iroute).ipoint(selectedin)).Distance Then selectedin = i
    Next i
    If point(route(iroute).ipoint(1)).Distance < point(route(iroute).ipoint(route(iroute).Npoints)).Distance Then
        nguarda = nguarda + 1
        guarda(nguarda) = terminal
        For i = selectedin To route(iroute).Npoints
            nguarda = nguarda + 1
            guarda(nguarda) = route(iroute).ipoint(i)
        Next i
    Else
        For i = 1 To selectedin
            nguarda = nguarda + 1
            guarda(nguarda) = route(iroute).ipoint(i)
        Next i
        nguarda = nguarda + 1
        guarda(nguarda) = terminal
    End If
    route(iroute).Npoints = nguarda
    ReDim route(iroute).ipoint(nguarda)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i
    RouteCaminhoMínimo (iroute)
    EsticaAteTerminal = True
End Function
Function EntranoTerminal(iroute As Integer, terminal As Long) As Boolean
Dim are As Boolean
Dim netdist As Single
Dim guarda(300)
    If Distância(route(iroute).ipoint(1), terminal) < 600 Or Distância(route(iroute).ipoint(route(iroute).Npoints), terminal) < 600 Then
        EsticaAteTerminal iroute, terminal
    Else
        For i = 1 To route(iroute).Npoints
            point(route(iroute).ipoint(i)).Distance = 100000
        Next i
        N = NLinksInShortestPathDijkstraTO(terminal, route(iroute).ipoint(1), are, netdist)
        If Not are Then MsgBox "don't be scared"
        selectedin = 1
        For i = 2 To route(iroute).Npoints
            If point(route(iroute).ipoint(i)).Distance < point(route(iroute).ipoint(selectedin)).Distance Then selectedin = i
        Next i
        For i = 1 To route(iroute).Npoints
            point(route(iroute).ipoint(i)).Distance = 100000
        Next i
        N = NLinksInShortestPathDijkstra(terminal, route(iroute).ipoint(route(iroute).Npoints), are, netdist)
        If Not are Then MsgBox "don't be scared"
        selectedout = route(iroute).Npoints
        For i = route(iroute).Npoints - 1 To selectedin Step -1
            If point(route(iroute).ipoint(i)).Distance < point(route(iroute).ipoint(selectedout)).Distance Then selectedout = i
        Next i
        nguarda = 0
        
        For i = 1 To selectedin
            nguarda = nguarda + 1
            guarda(nguarda) = route(iroute).ipoint(i)
        Next i
        nguarda = nguarda + 1
        guarda(nguarda) = terminal
        For i = selectedout To route(iroute).Npoints
            nguarda = nguarda + 1
            guarda(nguarda) = route(iroute).ipoint(i)
        Next i
        route(iroute).Npoints = nguarda
        ReDim route(iroute).ipoint(nguarda)
        For i = 1 To nguarda
            route(iroute).ipoint(i) = guarda(i)
        Next i
        RouteCaminhoMínimo (iroute)
    End If
    EntranoTerminal = True
    'TiradoCentro
    For i = 1 To route(iroute).Npoints
        If Distância(route(iroute).ipoint(i), PointNamed(8014)) > 1500 And point(route(iroute).ipoint(i)).Limite < 4 Then
            StartP = i
            Exit For
        End If
    Next i
    For i = route(iroute).Npoints To 1 Step -1
        If Distância(route(iroute).ipoint(i), PointNamed(8014)) > 1500 And point(route(iroute).ipoint(i)).Limite < 4 Then
            EndP = i
            Exit For
        End If
    Next i
    nguarda = 0
    For i = StartP To EndP
        nguarda = nguarda + 1
        guarda(nguarda) = route(iroute).ipoint(i)
    Next i
    route(iroute).Npoints = nguarda
    ReDim route(iroute).ipoint(nguarda)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i

End Function
Sub AllRoutesCaminhoMínimo()
    For i = 1 To nRoutes
        RouteCaminhoMínimo (i)
    Next i
End Sub
Sub PrintMissingLinks()
For iroute = 1 To nRoutes
    If Not onlyChanges Or (onlyChanges And (i > OldNRoutes Or route(iroute).changed)) Then
        LastPoint = 0
        For j = 1 To route(iroute).Npoints
            If route(iroute).ipoint(j) = 0 And Not NeverMindNodes Then
                If MsgBox("Route: " & route(iroute).number & " should pass in unknown point " & point(route(iroute).ipoint(j)).Name & ". The point will be ignored if proceed, proceed anyway? ", vbYesNo, "WriteM2Routes") = vbNo Then End
            ElseIf route(iroute).ipoint(j) = LastPoint And Not NeverMindNodes Then
                If MsgBox("Route: " & route(iroute).number & " passes twice in known point " & point(route(iroute).ipoint(j)).Name & ". The point will be ignored if proceed, proceed anyway? ", vbYesNo, "WriteM2Routes") = vbNo Then End
            Else
                If LastPoint <> 0 Then
                    il = Get_Link(route(iroute).ipoint(LastPoint), route(iroute).ipoint(j))
                    If il = 0 And route(iroute).ipoint(LastPoint) <> route(iroute).ipoint(j) Then
                        If Not CreateNeededLinksandModes Then
                            If MsgBox("Route: " & route(iroute).number & " should pass in unknown link between points " & _
                                      point(route(iroute).ipoint(LastPoint)).Name & " and " & point(route(iroute).ipoint(j)).Name & _
                                      "." & Chr(13) & " The link will be created if proceed, proceed anyway? ", vbYesNo, "ReadM2Routes") = vbNo Then End
                        End If
                        If LastPoint <> route(iroute).ipoint(j) Then
                            Dlink.op = route(iroute).ipoint(LastPoint)
                            Dlink.dp = route(iroute).ipoint(j)
                            Dlink.Extension = Distância(Dlink.op, Dlink.dp)
                            Dlink.tipo = 1000
                            Dlink.modes = route(iroute).mode
                            Dlink.Lanes = 0
                            il = Add_Link(Dlink)
                        End If
                    End If
                    If Not LinkhasMode(il, route(iroute).mode) Then
                        If Not CreateNeededLinksandModes Then
                            If MsgBox("Route: " & route(iroute).number & " Link " & _
                                      point(route(iroute).ipoint(LastPoint)).Name & " and " & point(route(iroute).ipoint(j)).Name & _
                                      " doesn't take mode " & route(iroute).mode & "." & Chr(13) & "The mode will be add if proceed, proceed anyway? ", vbYesNo, "ReadM2Routes") = vbNo Then End
                            End If
                        link(il).modes = link(il).modes & route(iroute).mode
                        If link(il).tipo < 1000 Then link(il).tipo = link(il).tipo + 1000
                    End If
                End If
                LastPoint = j
            End If
        Next j
        If route(iroute).path = "" Then route(iroute).path = "no"
        If route(iroute).lay = "" Then route(iroute).lay = 5
    End If
Next iroute
End Sub
Sub Read_Move_Command_file(fullfile)
Dim stringline As String
Open fullfile For Input As #1
Do While Not EOF(1)
    Line Input #1, stringline
    N = CountWords(stringline)
    If N > 0 And Splitword(1) <> "c" Then
        If Splitword(1) = "move" Then Debug.Print Move_Route_CommandLine(stringline)
    End If
Loop
Close #1

End Sub
Sub Duzinho()
Dim stringline As String
    Open "C:\Users\a\Documents\Professional\Emmeteste\New_Project\Database\TypeRev" For Input As #1
    Do While Not EOF(1)
        Line Input #1, stringline
        N = CountWords(stringline)
        il = Get_Link(PointNamed(Val(Splitword(1))), PointNamed(Val(Splitword(2))))
        If il = 0 Then
            MsgBox "Chevete"
        Else
            link(il).tipo = Int(link(il).tipo / 100) + Val(Splitword(3))
            link(il).isM2 = 3
        End If
    Loop
    Close #1
End Sub
Sub ReplaceNodeinRoute(FromPointName As Long, ToPointName As Long)
i590 = PointNamed(FromPointName)
i59 = PointNamed(ToPointName)
If i590 <> 0 Then
    If i59 = 0 Then MsgBox "tem 590 mas nao tem 59!"
    For ir = 1 To point(i590).nRoutes
        iroute = point(i590).iroute(ir)
        For i = 1 To route(iroute).Npoints
            If route(iroute).ipoint(i) = i590 Then route(iroute).ipoint(i) = i59
        Next i
    Next ir
End If
End Sub
Sub ExtendROute(iroute)
Dim touch As Boolean
If point(route(iroute).ipoint(1)).DistânciadaRede < 150 Then
    If point(route(iroute).ipoint(1)).Estação <> 1 Then
        touch = False
        ENE = NLinksInShortestPath(point(route(iroute).ipoint(1)).Closer, route(iroute).ipoint(1), touch)
        If touch Then
            ReDim Preserve route(iroute).ipoint(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).para(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).dwt(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).ttf(route(iroute).Npoints + ENE)
            For ip = route(iroute).Npoints + ENE To ENE + 1 Step -1
                route(iroute).ipoint(ip) = route(iroute).ipoint(ip - ENE)
                route(iroute).para(ip) = route(iroute).para(ip - ENE)
                route(iroute).dwt(ip) = route(iroute).dwt(ip - ENE)
                route(iroute).ttf(ip) = route(iroute).ttf(ip - ENE)
            Next ip
            For ip = ENE To 1 Step -1
                route(iroute).ipoint(ip) = link(LinkList(ip)).op
            Next ip
            route(iroute).Npoints = route(iroute).Npoints + ENE
        Else
            Debug.Print route(iroute).number & " requires integration attention on beggining"
        End If
    End If
End If
If point(route(iroute).ipoint(route(iroute).Npoints)).DistânciadaRede < 150 Then
    If point(route(iroute).ipoint(route(iroute).Npoints)).Estação <> 1 Then
        touch = False
        ENE = NLinksInShortestPath(route(iroute).ipoint(route(iroute).Npoints), point(route(iroute).ipoint(route(iroute).Npoints)).Closer, touch)
        If touch Then
            ReDim Preserve route(iroute).ipoint(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).para(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).dwt(route(iroute).Npoints + ENE)
            ReDim Preserve route(iroute).ttf(route(iroute).Npoints + ENE)
            For ip = 1 To ENE
                route(iroute).ipoint(route(iroute).Npoints + ip) = link(LinkList(ip)).dp
            Next ip
            route(iroute).Npoints = route(iroute).Npoints + ENE
        Else
            Debug.Print route(iroute).number & " requires integration attention on beggining"
        End If
    End If
End If
End Sub
Sub RouteCaminhoMínimo(iroute As Integer)
Dim PANT As Long
Dim P As Long
Dim guarda(300) As Long
Dim Arelinked As Boolean
Dim NetDis As Single
    nguarda = 1
    guarda(nguarda) = route(iroute).ipoint(1)
    For ipon = 1 To route(iroute).Npoints - 1
        PANT = route(iroute).ipoint(ipon)
        P = route(iroute).ipoint(ipon + 1)
        il = Get_Link(PANT, P)
        If il = 0 Then
            N = NLinksInShortestPathDijkstra(PANT, P, Arelinked, NetDis)
            If Not Arelinked Then MsgBox "Hello, should be nothing serious"
            For i = 1 To N
                nguarda = nguarda + 1
                guarda(nguarda) = link(LinkList(i)).dp
            Next i
        Else
            nguarda = nguarda + 1
            guarda(nguarda) = P
        End If
    Next ipon
    route(iroute).Npoints = nguarda
    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i
End Sub
Function Move_Route_CommandLine(ByVal stringline As String) As Integer
' Command line: "move A1 A2 A3 to B1 B2 B3
' Replaces itinerary of all routes that have A1 A2... AN by B1 B2 B3 ... BN
' This routine does not check if the links or points FROM or TO exists.
' if points don´t exist will use 0!
Dim LineFrom As Route_Type
Dim LineTo As Route_Type
Dim LineGuarda As Route_Type
Move_Route_CommandLine = 0
If Left(stringline, 5) <> "move " Then MsgBox "Chamada de Move_Route com linha inválida"
stringline = Mid(stringline, 5)
If CountStrings(stringline, "to") <> 2 Then MsgBox "Chamada de Move_Route com linha inválida"
LineFrom.Npoints = CountStringsB(Splitword(1), " ")
ReDim LineFrom.ipoint(LineFrom.Npoints)

For i = 1 To LineFrom.Npoints
    LineFrom.ipoint(i) = PointNamed(SplitwordB(i))
Next i

LineTo.Npoints = CountStringsB(Splitword(2), " ")
ReDim LineTo.ipoint(LineTo.Npoints)
For i = 1 To LineTo.Npoints
    LineTo.ipoint(i) = PointNamed(SplitwordB(i))
Next i

For i = 1 To point(LineFrom.ipoint(1)).nRoutes
    iroute = point(LineFrom.ipoint(1)).iroute(i)
    
    'Routine is here
    '1. Check if Route pass thru all FROM points in order, mark first and last
again:
    For j = 1 To route(iroute).Npoints
          If route(iroute).ipoint(j) = LineFrom.ipoint(1) Then
                bate = False
                If route(iroute).Npoints >= j - 1 + LineFrom.Npoints Then
                bate = True
                    For k = 1 To LineFrom.Npoints
                        If route(iroute).ipoint(j - 1 + k) <> LineFrom.ipoint(k) Then bate = False
                    Next k
                End If
                If bate Then
                    Move_Route_CommandLine = Move_Route_CommandLine + 1
                    '2. If routes pass thru all of them move route to guarda it till the end, makes coding here easier
                    LineGuarda.Npoints = route(iroute).Npoints + LineTo.Npoints - LineFrom.Npoints
                    ReDim LineGuarda.ipoint(LineGuarda.Npoints)
                    igua = 0
                    For k = 1 To j - 1
                        igua = igua + 1
                        LineGuarda.ipoint(igua) = route(iroute).ipoint(k)
                    Next k
                    For k = 1 To LineTo.Npoints
                        igua = igua + 1
                        LineGuarda.ipoint(igua) = LineTo.ipoint(k)
                    Next k
                    For k = j + LineFrom.Npoints To route(iroute).Npoints
                        igua = igua + 1
                        LineGuarda.ipoint(igua) = route(iroute).ipoint(k)
                    Next k
                    If igua <> LineGuarda.Npoints Then MsgBox " Check,please! "
                    route(iroute).Npoints = LineGuarda.Npoints
                    ReDim route(iroute).ipoint(route(iroute).Npoints)
                    For k = 1 To LineGuarda.Npoints
                        route(iroute).ipoint(k) = LineGuarda.ipoint(k)
                    Next k
                    GoTo again
                End If
          End If
    Next j
Next i

End Function
Sub EntrandonoTerminal()
Dim plan As Worksheet
Dim Dlink As Link_type
Dim Dpoint As Point_type
Dim nomim As String
Dim Pfrom As Long
Dim pinteg As Long
Dim areli As Boolean
Dim distsai As Single
Dim distentra As Single
Dim P As Long
Dim guarda(1000) As Long
'Criando os terminais
'OldNpoints = Npoints 'quando vai direto, tira isso
'OldNLinks = NLinks
Set plan = ThisWorkbook.Sheets("ParaMacro")
For iRow = 20 To 22
    icol = 20
    nomim = plan.Cells(iRow, icol)
    P2 = PointNamed(plan.Cells(iRow, icol + 3))
    P1 = PointNamed(plan.Cells(iRow, icol + 2))
    xis = (point(P1).x + point(P2).x) / 2
    ypis = (point(P1).y + point(P2).y) / 2
    pinteg = Add_Point(xis, ypis, 3, nomim)
    For jcol = icol + 4 To icol + 5
        Pfrom = PointNamed(plan.Cells(iRow, jcol))
        If Pfrom <> 0 Then
            Dlink.op = Pfrom
            Dlink.dp = pinteg
            Dlink.Extension = Distância(Pfrom, pinteg)
            Dlink.Lanes = 2
            Dlink.modes = "bcdip"
            Dlink.tipo = 13
            Dlink.vdf = 20
            A = Add_Link(Dlink)
        End If
    Next jcol
    For jcol = icol + 6 To icol + 7
        PTo = PointNamed(plan.Cells(iRow, jcol))
        If PTo <> 0 Then
            Dlink.op = pinteg
            Dlink.dp = PTo
            Dlink.Extension = Distância(Dlink.dp, pinteg)
            Dlink.Lanes = 2
            Dlink.modes = "bcdip"
            Dlink.vdf = 20
            A = Add_Link(Dlink)
        End If
    Next jcol
Next iRow
Set plan = Workbooks("Sistemas Tronco Alimentados.xls").Sheets("Cenário Conservador")
For icol = 1 To 41 Step 5
    P = PointNamed(plan.Cells(2, icol))
    iRow = 4
    roUtenumb = plan.Cells(iRow, icol)
    mult = 5
    CorreRate = 1
    If icol = 41 Then
        CorreRate = 0.1
        mult = 2
    End If
    Do While roUtenumb <> ""
        iroute = 0
        For ir = 1 To nRoutes
            If route(ir).number = roUtenumb Then
                iroute = ir
                Exit For
            End If
        Next ir
        If ir = 0 Then
            plan.Cells(iRow, icol + 3) = "nope"
        Else
            cutin = 0
            distcomp = 100000
            cutout = 0
            NP = route(iroute).Npoints
            For ip = 1 To route(iroute).Npoints
                dista = Distância(Val(P), route(iroute).ipoint(ip))
                If dista < distcomp Then distcomp = dista
            Next ip
            distin = 100000
            distout = 100000
            Divideyet = False
            For ip = 1 To route(iroute).Npoints
                ipor = route(iroute).ipoint(ip)
                If point(ipor).Name < 10000 Then
                    dista = Distância(Val(P), Val(ipor))
                    If dista < mult * distcomp Then
                        entra = NLinksInShortestPath(ipor, P, areli, distentra)
                        If distentra <= distin Then
                            distin = distentra
                            cutin = ip
                            If entra = 1 And Not Divideyet Then distcomp = distcomp / 3: Divideyet = True
                        End If
                    End If
                    If dista < mult * distcomp Then
                        sai = NLinksInShortestPath(P, ipor, areli, distsai)
                        If distsai < distout Then
                            distout = distsai
                            cutout = ip
                            If sai = 1 And Not Divideyet Then distcomp = distcomp / 3: Divideyet = True
                        End If
                    End If
                End If
            Next ip
            ip = 0
            If Left(route(iroute).number, 3) = "314" Then
                H = ello
            End If
            If True Then 'false Then
                If Distância(route(iroute).ipoint(1), P) < 400 Then
                    'Parte do Terminal para cutout
                    N1 = NLinksInShortestPath(P, route(iroute).ipoint(cutout), areli, distentra)
                    For i = 1 To N1
                        ip = ip + 1
                        guarda(ip) = link(LinkList(i)).op
                    Next i
                    If Distância(route(iroute).ipoint(route(iroute).Npoints), P) < 400 Then
                        'entra no terminal ao final. Precisa verificar se cutin não ficou no começo da linha
                        If cutin <= cutout Then
                            MsgBox "ROute " & route(iroute).number
                        Else
                            For i = cutout To cutin
                                ip = ip + 1
                                guarda(ip) = route(iroute).ipoint(i)
                            Next i
                            N1 = NLinksInShortestPath(route(iroute).ipoint(cutin), P, areli, distentra)
                            For i = 1 To N1
                                ip = ip + 1
                                guarda(ip) = link(LinkList(i)).dp
                            Next i
                        End If
                    Else
                        'Não entra no terminal ao final
                        For i = cutout To route(iroute).Npoints
                            ip = ip + 1
                            guarda(ip) = route(iroute).ipoint(i)
                        Next i
                    End If
                Else
                    'Entra no terminal apartir de cutin
                    For i = 1 To cutin
                        ip = ip + 1
                        guarda(ip) = route(iroute).ipoint(i)
                    Next i
                    N1 = NLinksInShortestPath(route(iroute).ipoint(cutin), P, areli, distentra)
                    For i = 1 To N1
                        ip = ip + 1
                        guarda(ip) = link(LinkList(i)).dp
                    Next i
                    If Distância(route(iroute).ipoint(route(iroute).Npoints), P) < 400 Then
                        'não sai do terminal no terminal
                        'DONE
                    Else
                        'sai do terminal
                        N1 = NLinksInShortestPath(P, route(iroute).ipoint(cutout), areli, distentra)
                        For i = 1 To N1
                            ip = ip + 1
                            guarda(ip) = link(LinkList(i)).dp
                        Next i
                        'Termina a linha
                        For i = cutout + 1 To route(iroute).Npoints
                            ip = ip + 1
                            guarda(ip) = route(iroute).ipoint(i)
                        Next i
                    End If
                End If
                route(iroute).Npoints = ip
                ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
                For q = 1 To route(iroute).Npoints
                     route(iroute).ipoint(q) = guarda(q)
                Next q
                Debug.Print point(P).Name & " "; route(iroute).number
                route(iroute).checked = True
            Else
                plan.Cells(iRow, icol + 2) = distin
                plan.Cells(iRow, icol + 3) = distout
            End If
        End If
        iRow = iRow + 1
        roUtenumb = plan.Cells(iRow, icol)
    Loop
Next icol
End Sub

