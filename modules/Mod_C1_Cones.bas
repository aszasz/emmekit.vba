Attribute VB_Name = "Mod_C1_Cones"
Option Explicit

'Public NGroups As Integer
'Public TableofIntegration(50, 50) As Single
'Public MetrosPorRealapé As Single
'Public TableofIntegration2(50, 50) As Single
'Public RoutesFirstEncounter(2500, 2500) As Long 'This creates a huge table
'Public RoutesNeedEncounter(2500, 2500) As Long
'Public RoutesEncounter(2500, 2500) As Long

'Public RoutesDistanceBeforeAfterFirstEncounter(2500, 2500) As Long

' Integration
Dim NFareGroups As Integer
Dim FareGroupCode() As String
Dim NRoutesInGroup() As Integer
Dim TableofIntegration(50, 50) As Single 'Fare prices to first boarding and integration
Dim MetrosPorRealapé As Single 'Distanceper$walking: 1.00$ is equivalent to walking this distance
Sub Run_Cones()
'Called from a button in Worksheet plan
'For each collumn marked with X in checkRow will call Cones
Dim plan As Worksheet
Set plan = ThisWorkbook.Sheets("Cones_tool")

'typed arguments to pass to Cones
Dim msgCell As Range

'locations in plan where to find arguments
Dim startCol As Integer: startCol = 3
Dim endCol As Integer: endCol = 104

Dim folderRow As Integer: folderRow = 3
Dim netINrow  As Integer: netINrow = 5
Dim transitINrow  As Integer: transitINrow = 6
Dim walkingModeRow  As Integer: walkingModeRow = 8
Dim specialLinkTypeRow  As Integer: specialLinkTypeRow = 9
Dim coneHeightRow As Integer: coneHeightRow = 10
Dim chargingModeRow  As Integer: chargingModeRow = 11
Dim walkingCostRow  As Integer: walkingCostRow = 12
Dim fareTableRow As Integer: fareTableRow = 13
Dim groupListRow As Integer: groupListRow = 14
Dim penaltyRow As Integer: penaltyRow = 15
Dim OnetransferRow As Integer: OnetransferRow = 16
Dim OnlyChangesRow As Integer: OnlyChangesRow = 18
Dim netOUTrow  As Integer: netOUTrow = 19
Dim transitOUTRow  As Integer: transitOUTRow = 20


Dim checkRow As Integer: checkRow = 21
Dim messageRow As Integer: messageRow = 22

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
        Call Cones(folder & plan.Cells(netINrow, icol), folder & plan.Cells(transitINrow, icol), folder & plan.Cells(netOUTrow, icol), folder & plan.Cells(transitOUTRow, icol), _
                    plan.Cells(fareTableRow, icol), _
                    plan.Cells(groupListRow, icol), _
                    plan.Cells(walkingModeRow, icol), _
                    plan.Cells(chargingModeRow, icol), _
                    plan.Cells(walkingCostRow, icol), _
                    plan.Cells(specialLinkTypeRow, icol), _
                    plan.Cells(penaltyRow, icol), _
                    plan.Cells(coneHeightRow, icol), _
                    plan.Cells(OnetransferRow, icol), _
                    plan.Cells(OnlyChangesRow, icol), _
                    msgCell)
    End If
Next icol

Exit Sub:
NoParam:
    MsgBox "Fail to Load Network Parameters"
Exit Sub
Other:
    
End Sub
' Transform IN network and route to OUT network and routes, creating integration cones
' where integrated routes, according to provided tables (in IntFare and LineGroupList)
' stops only inside the cones they are diverted to
' - Users can acess boarding point only through auxiliary mode chargeMode
'         - fare price is placed in ul2 (as provided by table IntFare)
' - New boarding nodes will try to have the same number end + 100,000*group and alighting nodes + 100,000*group1 + 50,000
' - Will add integration links between cones from points linked by links with ty IntegrationLinkType
'         - with an additional walking cost given by equivWalkDistance (in ul2)
' - Will add transferPenaltyCost to ul2
' Add log to MsgCell
Sub Cones(FileBaseNetIN As String, FileRouteIN As String, FileBaseNetOUT As String, FileRouteOUT As String, _
          IntFareRangeString As String, _
          LineGroupListRange As String, _
          walkmode As String, _
          chargemode As String, _
          equivWalkDist As Single, _
          integrationLinkType As Integer, _
          transferPenaltyCost As Single, _
          coneHeight_in_meters As Single, _
          OnlyONETransferX As String, _
          OnlyOutputChanges As String, _
          msgCell As Range)

marker = 0
Dim i As Long
Dim iroute As Integer
Dim IntFare As Range

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
CountRoutesinNetwork ' this is called to fill links and nodes positions (used to split node later)

'check walkmode
If Len(walkmode) <> 1 Then
    msgCell = msgCell & Chr(10) & "Exiting early: Walking mode (" & walkmode & ") should be only one letter"
    Exit Sub
ElseIf Asc(walkmode) <= Asc("a") Or Asc(walkmode) >= Asc("z") Then
    msgCell = msgCell & Chr(10) & "Warning: Expected walking mode (" & walkmode & ") to be a non capital letter"
End If
Dim sum As Long
Dim ilink As Long
sum = 0
For ilink = 1 To NLinks
    If LinkhasMode(ilink, walkmode) Then sum = sum + 1
Next ilink
msgCell = msgCell & Chr(10) & " Found " & sum & " links with walking mode (" & walkmode & ")"

'check integrationlinktype
If Not IsNumeric(integrationLinkType) Then
    msgCell = msgCell & Chr(10) & "Exiting early: Integration link type (" & integrationLinkType & ") should be numeric"
    Exit Sub
ElseIf integrationLinkType >= 1000 Or integrationLinkType <= 0 Then
    msgCell = msgCell & Chr(10) & "Warning: Expected integration link type (" & integrationLinkType & ") to be between 1 and 999"
End If
sum = 0
For ilink = 1 To NLinks
    If link(ilink).tipo = integrationLinkType Then sum = sum + 1
Next ilink
msgCell = msgCell & Chr(10) & " Found " & sum & " integration links (type=" & integrationLinkType & ")"

'check chargemode
If Len(chargemode) <> 1 Then
    msgCell = msgCell & Chr(10) & "Exiting early: Charge mode (" & chargemode & ") should be only one letter"
    Exit Sub
ElseIf Asc(chargemode) <= Asc("a") Or Asc(chargemode) >= Asc("z") Then
    msgCell = msgCell & Chr(10) & "Warning: Expected auxiliary charging mode (" & chargemode & ") to be a non capital letter"
End If
sum = 0
For ilink = 1 To NLinks
    If LinkhasMode(ilink, chargemode) Then sum = sum + 1
Next ilink
If sum > 0 Then msgCell = msgCell & Chr(10) & "Warning:  Found already " & sum & " links with charging mode (" & chargemode & ") in the network"

'' Read Fare Integration
msgCell = msgCell & Chr(10) & "Reading Integration Table    " & Format(Now(), "hh:mm:ss") & " ..."
On Error GoTo NoFareRange
Set IntFare = String_to_Range(IntFareRangeString)
On Error GoTo 0
If InputTableofIntegration(IntFare, msgCell) Then
    msgCell = msgCell & Chr(10) & "... DONE: " & NFareGroups & " faregroups  ("
        For i = 1 To NFareGroups - 1
            msgCell = msgCell & FareGroupCode(i) & ", "
    Next i
    msgCell = msgCell & FareGroupCode(i) & ")"
Else
    msgCell = msgCell & Chr(10) & "Fail to Read Fares in: " & IntFareRangeString
    Exit Sub
End If

'Put routes in faregroups
msgCell = msgCell & Chr(10) & "Reading routes fare groups    " & Format(Now(), "hh:mm:ss") & " ..."
If Not AssignRouteGroupsFromField(LineGroupListRange, msgCell) Then
    msgCell = msgCell & Chr(10) & "Fail to Read Transit Lines Fare Groups in: " & LineGroupListRange
    Exit Sub
Else
    msgCell = msgCell & Chr(10) & "... DONE: "
        For i = 0 To NFareGroups
            msgCell = msgCell & Chr(10) & " " & FareGroupCode(i) & " --> " & NRoutesInGroup(i) & " lines"
    Next i
End If

'Mark Points to Expand
    msgCell = msgCell & Chr(10) & "Selected " & NSelected_Nodes_to_Cone & " points to expand for integration "


'Create all cones nodes
msgCell.Value = msgCell.Value & Chr(10) & "Creating new nodes    " & Format(Now(), "hh:mm:ss") & " ..."
For i = 1 To Npoints
    If point(i).selected Then
        SplitPoint i, coneHeight_in_meters
    End If
Next i
msgCell = msgCell & Chr(10) & "... DONE: " & Npoints - OldNpoints & " new nodes ( total = " & Npoints & " )"


Dim onlyOneTransfer As Boolean
onlyOneTransfer = (Format(OnlyONETransferX, ">") = "X")

'Create all cones links
msgCell.Value = msgCell.Value & Chr(10) & "Creating new links    " & Format(Now(), "hh:mm:ss") & " ..."
For i = 1 To Npoints
    If point(i).selected Then
        RelinkPoint i, equivWalkDist, transferPenaltyCost, walkmode, chargemode, integrationLinkType, onlyOneTransfer
    End If
Next i

msgCell = msgCell & Chr(10) & "... DONE: " & NLinks - OldNLinks & " new links ( total = " & NLinks & " )"

If onlyOneTransfer Then
    msgCell.Value = msgCell.Value & Chr(10) & "Duplicating and Rerouting integrated transit lines   " & Format(Now(), "hh:mm:ss") & " ..."
Else
    msgCell.Value = msgCell.Value & Chr(10) & "Rerouting transit lines   " & Format(Now(), "hh:mm:ss") & " ..."
End If

sum = 0
For iroute = 1 To nRoutes
    If route(iroute).igroup > 0 Then
        If onlyOneTransfer Then
            RerouteONETRANSFER iroute
        Else
            Reroute iroute
        End If
    End If
Next iroute

sum = 0
For iroute = 1 To nRoutes
    sum = sum + route(iroute).Npoints 'check to see if should report on less per route, there is a virtual at the end
Next iroute

msgCell = msgCell & Chr(10) & "... DONE: " & nRoutes - OldNRoutes & " new routes, total transit segments = " & sum & " )"

Dim onlyChanges As Boolean
onlyChanges = (Format(OnlyOutputChanges, ">") = "X")

msgCell.Value = msgCell.Value & Chr(10) & "Writting routes  " & Format(Now(), "hh:mm:ss") & " ..."
WriteEmmeRoutes FileRouteOUT, False, True, True, onlyChanges

msgCell.Value = msgCell.Value & Chr(10) & "Writting changed network  " & Format(Now(), "hh:mm:ss") & " ..."
WriteEmmeNetwork FileBaseNetOUT, onlyChanges


msgCell.Value = msgCell.Value & Chr(10) & "...DONE (don't forget to add  " & chargemode & " mode to your scenario!" & Format(Now(), "hh:mm:ss") & " ..."
Exit Sub
NoFareRange:
    msgCell = msgCell & Chr(10) & "Could not locate" & IntFareRangeString
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

Function InputTableofIntegration(FirstCell As Range, msgCell As Range) As Boolean
'To see specification of what this is supposed to read: check sheet "Sample_Data_for_Cones_tool"
Dim plan As Worksheet
Dim iRow As Integer, icol As Integer
Dim startRow As Integer, startCol As Integer
Dim i As Integer, j As Integer

Set plan = FirstCell.Worksheet
startRow = FirstCell.Row
startCol = FirstCell.Column

InputTableofIntegration = True

If FirstCell.Value <> "FARE_PRICES" Then
    msgCell = msgCell & Chr(10) & "Check if Integration Fare range is correct, missing check expression   'FARE_PRICES'"
    InputTableofIntegration = False
    Exit Function
End If

'Learn how many groups there is and their group code
iRow = startRow + 2
icol = startCol
NFareGroups = 0
ReDim FareGroupCode(0)
FareGroupCode(0) = "NFI"
Do While plan.Cells(iRow, icol) <> ""
    NFareGroups = NFareGroups + 1
    ReDim Preserve FareGroupCode(NFareGroups)
    FareGroupCode(NFareGroups) = plan.Cells(iRow, icol)
    iRow = iRow + 1
Loop


'Fill TableOfIntegration from 1 to 2 => Saves price do board in Group i when coming from j
'In the spreadsheet: row = boarding to route group (grupo onde embarco)
'In the variable TableofIntegration: row = route group coming from (grupo de onde parti)
startRow = startRow + 1
startCol = startCol + 1
For i = 0 To NFareGroups
    For j = 0 To NFareGroups
        If Not IsNumeric(plan.Cells(startRow + i, startCol + j)) Then
            InputTableofIntegration = False
            msgCell = msgCell & Chr(10) & "Not a number in " & plan.Name & "!" & plan.Cells(startRow + i, startCol + j).Address
        ElseIf plan.Cells(startRow + i, startCol + j) < 0 And plan.Cells(startRow + i, startCol + j) <> -1 Then
            msgCell = msgCell & Chr(10) & "Unexpected value in " & plan.Name & "!" & plan.Cells(startRow + i, startCol + j).Address
            InputTableofIntegration = False
        Else
            TableofIntegration(j, i) = plan.Cells(startRow + i, startCol + j)
        End If
    Next j
Next i
End Function
Function AssignRouteGroupsFromField(LineGroupListRange As String, msgCell As Range)
    Dim iroute As Integer
    Dim grouprange As Range
    AssignRouteGroupsFromField = True
    For iroute = 1 To nRoutes
        route(iroute).group = 0
        If LineGroupListRange = "ut1" Then route(iroute).group = route(iroute).t1
        If LineGroupListRange = "ut2" Then route(iroute).group = route(iroute).t2
        If LineGroupListRange = "ut3" Then route(iroute).group = route(iroute).t3
        If LineGroupListRange = "mode" Then route(iroute).group = route(iroute).mode
    Next iroute
    
    If Not (LineGroupListRange = "ut1" Or LineGroupListRange = "ut2" Or LineGroupListRange = "ut3" Or LineGroupListRange = "mode") Then
        On Error GoTo NoRange
        Set grouprange = String_to_Range(LineGroupListRange)
        On Error GoTo 0
        
        Dim plan As Worksheet
        Dim iRow As Integer, icol As Integer
        'To see specification of what this is supposed to read: check sheet "Sample_Data_for_Cones_tool"
        Set plan = grouprange.Worksheet
        iRow = grouprange.Row
        icol = grouprange.Column
        If grouprange.Value <> "Line" Then
            msgCell = msgCell & Chr(10) & "Check if Integration Fare range is correct, missing check expression   'Line'"
            AssignRouteGroupsFromField = False
            Exit Function
        End If
        iRow = iRow + 1
        Do While plan.Cells(iRow, icol) <> ""
            iroute = Get_Route(plan.Cells(iRow, icol))
            If iroute <> 0 Then
                route(iroute).group = plan.Cells(iRow, icol + 1)
            Else
                msgCell = msgCell & Chr(10) & "Warning: Line " & plan.Cells(iRow, icol) & " missing (" & plan.Name & "!" & plan.Cells(iRow, icol) & ")"
            End If
            iRow = iRow + 1
        Loop
    End If
    ReDim NRoutesInGroup(NFareGroups)
    For iroute = 1 To nRoutes
        route(iroute).igroup = GetFareGroup(route(iroute).group)
        NRoutesInGroup(route(iroute).igroup) = NRoutesInGroup(route(iroute).igroup) + 1
    Next iroute
    Exit Function
NoRange:
    msgCell = msgCell & Chr(10) & "Could not convert to range" & grouprange
    
End Function
Function GetFareGroup(groupCode As String) As Integer
    For GetFareGroup = 1 To NFareGroups
        If groupCode = FareGroupCode(GetFareGroup) Then Exit Function
    Next GetFareGroup
    GetFareGroup = 0
End Function
'If a node is stop of an integrated line, this routine will place Point().selected =true
Function NSelected_Nodes_to_Cone() As Long
'Run all integrated routes from begining to end, if para = "+" then point is selected
UnSelect_all_points
Dim i As Long, il As Long
Dim iroute As Integer
For iroute = 1 To nRoutes
    If route(iroute).igroup <> 0 Then
        For i = 1 To route(iroute).Npoints
            If route(iroute).para(i) = "+" Then
                point(route(iroute).ipoint(i)).selected = True
            End If
        Next i
    End If
Next iroute
For i = 1 To Npoints
    If point(i).selected Then NSelected_Nodes_to_Cone = NSelected_Nodes_to_Cone + 1
Next i
End Function
Function Get_First_Avaiable_Board_Name(ByVal ipoint As Long, ByVal grupo As Integer) As String
'This function tries names where both boarding and alighting names (+50,000) are free to match ending number
Dim conta As Integer, try As Integer, prefixo As Integer '
Dim Candi As String

Get_First_Avaiable_Board_Name = ""
conta = 0
Do While Get_First_Avaiable_Board_Name = "" And conta < 1000
    prefixo = grupo
    If grupo > 9 Then prefixo = prefixo - 10
    Candi = prefixo * 100000 + ipoint
    try = 0
    Do While (PointNamed(Candi) <> 0 Or PointNamed(Candi + 50000) <> 0) And try < 10
        prefixo = prefixo + 1
        If prefixo > 9 Then prefixo = prefixo - 9
        Candi = prefixo * 100000 + ipoint
        try = try + 1
    Loop
    If PointNamed(Candi) = 0 And PointNamed(Candi + 50000) = 0 Then Get_First_Avaiable_Board_Name = Candi
    ipoint = ipoint + 1
    conta = conta + 1
Loop
If Get_First_Avaiable_Board_Name = "" Then MsgBox "still cant get a name"
End Function
'Splits a point into a cone
Sub SplitPoint(ipoint As Long, Optional cone_height_in_meters As Single = 25)
Dim auxpoint As Point_type
Dim CandiName As String
Dim CAngle As Single 'for cone central angle
Dim RelX As Integer
Dim RelY As Integer
Dim jroute As Integer, jrpos As Integer

'Reset for safety
point(ipoint).HasCone = True
ReDim point(ipoint).SubPointAlight(NFareGroups)
ReDim point(ipoint).SubPointBoard(NFareGroups)
Dim j As Integer
For j = 1 To NFareGroups
    point(ipoint).SubPointAlight(j) = 0
    point(ipoint).SubPointBoard(j) = 0
Next j

For j = 1 To point(ipoint).nRoutes
    jroute = point(ipoint).iroute(j)
    jrpos = point(ipoint).iRoutePosition(j)
    If route(jroute).para(jrpos) = "+" Then
        point(ipoint).SubPointAlight(route(jroute).igroup) = 1 ' at this point 1 means group needs cone
    End If
Next j

RelY = 1
CAngle = GetAngletoSplitPoint(ipoint)

'Create Cones in growing order of fare groups
For j = 1 To NFareGroups
    If point(ipoint).SubPointAlight(j) = 1 Then 'it has
        CandiName = Get_First_Avaiable_Board_Name(point(ipoint).Name, j)
'        point(ipoint).Toexpand = True
        auxpoint = PointInsertPosition(point(ipoint), CAngle, -1, RelY, cone_height_in_meters * approx_one_meter_in_network)
        Do While Get_Point(auxpoint.x, auxpoint.y, True, 0, CandiName) < Npoints 'need not to be exactly uppon other point
            auxpoint.x = auxpoint.x - approx_one_meter_in_network
        Loop
        point(ipoint).SubPointBoard(j) = Npoints
        point(Npoints).t3 = -j
        point(Npoints).isM2 = 2
        CandiName = CandiName + 50000
        auxpoint = PointInsertPosition(point(ipoint), CAngle, 1, RelY, cone_height_in_meters * approx_one_meter_in_network)
        Do While Get_Point(auxpoint.x, auxpoint.y, True, 0, CandiName) < Npoints
            auxpoint.x = auxpoint.x + approx_one_meter_in_network
        Loop
        point(ipoint).SubPointAlight(j) = Npoints
        point(Npoints).t3 = -j - 50
        point(Npoints).isM2 = 2
        RelY = RelY + 1
    End If
Next j

End Sub
'Links cones points
Sub RelinkPoint(i As Long, EquivWalkDistance As Single, _
                penaltyCost As Single, walkmode As String, chargemode As String, _
                linkIntType As Integer, onlyOneTransfer As Boolean)
'Presumes all points that will have cones are split allready
    
    Dim Dlink As Link_type 'basic framework of cone links
    Dlink.Lanes = 2
    Dlink.nRoutes = 0
    ReDim Dlink.iroute(0)
    Dlink.t1 = 360 '<-- this can to be used as speed, so bus don't waste time... need to have consistent ttf
    Dlink.t2 = 0
    Dlink.t3 = 0
    Dlink.isM2 = 2
    Dlink.Extension = 0.01 'minimal extension that don't get truncated to zero... this may have changed
    ' this would be an extra extension of 30 meters (if that the unit) may represent extra time on station, keeping speed
     
    'Look for integration connections
    Dim j As Integer, k As Integer, kk As Integer
    Dim NodeI(100) As Long ' list of main nodes that have to be integrated with i
    Dim NNodesI As Integer
    NNodesI = 0
    NodeI(0) = i 'set that this point integrates to itself
    Dim potli As Long 'potential node to link
    Dim tipol As Integer, tipoii As Integer ' just to save link type and improve readability
    Dim already As Boolean
    For j = 1 To point(i).NLinksDaqui
        tipol = link(point(i).iLinkDaqui(j)).tipo
        If tipol = linkIntType Then
            'Found a node to link with
            potli = link(point(i).iLinkDaqui(j)).dp
            
            'check if destination node is not listed already
            already = False
            For k = 1 To NNodesI
                If NodeI(k) = potli Then
                    already = True
                    Exit For
                End If
            Next k
            
            If Not already Then
                'is not on the list, then add it
                NNodesI = NNodesI + 1
                NodeI(NNodesI) = potli
            End If
        End If
    Next j
    
    'Place links
    Dim NumP1 As Long 'to save boarding point name
    Dim NumP51 As Long 'to save alight point name
    Dim LinkOn As Long 'link number between cones
    Dim og As Integer, dg As Integer 'origin group, destiny group
    For og = 1 To NFareGroups
        If point(i).HasCone Then
            NumP1 = point(i).SubPointBoard(og)
            If NumP1 > 0 Then 'marked by split point, if exists has to be expanded
                'Set the walking in link (vehicle modes plus auxtransit chargemode for several transits, only transit if one transfer)
                Dlink.op = i
                Dlink.dp = NumP1
                Dlink.Extension = 0.01
                Dlink.modes = chargemode
                If onlyOneTransfer Then
                    Dlink.modes = "" 'let charge (and walk) mode out so there is no first boardings there
                    Dlink.t3 = 0
                End If
                Dlink.tipo = 800 + og 'uses the 800-900 range for types
                Dlink.t3 = TableofIntegration(0, og)
                Call Add_Link(Dlink)
                
                If onlyOneTransfer Then
                'Add exit link for second route
                    Dlink.op = NumP1
                    Dlink.dp = i
                    Dlink.modes = "" 'don´t need this line, but let us keep this clear
                    Dlink.tipo = 820 + og
                    Dlink.t3 = 0
                    Call Add_Link(Dlink)
                End If
                
                ' Link between points (only vehicle, this is done to save transit segments which is more limited than links),
                ' only needed if not only one transfer
                NumP51 = point(i).SubPointAlight(og)
                If NumP51 = 0 Then MsgBox "OOps: need debug!" 'if boarding was created, then the other should as well
                If Not onlyOneTransfer Then
                    Dlink.op = NumP1
                    Dlink.dp = NumP51
                    Dlink.tipo = 700 + og
                    Dlink.t3 = 0
                    Dlink.modes = ""
                    Call Add_Link(Dlink)
                End If
                
                'Exiting link (vehicle in all cases, pedestrian if unlimited transfer, charge firstboarding if only one)
                Dlink.op = NumP51
                Dlink.dp = i
                Dlink.tipo = 850 + og
                Dlink.modes = walkmode
                If onlyOneTransfer Then
                    Dlink.modes = chargemode
                    Dlink.t3 = TableofIntegration(0, og) 'charge full fare for exit
                End If
                Call Add_Link(Dlink)
                
                If onlyOneTransfer Then 'first line needs acess directly
                    Dlink.t3 = 0
                    Dlink.op = i
                    Dlink.dp = NumP51
                    Dlink.tipo = 870 + og
                    Dlink.modes = ""
                    Call Add_Link(Dlink)
                End If
                
                
                'Integration Links
                For k = 0 To NNodesI
                    'set extension
                    If i = NodeI(k) Then
                        Dlink.Extension = 0.01
                    Else
                        LinkOn = Get_Link(i, NodeI(k))
                        If LinkOn = 0 Then MsgBox "Debug: link special type on missin"
                        Dlink.Extension = link(LinkOn).Extension
                    End If
                    If Dlink.Extension < 0.01 Then Dlink.Extension = 0.01
                    
                    'Integration links are drawn from "right to left"
                    For dg = 1 To NFareGroups
                        If point(NodeI(k)).HasCone Then
                            NumP1 = point(NodeI(k)).SubPointBoard(dg)
                            If NumP1 > 0 Then ' if the point exist, then there is buses of destiny group dg here
                                If TableofIntegration(dg, og) <> -1 Then 'not all existing fare groups integrate with each other
                                    Dlink.op = NumP51
                                    Dlink.dp = NumP1
                                    Dlink.tipo = 900 + dg
                                    Dlink.modes = chargemode
                                    Dlink.t3 = TableofIntegration(dg, og)
                                    If onlyOneTransfer Then Dlink.t3 = Dlink.t3 + TableofIntegration(0, og) 'as is charge full fare for exit
                                    If EquivWalkDistance > 0 And Dlink.Extension > 0.01 Then
                                        Dlink.t3 = Dlink.t3 + Dlink.Extension / EquivWalkDistance
                                    End If
                                    Dlink.t3 = Dlink.t3 + penaltyCost
                                    If Get_Link(NumP51, NumP1) <> 0 Then MsgBox "OOps: need debug! Trying to create existing integration link"
                                    Call Add_Link(Dlink)
                                End If
                            End If 'NumP1 > 0
                        End If 'point(NodeI(k)).HasCone
                    Next dg
                Next k
            End If 'NumP1>0
        End If 'HasCone
    Next og
End Sub
Sub Reroute(iroute As Integer)
'Redraw routes that need to go inside integration cones
' and checks if links and nodes of the cones that are supposed to exist
' really do.
' It only changes the boarding/alighting set
' but it can easily change dwt and ttf if a consistent set up is made

If route(iroute).igroup = 0 Then
    Exit Sub
End If


Dim i As Long, ipoint As Long 'current route point observed
Dim P1 As Long, P2 As Long 'save the cone points to go into
Dim il As Long 'to check if link exist

'guarda means save, para means stop
Dim guarda(1000) As Long
Dim guardapara(1000) As String
Dim guardadwt(1000) As Single
Dim guardattf(1000) As Integer
Dim nguarda As Integer

' This routine follows the route along transfering nodes to guarda,
' when it finds a stop, adds the tour inside the cone to guarda
' it is supposed to be there after checking... and finally transfer
' guarda to the route

nguarda = 0
For ipoint = 1 To route(iroute).Npoints
    i = route(iroute).ipoint(ipoint)
    If route(iroute).para(ipoint) = "+" Then 'it is a stop
        P1 = point(i).SubPointBoard(route(iroute).igroup)
        P2 = point(i).SubPointAlight(route(iroute).igroup)
        If P1 = 0 Or P2 = 0 Then MsgBox " Missing P1 ou P2!"
        
        'The main point (on the street), is maintained, but stop forbidden
        If ipoint > 1 Then 'First route point does not enters the cone, starts inside
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = "#"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If
        
        'Boarding point =P1
        nguarda = nguarda + 1
        guarda(nguarda) = P1
        guardapara(nguarda) = "<"
        guardadwt(nguarda) = route(iroute).dwt(ipoint)
        guardattf(nguarda) = route(iroute).ttf(ipoint)
        
        'allighting point =P2
        nguarda = nguarda + 1
        guarda(nguarda) = P2
        guardapara(nguarda) = ">"
        guardadwt(nguarda) = route(iroute).dwt(ipoint)
        guardattf(nguarda) = route(iroute).ttf(ipoint)
        
        'Return to the street (only if it is not the last point
        If ipoint < route(iroute).Npoints Then
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = "#"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If

    Else ' not a stop, keep it that way
        nguarda = nguarda + 1
        guarda(nguarda) = i
        guardapara(nguarda) = "#"
        guardadwt(nguarda) = route(iroute).dwt(ipoint)
        guardattf(nguarda) = route(iroute).ttf(ipoint)
    End If
Next ipoint

'Reroute route to what was saved in guarda
route(iroute).Npoints = nguarda
ReDim route(iroute).ipoint(route(iroute).Npoints)
ReDim route(iroute).para(route(iroute).Npoints)
ReDim route(iroute).dwt(route(iroute).Npoints)
ReDim route(iroute).ttf(route(iroute).Npoints)

For ipoint = 1 To route(iroute).Npoints
        route(iroute).ipoint(ipoint) = guarda(ipoint)
        route(iroute).para(ipoint) = guardapara(ipoint)
        route(iroute).dwt(ipoint) = guardadwt(ipoint)
        route(iroute).ttf(ipoint) = guardattf(ipoint)
        If ipoint > 1 Then ' Check if the link exist and add mode if needed (which is in the cones)
            il = Get_Link(route(iroute).ipoint(ipoint - 1), route(iroute).ipoint(ipoint))
            If il = 0 Then MsgBox "Debug: Route into missing link! when rerouting"
            If Not LinkhasMode(il, route(iroute).mode) Then
                link(il).modes = link(il).modes & route(iroute).mode
            End If
        End If
Next ipoint
route(iroute).changed = True
End Sub
Sub RerouteONETRANSFER(ByVal iroute As Integer)
' Duplicates the given iroute and change both to go inside integration cones
' (checks if links and nodes of the cones that are supposed to exist really do.)
' It only changes the boarding/alighting set
' but it can easily change dwt and ttf if a consistent set up is made


If route(iroute).igroup = 0 Then
    Exit Sub
End If

Dim i As Long 'current route point observed
Dim P1 As Long, P2 As Long 'save the cone points to go into
Dim il As Long 'to check if link exists

'guarda means save, para means stop
Dim guarda(1000) As Long
Dim guardapara(1000) As String
Dim guardadwt(1000) As Single
Dim guardattf(1000) As Integer
Dim nguarda As Integer


'Copy transit line
Dim NP As Integer
NP = route(iroute).Npoints
nRoutes = nRoutes + 1
ReDim Preserve route(nRoutes)
ReDim route(nRoutes).ipoint(NP)
ReDim route(nRoutes).dwt(NP)
ReDim route(nRoutes).ttf(NP)
ReDim route(nRoutes).para(NP)
route(nRoutes) = route(iroute) ' new route = iroute

'Change number and description of second route
Dim tryname As String
Dim lettercode As Integer
lettercode = Asc("X") - 1

Do
lettercode = lettercode + 1
If lettercode > Asc("Z") Then lettercode = Asc("A")
If Len(route(nRoutes).number) < 6 Then
    tryname = Chr(lettercode) & route(iroute).number
Else
    tryname = Chr(lettercode) & Right(route(iroute).number, 5)
End If
If lettercode = Asc("X") - 1 Then tryname = InputBox("Could not find a name to double line " & route(iroute).Name & "Please enter one bellow", "UNLIKELY BREAK")
Loop While Get_Route(tryname) <> 0

route(nRoutes).number = tryname
route(nRoutes).Name = "2_" & route(iroute).number & "_" & route(iroute).Name

Dim ipoint As Long
nguarda = 0
For ipoint = 1 To route(iroute).Npoints
    i = route(iroute).ipoint(ipoint)
    If route(iroute).para(ipoint) = "+" Then 'its stop
        P1 = point(i).SubPointBoard(route(iroute).igroup)
        P2 = point(i).SubPointAlight(route(iroute).igroup)
        If P1 = 0 Or P2 = 0 Then MsgBox " Sem P1 ou P2!"
        ' Keep the same point and goes to the right leg for alight
        If ipoint > 1 Then 'unless it is the first
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = "#"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
            
            nguarda = nguarda + 1
            guarda(nguarda) = P2
            guardapara(nguarda) = ">"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If
        'Back to the street, only board allowed
        If ipoint < route(iroute).Npoints Then
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = "<"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If
    Else 'it is not a stop
        nguarda = nguarda + 1
        guarda(nguarda) = i
        guardapara(nguarda) = "#"
        guardadwt(nguarda) = route(iroute).dwt(ipoint)
        guardattf(nguarda) = route(iroute).ttf(ipoint)
    End If
Next ipoint

'Reroute route to what was saved in guarda
route(iroute).Npoints = nguarda
ReDim route(iroute).ipoint(route(iroute).Npoints)
ReDim route(iroute).para(route(iroute).Npoints)
ReDim route(iroute).dwt(route(iroute).Npoints)
ReDim route(iroute).ttf(route(iroute).Npoints)

For ipoint = 1 To route(iroute).Npoints
        route(iroute).ipoint(ipoint) = guarda(ipoint)
        route(iroute).para(ipoint) = guardapara(ipoint)
        route(iroute).dwt(ipoint) = guardadwt(ipoint)
        route(iroute).ttf(ipoint) = guardattf(ipoint)
        If ipoint > 1 Then ' Check if the link exist and add mode if needed (which is in the cones)
            il = Get_Link(route(iroute).ipoint(ipoint - 1), route(iroute).ipoint(ipoint))
            If il = 0 Then MsgBox "Debug: Route into missing link! when rerouting"
            If Not LinkhasMode(il, route(iroute).mode) Then
                link(il).modes = link(il).modes & route(iroute).mode
            End If
        End If
Next ipoint
route(iroute).changed = True

'second route
iroute = nRoutes
nguarda = 0
route(iroute).t1 = 0

For ipoint = 1 To route(iroute).Npoints
    i = route(iroute).ipoint(ipoint)
    If route(iroute).para(ipoint) = "+" Then ' stop
        P1 = point(i).SubPointBoard(route(iroute).igroup)
        P2 = point(i).SubPointAlight(route(iroute).igroup)
        If P1 = 0 Or P2 = 0 Then MsgBox " Sem P1 ou P2!"
        'keep street point, unless it is the first, only alight here
        If ipoint > 1 Then
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = ">"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If
        ' into 'left leg' and back
        If ipoint < route(iroute).Npoints Then
            ' this it is the only place you can board this route
            nguarda = nguarda + 1
            guarda(nguarda) = P1
            guardapara(nguarda) = "<"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
            ' back to the street
            nguarda = nguarda + 1
            guarda(nguarda) = i
            guardapara(nguarda) = "#"
            guardadwt(nguarda) = route(iroute).dwt(ipoint)
            guardattf(nguarda) = route(iroute).ttf(ipoint)
        End If
    Else 'not a stop
        nguarda = nguarda + 1
        guarda(nguarda) = i
        guardapara(nguarda) = "#"
        guardadwt(nguarda) = route(iroute).dwt(ipoint)
        guardattf(nguarda) = route(iroute).ttf(ipoint)
    End If
Next ipoint

'Reroute second route to what was saved in guarda
route(iroute).Npoints = nguarda
ReDim route(iroute).ipoint(route(iroute).Npoints)
ReDim route(iroute).para(route(iroute).Npoints)
ReDim route(iroute).dwt(route(iroute).Npoints)
ReDim route(iroute).ttf(route(iroute).Npoints)

For ipoint = 1 To route(iroute).Npoints
        route(iroute).ipoint(ipoint) = guarda(ipoint)
        route(iroute).para(ipoint) = guardapara(ipoint)
        route(iroute).dwt(ipoint) = guardadwt(ipoint)
        route(iroute).ttf(ipoint) = guardattf(ipoint)
        If ipoint > 1 Then ' Check if the link exist and add mode if needed (which is in the cones)
            il = Get_Link(route(iroute).ipoint(ipoint - 1), route(iroute).ipoint(ipoint))
            If il = 0 Then MsgBox "Debug: Route into missing link! when rerouting"
            If Not LinkhasMode(il, route(iroute).mode) Then
                link(il).modes = link(il).modes & route(iroute).mode
            End If
        End If
Next ipoint

End Sub
Sub FillAtende(MaxDistanceDiference As Integer)
Dim arethey As Boolean
Dim netdist As Single
Dim bestdist As Single
Dim iroute As Integer
    For izone = FirstZona To LastZona
        If PointNamed(izone) = 0 Then
            For iroute = 1 To nRoutes
                Atende(izone, iroute) = 0
            Next iroute
        Else
            For i = 1 To point(PointNamed(izone)).NLinksDaqui
                link(point(PointNamed(izone)).iLinkDaqui(i)).selected = True
            Next i
            N = NLinksInShortestPathWithinSelectedLinks(PointNamed(izone), 0, arethey, bestdist, True, 3000, True)
            If arethey Then MsgBox "Uptobusstoproblem!"
            N = NLinksInShortestPathWithinSelectedLinks(PointNamed(izone), 0, arethey, netdist, True, 3000)
            If arethey Then MsgBox "Uptodist problem!"
            For i = 1 To point(PointNamed(izone)).NLinksDaqui
                link(point(PointNamed(izone)).iLinkDaqui(i)).selected = False
            Next i
            
            For i = 1 To nRoutes
                If route(i).mode = "y" Then
                    For j = 2 To route(i).Npoints
                        gl = Get_Link(route(i).ipoint(j - 1), route(i).ipoint(j))
                        If gl <> 0 And link(gl).modes <> "y" Then
                            route(i).mode = "yo"
                            Exit For
                        End If
                    Next j
                End If
            Next i
            
            For iroute = 1 To nRoutes
                point(0).Distance = bestdist + MaxDistanceDiference
                Atende(izone, iroute) = 0
                If route(iroute).mode <> "yo" Then
                    For i = 1 To route(iroute).Npoints
                        ipoint = route(iroute).ipoint(i)
                        If point(ipoint).marker = marker _
                        And point(ipoint).Distance < point(Atende(izone, iroute)).Distance _
                        And route(iroute).para(i) <> "#" Then
                            Atende(izone, iroute) = ipoint
                        End If
                    Next i
                End If
            Next iroute
            For iroute = 1 To nRoutes
                If route(iroute).mode = "yo" Then route(iroute).mode = "y"
            Next iroute
            
'            If Int(izone / 50) = izone / 50 Then Debug.Print izone & " " & Format(Now, "hh:mm:ss")
        End If
    Next izone
End Sub
Sub MarkPointstoExpand()
'Routine that marks all points that need to be expanded
Dim Jávi(3000) As Boolean

For i = 1 To Npoints
    point(i).Toexpand = False
Next i
For i = 1 To Npoints
    'First: all centroids exits are to be expanded
    If point(i).isM2 = 1 Then
        For j = 1 To point(i).NLinksDaqui
            point(link(point(i).iLinkDaqui(j)).dp).Toexpand = True
        Next j
    End If
    'Every Point that has a link type=777 is to be expanded
    For j = 1 To point(i).NLinksDaqui
        tipol = link(point(i).iLinkDaqui(j)).tipo
        If tipol = 777 Then
            point(link(point(i).iLinkDaqui(j)).dp).Toexpand = True
        End If
    Next j
    'Every metro station is to be expanded
    If point(i).Isstation <> "" Then
        point(i).Toexpand = True
    End If
Next i
'Marca para expansão todos os pontos que são 1os encontros de linhas
For i = 1 To nRoutes
    'Start seing no routes
    If route(i).t3 > 0 Then
        For j = 1 To nRoutes
             Jávi(j) = False
             If route(j).t3 = 0 Then Jávi(j) = True
        Next j
        For j = 1 To route(i).Npoints
            For k = 1 To point(route(i).ipoint(j)).nRoutes
                If Not Jávi(point(route(i).ipoint(j)).iroute(k)) And TableofIntegration(route(i).t3, route(point(route(i).ipoint(j)).iroute(k)).t3) <> -1 Then
                    Jávi(point(route(i).ipoint(j)).iroute(k)) = True
                    point(route(i).ipoint(j)).Toexpand = True
                End If
            Next k
        Next j
    End If
Next i
End Sub
Sub RouteCaminhoMínimoWithinSelectedLinks(iroute As Integer)
Dim PANT As Long
Dim P As Long
Dim guarda(2000) As Long
Dim guardapara(2000) As String
Dim Arelinked As Boolean
Dim NetDis As Single
    nguarda = 1
    guarda(nguarda) = route(iroute).ipoint(1)
    guardapara(nguarda) = route(iroute).para(1)
    For ipon = 1 To route(iroute).Npoints - 1
        PANT = route(iroute).ipoint(ipon)
        P = route(iroute).ipoint(ipon + 1)
'        Ppara = RoUtE(iroute).para(ipon + 1)
        If point(P).Name = "28812" Then
            A = B
        End If
        il = Get_Link(PANT, P)
        If il = 0 Then
            N = NLinksInShortestPathWithinSelectedLinks(PANT, P, Arelinked, NetDis, True)
            If Not Arelinked Then
                Arelinked = MsgBox("Route " & route(iroute).number & " can´t go from " & point(PANT).Name & " to " & point(P).Name & " Tentar ir ao ponto mais próximo? Não ignora o ponto de destino.", vbYesNo, "Route Caminho Mínimo within Selected Links") = vbYes
'                If ipon = Route(iroute).Npoints - 1 Then Arelinked = True: Debug.Print Route(iroute).number & " can´t go from " & point(PANT).NAME & " to " & point(P).NAME
            End If
            If Arelinked Then
                For i = 1 To N
                    nguarda = nguarda + 1
                    guarda(nguarda) = link(LinkList(i)).dp
                    guardapara(nguarda) = Ppara
                Next i
            End If
        Else
            nguarda = nguarda + 1
            guarda(nguarda) = P
            guardapara(nguarda) = Ppara
        End If
    Next ipon
    route(iroute).Npoints = nguarda
    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
    ReDim Preserve route(iroute).para(route(iroute).Npoints)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
        route(iroute).para(i) = guardapara(i)
    Next i
End Sub
Sub LoadVolauVolax(fullfile As String)
Dim stringline As String
    If Dir(fullfile) <> "" Then
        Open fullfile For Input As #1
            Do While Not EOF(1)
                Line Input #1, stringline
                If CountWords(stringline) = 5 Then
                    If IsNumeric(Splitword(1)) And IsNumeric(Splitword(2)) Then
                        il = Get_Link(PointNamed(Splitword(1)), PointNamed(Splitword(2)))
                        If il <> 0 Then
                            link(il).volau = Val(Splitword(3))
                            link(il).volax = Val(Splitword(4))
                        End If
                    End If
                End If
            Loop
        Close #1
    End If
End Sub

Sub Update_Encounters()
Dim jroute As Integer
Dim RE As Long
Debug.Print "Start Update Encounters " & Format(Now, "hh:mm:ss")
For i = 1 To nRoutes
    For j = 1 To nRoutes
        RoutesEncounter(i, j) = 0
    Next j
Next i

For iroute = 1 To nRoutes
    For ipoint = 1 To route(iroute).Npoints
        If route(iroute).para(ipoint) = "+" Then
            i = route(iroute).ipoint(ipoint)
            For kk = 0 To point(i).NLinksDaqui
                i2 = 0
                If kk = 0 Then
                    i2 = i
                ElseIf link(point(i).iLinkDaqui(kk)).tipo = 777 Then
                    i2 = link(point(i).iLinkDaqui(kk)).dp
                End If
                If i2 <> 0 Then
                    For k = 1 To point(i2).nRoutes
                        jroute = point(i2).iroute(k)
                        jrposi = point(i2).iRoutePosition(k)
                        If route(jroute).para(jrposi) = "+" Then
                            If RoutesEncounter(iroute, jroute) = 0 Then
                                'Como seguimos a linha iroute em ordem direta, caso exista o encontro...
                                ' já se sabe que foi anterior
                                RoutesEncounter(iroute, jroute) = i2
                            End If
                            'Já no sentido contrário (da linha jroute para linha iroute)
                            'verifica-se, caso exista, que o encontro registrado agora é mais cedo na linha jroute
                            If RoutesEncounter(jroute, iroute) = 0 Then
                                RoutesEncounter(jroute, iroute) = i2
                            ElseIf iroute <> jroute Then
                                RE = RoutesEncounter(jroute, iroute)
                                If jrposi < point(RE).iRoutePosition(DoesRoutepasshere(jroute, RE)) Then
                                    RoutesEncounter(jroute, iroute) = i2
                                End If
                            End If
                        End If
                    Next k
                End If
            Next kk
        End If
    Next ipoint
Next iroute
Debug.Print "End Update Encounters " & Format(Now, "hh:mm:ss")
End Sub
Sub Fill_First_Encounter()
Dim iroute As Integer
Dim jroute As Integer
Dim kroute As Integer
Dim FirstE As Long
Dim i As Long
    CountRoutesinNetwork
    For i = 1 To nRoutes
        For j = 1 To nRoutes
            RoutesFirstEncounter(i, j) = -1
        Next j
    Next i
    
    For i = 1 To Npoints
        If Not point(i).Deleta Then
        For j = 1 To point(i).nRoutes
            jroute = point(i).iroute(j)
            If route(jroute).t3 <> 0 Then
                For kk = 0 To point(i).NLinksDaqui
                    i2 = 0
                    If kk = 0 Then
                        i2 = i
                    ElseIf link(point(i).iLinkDaqui(kk)).tipo = 777 Then
                        i2 = link(point(i).iLinkDaqui(kk)).dp
                    End If
                    If i2 <> 0 Then
                        For k = j + 1 To point(i2).nRoutes
                            kroute = point(i2).iroute(k)
                            If TableofIntegration(route(jroute).t3, route(kroute).t3) <> -1 And route(kroute).t3 <> 0 Then
                                FirstE = RoutesFirstEncounter(jroute, kroute)
                                If FirstE = -1 Then
                                    RoutesFirstEncounter(jroute, kroute) = i2
                                Else
                                    If point(i2).iRoutePosition(j) < point(FirstE).iRoutePosition(DoesRoutepasshere(jroute, FirstE)) Then
                                        RoutesFirstEncounter(jroute, kroute) = i2
                                    End If
                                End If
                                'doutro lado
                                FirstE = RoutesFirstEncounter(kroute, jroute)
                                If FirstE = -1 Then
                                    RoutesFirstEncounter(kroute, jroute) = i2
                                Else
                                    If point(i2).iRoutePosition(k) < point(FirstE).iRoutePosition(DoesRoutepasshere(kroute, FirstE)) Then
                                        RoutesFirstEncounter(jroute, kroute) = i2
                                    End If
                                End If
                            End If
                        Next k
                    End If
                Next kk
            End If
        Next j
        End If
'        If Int(i / 500) = i / 500 Then Debug.Print i & " " & Format(Now, "hh:mm:ss")

    Next i
    For i = 1 To Npoints
        point(i).NIntegra = 0
    Next i
    
    ttfe = 0
    ttpp = 0
    For iroute = 1 To nRoutes
        ttpp = ttpp + route(iroute).Npoints
        For jroute = 1 To nRoutes
            FirstE = RoutesFirstEncounter(iroute, jroute)
            If FirstE <> -1 Then
                ttfe = ttfe + 1
                point(FirstE).NIntegra = point(FirstE).NIntegra + 1
            End If
        Next jroute
    Next iroute
    Debug.Print "TotalSegs=" & ttpp + 3 * ttfe & "(" & ttpp & "+ 3 * " & ttfe & ")"
    


    
End Sub
Sub Revê_Acessos_Estações()
Dim CandiLink As Link_type
Dim SelPoint As Long
Dim i As Long
CandiLink.isM2 = 2
CandiLink.modes = "p"
CandiLink.tipo = 777
CandiLink.Lanes = 0
CandiLink.vdf = 0


' Elimina todos os conectores de estações que não são da linha, mode correto em si
For i = 1 To Npoints
    If point(i).Isstation = "m" Or point(i).Isstation = "t" Or point(i).Isstation = "y" Then
        'o ponto é estação.
        'todas as linhas não expressas (modo da estação) param aqui
        Ntodel = 0
        For j = 1 To point(i).NLinksDaqui
            il = point(i).iLinkDaqui(j)
            If link(il).modes <> point(i).Isstation Then
                link(il).isM2 = 0
                Ntodel = Ntodel + 1
                LinkList(Ntodel) = il
            End If
        Next j
        For j = 1 To point(i).NLinksPraKa
            il = point(i).iLinkPraKa(j)
            If link(il).modes <> point(i).Isstation Then
                link(il).isM2 = 0
                Ntodel = Ntodel + 1
                LinkList(Ntodel) = il
            End If
        Next j
        For j = 1 To Ntodel
            Delete_Link (LinkList(j))
        Next j
        
        point(i).NIntegra = 0 'número total de linhas fora do sistema (trilhos) que se integram a estação, via link type 777
        'Começa sem integrações
        For j = 1 To nRoutes
            route(j).checked = False    'interessa atender
        Next j
        Dim DistToTryIntegation As Single
        DistToTryIntegation = 350
        Do While DistToTryIntegation < 1500 And point(i).NIntegra = 0
            If point(i).STname = "QUEI" Then
                A = B
            End If
            NtoSort = NAutoPointsListedInARadius(point(i), DistToTryIntegation, True)
            PointHeapSortbyDistance i, True
            N = NtoSort
            For j = 1 To N
                If PointList(j) <> i Then
                    point(PointList(j)).NIntegra = 0
                    For k = 1 To point(PointList(j)).nRoutes
                        iroute = point(PointList(j)).iroute(k)
                        If Not route(iroute).checked Then
                            point(PointList(j)).NIntegra = point(PointList(j)).NIntegra + 1
                        End If
                    Next k
                    'if this point brings more then 20% new lines for integration with the station, the link 777 must be created
                    If point(PointList(j)).NIntegra > 0.2 * point(i).NIntegra Then
                        SelPoint = PointList(j)
                        CandiLink.op = i
                        CandiLink.dp = SelPoint
                        CandiLink.Extension = Distância(i, SelPoint, True) * meter_per_extension_unit
                        If Get_Link(CandiLink.op, CandiLink.dp) = 0 Then
                            Add_Link CandiLink ': Debug.Print point(CandiLink.op).NAME & "-->" & point(CandiLink.dp).NAME
                        Else
                            MsgBox "Trying to create access to station twice"
                        End If
                        CandiLink.dp = i
                        CandiLink.op = SelPoint
                        If Get_Link(CandiLink.op, CandiLink.dp) = 0 Then
                            Add_Link CandiLink
                        Else
                            MsgBox "Trying to create access to station twice"
                        End If
                        For k = 1 To point(SelPoint).nRoutes
                            iroute = point(SelPoint).iroute(k)
                            route(iroute).checked = True
                            route(iroute).para(point(SelPoint).iRoutePosition(k)) = "+"
                            'Aqui deve encher route_first_encounter
                            'que precisa ter sido zerada antes dessa rotina e não zerada depois,
                            'no que se refere as integrações do grupo (m,t,y com b)
'                            For l = 1 To point(i).NRoutes
'                                jroute = point(i).iroute(l)
'                                RoutesFirstEncounter(iroute, jroute) = SelPoint
'                                RoutesFirstEncounter(jroute, iroute) = SelPoint
'                            Next l
                        Next k
                        point(i).NIntegra = point(i).NIntegra + point(PointList(j)).NIntegra
                    Else
                        point(PointList(j)).NIntegra = 0
                    End If
                End If
            Next j
            DistToTryIntegation = DistToTryIntegation + 500
        Loop
    End If
    
Next i
    
    ' Adiciona links type 777 : PAVUNA 1021 1184
    '                         : TRIAGEM 1059 1214  <-- !!! ERA PARA SER 1029
    '                         : SÃO CRISTÓVÃO 1025 1203
    '                         : CENTRAL 1005 1119
    '                         : MADUREIRA 27756 1156
    '                         : MAGN 1159 27757
    '                         : PENHA 27771 1185
    '                         : CAR 27763 1031
    '                         : DEOD 28118 1128
    '                         : BAND 28103 27741
    '                         : CURUC 28105 27738
    '                         : ALVORADA 28010 27733
    '                         : ALENDE 28017 28109
    '                         : JD OCEANICO: 1228 28002

        CandiLink.op = PointNamed(1021): CandiLink.dp = PointNamed(1184): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp) * meter_per_extension_unit: If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(1059): CandiLink.dp = PointNamed(1214): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(1025): CandiLink.dp = PointNamed(1203): CandiLink.Extension = 220: If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(1005): CandiLink.dp = PointNamed(1119): CandiLink.Extension = 60: If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(27756): CandiLink.dp = PointNamed(1156): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(1159): CandiLink.dp = PointNamed(27757): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(27771): CandiLink.dp = PointNamed(1185): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(27763): CandiLink.dp = PointNamed(1031): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28118): CandiLink.dp = PointNamed(1128): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28103): CandiLink.dp = PointNamed(27741): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28105): CandiLink.dp = PointNamed(27738): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28010): CandiLink.dp = PointNamed(27733): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28017): CandiLink.dp = PointNamed(28109): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(1228): CandiLink.dp = PointNamed(28002): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(1021): CandiLink.op = PointNamed(1184): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(1059): CandiLink.op = PointNamed(1214): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(1025): CandiLink.op = PointNamed(1203): CandiLink.Extension = 220: If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(1005): CandiLink.op = PointNamed(1119): CandiLink.Extension = 60: If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(27756): CandiLink.op = PointNamed(1156): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(1159): CandiLink.op = PointNamed(27757): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(27771): CandiLink.op = PointNamed(1185): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(27763): CandiLink.op = PointNamed(1031): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(28118): CandiLink.op = PointNamed(1128): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(28103): CandiLink.op = PointNamed(27741): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(28105): CandiLink.op = PointNamed(27738): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(28010): CandiLink.op = PointNamed(27733): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.dp = PointNamed(28017): CandiLink.op = PointNamed(28109): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink
        CandiLink.op = PointNamed(28002): CandiLink.dp = PointNamed(1228): CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True): If CandiLink.op * CandiLink.dp <> 0 Then Add_Link CandiLink


End Sub
Sub Revê_Connectors()
' REVISÃO DE CONECTORES DE ZONA (VERIFICA IDA e REPETE VOLTA)
    'conectores que são fora da rede de autos -> passam para a rede de autos mais próxima
    ' = deleta e cria novo
    'inclui o conector mais curto
    'apartir do mais curto passa para os seguintes:
        ' ELIMINA OS CONECTORES EM QUE A DISTÂNCIA ATÉ O MESMO PONTO PELA REDE (e conectore já colocados):
        '   - for até 200 metros maior que pelo link (não preciso de conectores para economizar 200 metros)
        '   - OU for até 40% maior que a distância pelo link (nem para economizar 30% da distância pela rede)
Dim il As Long
Dim il2 As Long
Dim CandiLink As Link_type
Dim arethey As Boolean
Dim netdist As Single
For il = 1 To NLinks
    link(il).selected = LinkhasMode(il, "a")
    If link(il).isM2 = 1 Then link(il).selected = False
Next il
For i = 1 To Npoints
    If point(i).isM2 = 1 Then
        If point(i).Name = "60" Then
            A = B
        End If
        For j = 1 To point(i).NLinksDaqui
            il = point(i).iLinkDaqui(j)
            il2 = Get_Link(link(il).dp, link(il).op)
            If il2 = 0 Then MsgBox "Missing Return Connector"
            If link(il2).Extension <> link(il).Extension Then MsgBox "Different return connector length"
'            If point(Link(il).dp).Isstation <> "" Then
'                MsgBox "Link Direto: da zona " & point(i).NAME & " para a estação " & point(Link(il).dp).STname
            If point(link(il).dp).isM2 = 1 Then
                link(il).isM2 = 0
                If il2 <> 0 Then link(il2).isM2 = 0
            ElseIf Not point(link(il).dp).IsAutoNetwork Then
                NP = Get_Closer_Auto_Point(point(link(il).dp), 1000, 1000, True)
                CandiLink = link(il)
                CandiLink.modes = "ap"
                CandiLink.selected = False
                CandiLink.dp = NP
                If Get_Link(CandiLink.op, CandiLink.dp) = 0 And NP <> 0 Then Add_Link CandiLink
                link(il).isM2 = 0
'                Link(il).Selected=
'                   Delete_Link (il)
                If il2 <> 0 Then
                    CandiLink = link(il2)
                    CandiLink.modes = "ap"
                    CandiLink.op = NP
                    CandiLink.selected = False
                    If Get_Link(CandiLink.op, CandiLink.dp) = 0 And NP <> 0 Then Add_Link CandiLink
                    link(il2).isM2 = 0
'                Link(il).Selected
'                   Delete_Link (il2)
                End If
            End If
        Next j
        shortestlink = 1
        Do While shortestlink <> 0
            'Procura o shortest entre os não selecionados e não deletados (IsM2=0)
            shortestlink = 0
            link(0).Extension = 100000
            For j = 1 To point(i).NLinksDaqui
                il = point(i).iLinkDaqui(j)
                If Not link(il).selected And link(il).isM2 <> 0 Then
                    If link(il).Extension < link(shortestlink).Extension Then shortestlink = il
                End If
            Next j
            link(0).Extension = 0
            If shortestlink <> 0 Then
                N = NLinksInShortestPathWithinSelectedLinks(i, link(shortestlink).dp, arethey, netdist, True)
                If Not arethey Then
                    link(shortestlink).selected = True
                ElseIf netdist - link(shortestlink).Extension > 200 And netdist > 2 * link(shortestlink).Extension Then
                    link(shortestlink).selected = True
                Else
                    il2 = Get_Link(link(shortestlink).dp, link(shortestlink).op)
                    link(shortestlink).isM2 = 0
                    If il2 <> 0 Then link(il2).isM2 = 0
                End If
            End If
        Loop
        For j = 1 To point(i).NLinksDaqui
            il = point(i).iLinkDaqui(j)
            link(il).selected = False
        Next j
        For j = 1 To point(i).NLinksPraKa
            il = point(i).iLinkPraKa(j)
            link(il).selected = False
        Next j
'        If Int(i / 100) * 100 = i Then Debug.Print i & Format(Now, " hh:mm:ss")
    End If
Next i

End Sub
Sub Revê_Connectors_from_scratch()
' REVISÃO DE CONECTORES DE ZONA (VERIFICA IDA e REPETE VOLTA)
    'conectores que são fora da rede de autos -> passam para a rede de autos mais próxima
    ' = deleta e cria novo
    'inclui o conector mais curto
    'apartir do mais curto passa para os seguintes:
        ' ELIMINA OS CONECTORES EM QUE A DISTÂNCIA ATÉ O MESMO PONTO PELA REDE (e conectore já colocados):
        '   - for até 200 metros maior que pelo link (não preciso de conectores para economizar 200 metros)
        '   - OU for até 40% maior que a distância pelo link (nem para economizar 30% da distância pela rede)
Dim i As Long
Dim il As Long
Dim CandiLink As Link_type
Dim arethey As Boolean
Dim netdist As Single
Dim Candi As Long
For i = 1 To Npoints
    If point(i).isM2 = 1 Then
        point(i).radius = 0
        For j = 1 To point(i).NLinksDaqui
            If link(point(i).iLinkDaqui(j)).Extension > point(i).radius Then point(i).radius = link(point(i).iLinkDaqui(j)).Extension
        Next j
    End If
Next i
For il = 1 To NLinks
    link(il).selected = LinkhasMode(il, "a")
    If link(il).isM2 = 1 Then
        link(il).isM2 = 0
        Delete_Link (il)
    End If
Next il
CandiLink.isM2 = 1
CandiLink.selected = True
CandiLink.tipo = 99
CandiLink.Lanes = 1
CandiLink.vdf = 99
CandiLink.modes = "ap"
CandiLink.t2 = 99999
CandiLink.t3 = 10

For i = 1 To Npoints
    If point(i).isM2 = 1 Then
        If point(i).Name = "42" Then
            A = B
        End If
        N = NAutoPointsListedInARadius(point(i), point(i).radius, True)
        NtoSort = N
        PointHeapSortbyDistance i, True
        For j = 1 To N
            Candi = PointList(j)
            candiminVDF = point(Candi).MinVDF
            candidist = Distância(i, Candi, True)
            ' Testa se deve existir um conector entre Point(i)(= a zona) e candi
            Nwhatever = NLinksInShortestPathWithinSelectedLinks(i, Candi, arethey, netdist, True)
            incluilink = False
            If Not arethey Then
                incluilink = True
            ElseIf netdist - candidist > 150 And ((candiminVDF > 3 And candidist < 0.75 * netdist) Or (candidist < 0.3 * netdist)) Then
                incluilink = True
            End If
            If incluilink Then
                CandiLink.op = i
                CandiLink.dp = Candi
                CandiLink.Extension = candidist
                NL = Add_Link(CandiLink)
                A = B
            End If
            ' Testa se deve existir um conector entre candi e Point(i),(= a zona)
            Nwhatever = NLinksInShortestPathWithinSelectedLinks(Candi, i, arethey, netdist, True)
            incluilink = False
            If Not arethey Then
                incluilink = True
            ElseIf netdist - candidist > 150 And candidist < 0.75 * netdist Then
                incluilink = True
            End If
            If incluilink Then
                CandiLink.op = Candi
                CandiLink.dp = i
                CandiLink.Extension = candidist
                NL = Add_Link(CandiLink)
            End If
        Next j
        For j = 1 To point(i).NLinksDaqui
            il = point(i).iLinkDaqui(j)
            link(il).selected = False
        Next j
        For j = 1 To point(i).NLinksPraKa
            il = point(i).iLinkPraKa(j)
            link(il).selected = False
        Next j
        If Int(i / 100) * 100 = i Then Debug.Print i & Format(Now, " hh:mm:ss")
    End If
Next i
End Sub
Sub Elimina_Nós() ' E MARCA UI3
'Elimina nós que podem ser eliminados, muito próximos e só com dois vizinhos (tanto auto quanto bus)
'reforma os links para esses nós
'passa eventuais conectores e links auxiliares para o nó vizinho
'tira nós das linhas
Dim i As Long
Dim ilink As Long
Dim ipoint As Long
Dim CandiP As Point_type
Dim CandiLink As Link_type

For i = 1 To NLinks
    If link(i).t3 = 0 Then
        link(i).t3 = 27
    End If
Next i
elimseg = 0
conta = 0
For i = 1 To Npoints
    If point(i).isM2 = 1 Then
        point(i).t3 = 1
    Else
        point(i).t3 = 0
        If Len(point(i).Isstation) > 1 Then
            MsgBox point(i).Name & "  " & point(i).Isstation
        End If
        If Val(point(i).Name) > 30000 Then
            For k = 1 To point(i).nRoutes
                Remove_Point_from_Route i, point(i).iroute(k)
            Next k
            For k = 1 To point(i).NLinksDaqui
                Delete_Link (point(i).iLinkDaqui(1))
            Next k
            For k = 1 To point(i).NLinksPraKa
                Delete_Link (point(i).iLinkPraKa(1))
            Next k
            point(i).isM2 = 0
        End If
        If Len(point(i).Isstation) = 1 Then
            If point(i).IsAutoNetwork Then Debug.Print point(i).Name & " É AUTO E " & point(i).Isstation
            Select Case point(i).Isstation
                Case "b"
                    point(i).t3 = 3
                Case "t"
                    point(i).t3 = 4
                Case "m"
                    point(i).t3 = 5
                Case "y"
                    point(i).t3 = 6
            End Select
        End If
        If point(i).IsAutoNetwork Then point(i).t3 = 2
        If point(i).Name = "22424" Then
            A = B
        End If
        If point(i).t3 = 2 Then
            If point(i).nRoutes > 1 Then point(i).t3 = 3
            If NAutoVizinhos(i) = 2 Then
                deflexão = DifAngle(Angle(point(PointList(1)), point(i)), Angle(point(i), point(PointList(2))))
                If Distância(i, PointList(2), True) < 50 Then deflexão = 0
                If Distância(i, PointList(1), True) < 50 Then deflexão = 0
                If deflexão < 45 Then
                    If Get_Link(PointList(1), PointList(2)) = 0 Then
                        il1 = Get_Link(PointList(1), i)
                        il2 = Get_Link(i, PointList(2))
                        If il1 <> 0 And il2 <> 0 Then
                            CandiLink = link(il1)
                            If link(il1).Extension < link(il2).Extension Then CandiLink = link(il2)
                            CandiLink.Extension = link(il1).Extension + link(il2).Extension
                            CandiLink.t3 = CandiLink.Extension / (link(il1).Extension / link(il1).t3 + link(il2).Extension / link(il2).t3)
                            CandiLink.peda = link(il1).peda + link(il2).peda
                            CandiLink.op = PointList(1)
                            CandiLink.dp = PointList(2)
                            Delete_Link (il1)
                            Delete_Link (il2)
                            If Get_Link(CandiLink.op, CandiLink.dp) = 0 Then
                                bub = Add_Link(CandiLink)
                            Else
                                'poderia adicionar faixas
                            End If
                        End If
                        il1 = Get_Link(PointList(2), i)
                        il2 = Get_Link(i, PointList(1))
                        If il1 <> 0 And il2 <> 0 Then
                            CandiLink = link(il1)
                            If link(il1).Extension < link(il2).Extension Then CandiLink = link(il2)
                            CandiLink.Extension = link(il1).Extension + link(il2).Extension
                            CandiLink.t3 = CandiLink.Extension / (link(il1).Extension / link(il1).t3 + link(il2).Extension / link(il2).t3)
                            CandiLink.peda = link(il1).peda + link(il2).peda
                            CandiLink.op = PointList(2)
                            CandiLink.dp = PointList(1)
                            Delete_Link (il1)
                            Delete_Link (il2)
                            If Get_Link(CandiLink.op, CandiLink.dp) = 0 Then
                                bub = Add_Link(CandiLink)
                            Else
                                'poderia adicionar faixas
                            End If
                        End If
                        'Tira Ponto das linhas
                        For k = 1 To point(i).nRoutes
                            Remove_Point_from_Route i, point(i).iroute(k)
                        Next k
                        'Eventuais links remanescentes para o ponto (not AutoNetwork) são movidos
                        nll = 0
                        For k = 1 To point(i).NLinksDaqui
                            il = point(i).iLinkDaqui(k)
                            CandiLink = link(il)
                            CandiLink.op = PointList(1)
                            If Distância(PointList(1), CandiLink.dp, True) > Distância(PointList(2), CandiLink.dp, True) Then CandiLink.op = PointList(2)
                            CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True) * meter_per_extension_unit
                            If Get_Link(CandiLink.op, CandiLink.dp) = 0 And CandiLink.op <> CandiLink.dp Then Add_Link CandiLink
                            link(il).isM2 = 0
                            nll = nll + 1
                            LinkList(nll) = il
                        Next k
                        For k = 1 To point(i).NLinksPraKa
                            il = point(i).iLinkPraKa(k)
                            CandiLink = link(il)
                            CandiLink.dp = PointList(1)
                            If Distância(PointList(1), CandiLink.op, True) > Distância(PointList(2), CandiLink.op, True) Then CandiLink.dp = PointList(2)
                            CandiLink.Extension = Distância(CandiLink.op, CandiLink.dp, True) * meter_per_extension_unit
                            If Get_Link(CandiLink.op, CandiLink.dp) = 0 And CandiLink.op <> CandiLink.dp Then Add_Link CandiLink
                            link(il).isM2 = 0
                            nll = nll + 1
                            LinkList(nll) = il
                        Next k
                        ' deletion must be done out of the for next
                        For k = 1 To nll
                            Delete_Link (LinkList(k))
                        Next k
                        point(i).isM2 = 0
                        elimseg = elimseg + point(i).nRoutes
                        conta = conta + 1
                    End If
                End If
               ' Print #3, point(i).NAME
            End If
        End If
    End If
Next i
Debug.Print elimseg; conta
End Sub


