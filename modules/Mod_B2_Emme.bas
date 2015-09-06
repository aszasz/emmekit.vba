Attribute VB_Name = "Mod_B2_Emme"
Option Explicit
'Public veicXkm As Single
'Public veicXhora As Single
'Public volau As Single

'To keep track of how was the network after load
'so only changes can be written
Public OldNpoints As Long
Public OldNLinks As Long
Public OldNRoutes As Integer
'Public paxXkm As Single
'Public paxXhora As Single
'Public voltr As Single
'Public overcong As Single
'Public cong As Single
'Public undercong As Single
'Public TotalPublicTrips As Single
'Public TotalBoardings As Single

' Eventually we had problems with decimal exporting coordinates from EMME due to language settings in Windows...
' So, instead of running a corrector in the EMME output we copied and pasted values from EMME to a spreadsheet and have the reading from here.
Sub ReadM2PointsfromTable(plan As Worksheet, startRow As Long, startCol As Integer, Optional CentroidsAreAboveNode As Integer = 0, Optional CentroidsAreBellowNode As Integer = 1000)
    Dim CandiPoint As Point_type
    iRow = startRow
    Do While plan.Cells(iRow, startCol) <> ""
        icol = startCol
        CandiPoint.Name = plan.Cells(iRow, icol): icol = icol + 1
        If Val(CandiPoint.Name) < CentroidsAreBellowNode And Val(CandiPoint.Name) > CentroidsAreAboveNode Then
            CandiPoint.isM2 = 1
        Else
            CandiPoint.isM2 = 2
        End If
        CandiPoint.x = plan.Cells(iRow, icol): icol = icol + 1
        CandiPoint.y = plan.Cells(iRow, icol): icol = icol + 1
        CandiPoint.t1 = plan.Cells(iRow, icol): icol = icol + 1
        CandiPoint.t2 = plan.Cells(iRow, icol): icol = icol + 1
        CandiPoint.t3 = plan.Cells(iRow, icol): icol = icol + 1
        CandiPoint.STname = plan.Cells(iRow, icol): icol = icol + 1
        TP = Npoints 'check how many points already exists
        ip = Get_Point(CandiPoint.x, CandiPoint.y, True, CandiPoint.isM2, CandiPoint.Name) 'get the point number, (True for: create it if does not exist)
        If ip <= TP Then 'if ip<=TP, the point already existed, so something is wrong
            Debug.Print point(ip).Name & " " & CandiPoint.Name
            ip = Add_Point(CandiPoint.x, CandiPoint.y, CandiPoint.isM2, CandiPoint.Name)
        End If
'        Else 'we're cool the point was created, need to be filled
        point(ip).isM2 = CandiPoint.isM2
        point(ip).t1 = CandiPoint.t1
        point(ip).t2 = CandiPoint.t2
        point(ip).t3 = CandiPoint.t3
        point(ip).STname = CandiPoint.STname
        point(ip).Name = CandiPoint.Name
        iRow = iRow + 1
    Loop
End Sub
Sub ReadM2NetworkFile(fullfile As String, msgCell As Range)
'Return String is an exit message
    Dim stringline As String
    Dim CandiPoint As Point_type
    Dim CandiLink As Link_type
    Dim ip, il, TP, N As Long
    Dim shortFileName As String
    N = CountStrings(fullfile, "\")
    shortFileName = Splitword(N)
    stringline = ""
    Open fullfile For Input As #1
    Do While Not EOF(1)
        If Left(stringline, 1) <> "t" Then Line Input #1, stringline
        N = CountWords(stringline)
        If Splitword(1) = "t" And Splitword(2) = "nodes" Then
'            If N = 3 And Splitword(3) = "init" Then ResetNodes
            stringline = ""
            Do While Not EOF(1) And Left(stringline, 1) <> "t"
                Line Input #1, stringline
                If Left(stringline, 1) <> "c" And Left(stringline, 1) <> "t" And stringline <> "" Then
                    CandiPoint = Read_M2Node(stringline) 'saves the reading in CandiPoint
                    If CandiPoint.marker = 1 Then   'if the reading was ok then
                        TP = Npoints 'check how many points already exists
                        ip = Get_Point(CandiPoint.x, CandiPoint.y, True, CandiPoint.isM2, CandiPoint.Name) 'get the point number, (True for: create it if does not exist)
                        If ip <= TP Then 'if ip<=TP, the point already existed, so something is wrong
                            Debug.Print point(ip).Name & " " & Splitword(2)
                            CandiPoint.x = CandiPoint.x + 25 * (360 / 40000000)
                            ip = Get_Point(CandiPoint.x, CandiPoint.y, True, CandiPoint.isM2, CandiPoint.Name) 'get the point number, (True for: create it if does not exist)
                        End If
                        point(Npoints).isM2 = CandiPoint.isM2
                        point(Npoints).t1 = CandiPoint.t1
                        point(Npoints).t2 = CandiPoint.t2
                        point(Npoints).t3 = CandiPoint.t3
                        point(Npoints).STname = CandiPoint.STname
                        point(Npoints).Name = CandiPoint.Name
                    ElseIf msgCell <> vbNull Then
                        msgCell = msgCell & Chr(10) & "  Fail to read node in line--> " & stringline
                    End If
                End If
            Loop
        ElseIf Splitword(1) = "t" And Splitword(2) = "links" Then
'            If N = 3 And Splitword(3) = "init" Then ResetLinks
            stringline = ""
            Do While Left(stringline, 1) <> "t" And Not EOF(1)
                Line Input #1, stringline
                If Left(stringline, 1) <> "c" And Left(stringline, 1) <> "t" And stringline <> "" Then
                    CandiLink = Read_M2Link(stringline) 'saver reading in CandiLink
                    If CandiLink.Auxiliar = 1 Then
                        il = Get_Link(CandiLink.op, CandiLink.dp)
                        If il <> 0 Then
                            If msgCell <> vbNull Then msgCell = msgCell & Chr(10) & "could not create/find link in line --> " & stringline & Chr(13)
                        Else
                            Call Add_Link(CandiLink)
                            ReDim link(NLinks).Cosmetic(0)
                            link(NLinks).Cosmetic(0) = 0
                            ReDim link(NLinks).iroute(0)
                            link(NLinks).iroute(0) = 0
                        End If
                    End If
                End If
            Loop
        ElseIf stringline <> "" And Left(stringline, 1) <> "c" Then
            If MsgBox("Unable to process line:" & Chr(13) & stringline & Chr(13) & "Ignore?", vbYesNo, "ReadM2NetworkFile:" & shortFileName) = vbYes Then
                msgCell = msgCell & Chr(10) & "Ignoring line --> " & stringline & Chr(13)
                End
            Else
                msgCell = msgCell & Chr(10) & "Ending, not able to process line --> " & stringline & Chr(13)
            End If
        End If
    Loop
    Close #1
End Sub
Function Read_M2Node(stringline As String) As Point_type
' This function saves the EMME2Node in Read_M2Node with .mark=1 if the reading was OK,
' it does not insert the point in Mpoints and network for finding
    Dim isM2 As Integer
    Dim xcord As Double
    Dim ycord As Double
    Read_M2Node.marker = 0
    Dim N As Integer
    N = CountWords(stringline)
    If N <> 8 Then Exit Function
    'IS CENTROID
    Select Case Splitword(1)
        Case "a"
            Read_M2Node.isM2 = 2
        Case "a*"
            Read_M2Node.isM2 = 1
        Case Else
            Exit Function
    End Select
    'X COORDINATE
    If Not IsNumeric(Splitword(3)) Then
        Exit Function
    Else
        xcord = Val(Splitword(3))
        If xcord = 0 Then Exit Function
        Read_M2Node.x = xcord
    End If
    'Y COORDINATE
    If Not IsNumeric(Splitword(4)) Then
        Exit Function
    Else
        ycord = Val(Splitword(4))
        If ycord = 0 Then Exit Function
        Read_M2Node.y = ycord
    End If
    Read_M2Node.t1 = Val(Splitword(5))
    Read_M2Node.t2 = Val(Splitword(6))
    Read_M2Node.t3 = Val(Splitword(7))
    Read_M2Node.Name = Splitword(2)
    Read_M2Node.STname = Splitword(8)
    Read_M2Node.marker = 1
End Function
Function Read_M2Link(stringline As String) As Link_type
    Read_M2Link.Auxiliar = 0
    Dim N As Integer
    Dim P As Long
    N = CountWords(stringline)
    If N <> 11 Then MsgBox "hELLO": Exit Function
    If Splitword(1) <> "a" Then Exit Function
    
    'ORIGIN NODE
    If Not IsNumeric(Splitword(3)) Then
        Exit Function
    Else
        P = PointNamed(Splitword(2))
        If P <> 0 Then
            Read_M2Link.op = P
        Else
 '           MsgBox ("Origin Point Not Found")
        End If
    End If
    'DESTINY NODE
    If Not IsNumeric(Splitword(3)) Then
        Exit Function
    Else
        P = PointNamed(Splitword(3))
        If P <> 0 Then
            Read_M2Link.dp = P
        Else
  '          MsgBox ("Destiny Point Not Found")
        End If
    End If
    If point(Read_M2Link.op).isM2 = 1 Or point(Read_M2Link.dp).isM2 = 1 Then
        Read_M2Link.isM2 = 1
    Else
        Read_M2Link.isM2 = 2
    End If
    Read_M2Link.Extension = Val(Splitword(4))
'    Read_M2Link.Length = val(Splitword(4)) * meter_per_extension_unit
    Read_M2Link.modes = Splitword(5)
    Read_M2Link.tipo = Splitword(6)
    Read_M2Link.Lanes = Val(Splitword(7))
    Read_M2Link.vdf = Splitword(8)
    Read_M2Link.t1 = Val(Splitword(9))
    Read_M2Link.t2 = Splitword(10)
    Read_M2Link.t3 = Val(Splitword(11))
    Read_M2Link.Auxiliar = 1
End Function
Sub ReadM2RoutesFile(fullfile As String, msgCell As Range, _
                    Optional NeverMindNodes As Boolean = False, _
                    Optional CreateNeededLinksandModes As Boolean = False)
    'Presumes network is read
    'The  routine ignore a node that appear twice in sequence (happens when it throws out virutal nodes)
    'if NeverMindNodes is set false, it will inform errors in reading nodes and links as it is expected
    'if CreateNeedLinks is set false, it will inform errors (absent links) in links based in NeverMindNodes
    '    - if both are false, it will work as it was original routine... that's when the routine care for the nodes and links
    '    - if both are true, will create the needed network for input the routes (and eventually a route can have no nodes)
    '    - if first is false and second true: it won't demand nodes, but once they exist creates new links
    '    - if first is true, but second is false: it will input the node, if it exists, even without the link
    
    Dim stringline As String 'for calling Splitword and SplitwordB
    Dim Dlink As Link_type
    Dim il As Long, P  As Long
    Dim N, i As Integer
    Dim shortFileName, FirstCh
    Dim para As String 'temp input for allow board/alight
    Dim tempopara As Single 'temp input for dwelltima
    Dim functempo As Integer 'temporary value for ttf
    Dim functempot As Integer 'temporary value for ttft
    Dim functempol As Integer 'temporary value for ttfl
    Dim us1 As Single
    Dim us2 As Single
    Dim us3 As Single
    Dim keep_going As Boolean
    N = CountStrings(fullfile, "\")
    shortFileName = Splitword(N)
    Open fullfile For Input As #1
    Do While Not EOF(1)
        Line Input #1, stringline
        FirstCh = Left(stringline, 1)
        If FirstCh = "c" Or (FirstCh = "t" And Left(stringline, 2) <> "tt") Then
        
        ElseIf FirstCh = "a" Then
            If CountStrings(stringline, "'") <> 5 Then MsgBox "Hello, check me"
            If Splitword(1) <> "a" Then MsgBox "Hello, check me"
            nRoutes = nRoutes + 1
            ReDim Preserve route(nRoutes)
            route(nRoutes).number = Splitword(2)
            route(nRoutes).Name = Splitword(4)
            If CountWords(Splitword(3) & " " & Splitword(5)) <> 7 Then MsgBox "Hello, check me"
            route(nRoutes).mode = Splitword(1)
            route(nRoutes).vehicle = Splitword(2)
            route(nRoutes).headway = Val(Splitword(3))
            route(nRoutes).speed = Val(Splitword(4))
            route(nRoutes).t1 = Val(Splitword(5))
            route(nRoutes).t2 = Val(Splitword(6))
            route(nRoutes).t3 = Splitword(7)
            'emme standard values when nothing said
            para = "+"     '=boarding and alighting allowed
            functempo = 0  '=ttf (standard transit time function in emme)
            tempopara = 0  '=dwt (dwell time in emme)
            functempot = 0  '=tttf
            functempol = 0  '=ttfl
            us1 = 0
            us2 = 0
            us3 = 0
            route(nRoutes).Npoints = 0
            route(nRoutes).HasPara = True
            route(nRoutes).HasDwt = True
            route(nRoutes).HasTtf = True
            route(nRoutes).HasTtfL = True
            route(nRoutes).HasTtf = True
            route(nRoutes).HasTtfT = True
            route(nRoutes).HasUs1 = True
            route(nRoutes).HasUs2 = True
            route(nRoutes).HasUs3 = True
            ReDim route(nRoutes).ipoint(route(nRoutes).Npoints)
            ReDim route(nRoutes).para(route(nRoutes).Npoints)
            ReDim route(nRoutes).dwt(route(nRoutes).Npoints)
            ReDim route(nRoutes).ttf(route(nRoutes).Npoints)
            ReDim route(nRoutes).ttfT(route(nRoutes).Npoints)
            ReDim route(nRoutes).ttfL(route(nRoutes).Npoints)
            ReDim route(nRoutes).us1(route(nRoutes).Npoints)
            ReDim route(nRoutes).us2(route(nRoutes).Npoints)
            ReDim route(nRoutes).us3(route(nRoutes).Npoints)
            route(nRoutes).para(route(nRoutes).Npoints) = para
            route(nRoutes).dwt(route(nRoutes).Npoints) = tempopara
            route(nRoutes).ttf(route(nRoutes).Npoints) = functempo
            route(nRoutes).ttfT(route(nRoutes).Npoints) = functempot
            route(nRoutes).ttfL(route(nRoutes).Npoints) = functempol
            route(nRoutes).us1(route(nRoutes).Npoints) = us1
            route(nRoutes).us2(route(nRoutes).Npoints) = us2
            route(nRoutes).us3(route(nRoutes).Npoints) = us3
        Else
            For i = 1 To CountWords(stringline)
                If CountStringsB(Splitword(i), "=") = 2 Then
                    If SplitwordB(1) = "dwt" Then
                        para = Left(SplitwordB(2), 1)
                        If IsNumeric(para) Or para = "." Then para = "+"
                        If Left(SplitwordB(2), 1) = "#" Or Left(SplitwordB(2), 1) = ">" Or Left(SplitwordB(2), 1) = "<" Or Left(SplitwordB(2), 1) = "+" Then SplitwordB(2) = Mid(SplitwordB(2), 2)
                        tempopara = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "ttf" Then
                        functempol = Val(SplitwordB(2))
                        functempot = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "ttfl" Then
                        functempol = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "ttft" Then
                        functempot = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "us1" Then
                        us1 = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "us2" Then
                        us2 = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "us3" Then
                        us3 = Val(SplitwordB(2))
                    ElseIf SplitwordB(1) = "lay" Then
                        route(nRoutes).lay = SplitwordB(2)
                    ElseIf SplitwordB(1) = "path" Then
                        route(nRoutes).path = SplitwordB(2)
                    End If
                ElseIf IsNumeric(Splitword(i)) Then
                    P = PointNamed(Splitword(i))
                    If P <> 0 And P <> route(nRoutes).ipoint(route(nRoutes).Npoints) Then
                        route(nRoutes).Npoints = route(nRoutes).Npoints + 1
                        ReDim Preserve route(nRoutes).ipoint(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).para(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).dwt(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).ttf(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).ttfT(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).ttfL(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).us1(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).us2(route(nRoutes).Npoints)
                        ReDim Preserve route(nRoutes).us3(route(nRoutes).Npoints)
                        route(nRoutes).ipoint(route(nRoutes).Npoints) = P
                        route(nRoutes).para(route(nRoutes).Npoints) = para
                        route(nRoutes).dwt(route(nRoutes).Npoints) = tempopara
                        route(nRoutes).ttfT(route(nRoutes).Npoints) = functempot
                        route(nRoutes).ttfL(route(nRoutes).Npoints) = functempol
                        route(nRoutes).us1(route(nRoutes).Npoints) = us1
                        route(nRoutes).us2(route(nRoutes).Npoints) = us2
                        route(nRoutes).us3(route(nRoutes).Npoints) = us3
                        'This have been controversial: should a line that passes twice at the same point be entered twice... it is annoyng when used for integration
                        'In this version register only once
                        If point(P).iroute(point(P).nRoutes) <> nRoutes Then
                            point(P).nRoutes = point(P).nRoutes + 1
                            ReDim Preserve point(P).iroute(point(P).nRoutes)
                            point(P).iroute(point(P).nRoutes) = nRoutes
                        End If
                        If route(nRoutes).Npoints >= 2 Then
                            'check link existence and if mode is allowed
                            il = Get_Link(route(nRoutes).ipoint(route(nRoutes).Npoints - 1), route(nRoutes).ipoint(route(nRoutes).Npoints))
                            If il = 0 Then
                                keep_going = False
                                If Not CreateNeededLinksandModes And Not NeverMindNodes Then
                                    If MsgBox("Route: " & route(nRoutes).number & " should pass in unknown link between points " & _
                                              point(route(nRoutes).ipoint(route(nRoutes).Npoints)).Name & " and " & Splitword(i) & _
                                              "." & Chr(13) & " The link will be created if proceed, proceed anyway? ", vbYesNo, "ReadM2Routes:" & shortFileName) = vbNo Then End
                                    keep_going = True
                                End If
                                If CreateNeededLinksandModes Or keep_going Then
                                    Dlink.op = route(nRoutes).ipoint(route(nRoutes).Npoints - 1)
                                    Dlink.dp = route(nRoutes).ipoint(route(nRoutes).Npoints)
                                    Dlink.Extension = Distância(Dlink.op, Dlink.dp) / meter_per_extension_unit
                                    Dlink.tipo = 1000
                                    Dlink.modes = route(nRoutes).mode
                                    Dlink.Lanes = 0
                                    il = Add_Link(Dlink)
                                End If
                            End If
                            If Not LinkhasMode(il, route(nRoutes).mode) Then
                                keep_going = False
                                If Not CreateNeededLinksandModes And Not NeverMindNodes Then
                                    If MsgBox("Route: " & route(nRoutes).number & " Link " & _
                                              point(route(nRoutes).ipoint(route(nRoutes).Npoints)).Name & " and " & Splitword(i) & _
                                              " doesn't take mode " & route(nRoutes).mode & "." & Chr(13) & "The mode will be add if proceed, proceed anyway? ", vbYesNo, "ReadM2Routes:" & shortFileName) = vbNo Then End
                                    keep_going = True
                                End If
                                If CreateNeededLinksandModes Or keep_going Then
                                    link(il).modes = link(il).modes & route(nRoutes).mode
                                    If link(il).tipo < 1000 Then link(il).tipo = link(il).tipo + 1000
                                End If
                            End If
                        End If
                    End If
                Else
                    If MsgBox("Route: " & route(nRoutes).number & " has unknown word " & Splitword(i) & ". The sentence will be ignored if proceed, proceed anyway? ", vbYesNo, "ReadM2Routes:" & shortFileName) = vbNo Then End
                End If
            Next i
        End If
    Loop
    Close #1
End Sub
Sub Register_Route_Links(iroute As Integer)
route(iroute).NLinks = 0
For i = 2 To route(iroute).Npoints
    il = Get_Link(route(iroute).ipoint(i - 1), route(iroute).ipoint(i))
    If il <> 0 Then
        route(iroute).NLinks = route(iroute).NLinks + 1
        ReDim Preserve route(iroute).ilink(route(iroute).NLinks)
        route(iroute).ilink(route(iroute).NLinks) = il
    Else
        MsgBox "Missing link #" & il - 1 & " in route " & route(iroute).Name _
        & "(" & iroute & ") from point " & point(route(iroute).ipoint(i - 1)).Name _
        & " to " & point(route(iroute).ipoint(i)).Name, , "Register_Route_Links"
    End If
Next i
End Sub
Sub Register_Routes_in_Network(Optional RegisterOnlyInSelectedpoints As Boolean = False, Optional RegisterOnlySelectedRoutes As Boolean = False, Optional islonglat As Boolean = False)
'Resets routes in nodes and links from network and run all routes to fill it.
'Plus review route position and extension to each point and total route extension
'Calculate extension based on link extension
'if uses a non-existing link, adds Euclidian distance between those points with NO WARNINGS
'if a route goes twice on the same point, register it twice (and distance related to each pass)
'routines that want to check it, just need to see if next route is the same as all the passages of
'each route are always together in a point or link
For i = 1 To Npoints
    If point(i).selected Or Not RegisterOnlyInSelectedpoints Then point(i).nRoutes = 0
Next i
For i = 1 To NLinks
    If link(i).selected Or Not RegisterOnlyInSelectedpoints Then link(i).nRoutes = 0
Next i
For iroute = 1 To nRoutes
    If route(iroute).selected Or Not RegisterOnlySelectedRoutes Then
        route(iroute).Extension = 0
        For i = 1 To route(iroute).Npoints
            If point(route(iroute).ipoint(i)).selected Or Not RegisterOnlyInSelectedpoints Then
                point(route(iroute).ipoint(i)).nRoutes = point(route(iroute).ipoint(i)).nRoutes + 1
                ReDim Preserve point(route(iroute).ipoint(i)).iroute(point(route(iroute).ipoint(i)).nRoutes)
                ReDim Preserve point(route(iroute).ipoint(i)).iRoutePosition(point(route(iroute).ipoint(i)).nRoutes)
                ReDim Preserve point(route(iroute).ipoint(i)).iRouteDistance(point(route(iroute).ipoint(i)).nRoutes)
                point(route(iroute).ipoint(i)).iroute(point(route(iroute).ipoint(i)).nRoutes) = iroute
                point(route(iroute).ipoint(i)).iRoutePosition(point(route(iroute).ipoint(i)).nRoutes) = i
            End If
            If i > 1 Then
                il = Get_Link(route(iroute).ipoint(i - 1), route(iroute).ipoint(i))
                If il <> 0 Then
                    If link(il).selected Or Not RegisterOnlyInSelectedpoints Then
                        route(iroute).Extension = route(iroute).Extension + link(il).Extension
                        link(il).nRoutes = link(il).nRoutes + 1
                        ReDim Preserve link(il).iroute(link(il).nRoutes)
                        link(il).iroute(link(il).nRoutes) = iroute
                    End If
                Else
                    route(iroute).Extension = route(iroute).Extension + Distância(route(iroute).ipoint(i - 1), route(iroute).ipoint(i)) / meter_per_extension_unit
                End If
                If point(route(iroute).ipoint(i)).selected Or Not RegisterOnlyInSelectedpoints Then point(route(iroute).ipoint(i)).iRouteDistance(point(route(iroute).ipoint(i)).nRoutes) = route(iroute).Extension
            End If
        Next i
    End If
Next iroute
End Sub

Sub CountRoutesinNetwork(Optional CountOnlyifT3LessThan0 As Boolean = False, Optional EvenRepeated As Boolean = False)
Dim i As Long, il As Long
Dim iroute As Integer
For i = 1 To Npoints
    point(i).nRoutes = 0
Next i
For i = 1 To NLinks
    link(i).nRoutes = 0
Next i
For iroute = 1 To nRoutes
    If (route(iroute).t3 < 0 Or Not CountOnlyifT3LessThan0) And Not route(iroute).deleted Then
        route(iroute).Extension = 0
        For i = 1 To route(iroute).Npoints
            If point(route(iroute).ipoint(i)).iroute(point(route(iroute).ipoint(i)).nRoutes) <> iroute Or Not EvenRepeated Then
                point(route(iroute).ipoint(i)).nRoutes = point(route(iroute).ipoint(i)).nRoutes + 1
                ReDim Preserve point(route(iroute).ipoint(i)).iroute(point(route(iroute).ipoint(i)).nRoutes)
                ReDim Preserve point(route(iroute).ipoint(i)).iRoutePosition(point(route(iroute).ipoint(i)).nRoutes)
                ReDim Preserve point(route(iroute).ipoint(i)).iRouteDistance(point(route(iroute).ipoint(i)).nRoutes)
                point(route(iroute).ipoint(i)).iroute(point(route(iroute).ipoint(i)).nRoutes) = iroute
                point(route(iroute).ipoint(i)).iRoutePosition(point(route(iroute).ipoint(i)).nRoutes) = i
            End If
            If i > 1 Then
                il = Get_Link(route(iroute).ipoint(i - 1), route(iroute).ipoint(i))
                If il <> 0 Then
                    route(iroute).Extension = route(iroute).Extension + link(il).Extension
                    If link(il).iroute(link(il).nRoutes) <> iroute Then
                        link(il).nRoutes = link(il).nRoutes + 1
                        ReDim Preserve link(il).iroute(link(il).nRoutes)
                        link(il).iroute(link(il).nRoutes) = iroute
                    End If
                End If
            End If
            point(route(iroute).ipoint(i)).iRouteDistance(point(route(iroute).ipoint(i)).nRoutes) = route(iroute).Extension
        Next i
    End If
Next iroute
End Sub
Function Route_Distance_Run(iroute As Integer, startpos As Integer, endpos As Integer) As Single
Route_Distance_Run = 0
For i = startpos + 1 To endpos
    il = Get_Link(route(iroute).ipoint(i - 1), route(iroute).ipoint(i))
    Route_Distance_Run = Route_Distance_Run + link(il).Extension
Next i
End Function
Sub ResetRoutes()
nRoutes = 0
ReDim route(nRoutes)
CountRoutesinNetwork
End Sub
Function Get_Route(route_number As String) As Integer
Get_Route = 0
Dim i As Integer
For i = 1 To nRoutes
    If route(i).number = route_number Then
        Get_Route = i: Exit For
    End If
Next i
End Function
Function DoesRoutepasshere(iroute As Integer, ipoint As Long) As Integer
DoesRoutepasshere = 0
Dim i As Integer
For i = 1 To point(ipoint).nRoutes
    If point(ipoint).iroute(i) = iroute Then DoesRoutepasshere = i: Exit Function
Next i
End Function
Function DoesRouteSTOPhere(iroute As Integer, ipoint As Long) As Integer
DoesRouteSTOPhere = 0
For i = 1 To point(ipoint).NStopRoutes
    If point(ipoint).iStopRoute(i) = iroute Then DoesRouteSTOPhere = i: Exit Function
Next i
End Function
Function Get_iroute_pointpos(ipoint As Long, iroute As Integer, Optional ByVal NPass As Integer = 1) As Integer
    For i = 1 To point(ipoint).nRoutes
        If point(ipoint).iroute(i) = iroute Then
            NPass = NPass - 1
            If NPass = 0 Then
                Get_iroute_pointpos = i
                Exit Function
            End If
        End If
    Next i
End Function
Function Get_iroute_stopPos(ipoint As Long, iroute As Integer, Optional ByVal NPass As Integer = 1) As Integer
    For i = 1 To point(ipoint).NStopRoutes
        If point(ipoint).iStopRoute(i) = iroute Then
            NPass = NPass - 1
            If NPass = 0 Then
                Get_iroute_stopPos = i
                Exit Function
            End If
        End If
    Next i
End Function

Sub SimplifyBaseNetwork(Optional DeletionFullFilename As String = "", Optional AboveisVirtual As Long = 10000)
Dim dummylink As Link_type
Dim GuardaPoint() As Point_type
Dim guardaLink() As Linktype_type
ReDim GuardaPoint(Npoints)
ReDim guardaLink(NLinks)
'Guarda
For i = 1 To Npoints
    GuardaPoint(i) = point(i)
Next i
For i = 1 To NLinks
    guardaLink(i) = link(i)
Next i
'Reseting
NP = 0
NL = 0
Npoints = 0
NLinks = 0
ReDim point(0)
ReDim link(0)
Dim iroute As Integer
For i = 1 To MAXNUMNOMES
    PointNamed(i) = 0
Next i
For i = 1 To DIM_XPOINTFINDER
    XPointFinder(i) = 0
Next i
For i = 1 To DIM_YPOINTFINDER
    YPointFinder(i) = 0
Next i

If DeletionFullFilename <> "" Then Open DeletionFullFilename For Output As #1
For i = 1 To NP
    If GuardaPoint(i).Name < AboveisVirtual Then
        Add_Point GuardaPoint(i).x, GuardaPoint(i).y, GuardaPoint(i).isM2, GuardaPoint(i).Name
    End If
Next i
If DeletionFullFilename <> "" Then Print #1, "t links"
For i = 1 To NL
    If GuardaPoint(link(i).op).Name < AboveisVirtual And GuardaPoint(link(i).dp).Name < AboveisVirtual Then
        guardaLink(i).op = PointNamed(GuardaPoint(link(i).op).Name)
        guardaLink(i).dp = PointNamed(GuardaPoint(link(i).dp).Name)
        Add_Link (guardaLink(i))
    ElseIf DeletionFullFilename <> "" Then
        Print #1, "d " & GuardaPoint(link(i).op).Name & "   " & GuardaPoint(link(i).dp).Name
    End If
Next i
If DeletionFullFilename <> "" Then
    Print #1, "t nodes"
    For i = 1 To NP
        If GuardaPoint(i).Name > AboveisVirtual Then
            Print #1, "d " & GuardaPoint(i).Name
        End If
    Next i
End If
If DeletionFullFilename <> "" Then Close #1

'MoveRoutestonewpoints
For i = 1 To nRoutes
    lastvalidpoint = 0
    For j = 1 To route(i).Npoints
        CandiPoint = PointNamed(GuardaPoint(route(i).ipoint(j)).Name)
        If CandiPoint <> 0 And CandiPoint <> route(i).ipoint(lastvalidpoint) Then
            lastvalidpoint = lastvalidpoint + 1
            route(i).ipoint(lastvalidpoint) = CandiPoint
        End If
    Next j
    route(i).Npoints
    ReDim Preserve route(i).ipoint(route(i).Npoints)
Next i
CountRoutesinNetwork
End Sub
Sub Remove_Point_from_Route(ipoint As Long, iroute As Integer)
Dim i As Integer
i = iroute
lastvalidpoint = 0
For j = 1 To route(i).Npoints
    If route(i).ipoint(j) <> ipoint Then
        lastvalidpoint = lastvalidpoint + 1
        route(i).ipoint(lastvalidpoint) = route(i).ipoint(j)
        route(i).para(lastvalidpoint) = route(i).para(j)
    End If
Next j
route(i).Npoints = lastvalidpoint
ReDim Preserve route(i).ipoint(route(i).Npoints)
ReDim Preserve route(i).para(route(i).Npoints)
End Sub
Sub SimplifyAllRoutes(Optional AboveisVirtual As Long = 10000)  'Eliminates nodes above 'AboveIsVirtual' from all routes
Dim iroute As Integer
For iroute = 1 To nRoutes 'For all Routes
    SimplifyRoute iroute, AboveisVirtual
Next iroute
End Sub
Sub Fill_AllRoutes_ilinks(Optional onlyincludeLinktypebellow As Integer = 10000, Optional arealimite As Integer = 0)
Dim i As Integer
For i = 1 To nRoutes
    Fill_Route_ilinks i, onlyincludeLinktypebellow, arealimite
Next i
End Sub
Sub Fill_Route_ilinks(iroute As Integer, Optional onlyincludeLinktypebellow As Integer = 10000, Optional arealimite As Integer = 0)
    route(iroute).NLinks = 0
    route(iroute).Extension = 0
    route(iroute).MPaxKm = 0
    route(iroute).MTime = 0
    route(iroute).MPaxHour = 0
    If route(iroute).HasVoltr Then route(iroute).MMaxVol = 1
    RouteFreq = 60 / route(iroute).headway
    arealimite = 0
    If route(iroute).mode = "i" Then limited = arealimite
    For ipoint = 2 To route(iroute).Npoints
        il = Get_Link(route(iroute).ipoint(ipoint - 1), route(iroute).ipoint(ipoint))
        If link(il).tipo < onlyincludeLinktypebellow And point(link(il).op).Limite >= limited And point(link(il).dp).Limite >= limited Then
            route(iroute).NLinks = route(iroute).NLinks + 1
            ReDim Preserve route(iroute).ilink(route(iroute).NLinks)
            route(iroute).ilink(route(iroute).NLinks) = il
            If route(iroute).HasVoltr Then
                ReDim Preserve route(iroute).linkvoltr(route(iroute).NLinks)
                ReDim Preserve route(iroute).linktimetr(route(iroute).NLinks)
                route(iroute).linkvoltr(route(iroute).NLinks) = route(iroute).voltr(ipoint)
                route(iroute).linktimetr(route(iroute).NLinks) = route(iroute).timetr(ipoint)
                route(iroute).MTime = route(iroute).MTime + route(iroute).timetr(ipoint)
                If route(iroute).linkvoltr(route(iroute).NLinks) > route(iroute).MMaxVol Then route(iroute).MMaxVol = route(iroute).linkvoltr(route(iroute).NLinks)
                route(iroute).MPaxKm = route(iroute).MPaxKm + route(iroute).voltr(ipoint) * link(il).Extension / 1000 ' (1000 meters per Km)
                route(iroute).MPaxHour = route(iroute).MPaxHour + route(iroute).voltr(ipoint) * route(iroute).timetr(ipoint) / 60
            End If
            route(iroute).Extension = route(iroute).Extension + link(il).Extension
        End If
    Next ipoint
End Sub
Function RemoveRouteFromPoint(iroute As Integer, ipoint As Long) As Boolean
For i = 1 To point(ipoint).nRoutes
    If point(ipoint).iroute(i) = iroute Then
        For j = i To point(ipoint).nRoutes - 1
            point(ipoint).iroute(j) = point(ipoint).iroute(j + 1)
        Next j
        point(ipoint).nRoutes = point(ipoint).nRoutes - 1
        ReDim Preserve point(ipoint).iroute(point(ipoint).nRoutes)
        RemoveRouteFromPoint = True
        Exit Function
    End If
Next i
End Function
Function Add_Route_OtherPointList(number As String, Name As String, mode As String, vehicle As Integer, headway As Single, speed As Single, NinPointsList As Integer, Optional ut1 As Single = 0, Optional ut2 As Single = 0, Optional ut3 As Single = 0) As Integer
            nRoutes = nRoutes + 1
            ReDim Preserve route(nRoutes)
            route(nRoutes).number = number
            route(nRoutes).Name = Name
            If Len(mode) > 1 Then MsgBox "Modes com mais de 1 digito!"
            route(nRoutes).mode = mode
            route(nRoutes).vehicle = vehicle
            route(nRoutes).headway = headway
            route(nRoutes).speed = speed
            route(nRoutes).t1 = ut1
            route(nRoutes).t2 = ut2
            route(nRoutes).t3 = ut3
            route(nRoutes).Npoints = NinPointsList
            ReDim route(nRoutes).ipoint(route(nRoutes).Npoints)
            For i = 1 To route(nRoutes).Npoints
                route(nRoutes).ipoint(i) = OtherPointList(i)
                point(route(nRoutes).ipoint(i)).nRoutes = point(route(nRoutes).ipoint(i)).nRoutes + 1
                ReDim Preserve point(route(nRoutes).ipoint(i)).iroute(point(route(nRoutes).ipoint(i)).nRoutes)
                point(route(nRoutes).ipoint(i)).iroute(point(route(nRoutes).ipoint(i)).nRoutes) = nRoutes
            Next i
            RouteCaminhoMínimo (nRoutes)
            Add_Route_OtherPointList = nRoutes
End Function
Function Add_Route(number As String, Name As String, mode As String, vehicle As Integer, headway As Single, speed As Single, NinPointsList As Integer, Optional ut1 As Single = 0, Optional ut2 As Single = 0, Optional ut3 As Single = 0) As Integer
            nRoutes = nRoutes + 1
            ReDim Preserve route(nRoutes)
            route(nRoutes).number = number
            route(nRoutes).Name = Name
            If Len(mode) > 1 Then MsgBox "Modes com mais de 1 digito!"
            route(nRoutes).mode = mode
            route(nRoutes).vehicle = vehicle
            route(nRoutes).headway = headway
            route(nRoutes).speed = speed
            route(nRoutes).t1 = ut1
            route(nRoutes).t2 = ut2
            route(nRoutes).t3 = ut3
            route(nRoutes).Npoints = NinPointsList
            ReDim route(nRoutes).ipoint(route(nRoutes).Npoints)
            For i = 1 To route(nRoutes).Npoints
                route(nRoutes).ipoint(i) = PointList(i)
            Next i
            RouteCaminhoMínimo (nRoutes)
End Function
Sub Add_points_to_route_OtherPointList(iroute As Integer, NinPointsList As Integer)
    ReDim Preserve route(iroute).ipoint(route(nRoutes).Npoints + NinPointsList)
    For i = 1 To NinPointsList
        route(iroute).ipoint(route(iroute).Npoints + i) = OtherPointList(i)
    Next i
    route(iroute).Npoints = route(iroute).Npoints + NinPointsList
    RouteCaminhoMínimo (nRoutes)
End Sub
Sub AddWalkingLinks(Optional AboveisVirtual = 10000)
Dim dummylink As Link_type
Dim il As Long
    For il = 1 To NLinks
        If LinkhasMode(il, "p") And link(il).tipo Mod 100 <> 8 Then
            li = Get_Link(link(il).dp, link(il).op)
            If li = 0 Then
                If Val(point(link(il).op).Name) < AboveisVirtual And Val(point(link(il).dp).Name) < AboveisVirtual Then
                dummylink = link(il)
                dummylink.modes = "p"
                dummylink.Lanes = 2
                dummylink.vdf = 1
                dummylink.tipo = 17
                dummylink.dp = link(il).op
                dummylink.op = link(il).dp
                bico = Add_Link(dummylink)
                End If
            End If
        End If
    Next il
End Sub
Sub SimplifyRoute(iroute As Integer, Optional AboveisVirtual As Long = 10000, Optional RemoveFromNodes As Boolean = False)
Dim i As Integer
i = iroute
lastvalidpoint = 0
For j = 1 To route(i).Npoints
    If point(route(i).ipoint(j)).Name < AboveisVirtual And route(i).ipoint(j) <> route(i).ipoint(lastvalidpoint) Then
        lastvalidpoint = lastvalidpoint + 1
        route(i).ipoint(lastvalidpoint) = route(i).ipoint(j)
    ElseIf point(route(i).ipoint(j)).Name < AboveisVirtual Then
        If RemoveFromNodes Then RemoveRouteFromPoint i, route(i).ipoint(j)
    End If
Next j
route(i).Npoints = lastvalidpoint
ReDim Preserve route(i).ipoint(route(i).Npoints)
Exit Sub

'The bellow is a weird code for the samething above... how come? After Test remove
nextpoint = 1 'start check in 1
lastvalidpoint = 0 'last valid point
While nextpoint <= route(iroute).Npoints 'repeat till the last route point, including the last
    While point(route(iroute).ipoint(nextpoint)).Name > AboveisVirtual And nextpoint < route(iroute).Npoints 'the last point is not checked
        If RemoveFromNodes Then RemoveRouteFromPoint iroute, route(iroute).ipoint(nextpoint)
        nextpoint = nextpoint + 1
    Wend
    'found nextpoint valid (whose name is bellow the parameter) OR IS THE LAST POINT (may or may not be valid)
    If route(iroute).ipoint(nextpoint) <> route(iroute).ipoint(lastvalidpoint) And point(route(iroute).ipoint(nextpoint)).Name < AboveisVirtual Then
   'If                        this point is not a return after virtual nodes   And        is valid THIS IS FOR CHECK THE LAST NODE ONLY
   'then includes the point in the position
        lastvalidpoint = lastvalidpoint + 1
        route(iroute).ipoint(lastvalidpoint) = route(iroute).ipoint(nextpoint)
    End If
    nextpoint = nextpoint + 1
Wend
route(iroute).Npoints = lastvalidpoint
ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
End Sub
Sub WriteEmmeRoutes(fullfilename As String, Optional NeverMindNodes As Boolean = False, Optional CreateNeededLinksandModes As Boolean = False, Optional IncludeDWT As Boolean = True, Optional onlyChanges As Boolean = False, Optional MarkT3 As Integer = 0, Optional apenda As Boolean = False)
'Presumes network is read
'if NeverMindNodes is set false, it will inform errors in reading nodes and links as it was in original routine
'if CreateNeedLinks is set false, it will inform errors in links based in NeverMindNodes:
'    - if both are false, it will work as it was original routine
'    - if first is false and second is true: will demand existing nodes, but creates new links or modes (link().tipo=1000+link().tipo)
'    - if first is true and second is false: will ignore absent nodes, but once the nodes are there, will demmand a link for/from them (with the right mode) to place the route
'    - if both are true, will create the needed network for input the routes (and eventually a route can have no nodes)
'IF OnlyChanges, will be based on the parameter OldNRoutes and Route().change, to see what is new
'                will route has not changed it won't be included, if route changed it will be deleted and included again
'                if markt3<>0 will use the value for marking the route is modified in ul3
Dim stringline As String, initstring As String
Dim shortFileName As String
Dim j As Integer, i As Integer, N As Integer
Dim print3 As Integer
Dim Dlink As Link_type
Dim lastpos As Integer
Dim il As Long, iroute As Integer
Dim LastPoint As Long
'check

For iroute = 1 To nRoutes
    If Not onlyChanges Or (onlyChanges And (iroute > OldNRoutes Or route(iroute).changed)) Then
        LastPoint = 0
        For j = 1 To route(iroute).Npoints
            If route(iroute).ipoint(j) = 0 And Not NeverMindNodes Then
                If MsgBox("Route: " & route(iroute).number & " should pass in unknown point " & point(route(iroute).ipoint(j)).Name & ". The point will be ignored if proceed, proceed anyway? ", vbYesNo, "WriteM2Routes") = vbNo Then End
            ElseIf route(iroute).ipoint(j) = route(iroute).ipoint(LastPoint) And Not NeverMindNodes Then
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
                            Dlink.Extension = Distância(Dlink.op, Dlink.dp) / meter_per_extension_unit
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
N = CountStrings(fullfilename, "\")
shortFileName = Splitword(N)
stringline = ""
If apenda Then
    Open fullfilename For Append As #1
    stringline = "Appending "
Else
    Open fullfilename For Output As #1
    stringline = "Writting "
End If
If onlyChanges Then
    Print #1, "c " & stringline & " only changes from running " & ThisWorkbook.Name & " in " & Now
    initstring = ""
Else
    Print #1, "c " & stringline & " full network from running " & ThisWorkbook.Name & " in " & Now
    initstring = " init "
End If
Print #1, "t lines" & initstring
For i = 1 To nRoutes
    If Not onlyChanges Or (onlyChanges And (i > OldNRoutes Or route(i).changed)) Then
        If onlyChanges And route(i).changed And i <= OldNRoutes Then Print #1, "d '" & route(i).number & "' "
        If Not route(i).deleted Then
            print3 = route(i).t3
            If MarkT3 <> 0 And (route(i).changed Or i > OldNRoutes) Then print3 = MarkT3
            stringline = "a" & "'" & route(i).number & "' " & route(i).mode
            stringline = stringline & "  " & route(i).vehicle
            stringline = stringline & "  " & Format(route(i).headway, "00.00")
            stringline = stringline & "  " & Format(route(i).speed, "00.00")
            If Len(route(i).Name) = 0 Then route(i).Name = " "
            stringline = stringline & " '" & Left(route(i).Name, 42) & "'"
            stringline = stringline & " " & Format(route(i).t1, "0.00") & " "
            stringline = stringline & " " & route(i).t2 & " "
            stringline = stringline & " " & route(i).t3 & " "
            Print #1, stringline
            stringline = ""
            lastpos = 0
            Add_sequence stringline, "path=" & route(i).path, 1
            LastPoint = 0 'last valid point printed
            For j = 1 To route(i).Npoints
                If route(i).ipoint(j) <> 0 And route(i).ipoint(j) <> route(i).ipoint(LastPoint) Then
    '                stringline = stringline & Right(Space(npos * 8) & sequence, npos * 8)
                    If IncludeDWT Then
                        If (route(i).dwt(LastPoint) <> route(i).dwt(j) Or route(i).para(LastPoint) <> route(i).para(j)) Then
                            Add_sequence stringline, "dwt=" & route(i).para(j) & route(i).dwt(j), 1
                        End If
                        If route(i).ttf(LastPoint) <> route(i).ttf(j) Then
                            Add_sequence stringline, "ttf=" & route(i).ttf(j), 1
                        End If
                    End If
                    Add_sequence stringline, point(route(i).ipoint(j)).Name, 1
                    LastPoint = j
                End If
            Next j
            Add_sequence stringline, "lay=" & route(i).lay, 1
            Print #1, " " & stringline
        End If
    End If
Next i
Close #1
End Sub
Sub Add_sequence(stringline As String, sequence As String, filenumber As Integer, Optional BlockLength As Integer = 8, Optional Nblocks As Integer = 8)
'This routine manages the printing of the stringline in blocks for WriteEmmeRoutes
    'the first char has to be space, then each sequence must be separated by space(up to col 73)
    'we chose to use the emme exit standard of 8 positions with 8 chars, but other arrangements can be made
    'but the first char has to be space to separate one from the other, so sequences can have effectivelly 7 chars for one position
    '                                                                                                     15 char for two positions
    'stringline is only to be written when it uses 8 positions
    'so when started, lastpos indicates the last used position of string line
    'lastpos indicates the last used position in the string
    'then npos gives the number of positions sequence needs to use: if it has 8 chars, needs 2 because of required leading space
    Dim lastpos As String
    Dim npos As Integer
    lastpos = Int((Len(stringline) - 1) / BlockLength) + 1
    npos = Int(Len(sequence) / BlockLength) + 1 'how many positions used by sequence
    If lastpos + npos > 8 Then               'whenever the new sequence is above position 8:
        Print #filenumber, " " & stringline  'print without including the sequence (that will go for the next line) but with lead space
        stringline = ""                      'and resets string line
        lastpos = 0
    End If
    stringline = stringline & Right(Space(npos * BlockLength) & sequence, npos * BlockLength)
End Sub
Sub WriteEmmeNetwork(fullfilename As String, Optional onlyChanges As Boolean = False, Optional apenda As Boolean = False, Optional MarkT3 As Single = 0, Optional NeverMindRoutes = False)
'IF OnlyChanges, will be based on the parameter OldNPoints and OldNLinks, to check what is new
'                will try to delete links with tipo=1000 or 100, depending of NeverMindRoutes
'                and consider that there were changes if tipo>1000 (will remove 1000 of tipo)
'                and will not print links with no modes
'                will delete (or not write) nodes that have .IsM2=100 DOES NOT CHECK FOR LINKS USING THE POINT
'                and consider that there were changes in nodes where .IsM2>100
'                if markt3<>0 will use the value for marking the link was modified in ul3
Dim N As Integer, i As Long
Dim shortFileName As String, stringline As String, initstring As String
Dim print3 As Single
Dim printipo As Integer

N = CountStrings(fullfilename, "\")
shortFileName = Splitword(N)
stringline = ""
CountRoutesinNetwork
If apenda Then
    Open fullfilename For Append As #2
    stringline = "Appending "
Else
    Open fullfilename For Output As #2
    stringline = "Writting"
End If
If onlyChanges Then
    Print #2, "c " & stringline & " only changes from running " & ThisWorkbook.Name & " in " & Now
    initstring = ""
Else
    Print #2, "c " & stringline & " full network from running " & ThisWorkbook.Name & " in " & Now
    initstring = " init "
End If
If onlyChanges Then
Print #2, "t links"
    For i = 1 To OldNLinks
        If link(i).isM2 = 0 Or link(i).tipo = 1000 Or link(i).modes = "" Then Print #2, "d " & point(link(i).op).Name & "  " & point(link(i).dp).Name
    Next i
End If
Print #2, "t nodes" & initstring
For i = 1 To Npoints
    print3 = point(i).t3
    If MarkT3 <> 0 And (i > OldNpoints Or point(i).isM2 > 100) Then print3 = MarkT3
        stringline = "a"
    If onlyChanges And i <= OldNpoints And point(i).isM2 > 100 Then stringline = "m"
    
    If point(i).isM2 = 1 Or point(i).isM2 = 101 Then
        stringline = stringline & "*"
    Else
        stringline = stringline & " "
    End If
    
    stringline = stringline & Right("    " & point(i).Name, 7)
    stringline = stringline & " " & point(i).x
    stringline = stringline & " " & point(i).y
    stringline = stringline & " " & Right("           " & Format(point(i).t1, "0.00"), 9)
    stringline = stringline & " " & Right("           " & Format(point(i).t2, "0.00"), 9)
    stringline = stringline & " " & Right("           " & Format(print3, "0.00"), 9)
    stringline = stringline & "  " & Right(point(i).STname, 4)
    
    If Not point(i).Deleta And point(i).isM2 <> 100 And point(i).isM2 > 0 And _
    (Not onlyChanges Or (onlyChanges And (i > OldNpoints Or point(i).isM2 > 100))) Then Print #2, stringline
Next i
Print #2,
Print #2, "t links" & initstring

For i = 1 To NLinks
    print3 = link(i).t3
    printipo = link(i).tipo
    If MarkT3 <> 0 And (i > OldNLinks Or link(i).tipo >= 1000) Then print3 = MarkT3
    stringline = Right("       " & point(link(i).op).Name, 7)
    stringline = stringline & Right("       " & point(link(i).dp).Name, 7)
    If link(i).tipo = 1000 Or link(i).isM2 = 0 Then
        If onlyChanges And i <= OldNLinks Then
            If link(i).nRoutes = 0 Or NeverMindRoutes Then
'                Print #2, "d " & stringline <-- this was moved to before adding points
            Else
                Print #2, "m " & stringline & " modes=-a"
            End If
        ElseIf link(i).nRoutes > 0 And Not NeverMindRoutes Then
                printipo = 100
        End If
    End If
    If printipo <> 1000 And link(i).modes <> "" Then
        If printipo > 1000 Then printipo = printipo - 1000
        stringline = stringline & " " & Right("       " & Format(link(i).Extension, "0.000"), 7)
        stringline = stringline & " " & Left(ordermodes(link(i).modes) & "          ", 10)
        stringline = stringline & " " & Right("       " & printipo, 4)
        stringline = stringline & " " & Format(link(i).Lanes, "0.0")
        stringline = stringline & " " & Right("         " & link(i).vdf, 7)
        stringline = stringline & " " & Right("           " & Format(link(i).t1, "0.00"), 9)
        stringline = stringline & " " & Right("           " & Format(link(i).t2, "0.00"), 9)
        stringline = stringline & " " & Right("           " & Format(print3, "0.00"), 9)
        If Not onlyChanges Or (onlyChanges And i > OldNLinks) Then
            If link(i).isM2 > 0 Then Print #2, "a " & stringline
        ElseIf onlyChanges And (i <= OldNLinks And link(i).tipo > 1000) Then
            Print #2, "m " & stringline
        End If
    End If
Next i

Dim yet As Boolean
'Delete nodes at the end of file
yet = False
For i = 1 To Npoints
    If point(i).isM2 = 100 And i <= OldNpoints And onlyChanges Then
        If Not yet Then Print #2, "t nodes": yet = True
        Print #2, "d " & point(i).Name
    End If
Next i

Print #2,  'empty line for safety
Close #2
End Sub
Sub WriteEmmeCosmeticNetwork(fullfilename As String, Optional onlyChanges As Boolean = False, Optional apenda As Boolean = False)
' Onlychanges not working yet
N = CountStrings(fullfilename, "\")
shortFileName = Splitword(N)
stringline = ""
CountRoutesinNetwork
If apenda Then
    Open fullfilename For Append As #2
    stringline = "Appending "
Else
    Open fullfilename For Output As #2
    stringline = "Writting"
End If
If onlyChanges Then
    Print #2, "c " & stringline & " only changes from running " & ThisWorkbook.Name & " in " & Now
    initstring = ""
Else
    Print #2, "c " & stringline & " full network from running " & ThisWorkbook.Name & " in " & Now
    initstring = " init "
End If
Print #2, "t linkvertices "

For i = 1 To NLinks
    If link(i).isM2 > 0 And link(i).delete = False Then
        Print #2, "r " & point(link(i).op).Name & " " & point(link(i).dp).Name
        For j = 1 To link(i).Ncosmetic
            Print #2, "a " & point(link(i).op).Name & " " & point(link(i).dp).Name & " " & j & " " & point(link(i).Cosmetic(j)).x & " " & point(link(i).Cosmetic(j)).y
        Next j
    End If
Next i
Print #2,  'empty line for safety
Close #2
End Sub
Function LinkhasMode(ilink As Long, mode As String) As Boolean
    Dim i As Integer
    LinkhasMode = False
    For i = 1 To Len(link(ilink).modes)
        If Mid(link(ilink).modes, i, 1) = mode Then LinkhasMode = True: Exit Function
    Next i
End Function
Function REmoveModefromlink(modes As String, mode As String) As String
    REmoveModefromlink = ""
    For i = 1 To Len(modes)
        found = False
        For j = 1 To Len(mode)
            If Mid(modes, i, 1) = Mid(mode, j, 1) Then found = True
        Next j
        If Not found Then REmoveModefromlink = REmoveModefromlink & Mid(modes, i, 1)
    Next i
End Function
Function ordermodes(modes As String) As String
    Dim i As Integer, j As Integer
    'Long expression only change positions... buble like
    For i = 1 To Len(modes)
        For j = i To Len(modes)
            If Mid(modes, j, 1) < Mid(modes, i, 1) Then modes = Left(modes, i - 1) & Mid(modes, j, 1) & Mid(modes, i + 1, j - i - 1) & Mid(modes, i, 1) & Mid(modes, j + 1)
        Next j
    Next i
    ordermodes = modes
End Function
Function SplitLinkPoint(ipoint As Long, jpoint As Long, Distance As Integer, Optional presid As Integer = 0, Optional IsLatLong As Boolean) As Long
Dim Dpoint As Point_type
Dim Dlink As Link_type
Dim ilink As Long
Dim jlink As Long
Dim ip As Long
Dim Ang As Single
Dim dista As Single
ilink = Get_Link(ipoint, jpoint)
jlink = Get_Link(jpoint, ipoint)
If ilink = 0 And jlink = 0 Then
    MsgBox "Can't split link that I don't know"
Else
    Ang = Angle(point(ipoint), point(jpoint))
    prop = Distance / Distância(ipoint, jpoint, IsLatLong)
    If prop > 1 Then prop = 0.9
    Dpoint = PointInsertPosition(point(ipoint), Ang, 0, 1, 0.01 + Distance, IsLatLong)
    ip = Add_Point(Dpoint.x, Dpoint.y, 2, Get_First_Avaiable_Name(30000, 999999))
    If ilink <> 0 Then
        Dlink = link(ilink)
        Dlink.t3 = presid
        Dlink.dp = ip
        Dlink.Extension = link(ilink).Extension * prop
        ilink1 = Add_Link(Dlink)
        Dlink.op = ip
        Dlink.dp = jpoint
        Dlink.Extension = link(ilink).Extension * (1 - prop)
        ilink2 = Add_Link(Dlink)
        If Not (Tira_Link(ipoint, ilink) And Tira_Link(jpoint, ilink)) Then MsgBox "Não tirou link"
        link(ilink).tipo = 1000
    End If
    If jlink <> 0 Then
        Dlink = link(jlink)
        Dlink.t3 = presid
        Dlink.dp = ip
        Dlink.Extension = link(ilink).Extension * (1 - prop)
        ilink3 = Add_Link(Dlink)
        Dlink.op = ip
        Dlink.dp = ipoint
        Dlink.Extension = link(ilink).Extension * prop
        ilink4 = Add_Link(Dlink)
        If Not (Tira_Link(ipoint, jlink) And Tira_Link(jpoint, jlink)) Then MsgBox "Não tirou link"
        link(jlink).tipo = 1000
    End If
    link(ilink).isM2 = 0
    link(jlink).isM2 = 0
    SplitLinkPoint = ip
    For iroute = 1 To point(ipoint).nRoutes
        For jroute = 1 To point(jpoint).nRoutes
            If point(ipoint).iroute(iroute) = point(jpoint).iroute(jroute) Then
                Update_Route_in_SplitLink point(ipoint).iroute(iroute), ipoint, jpoint, ip
            End If
        Next jroute
    Next iroute
End If
End Function
Sub Insert_middle_point_in_Routes_thru(iPointName As Long, JpointName As Long, MidPointName As Long, Optional OneWay As Boolean = False)
Dim iroute As Integer
For i = 1 To point(PointNamed(iPointName)).nRoutes
    iroute = point(PointNamed(iPointName)).iroute(i)
    Call Add_middle_point_in_Route(iroute, PointNamed(iPointName), PointNamed(JpointName), PointNamed(MidPointName), OneWay)
Next i
End Sub
Sub Shift_Route_Link(iroute As Integer, FromLink As Long, ToLink As Long)
For i = 1 To route(iroute).Npoints - 1
    If route(iroute).ipoint(i) = link(FromLink).op And route(iroute).ipoint(i + 1) = link(FromLink).dp Then
        route(iroute).ipoint(i) = link(ToLink).op
        route(iroute).ipoint(i + 1) = link(ToLink).dp
    End If
Next i
End Sub
Sub Add_middle_point_in_Route(iroute As Integer, ipoint As Long, jpoint As Long, MidPoint As Long, Optional OneWay As Boolean = False)
    Doitagain = True
    Do While Doitagain
        Doitagain = False
        For i = 1 To route(iroute).Npoints
            If route(iroute).ipoint(i) = ipoint And i > 1 And Not OneWay Then
                If route(iroute).ipoint(i - 1) = jpoint Then
                    route(iroute).Npoints = route(iroute).Npoints + 1
                    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
                    For j = route(iroute).Npoints To i + 1 Step -1
                        route(iroute).ipoint(j) = route(iroute).ipoint(j - 1)
                    Next j
                    route(iroute).ipoint(i) = MidPoint
                    Doitagain = True
                    Exit For
                End If
            End If
            If route(iroute).ipoint(i) = ipoint And i < route(iroute).Npoints Then
                If route(iroute).ipoint(i + 1) = jpoint Then
                    route(iroute).Npoints = route(iroute).Npoints + 1
                    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
                    For j = route(iroute).Npoints To i + 2 Step -1
                        route(iroute).ipoint(j) = route(iroute).ipoint(j - 1)
                    Next j
                    route(iroute).ipoint(i + 1) = MidPoint
                    Doitagain = True
                    Exit For
                End If
            End If
        Next i
    Loop
End Sub
Sub Update_Route_in_SplitLink(iroute As Integer, ipoint As Long, jpoint As Long, MidPoint As Long, Optional OneWay As Boolean = False)
Dim guarda(1500) As Long
For i = 1 To route(iroute).Npoints
    nguarda = nguarda + 1
    guarda(nguarda) = route(iroute).ipoint(i)
    If route(iroute).ipoint(i) = ipoint Then
        If i < route(iroute).Npoints Then
            If route(iroute).ipoint(i + 1) = jpoint Then
                nguarda = nguarda + 1
                guarda(nguarda) = MidPoint
'                point(MidPoint).NRoutes = point(MidPoint).NRoutes + 1
'                ReDim Preserve point(MidPoint).iroute(point(MidPoint).NRoutes)
'                point(MidPoint).iroute(point(MidPoint).NRoutes) = iroute
            End If
        End If
    End If
    If Not OneWay Then
        If route(iroute).ipoint(i) = jpoint Then
            If i < route(iroute).Npoints Then
                If route(iroute).ipoint(i + 1) = ipoint Then
                    nguarda = nguarda + 1
                    guarda(nguarda) = MidPoint
'                    point(MidPoint).NRoutes = point(MidPoint).NRoutes + 1
'                    ReDim Preserve point(MidPoint).iroute(point(MidPoint).NRoutes)
'                    point(MidPoint).iroute(point(MidPoint).NRoutes) = iroute
                End If
            End If
        End If
    End If
Next i
If nguarda <> route(iroute).Npoints Then
    route(iroute).Npoints = nguarda
    ReDim route(iroute).ipoint(route(iroute).Npoints)
    For i = 1 To nguarda
        route(iroute).ipoint(i) = guarda(i)
    Next i
End If
If route(iroute).NLinks > 0 Then
    Register_Route_Links (iroute)
End If
End Sub
Function Cut_Route(iroute As Integer, Optional FirstPoint As Long = 0, Optional LastPoint As Long = 0) As Boolean
'Corta a linha antes de firstpoint e depois de lastpoint
Dim i As Integer
Dim FROMPoint As Integer
Dim TOPoint As Integer
For i = 1 To route(iroute).Npoints
    If route(iroute).ipoint(i) = FirstPoint Then FROMPoint = i
    If route(iroute).ipoint(i) = LastPoint Then TOPoint = i
Next i

If FirstPoint = 0 Then FROMPoint = 1
If LastPoint = 0 Then TOPoint = route(iroute).Npoints

For i = FROMPoint To TOPoint
    route(iroute).ipoint(i - FROMPoint + 1) = route(iroute).ipoint(i)
    If route(iroute).HasPara Then route(iroute).para(i - FROMPoint + 1) = route(iroute).para(i)
    If route(iroute).HasPara Then route(iroute).dwt(i - FROMPoint + 1) = route(iroute).dwt(i)
    If route(iroute).HasTtf Then route(iroute).ttf(i - FROMPoint + 1) = route(iroute).ttf(i)
    If route(iroute).HasTtfT Then route(iroute).ttfT(i - FROMPoint + 1) = route(iroute).ttfT(i)
    If route(iroute).HasTtfL Then route(iroute).ttfL(i - FROMPoint + 1) = route(iroute).ttfL(i)
    If route(iroute).HasUs1 Then route(iroute).us1(i - FROMPoint + 1) = route(iroute).us1(i)
    If route(iroute).HasUs2 Then route(iroute).us2(i - FROMPoint + 1) = route(iroute).us2(i)
    If route(iroute).HasUs3 Then route(iroute).us3(i - FROMPoint + 1) = route(iroute).us3(i)
Next i

route(iroute).Npoints = TOPoint - FROMPoint + 1
ReDim Preserve route(iroute).ipoint(route(iroute).Npoints)
If route(iroute).HasPara Then ReDim Preserve route(iroute).para(route(iroute).Npoints)
If route(iroute).HasPara Then ReDim Preserve route(iroute).dwt(route(iroute).Npoints)
If route(iroute).HasTtf Then ReDim Preserve route(iroute).ttf(route(iroute).Npoints)
Cut_Route = True
End Function
Function Extend_Route_TO(iroute, NPointsinPointList) As Boolean
Dim j As Integer
Dim arethey As Boolean
Dim ENNE As Integer
Dim i As Integer
Dim ipoint As Integer
Extend_Route_TO = True
For i = 1 To NPointsinPointList
    ipoint = PointList(i)
    ENNE = NLinksInShortestPathWithinSelectedLinks(route(iroute).ipoint(route(iroute).Npoints), ipoint, arethey)
    If Not arethey Then MsgBox "Não é possível ligar o ponto final da linha " & route(iroute).number & " (ponto " & point(route(iroute).ipoint(route(iroute).Npoints)).Name & "). ao ponto " & point(ipoint).Name & ". A linha permanecerá com o mencionado ponto final.": Exit Function
    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints + ENNE)
    If route(iroute).HasPara Then ReDim Preserve route(iroute).para(route(iroute).Npoints + ENNE)
    If route(iroute).HasPara Then ReDim Preserve route(iroute).dwt(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtf Then ReDim Preserve route(iroute).ttf(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtfT Then ReDim Preserve route(iroute).ttfT(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtfL Then ReDim Preserve route(iroute).ttfL(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs1 Then ReDim Preserve route(iroute).us1(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs2 Then ReDim Preserve route(iroute).us2(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs3 Then ReDim Preserve route(iroute).us3(route(iroute).Npoints + ENNE)
    For j = route(iroute).Npoints + 1 To route(iroute).Npoints + ENNE
        route(iroute).ipoint(j) = link(LinkList(j - route(iroute).Npoints)).dp
        If route(iroute).HasPara Then route(iroute).para(j) = route(iroute).para(route(iroute).Npoints)
        If route(iroute).HasPara Then route(iroute).dwt(j) = route(iroute).dwt(route(iroute).Npoints)
        If route(iroute).HasTtf Then route(iroute).ttf(j) = route(iroute).ttf(route(iroute).Npoints)
        If route(iroute).HasTtfT Then route(iroute).ttfT(j) = route(iroute).ttfT(route(iroute).Npoints)
        If route(iroute).HasTtfL Then route(iroute).ttfL(j) = route(iroute).ttfL(route(iroute).Npoints)
        If route(iroute).HasUs1 Then route(iroute).us1(j) = route(iroute).us1(route(iroute).Npoints)
        If route(iroute).HasUs2 Then route(iroute).us2(j) = route(iroute).us2(route(iroute).Npoints)
        If route(iroute).HasUs3 Then route(iroute).us3(j) = route(iroute).us2(route(iroute).Npoints)
    Next j
    route(iroute).Npoints = route(iroute).Npoints + ENNE
Next i
End Function
Function Extend_Route_FROM(iroute, NPointsinPointList) As Boolean
Dim j As Integer
Dim arethey As Boolean
Dim ENNE As Integer
Extend_Route_FROM = True
Dim i As Integer
Dim ipoint As Integer
For i = 1 To NPointsinPointList
    ipoint = PointList(i)
    ENNE = NLinksInShortestPathWithinSelectedLinks(ipoint, route(iroute).ipoint(1), arethey)
    If Not arethey Then MsgBox "It is not possible to link  node " & point(ipoint).Name & " to first node of transit line " & route(iroute).number & " (node " & point(route(iroute).ipoint(1)).Name & "). Line will remain with its current initial node.": Exit Function
    ReDim Preserve route(iroute).ipoint(route(iroute).Npoints + ENNE)
    If route(iroute).HasPara Then ReDim Preserve route(iroute).para(route(iroute).Npoints + ENNE)
    If route(iroute).HasPara Then ReDim Preserve route(iroute).dwt(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtf Then ReDim Preserve route(iroute).ttf(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtfT Then ReDim Preserve route(iroute).ttfT(route(iroute).Npoints + ENNE)
    If route(iroute).HasTtfL Then ReDim Preserve route(iroute).ttfL(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs1 Then ReDim Preserve route(iroute).us1(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs2 Then ReDim Preserve route(iroute).us2(route(iroute).Npoints + ENNE)
    If route(iroute).HasUs3 Then ReDim Preserve route(iroute).us3(route(iroute).Npoints + ENNE)
    For j = route(iroute).Npoints + ENNE To ENNE + 1 Step -1
        route(iroute).ipoint(j) = route(iroute).ipoint(j - ENNE)
        If route(iroute).HasPara Then route(iroute).para(j) = route(iroute).para(j - ENNE)
        If route(iroute).HasPara Then route(iroute).dwt(j) = route(iroute).dwt(j - ENNE)
        If route(iroute).HasTtf Then route(iroute).ttf(j) = route(iroute).ttf(j - ENNE)
        If route(iroute).HasTtfT Then route(iroute).ttfT(j) = route(iroute).ttfT(j - ENNE)
        If route(iroute).HasTtfL Then route(iroute).ttfL(j) = route(iroute).ttfL(j - ENNE)
        If route(iroute).HasUs1 Then route(iroute).us1(j) = route(iroute).us1(j - ENNE)
        If route(iroute).HasUs2 Then route(iroute).us2(j) = route(iroute).us2(j - ENNE)
        If route(iroute).HasUs3 Then route(iroute).us3(j) = route(iroute).us2(j - ENNE)
    Next j
    For j = ENNE To 1 Step -1
        route(iroute).ipoint(j) = link(LinkList(j)).op
        If route(iroute).HasPara Then route(iroute).para(j) = route(iroute).para(ENNE + 1)
        If route(iroute).HasPara Then route(iroute).dwt(j) = route(iroute).dwt(ENNE + 1)
        If route(iroute).HasTtf Then route(iroute).ttf(j) = route(iroute).ttf(ENNE + 1)
        If route(iroute).HasTtfT Then route(iroute).ttfT(j) = route(iroute).ttfT(ENNE + 1)
        If route(iroute).HasTtfL Then route(iroute).ttfL(j) = route(iroute).ttfL(ENNE + 1)
        If route(iroute).HasUs1 Then route(iroute).us1(j) = route(iroute).us1(ENNE + 1)
        If route(iroute).HasUs2 Then route(iroute).us2(j) = route(iroute).us2(ENNE + 1)
        If route(iroute).HasUs3 Then route(iroute).us3(j) = route(iroute).us2(ENNE + 1)
    Next j
    route(iroute).Npoints = route(iroute).Npoints + ENNE
Next i
End Function
Sub SelectAllnotConnector()
    Dim i As Long
    For i = 1 To NLinks
        If link(i).isM2 = 2 Then
            link(i).selected = True
        Else
            link(i).selected = False
        End If
    Next i
End Sub

