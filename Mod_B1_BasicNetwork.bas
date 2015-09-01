Attribute VB_Name = "Mod_B1_BasicNetwork"
'This module was not originaly made for EMME Tools, but for Network Manipulations overall it started in Jakarta 2004
'The Jakarta version used to have a function called Link() that would return the object Link_type accepting
'inversion of origin and destination or negative number, this is not going to be used anymore as EMME has one-way links


Option Explicit

Public Const PI As Double = 3.14159265358979
Public Const meter_per_degree As Single = 40000000 / 360 ' this is aprox. for meridians (delta Y) and in the equator
Public approx_one_meter_in_network 'other modules may need to know the equivalence of one meter to place nodes
Public meter_per_extension_unit 'other modules may need to know the equivalence of one meter to link extension units as well
Public Const Infinite_Distance As Single = 100000000#  'the largest distance in the network, broadcast to other modules
                                          '100,000,000 is a really large distance in any practical unit in networks

'Network parameters... first thing to read, they are global

Dim NetIsLongLat As Boolean     'diferent methods for distance and angles if Xs and Ys are in Latitude and Longitude
Dim MIN_X As Single   'lesser x coordinate
Dim MAX_X As Single   'greater x coordinate
Dim MIN_Y As Single   'lesser y coordinate
Dim MAX_Y As Single   'greater y coordinate
Dim MAX_NODE_NUMBER As Long ' greatest node number allowed in emmebank
Dim X_SCREEN_STEP As Single  'distance to split screen in X NODE FINDER
Dim Y_SCREEN_STEP As Single  'distance to split screen in Y NODE FINDER
Dim DIM_X_POINTFINDER As Long 'size of XPointFinder Vector
Dim DIM_Y_POINTFINDER As Long 'size of YPointFinder Vector
Dim DIM_LINKLIST As Long 'size of LinkList Vectors
Dim NetUseLinkSearch As Boolean 'if there will be need to use square based on squares (spends 10 Megabytes of memory in 2 variables)


'Point_type is be used for a global vector named Point(), associated with
' a global vector named Link() of Link_type, and a global vector named Route() i.e., transit lines
' this point to each other to describe one transit net (as exported from EMME) to be handled
' thru functions that search across them.

'The functions that modify the network MUST update this vectors (letting the relations consistent)

' Transit lines are input into vector route of type Route_type
' Once route of route type is inserted, it fills the vectors link and point vectors too, while one change
' routes, must update those vectors... or do changes and then, run CountRoutesinNetwork, that does it for all routes
' (When we started to put this in classes, we shift to programming in haXe)

'points, links and routes are visible everywhere... so are the lists of them, so routines like
' - NPoints_within_radius
' - NLinks_in_shortest_path
' - NRoutes_in_pass_between

Public Npoints As Long
Public point() As Point_type

Public NLinks As Long
Public link() As Link_type

Public nRoutes As Integer
Public route() As Route_Type

'Lists to return selected elements
Public PointList() As Long
'Public OtherPointList() As Long
Public LinkList() As Long
Public RouteList() As Integer
'Public LinkListSPread() As Long 'para ShortestPath: Dijkstra não precisa

'PointNamed is a long vector that allows to find nodes by the name they have in EMME
' The following must be always true: point(PointNamed(i)).Name = i
Public PointNamed() As Long

Public marker As Long 'marker... always grows, it is public because other modules eventually want to mark points to their business


Type Point_type
    x As Double
    y As Double
    isM2 As Integer '0=is not EMME, 1=is centroid, 2=is not centroid
    t1 As Single 'ui1 in EMME
    t2 As Single 'ui2 in EMME
    t3 As Single 'ui3 in EMME
    Name As String '5 numbers = i in EMME (id)
    STname As String 'STation name = EMME label
    'Daqui is Portuguese to "FromHere", at some point will replace all...
    NLinksDaqui As Integer 'number of links from this node  (número de links que começam nesse nó)
    iLinkDaqui() As Long 'links starting in this node (os links que começam nesse nó)
    'PraKa is Portuguese to "ToHere" - at some point will replace all...
    NLinksPraKa As Integer
    iLinkPraKa() As Long 'os links que acabam no nó
    selected As Boolean
    nRoutes As Integer 'number of transit lines passing thru this node
    iroute() As Integer 'transit lines thru this node
    iRoutePosition() As Integer 'position this node is in the transit line
    iRouteDistance() As Single 'distance transit line run when reached this point
    NStopRoutes As Integer ' how many routes STOP on the node
    iStopRoute() As Integer 'number of i-th route to stop on the node
    iStopPosition() As Integer 'for the i-th route, how many stops done until stop on this node
    iStopDistance() As Single 'for the i-th route, how many km rode until stop on this node
    ' Total boardings and alightings for functions who use this
    auxBoard As Single
    auxAligth As Single
    ' to search network based on coordinates
    NextNodeX As Long 'next node in the ordered list (with X coordinate higher or equal this node) (nó com próxima coordenada X maior ou igual a deste nó)
    NextNodeY As Long 'node with Y coordinate higher or equal this node (nó com próxima coordenada Y maior ou igual a deste nó)
    NextNodeOnSameSquare As Long 'next and previous node that X and Y coordinates are in the same rectangle (SCREENDIVIDE)
    PreviousNodeOnSameSquare As Long ' (próximo nó e nó anterior com coordenada X e Y no mesmo quadrado (SCREENDIVIDE), para poder mover pontos e tirá-los da lista)
    marker As Single ' to mark a node, one must use the value of public var Mark and then increment Mark (so no need to reset as you know no other node has the same value... it should be a long int)
    
    ' to Integration Fare Routines
'    Isstation As String 'station to one mode... write the mode
    HasCone As Boolean
    SubPointBoard() As Long 'points to the cone
    SubPointAlight() As Long 'points to the cone
    
    ' for ShortestPath functions
    Distance As Single 'for function NLinksinShortestPath
    time As Single 'for function NLinksinShortestPath
    LastLink As Long 'for  NLinksinShortestPath
    
    Deleta As Boolean 'to let marked that this point once existed, but was deleted from the network
End Type

Type Link_type
    op As Long 'origin point
    dp As Long 'destiny point
    isM2 As Integer '0=is not, 1=is to centroid, 2=is not to centroid
    Extension As Single 'this is from EMME
    Length As Single 'copy of Extension divided by thousand(in km if other is in meters), when extension is changed to weight ShortestPaths
    time As Single 'to cross the link
    Auxiliar As Single
    modes As String
    Lanes As Integer
    tipo As Integer '= type in EMME (tipo is Portuguese for type)
    vdf As Integer
    t1 As Single ' ul1 in EMME
    t2 As Single ' ul2 in EMME
    t3 As Single ' ul3 in EMME
    
    timau As Single ' auto time from EMME
    volau As Single ' auto volume from EMME, usually input from report
    volad As Single ' adicional volume from EMME, usually input from report
    volax As Single ' auxiliar (pedestrian) volume from EMME
    voltr As Single ' transit volume from EMME (summ of voltr segments of lines thru this link)
    
    selected As Boolean ' for routine NLinksinshortestPathwithinSelected
    delete As Boolean '
    peda As Single 'toll (pedágio in Portuguese)
    marker As Long '
    tag As Integer
    
    nRoutes As Integer 'número de linhas que passam no link
    iroute() As Integer 'número das linhas que passam no link
    
    'dealing with shapes between nodes
    Ncosmetic As Integer 'Ncosmetic nodes in between
    Cosmetic() As Long  'cosmetic points
    
End Type

Type Route_Type
    number As String 'line in emme
    igroup As Integer 'used to group routes in different contexts
    group As String  'saved the name here as igroup eventually point to different grouping in different tools
    mode As String
    vehicle As Integer
    headway As Single
    speed As Single
    Name As String
    path As String
    lay As String
    t1 As Single ' ut1 in EMME
    t2 As Single ' ut2 in EMME
    t3 As Single ' ut3 in EMME
    Extension As Single
    
    marker As Long  'marker: just increases
    selected As Boolean 'one control
    checked As Boolean 'another control
    changed As Boolean 'mark if route had itinerary changes, to be wewritten
    deleted As Boolean 'mark that the route once existed, now is deleted
    
    Npoints As Integer ' nodes the line passes
    ipoint() As Long ' nodes list
    
    'Bellow is not loaded when reading EMME Network, need to call routine for doing it: Fill_Route_ilinks
    'Sub CountRoutesinNetwork only updates values for the Network
    'Sub RegisterRoutesinNetwork creates missing links, if missing
    NLinks As Integer 'one would expect this to be Npoints - 1, but eventually there were virtual nodes repeated
    ilink() As Long
    
    HasPara As Boolean 'Para means stop, true if there is the information bellow
    para() As String ' < board, > allight, + board and alight , # non-stop
    HasDwt As Boolean
    dwt() As Single  'dwt must mark stop time
    HasTtf As Boolean
    ttf() As Integer
    HasVoltr As Boolean
    voltr() As Single
    timetr() As Single
    linkvoltr() As Single 'total volume on the link, suposedly equal Link(Route(thisroute).ilink(thislink)).voltr
    linktimetr() As Single 'travel time on the link Link(Route(thisroute).ilink(thislink)).voltr
End Type

'This is used for search for links in areas... one link may cross a certain area of the screen, even there is no point on it.
Type LinkListperSquare_type
  SquareX As Integer
  SquareY As Integer
  ilink  As Long
  NextLinkOnTheSameSquare As Long
  PreviousLinkOnTheSameSquare As Long
End Type
Public NLinkFinders As Long
Public LinkFinder() As LinkListperSquare_type


'Public VehCap(30) As Integer

'Functions YFinder e XFinder, transform point coordenates of starting search point to the YXPointFinder e XPointFinder start position

'If FirstPointonSquare is implemented everywhere this can be removed
Dim XPointFinder() As Long
Dim YPointFinder() As Long
Dim UseLinkSearch() As Boolean

'More than 4 Mbyte for each bellow (1000 X 1000 4-byte), mostly empty: alone it is faster than going thru the borders,
' so tho be used when search for links is extense
Dim FirstLinkFinderOnSquare() As Long
Dim FirstPointOnSquare() As Long

' This routine is the first thing to call, set the size of some variables that
' are function of the network enviroment
Sub Load_Network_Parameters(Optional SheetName = "NET_PARAMETERS", Optional icol As Integer = 3, Optional iRow As Integer = 3)
Dim plan As Worksheet

'Reading set with optional parameters for named projects
Set plan = ThisWorkbook.Sheets(SheetName)
'icol = 3
'irow = 3

'Does the reading
NetIsLongLat = False
If Trim(Format(plan.Cells(iRow, icol), ">")) = "X" Then
    NetIsLongLat = True
    approx_one_meter_in_network = 1 / meter_per_degree
Else
    NetIsLongLat = False
    approx_one_meter_in_network = 1 / plan.Cells(iRow, icol)
End If
iRow = iRow + 2
MIN_X = plan.Cells(iRow, icol): iRow = iRow + 1
MAX_X = plan.Cells(iRow, icol): iRow = iRow + 1
MIN_Y = plan.Cells(iRow, icol): iRow = iRow + 1
MAX_Y = plan.Cells(iRow, icol): iRow = iRow + 1
X_SCREEN_STEP = plan.Cells(iRow, icol): iRow = iRow + 1
Y_SCREEN_STEP = plan.Cells(iRow, icol): iRow = iRow + 1
iRow = iRow + 1
MAX_NODE_NUMBER = plan.Cells(iRow, icol): iRow = iRow + 1
DIM_LINKLIST = plan.Cells(iRow, icol): iRow = iRow + 1
If Trim(Format(plan.Cells(iRow, icol), ">")) = "X" Then NetUseLinkSearch = True
    
'Sets the size of screeners
DIM_X_POINTFINDER = (MAX_X - MIN_X) / X_SCREEN_STEP
DIM_Y_POINTFINDER = (MAX_Y - MIN_Y) / Y_SCREEN_STEP
ReDim XPointFinder(DIM_X_POINTFINDER)
ReDim YPointFinder(DIM_Y_POINTFINDER)
ReDim FirstPointOnSquare(DIM_X_POINTFINDER, DIM_Y_POINTFINDER) As Long

If NetUseLinkSearch Then
    ReDim FirstLinkFinderOnSquare(DIM_X_POINTFINDER, DIM_Y_POINTFINDER) As Long
End If

ReDim PointNamed(MAX_NODE_NUMBER)
ReDim PointList(DIM_LINKLIST)
ReDim LinkList(DIM_LINKLIST)

End Sub

Sub ResetNetWork()
ResetNodes
ResetLinks
End Sub
Sub ResetNodes()
Dim i As Long, j As Long
For i = 0 To MAX_NODE_NUMBER
    PointNamed(i) = 0
Next i
For i = 0 To DIM_X_POINTFINDER
'    XPointFinder(i) = 0 '<--to be removed
    For j = 0 To DIM_Y_POINTFINDER
        FirstPointOnSquare(i, j) = 0
    Next j
Next i
'For i = 1 To DIM_Y_POINTFINDER  '<--to be removed when square take over
'    YPointFinder(i) = 0  '<--to be removed
'Next i  '<--to be removed
Npoints = 0
ReDim point(Npoints)
End Sub
Sub ResetLinks()
Dim i As Long
NLinks = 0
For i = 1 To Npoints
    point(i).NLinksDaqui = 0
    point(i).NLinksPraKa = 0
Next i
ReDim Links(NLinks)
If NetUseLinkSearch Then ResetLinkFinder
End Sub
Sub ResetLinkFinder()
Dim i As Long, j As Long
For i = 0 To DIM_X_POINTFINDER
    For j = 0 To DIM_Y_POINTFINDER
        FirstLinkFinderOnSquare(i, j) = 0
    Next j
Next i
NLinkFinders = 0
ReDim LinkFinder(NLinkFinders)
For i = 1 To NLinks
    If Not link(i).delete Then
        Add_Link_to_Link_Finder (i)
    End If
Next i
End Sub
Function XFinder(x As Double) As Long
' X é he coordinate
' XFinder returns the position in the sreen edge searcher for x
    XFinder = Int((x - MIN_X) / X_SCREEN_STEP)
    If XFinder < 0 Then XFinder = 0
    If XFinder > DIM_X_POINTFINDER Then XFinder = DIM_X_POINTFINDER
End Function
Function YFinder(y As Double) As Long
    YFinder = Int((y - MIN_Y) / Y_SCREEN_STEP)
    If YFinder < 0 Then YFinder = 0
    If YFinder > DIM_Y_POINTFINDER Then YFinder = DIM_Y_POINTFINDER
End Function
Sub UnSelect_all_points()
Dim i As Long
For i = 1 To Npoints
    point(i).selected = False
Next i
End Sub
Function Add_Point(ByVal Eastings As Double, ByVal Northings As Double, Optional isM2 As Integer = 0, Optional Name As String = "") As Long
    'Add_Point will return the number of the point added (=Npoints)
    'It also prepare insert the point in the finder
    Dim Chega As Boolean
    Dim Last_Node As Long
    Dim Next_Node As Long
    Add_Point = 0
    Npoints = Npoints + 1 'aumenta
    ReDim Preserve point(Npoints)
    point(Npoints).x = Eastings
    point(Npoints).y = Northings
    point(Npoints).isM2 = isM2
    point(Npoints).NLinksPraKa = 0
    point(Npoints).NLinksDaqui = 0
    point(Npoints).Name = Name
    point(Npoints).STname = Name
    If Name <> "" Then PointNamed(Name) = Npoints
    ReDim point(Npoints).iroute(0)
    Dim IntE, IntN As Integer
    IntE = XFinder(Eastings)
    IntN = YFinder(Northings)
    point(FirstPointOnSquare(IntE, IntN)).PreviousNodeOnSameSquare = Npoints
    point(Npoints).NextNodeOnSameSquare = FirstPointOnSquare(IntE, IntN)
    FirstPointOnSquare(IntE, IntN) = Npoints

    Add_Point = Npoints
End Function
Function Add_Link(dummylink As Link_type, Optional ResetRoutes As Boolean = True) As Long
    Dim op, dp As Long
    If dummylink.op = dummylink.dp Then
        Exit Function
    End If
    NLinks = NLinks + 1
    ReDim Preserve link(NLinks)
    link(NLinks) = dummylink
    'insert linkdaqui in origin point
    op = link(NLinks).op
    point(op).NLinksDaqui = point(op).NLinksDaqui + 1
    ReDim Preserve point(op).iLinkDaqui(point(op).NLinksDaqui)
    point(op).iLinkDaqui(point(op).NLinksDaqui) = NLinks
    'insert linkpraka in destiny point
    dp = link(NLinks).dp
    point(dp).NLinksPraKa = point(dp).NLinksPraKa + 1
    ReDim Preserve point(dp).iLinkPraKa(point(dp).NLinksPraKa)
    point(dp).iLinkPraKa(point(dp).NLinksPraKa) = NLinks
    Add_Link = NLinks
    If NetUseLinkSearch Then Add_Link_to_Link_Finder (NLinks)
End Function
'Remove_Link_FROM_here: remove ilink if it starts in ipoint, return true if found it and removed it
Function Tira_Link_Daqui(ipoint As Long, ilink As Long) As Boolean
    Tira_Link_Daqui = False
    For i = 1 To point(ipoint).NLinksDaqui
        If point(ipoint).iLinkDaqui(i) = ilink Then
            For j = i + 1 To point(ipoint).NLinksDaqui
                point(ipoint).iLinkDaqui(j - 1) = point(ipoint).iLinkDaqui(j)
            Next j
            point(ipoint).NLinksDaqui = point(ipoint).NLinksDaqui - 1
            Tira_Link_Daqui = True
            Exit Function
        End If
    Next i
End Function
'Remove_Link_TO_here: remove ilink if it ends in ipoint, return true if found it and removed it
Function Tira_Link_Praka(ipoint As Long, ilink As Long) As Boolean
    Tira_Link_Praka = False
    For i = 1 To point(ipoint).NLinksPraKa
        If point(ipoint).iLinkPraKa(i) = ilink Then
            For j = i + 1 To point(ipoint).NLinksPraKa
                point(ipoint).iLinkPraKa(j - 1) = point(ipoint).iLinkPraKa(j)
            Next j
            point(ipoint).NLinksPraKa = point(ipoint).NLinksPraKa - 1
            Tira_Link_Praka = True
            Exit Function
        End If
    Next i
End Function
' To remove link from ipoint, don´t care if it ends or starts on the given ipoint
Function Tira_Link(ipoint As Long, ilink As Long) As Boolean
    'if is a circular link this function will remove link twice, still works (one fail, other passes)
    Tira_Link = Tira_Link_Daqui(ipoint, ilink) Or Tira_Link_Praka(ipoint, ilink)
End Function
' To remove link without caring where is it from or to, but keep consistency
Function Delete_Link(il As Long) As Boolean
    Delete_Link = Tira_Link_Daqui(link(il).op, il) And Tira_Link_Praka(link(il).dp, il)
    link(il).delete = Delete_Link
    link(il).isM2 = 0
End Function
'DEGrees to RADians
Function degtorad(deg)
    degtorad = (deg / 180) * PI
End Function

'Return the distance in meters between P1 and P2 as Point_type
Function Point_Distance(P1 As Point_type, P2 As Point_type) As Single
    'it does not use precise earth curve for coordinated systems
    'CONSIDERING ONE Degree both in latitude and longitude we have
    ' difference for Reykjavik: (bellow 0,2% in large distances considering flattening)
    'acos:                    121,372.8
    'this formula:            121,376.3
    'considering flattening:  121,573.1
    'diference for São Paulo: (bellow 0,3% in large distances considering flattening)
    'acos:                    151,684.0
    'this formula:            151.685,5
    'considering flattening:  151.288,1
    'diference for Jakarta or Bogotá: (bellow 0,5% in large distances considering flattening)
    'acos:                    157.281,8
    'this formula:            157.282,8
    'considering flattening:  156.759,1
    'This difference comes down to 20 centimeters in 160 meters, near 0.12%
    
    DeltY = P1.y - P2.y
    DeltX = P1.x - P2.x
    If NetIsLongLat Then
        DeltY = DeltY * meter_per_degree
        Ymed = (P1.y + P2.y) / 2
        DeltX = DeltX * meter_per_degree / 360 * Cos(PI * Ymed / 180)
        'Acos formula: R1 * Acs(sin(degtorad(lat2)) * sin(degtorad(lat1)) + Cos(degtorad(lat2)) * Cos(degtorad(lat1)) * Cos(degtorad(long1 - long2)))
    End If
    Point_Distance = (DeltX ^ 2 + DeltY ^ 2) ^ 0.5
End Function
'Return the distance in meters between point(iponto) and point(jponto)
Function Distância(iponto As Long, jponto As Long)
    Distância = Point_Distance(point(iponto), point(jponto))
End Function
Function Distance(iponto As Long, jponto As Long)
    Distance = Point_Distance(point(iponto), point(jponto))
End Function
Function Point_Diference(P1 As Point_type, P2 As Point_type) As Point_type
Point_Diference.x = P1.x - P2.x
Point_Diference.y = P1.y - P2.y
End Function
Function Point_Summ(P1 As Point_type, P2 As Point_type) As Point_type
Point_Summ.x = P1.x + P2.x
Point_Summ.y = P1.y + P2.y
End Function

'Functions to look for points and links

Function NPointsListed(ByVal XLow As Double, ByVal YLow As Double, Optional ByVal XHigh As Double = -1, Optional ByVal YHigh As Double = -1, Optional Startfrom As Long) As Long
' Could be named NPoints listed on a square
' Returns to PointList() all points in the rectangle: x=XLow to x=XHigh,y=YLow to y=YHigh (obviously in the same unit the network)
' and returns how many there are
' It will be larger up to X_STEP and Y_SCREEN_STEP
' If no XHigh or YHigh are provided (=-1) gets all in the same squarenão são fornecidos (=-1), so o valor em questão interessa
' -it can add in NPointsListed, after StartFrom: such an extended list may have duplicates even if the the interval does not overlap, due to square sizes
    Dim mem As Double
    NPointsListed = Startfrom
    If XHigh = -1 Then XHigh = XLow
    If YHigh = -1 Then YHigh = YLow
    If XLow > XHigh Then 'coloca X em ordem
        mem = XHigh
        XHigh = XLow
        XLow = mem
    End If
    If YLow > YHigh Then 'coloca Y em ordem
        mem = YHigh
        YHigh = YLow
        YLow = mem
    End If
    Dim i As Integer, j As Integer
    Dim NextNode As Long
    For i = XFinder(XLow) To XFinder(XHigh)
        For j = YFinder(YLow) To YFinder(YHigh)
            NextNode = FirstPointOnSquare(i, j)
            Do While NextNode <> 0
                If Not point(NextNode).Deleta Then
                    'does not include deleted on the list
                    NPointsListed = NPointsListed + 1
                    PointList(NPointsListed) = NextNode
                End If
                NextNode = point(NextNode).NextNodeOnSameSquare
            Loop
        Next j
    Next i
End Function
Function NPointsListedInARadius(center As Point_type, ByVal radius_in_meters As Single, Startfrom As Long) As Long
    Dim XLow As Double
    Dim XHigh As Double
    Dim YLow As Double
    Dim YHigh As Double
    Dim radius As Single
    
    radius = radius_in_meters
    If NetIsLongLat Then
        radius = radius_in_meters / meter_per_degree
    End If
    
    NPointsListedInARadius = Startfrom
    
    XLow = center.x - radius
    XHigh = center.x + radius
    YLow = center.y - radius
    YHigh = center.y + radius
    Dim i As Integer
    For i = 1 To NPointsListed(XLow, YLow, XHigh, YHigh)
        dista = Point_Distance(center, point(PointList(i)), isLatLon)
        If dista <= radius Then
            NPointsListedInARadius = NPointsListedInARadius + 1
            PointList(NPointsListedInARadius) = PointList(i)
        End If
    Next i
End Function
Function NLinksListedonSquare(Eastings As Double, Northings As Double)
Dim IntE, IntN As Integer
    IntE = XFinder(Eastings)
    IntN = YFinder(Northings)
    NextOnSame = FirstLinkFinderOnSquare(IntE, IntN)
    Do While NextOnSame <> 0
        If Not link(LinkFinder(NextOnSame).ilink).delete Then
            NLinksListedonSquare = NLinksListedonSquare + 1
            LinkList(NLinksListedonSquare) = LinkFinder(NextOnSame).ilink
        End If
        NextOnSame = LinkFinder(NextOnSame).NextLinkOnTheSameSquare
    Loop
End Function
Function NLinksinARange(Eastings As Double, Northings As Double, Optional radius_in_meters As Single = -1, _
                        Optional OnlyAmongSelected As Boolean = False, _
                        Optional Startfrom As Long = 0) As Integer
Dim XLow As Double
Dim XHigh As Double
Dim YLow As Double
Dim YHigh As Double
If Range = -1 Then
   xradius = X_SCREEN_STEP
   yradius = X_SCREEN_STEP
ElseIf NetIsLongLat Then
    yradius = radius_in_meters / meter_per_degree
    xradius = radius_in_meters / meter_per_degree * Cos(Abs(degtorad(Northings)))
End If

NLinksinARange = Startfrom
marker = marker + 1
XLow = Eastings - xradius
XHigh = Eastings + xradius
YLow = Northings - yradius
YHigh = Northings + yradius
IYLow = YFinder(YLow)
IYHigh = YFinder(YHigh)
Dim IntE, IntN As Integer
For IntE = XFinder(XLow) To XFinder(XHigh)
    For IntN = IYLow To IYHigh
        NextOnSame = FirstLinkFinderOnSquare(IntE, IntN)
        Do While NextOnSame <> 0
            If Not link(LinkFinder(NextOnSame).ilink).delete Then
                If Not OnlyAmongSelected Or link(LinkFinder(NextOnSame).ilink).selected Then
                    If link(LinkFinder(NextOnSame).ilink).marker <> marker Then
                        NLinksinARange = NLinksinARange + 1
                        LinkList(NLinksinARange) = LinkFinder(NextOnSame).ilink
                    End If
                End If
            End If
            NextOnSame = LinkFinder(NextOnSame).NextLinkOnTheSameSquare
        Loop
    Next IntN
Next IntE
End Function
Function Add_Link_to_Link_Finder(il As Long)
Dim Xini As Double, Xfim As Double, Yini As Double, Yfim As Double
Xini = point(link(il).op).x
Xfim = point(link(il).dp).x
If Xini < Xfim Then 'points east
    Yini = point(link(il).op).y
    Yfim = point(link(il).dp).y
Else 'point west: change Xini<-->Xfim, so we do it west-->east
    Xini = point(link(il).dp).x
    Xfim = point(link(il).op).x
    Yini = point(link(il).dp).y
    Yfim = point(link(il).op).y
End If

Dim inclina, ystep, DeltaX, DeltaY As Single
DeltaX = Xfim - Xini 'non-negative
DeltaY = Yfim - Yini
If DeltaX = 0 Then DeltaX = X_SCREEN_STEP / 1000 'it is vertical, we need a small inclination for the bellow without doing ifs
inclina = DeltaY / DeltaX
ystep = 1
If inclina < 0 Then ystep = -1 'axe northwest/southeast, we will move to squares bellow

Dim first As Long
Dim LastX As Double, NowX As Double, LastY As Double, NowY As Double
Dim IYLast As Long, IYNow As Long
Dim i As Integer, j As Integer
NowX = Xini
NowY = Yini
Do
    LastX = NowX
    LastY = NowY
    NowX = LastX + X_SCREEN_STEP
    If NowX > Xfim Then NowX = Xfim
    NowY = LastY + (NowX - LastX) * inclina
    IYLast = YFinder(LastY)
    IYNow = YFinder(NowY)
    For i = XFinder(LastX) To XFinder(NowX)  ' XFinder(LastX) - 1 To XFinder(NowX) + 1
        For j = IYLast To IYNow Step ystep
            first = FirstLinkFinderOnSquare(i, j)
            NLinkFinders = NLinkFinders + 1
            ReDim Preserve LinkFinder(NLinkFinders)
            LinkFinder(NLinkFinders).SquareX = i
            LinkFinder(NLinkFinders).SquareY = j
            LinkFinder(NLinkFinders).ilink = il
            LinkFinder(NLinkFinders).NextLinkOnTheSameSquare = first
            LinkFinder(first).PreviousLinkOnTheSameSquare = NLinkFinders
            FirstLinkFinderOnSquare(i, j) = NLinkFinders
            'LinkFinder(NLinkFinders).PreviousLinkOnTheSameSquare = 0
        Next j
     Next i
Loop While NowX < Xfim

End Function


' Return the node number for given coordinates, if exists...
' If the node does not exist, it will be created unless CreatePoint is passed as false
'
' Optional TagRange: when loading a second network uppon the first (to join them)
'     - consider nodes within a range distance to return
'     - when creating point, create it aligned with the closest link in range, spliting that link
Function Get_Point(Eastings As Double, Northings As Double, _
                    Optional CreatePoint As Boolean = True, _
                    Optional isM2 As Integer = 0, _
                    Optional Name As String = "", _
                    Optional TagRangeinMeters As Single = 0) As Long
    
    Dim linkdis As Single
    Dim TouchOID As Integer
    Dim pon As Point_type
    Dim xis As Point_type ' it is for return
    Dim i As Long, il As Long
    Dim disting As Single
    
    Get_Point = 0
    For i = 1 To NPointsListed(Eastings, Northings)
        If point(PointList(i)).x = Eastings And point(PointList(i)).y = Northings Then
            Get_Point = PointList(i)
            Exit Function
        End If
    Next i
    
'Exact point not found yet
    If TagRangeinMeters > 0 Then     ' search for nearby points as requested
        Dim dop, ddp, DD As Single
        linkdis = TagRangeinMeters * 5  '5 times provide a r
        il = Get_Closest_Link_in_All(Eastings, Northings, xis, linkdis, TagRangeinMeters, TouchOID)
        If il > 0 And linkdis <= TagRangeinMeters Then
            If TouchOID = -1 Then Get_Point = link(il).op: Exit Function
            If TouchOID = 1 Then Get_Point = link(il).dp: Exit Function
            If TouchOID = 0 Then
                pon.x = Eastings
                pon.y = Northings
                dop = Point_Distance(xis, point(link(il).op))
                ddp = Point_Distance(xis, point(link(il).dp))
                If dop < ddp Then
                    DD = dop
                    Get_Point = link(il).op
                Else
                    DD = ddp
                    Get_Point = link(il).dp
                End If
                If DD < TagRangeinMeters Then
                    Exit Function
                Else
                    If CreatePoint Then Get_Point = Add_Point_Spliting_Link(il, Eastings, Northings, isM2, Name, True)
                    Exit Function
                End If
            End If
        End If
    ElseIf CreatePoint Then
        Get_Point = Add_Point(Eastings, Northings, isM2, Name)
    End If
    
End Function
'Delete ilink, and add two new links in place
Function Add_Point_Spliting_Link(ilink As Long, CandiEastings As Double, CandiNorthings As Double, _
                            Optional isM2 As Integer = 0, Optional Name As String = "", _
                            Optional UpdateRoutes As Boolean = True) As Long
'up to this version copies t1, t2 and t3 from original links
Dim CandiP As Point_type
Dim Dpoint As Point_type
Dim Dlink As Link_type
Dim jlink As Long
Dim ip As Long
Dim Ang As Single 'for angle
Dim dista As Single
CandiP.x = CandiEastings
CandiP.y = CandiNorthings
jlink = Get_Link(link(ilink).dp, link(ilink).op)

If ilink = 0 And jlink = 0 Then
    MsgBox "Can't split link that I don't know"
Else
    D = ObPoint_to_ObLink_Distance(CandiP, link(ilink), Dpoint, isLatLon)
    ip = Add_Point(Dpoint.x, Dpoint.y, isM2, Name)
    'prop = proportion of distance till insert point over link distance
    prop = Point_Distance(point(link(ilink).op), point(ip)) / Point_Distance(point(link(ilink).op), point(link(ilink).dp), isLatLon)
    If prop >= 1 Then MsgBox "HELLo": prop = 0.99
    If prop = 0 Then MsgBox "HELLo": prop = 0.01
    If ilink <> 0 Then
        Dlink = link(ilink)
        Dlink.dp = ip
        Dlink.Extension = link(ilink).Extension * prop
        ilink1 = Add_Link(Dlink)
        Dlink.op = ip
        Dlink.dp = link(ilink).dp
        Dlink.Extension = link(ilink).Extension * (1 - prop)
        ilink2 = Add_Link(Dlink)
        If Not (Tira_Link_Daqui(link(ilink).op, ilink) And Tira_Link_Praka(link(ilink).dp, ilink)) Then MsgBox "Não tirou link"
        link(ilink).tipo = 1000
        For W = 1 To link(ilink).nRoutes
            point(ip).nRoutes = point(ip).nRoutes + 1
            ReDim Preserve point(ip).iroute(point(ip).nRoutes)
            point(ip).iroute(point(ip).nRoutes) = link(ilink).iroute(W)
        Next W
    End If
    If jlink <> 0 Then
        Dlink = link(jlink)
        Dlink.dp = ip
        Dlink.Extension = link(ilink).Extension * (1 - prop)
        ilink3 = Add_Link(Dlink)
        Dlink.op = ip
        Dlink.dp = link(jlink).dp
        Dlink.Extension = link(ilink).Extension * prop
        ilink4 = Add_Link(Dlink)
        If Not (Tira_Link_Daqui(link(jlink).op, jlink) And Tira_Link_Praka(link(jlink).dp, jlink)) Then MsgBox "Remotion failed when splitting link"
        link(jlink).tipo = 1000
        For W = 1 To link(jlink).nRoutes
            point(ip).nRoutes = point(ip).nRoutes + 1
            ReDim Preserve point(ip).iroute(point(ip).nRoutes)
            point(ip).iroute(point(ip).nRoutes) = link(jlink).iroute(W)
        Next W
    End If
    link(ilink).isM2 = 0
    link(ilink).delete = True
    link(jlink).isM2 = 0
    Add_Point_Spliting_Link = ip
    link(jlink).delete = True

    If UpdateRoutes Then
        For iroute = 1 To point(link(ilink).op).nRoutes
            For jroute = 1 To point(link(ilink).dp).nRoutes
                If point(link(ilink).op).iroute(iroute) = point(link(ilink).dp).iroute(jroute) Then
                    Call Update_Route_in_SplitLink(point(link(ilink).op).iroute(iroute), link(ilink).op, link(ilink).dp, ip)
                End If
            Next jroute
        Next iroute
    End If
End If
Add_Point_Spliting_Link = ip
End Function

Function Get_Closest_Link_in_All(Eastings As Double, Northings As Double, XisPoint As Point_type, _
                                 Optional bestdist As Single = 100, _
                                 Optional RangeinMeters = 100, _
                                 Optional Xcontact As Integer = 0) As Long
Dim DummyPoint1 As Point_type
Dim DummyPoint2 As Point_type
Dim Extreme As Integer
Dim RangeInCord As Single
RangeInCord = RangeinMeters
If NetIsLongLat Then
    RangeInCord = RangeinMeters / meter_per_degree
End If
DummyPoint1.x = Eastings
DummyPoint1.y = Northings
N = NLinksinARange(Eastings, Northings, RangeInCord)
For klisted = 1 To N
    i = LinkList(klisted)
    linkdis = ObPoint_to_ObLink_Distance(DummyPoint1, link(i), DummyPoint2, True, Extreme)
    If linkdis < bestdist Then
        bestdist = linkdis
        Get_Closest_Link_in_All = i
        Xcontact = Extreme
        XisPoint = DummyPoint2
    End If
Next klisted

End Function
Function Get_Link(po As Long, PD As Long) As Long
Dim i As Integer
    For i = 1 To point(po).NLinksDaqui
        If link(point(po).iLinkDaqui(i)).dp = PD Then
            Get_Link = point(po).iLinkDaqui(i)
            Exit Function
        End If
    Next i
End Function
Function GetAngletoSplitPoint(ipoint As Long) As Single
Dim i As Integer
Dim GreaterAngle As Single
ReDim NumToSort(200)
For i = 1 To point(ipoint).NLinksDaqui
    NumToSort(i) = Angle(point(ipoint), point(link(point(ipoint).iLinkDaqui(i)).dp))
Next i
For i = 1 To point(ipoint).NLinksPraKa
    NumToSort(i + point(ipoint).NLinksDaqui) = Angle(point(ipoint), point(link(point(ipoint).iLinkPraKa(i)).op))
Next i
NtoSort = point(ipoint).NLinksPraKa + point(ipoint).NLinksDaqui
NumHeapSort
GreaterAngle = NumToSort(1) + (360 - NumToSort(NtoSort))
GetAngletoSplitPoint = NumToSort(NtoSort) + GreaterAngle / 2
If GetAngletoSplitPoint > 360 Then GetAngletoSplitPoint = GetAngletoSplitPoint - 360
For i = 2 To NtoSort
    If NumToSort(i) - NumToSort(i - 1) > GreaterAngle Then
        GreaterAngle = NumToSort(i) - NumToSort(i - 1)
        GetAngletoSplitPoint = NumToSort(i - 1) + GreaterAngle / 2
    End If
Next i
End Function
Function GetRestrictedAngletoSplitPoint(ipoint As Long, angfrom As Single, angto As Single) As Single
ReDim NumToSort(100)
Dim Ang As Single
nn = 0
For i = 1 To point(ipoint).NLinksDaqui
    Ang = Angle(point(ipoint), point(link(point(ipoint).iLinkDaqui(i)).dp))
    If IsAngBetween(Ang, angfrom, angto) And point(link(point(ipoint).iLinkDaqui(i)).dp).Name < 10000 Then
        nn = nn + 1
        NumToSort(nn) = Ang
    End If
Next i
For i = 1 To point(ipoint).NLinksPraKa
    Ang = Angle(point(ipoint), point(link(point(ipoint).iLinkPraKa(i)).op)) And point(link(point(ipoint).iLinkPraKa(i)).op).Name < 10000
    If IsAngBetween(Ang, angfrom, angto) Then
        nn = nn + 1
        NumToSort(nn) = Ang
    End If
    NumToSort(i + point(ipoint).NLinksDaqui) = Angle(point(ipoint), point(link(point(ipoint).iLinkPraKa(i)).op))
Next i
nn = nn + 1
NumToSort(nn) = angfrom
nn = nn + 1
NumToSort(nn) = angto
NtoSort = nn 'public parameter to NumHeapSort
NumHeapSort
NumToSort(0) = NumToSort(NtoSort) - 360
GreaterAngle = 0
For i = 1 To NtoSort
    If NumToSort(i) - NumToSort(i - 1) > GreaterAngle And NumToSort(i) <> angfrom Then
        GreaterAngle = NumToSort(i) - NumToSort(i - 1)
        GetRestrictedAngletoSplitPoint = NumToSort(i - 1) + GreaterAngle / 2
    End If
Next i
End Function
Function IsAngBetween(Ang As Single, angfrom As Single, angto As Single) As Boolean
IsAngBetween = False
If angfrom >= angto Then
    If Ang > angfrom Or Ang < angto Then IsAngBetween = True
Else
    If Ang < angto And Ang > angfrom Then IsAngBetween = True
End If
End Function
Function PointInsertPosition(BasePoint As Point_type, CentralAngle As Single, RelX As Integer, RelY As Integer, ModDist_in_network_unit As Single) As Point_type
'This function returns the point in position for the insertion of a new point
'Central Angle is the direction the tree of points is being mounted in degrees
'RelX and RelY are the relative position in the tree, ModDist is the space between points (- to the right,+ to the left)
PointInsertPosition.x = BasePoint.x + RelY * ModDist_in_network_unit * Cos(CentralAngle * PI / 180) + RelX * ModDist_in_network_unit * Sin(CentralAngle * PI / 180) / 2
PointInsertPosition.y = BasePoint.y + RelY * ModDist_in_network_unit * Sin(CentralAngle * PI / 180) - RelX * ModDist_in_network_unit * Cos(CentralAngle * PI / 180) / 2
End Function
Function Angle(pos1 As Point_type, pos2 As Point_type) As Single
Dim distx As Double, disty As Double
distx = pos2.x - pos1.x
disty = pos2.y - pos1.y
Angle = 0
If distx = 0 And disty = 0 Then Angle = 10000: Exit Function
If distx = 0 And disty < 0 Then Angle = 270: Exit Function
If distx = 0 And disty > 0 Then Angle = 90: Exit Function
If distx > 0 And distx = 0 Then Angle = 0: Exit Function
If distx < 0 And disty = 0 Then Angle = 180: Exit Function
If distx > 0 And disty > 0 Then Angle = (Atn(disty / distx)) * 180 / PI: Exit Function
If distx > 0 And disty < 0 Then Angle = 360 + ((Atn(disty / distx)) * 180 / 3.14159265): Exit Function
If distx < 0 Then Angle = 180 + ((Atn(disty / distx)) * 180 / PI): Exit Function
End Function
Function DifAngle(angle1 As Single, angle2 As Single) As Single
If angle1 = 10000 Or angle2 = 10000 Then difang = 10000: Exit Function
maior = angle1
menor = angle2
If angle2 > angle1 Then
    maior = angle2
    menor = angle1
End If
DifAngle = maior - menor
If DifAngle > 180 Then DifAngle = 360 - DifAngle
End Function
Function DifAngleOrder(maior As Single, menor As Single) As Single
If maior = 10000 Or menor = 10000 Then DifAngleOrder = 10000: Exit Function
If maior < menor Then maior = maior + 360
DifAngleOrder = maior - menor
End Function
