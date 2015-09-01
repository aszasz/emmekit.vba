Attribute VB_Name = "Mod_A1_Test"
Sub UnitTestAngle()
Load_Network_Parameters
ResetNetWork
Dim p1 As Long, p2 As Long
p1 = Add_Point(10, 10)
p2 = Add_Point(5, 10)
p3 = Add_Point(5, 5)
Debug.Assert Angle(point(p1), point(p2)) = 180
Debug.Assert Angle(point(p3), point(p1)) = 45
Debug.Assert Angle(point(p1), point(p3)) = 225
End Sub

Sub UnitTestGetCentralAngle()
Load_Network_Parameters
ResetNetWork
Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
p1 = Add_Point(10, 10)
p2 = Add_Point(5, 10)
p3 = Add_Point(10, 5)
p4 = Add_Point(0, 0)
Dim Dummy As Link_type
Dummy.op = p1: Dummy.dp = p2
L1 = Add_Link(Dummy)
Dummy.op = p1: Dummy.dp = p3
L2 = Add_Link(Dummy)
Dummy.op = p4: Dummy.dp = p1
L3 = Add_Link(Dummy)

Debug.Assert GetAngletoSplitPoint(p1) = 45

p4 = Add_Point(0, 0)
Dummy.op = p4: Dummy.dp = p2
L3 = Add_Link(Dummy)
Dummy.op = p4: Dummy.dp = p3
L4 = Add_Link(Dummy)
Debug.Assert GetAngletoSplitPoint(p4) = 225



End Sub
Sub UnitTestInsertPosition()
Load_Network_Parameters
ResetNetWork

Dim p1 As Long, p2 As Point_type
p1 = Add_Point(100, 1000)
p2 = PointInsertPosition(point(p1), 60, 1, 1, 25 * approx_one_meter_in_network)
Debug.Assert p2.x = 112.5
Debug.Print p2.x
End Sub



'''''''''''''''Removed from type_link in Basic Network
'tipo in link    tipo As String ' - para ui3
'   1 "centroide"
'   2 "só auto"
'   3 "ônibus "
'   4 "barco  "
'   5 "trem"
'   6 "metro "
'   7 "BRT"
'   81 "acesso a seletiva sul"
'   -81 "egresso a seletiva sul"
'   82 "acesso da seletiva norte"
'   -82 "egresso da seletiva norte"
'   104 "estação de barco"
'   105 "estação de trem"
'   106 "estação de metrô"
'   107 "estação de BRT"
'    detour As Integer 'para entra e sair BRasil
'    radius As Single 'comprimento do maior conector
'    SubPointBoard(20) As Long
'    SubPointAlight(20) As Long
'    SubPointCount As Long
'    nroutegrupos As Integer 'quantidade de tipo de linhas que passam no nó
'    iroutegrupos() As Integer 'os tipos de linha que passam no nó
'    Freq(-2 To 0) As Single
'    NRou(-2 To 0) As Single


'   NIntegra As Integer  'número de linhas que precisam integrar no ponto... inicialmente para revisão dos corredores
'   Corredor As Long 'número do corredor a que pertence, alternativamente
                      'em MarcaCorredoresPointsandLinks 0 se não pertence a corredor,1 se pertence
'    DistânciadoCorredor As Integer 'abaixo de 600 metros... a distância de um ponto de corredor
'    DistânciadaRedeBRT As Integer 'distância da rede link tipo 8, corredor de ônibus em BH
'    DistânciadaRedeMetrô As Integer 'distância da rede link tipo 10, metrô em BH
'    NewP As Integer 'aponta para o novo ponto no corredor
'    Estação As Long '1=parada, 0=não para
'    DistânciadaRede As Single 'distância do ponto até algum corredor
'    DistânciadaEstação As Single 'distância do ponto no corredor até a estação
'    closer As Long 'ponto no caminho da estação, para guiar linhas
'    Toexpand As Boolean
'    bikelanedistance As Single 'para caminho mínimo NLinksinShortestPath
'    Limite As Integer 'para indicar a que perímetro pertence
'    IsAutoNetwork As Boolean 'para subir pesqvel
'    MinVDF As Integer ' para subir pesqvel
'    PesqVel As Long

'    'Specific to Trunk-feeder operations
'    Corredor As Integer '(corridor number, to build trunk lines, or move lines into corridor, 0 means not part of a corridor)
'    Troncal As Single ' (trunk: used to count trunk lines in the corridor)
'    Limite As Integer 'para indicar a que perímetro pertence
'    VelPesq As Single
'    HoraVelPesq As Long
'    ' bands to save mileage in a given distance from the corridor.
'    'kilometragem que o link percorre en cada faixa de distância do corredor
'    A As Single
'    B As Single
'    C As Single
'    D As Single
'    E As Single
'
'
'    'Specific to process Boarding and alighting Data from APC (Automated Passenger Counts) in Maryland
'    Ntrips As Integer
'    itrip() As Long 'position in trip table... can trace route and group
'    hourtrip() As Single 'decimal hour
'    speed() As Single 'estimated speed for trip
'    Weight() As Single 'to average speeds... shall be distance used to calculate speed
'    pax() As Single  'pax in link... resulting from add APC and calculate after it
'    NAPCtrips As Integer
'    APCitrip() As Long 'position in trip table... can trace route and group
'    APChourtrip() As Single 'decimal hour
'    APCSpeed() As Single 'estimated speed for trip
'    APCSpeedWeight() As Single 'to average speeds... shall be distance used to calculate speed
'    APCPax() As Single 'from APC when have it
'    Traveltime(16 To 96) As Single 'average time to cross the link, by quarter hour of the day
'    centro As Boolean 'é link da área central
'    'Pesquisas
'    NFreqs As Integer
'    PesqGrupo() As String
'    PesqFreq() As Single

''''''''''''''Removed from type_route
'    'M bellow is short for EMME: when reading summary transit report
'    MVehicle As Single
'    MVehCapacity As Single
'    MLength As Single
'    MFleet As Single
'    MTime As Single
'    MBoard As Single
'    MPaxHour As Single
'    MPaxKm As Single
'    MAvgLoad As Single
'    MMaxLoad As Single
'    MMaxVol As Single
'
'
'    linkcor() As Single ' ?
'
'    'this for MD project
'    Ndepartures As Integer
'    Departure() As Long
'    header As Integer
'    direction As Boolean 'True if headed to metro (first) then if headed to center
'    AMfleet As Single
'    PMfleet As Single
'    ExtensionInCorridor As Single
'    ReturnRoute As Integer 'pair return transit line, when exists
'    NewHeadway As Single
'    Troncaliza As Integer 'para função troncaliza
'    TroncalizaPOSA As Integer 'para função troncaliza:define se é ida ou volta (antes ou depois)
'    TroncalizaPOSA2 As Integer 'para função troncaliza:define se é ida ou volta (antes ou depois)
'    TroncalizaPOSB As Integer 'para função troncaliza
'
'
'    marker2 As Long  'marker
'    NExpand As Integer ' number of points with integration (+3 segments)
'    followroute() As Integer 'numero da linha a ser seguido entre o nó anterior e o próximo... criado para T5
'    DistAntes As Single
'    DistDurante As Single
'    DistDepois As Single
'    KmAntes As Single
'    KmDurante As Single
'    KmForaDurante As Single
'    KmDepois As Single
'    FirstStop As Integer
'    LastStop As Integer
'    ToPrint As Integer
'    TRANSCADID As String
'    TRANSCADNAME As String
'    ' to process surveys and find headway
'    NFreqs As Integer
'    Freq() As Single
'
'    NFOVs As Integer
'    ifovposition() As Integer
'    ifovhora() As Single
'    ifovfreq() As Single
'    ifovocup() As Single
'

'
'
'Type Route_Groups
'    id As Integer
'    Name As String
'    LongName As String
'    NHeaders As Integer
'    header() As String
'    HeadertoCenter() As String
'    nRoutes As Integer
'    iroute() As Integer
'    irouteHeader() As Integer
'    Ntrips As Integer
'    itrip() As Long
'    APCdTrips As Integer
'    sum_on As Single
'    sum_off As Single
'    pax_km As Single
'    pax_hour As Single
'    veh_km As Single
'    veh_hour As Single
'    sum_on_IN As Single
'    sum_off_IN As Single
'    sum_on_IN_times_remaining_km As Single
'    sum_off_IN_times_run_km As Single
'    pax_km_IN As Single
'    pax_hour_IN As Single
'    veh_km_IN As Single
'    veh_hour_IN As Single
'End Type
'Public NRouteGroups As Integer
'Public RouteGroup() As Route_Groups
'
'Type Trip_table
'    iroute As Integer
'    ideparture As Integer
'    trip_index As Long
'    operadia As Integer
'    Nstation As Integer
'    departure_time As Single
'    seq() As Single 'stopsequence, before processing
'    stop_point() As Single 'the stop/zone, not inroute
'    arrival_time() As Single 'in hours, before processing
'    dist_traveled() As Single 'as declared, before processing
'                              'later processed
'    time_traveled() As Single 'processed, in minutes after start
''    reach() As Single 'time to reach station from previous
'    sampleAPC As Integer
'    APC_arrival_time() As Single 'in hours, before processing
'    APC_departure_time() As Single 'in hours, before processing
'    pax_on() As Single
'    pax_off() As Single
'    pax_load() As Single
'End Type
'Public Ntrips As Long
'Public trip() As Trip_table


'Public CorreRate As Single
'Public Contapontosexpandidos As Integer


'for metro fare integration
'Public NEstaçõeS As Integer
'Public Estação(30) As Long 'stations
'Public EstaçãoIntegra(30, 3) As Boolean 'true if station integrates with route of group j


'Public PontodoCorredor(20, 150) As Long
'Public NovosPontos(2000) As Point_type
'Public NNovosPontos As Integer
'Public MetroLine(9) As Long

'RIO
'Public Const MINX As Single = -44.38 'menor coordinada X
'Public Const MAXX As Single = -42.35  'maior coordinada X
'Public Const MINY As Single = -23.16 'menor coordenada Y
'Public Const MAXY As Single = -22.2 'maior coordenada Y
'Public Const MAXNUMNOMES As Long = 1000000 'maior número de nó
'Public Const XSCREENYDIVIDE As Single = 0.001 'distância para dividir a tela em X para NODE FINDER
'Public Const YSCREENYDIVIDE As Single = 0.001 'distância para dividir a tela em Y para NODE FINDER
'MontCO
'Ride ON BoundingBox -77.4191,38.9357,-76.9368,39.2884
'Public Const MINX As Single = -78.4 'menor coordinada X
'Public Const MAXX As Single = -76.4  'maior coordinada X
'Public Const MINY As Single = 38.5 'menor coordenada Y
'Public Const MAXY As Single = 39.6 'maior coordenada Y
'Public Const MAXNUMNOMES As Long = 1000000 'maior número de nó
'Public Const XSCREENYDIVIDE As Single = 0.001 'distância para dividir a tela em X para NODE FINDER
'Public Const YSCREENYDIVIDE As Single = 0.001 'distância para dividir a tela em Y para NODE FINDER
'Public Const DIM_XPOINTFINDER As Long = (MAXX - MINX) / XSCREENYDIVIDE
'Public Const DIM_YPOINTFINDER As Long = (MAXY - MINY) / YSCREENYDIVIDE
'Teste inicial, deveria armazenar em uma hashtable, mas aqui cabe na memória (1100 x 2000)

'Sub Add_Freq_to_Route(iroute As Integer, Freq As Single)
'    route(iroute).NFreqs = route(iroute).NFreqs + 1
'    ReDim Preserve route(iroute).Freq(route(iroute).NFreqs)
'    route(iroute).Freq(route(iroute).NFreqs) = Freq
'End Sub
'Sub Add_Freq_to_Link(ilink As Long, nome As String, Freq As Single)
'    link(ilink).NFreqs = link(ilink).NFreqs + 1
'    ReDim Preserve link(ilink).PesqGrupo(link(ilink).NFreqs)
'    ReDim Preserve link(ilink).PesqFreq(link(ilink).NFreqs)
'    link(ilink).PesqGrupo(link(ilink).NFreqs) = nome
'    link(ilink).PesqFreq(link(ilink).NFreqs) = Freq
'End Sub
'Function Bike_Spread_From_Point(ipoint As Long) As Long
''Returns the point that was closed in the previous round... expands from there... and marks that point with -marker
''COM RESTRIÇÃO DE LINKS A PÉ (p.m.k) K, Links "bi" com peso CorreRate
''Presume que a extensão dos links é sempre positiva e que correrate é positiva
'    Dim P As Point_type  'origin point
'    Dim SP As Point_type 'spread point
'    Dim idp As Point_type 'variable destiny point
'    Dim thelink As Link_type
'    P = point(ipoint)
'    Bike_Spread_From_Point = 0
'    bestTime = 1000000
'    NListed = 0
'    Do While point(LinkList(NListed + 1)).marker = P.marker
'        If point(LinkList(NListed + 1)).time <= bestTime Then
'            bestTime = point(LinkList(NListed + 1)).time
'            Bike_Spread_From_Point = LinkList(NListed + 1)
'            Position = NListed + 1
'        End If
'        NListed = NListed + 1
'    Loop
'    If Bike_Spread_From_Point = 0 Then Exit Function
'    SP = point(Bike_Spread_From_Point)
'    LinkList(Position) = LinkList(NListed)
'    LinkList(NListed) = 0
'    NListed = NListed - 1
'    If Bike_Spread_From_Point = PointNamed(5078) Then
'     PA = RA
'    End If
'    For i = 1 To SP.NLinksDaqui
'        thelink = link(SP.iLinkDaqui(i))
'        If (thelink.modes <> "m" And thelink.isM2 <> 1 And thelink.tipo < 600) Or (thelink.isM2 = 1 And (thelink.op = ipoint Or point(thelink.op).isM2 <> 1)) Then
'            idp = point(thelink.dp)
'            If idp.marker <> P.marker Then
'                point(thelink.dp).time = 1000000
'                point(thelink.dp).marker = marker
'                NListed = NListed + 1
'                LinkList(NListed) = thelink.dp
'                IPos = 1
'            End If
'            linktime = thelink.Extension / 250
'            linkbikedist = 0
'            If LinkhasMode(SP.iLinkDaqui(i), "t") Then
'                linktime = thelink.Extension / 350
'                linkbikedist = thelink.Extension
'            End If
'            canditime = SP.time + linktime
'            If canditime < point(thelink.dp).time Then
'                point(thelink.dp).time = canditime
'                point(thelink.dp).Distance = point(thelink.op).Distance + thelink.Extension
'                point(thelink.dp).bikelanedistance = point(thelink.op).bikelanedistance + linkbikedist
'                point(thelink.dp).LastLink = SP.iLinkDaqui(i)
'            End If
'        End If
'    Next i
'End Function
'
'Function NLinksInBikePath(ByVal Opoint As Long, ByVal Dpoint As Long, Optional Arelinked As Boolean, Optional NetWorkDistance As Single, Optional bikelanedistance) As Integer
'    Dim NLast As Integer
'    Dim NLinkList As Integer
'    NetWorkDistance = 0
'    NLast = 0
'    If Opoint = Dpoint Then Arelinked = True: Exit Function
'    marker = marker + 1
'    LinkList(1) = Opoint
'    NLinkList = 1
'    point(Opoint).Distance = 0
'    point(Opoint).marker = marker
'    reach = 0
'    Do While NLinkList > 0 And reach < 10
'        For i = 1 To NLinkList
'            NLast = NBikeSpread_From_Point(point(LinkList(i)), marker, NLast)
'        Next i
'        For i = 1 To NLast
'            LinkList(i) = LinkListSPread(i)
'        Next i
'        NLinkList = NLast
'        NLast = 0
'        If point(Dpoint).marker = marker Then
'            reach = reach + 1
'        End If
'    Loop
'    If reach = 0 Then
'        Arelinked = False
'        bestdist = MAXX
'        For i = 1 To Npoints
'            If point(i).marker = marker Then
'                dist = Point_Distance(point(Dpoint), point(i))
'                If dist < bestdist Then
'                    startpoint = i
'                    bestdist = dist
'                End If
'            End If
'        Next i
'    Else
'        Arelinked = True
'        startpoint = Dpoint
'    End If
'    NLinksInBikePath = 0
'    NetWorkDistance = 0
'    Do While startpoint <> Opoint
'        NLinksInBikePath = NLinksInBikePath + 1
'        LinkList(NLinksInBikePath) = point(startpoint).LastLink
'        NetWorkDistance = NetWorkDistance + link(point(startpoint).LastLink).Extension
'        startpoint = link(point(startpoint).LastLink).op
'    Loop
'    For i = 1 To Int(NLinksInBikePath / 2)
'        memo = LinkList(i)
'        LinkList(i) = LinkList(NLinksInBikePath - i + 1)
'        LinkList(NLinksInBikePath - i + 1) = memo
'    Next i
'End Function
'Function NBikeSpread_From_Point(P As Point_type, marker As Long, Optional NStart As Integer = 0) As Integer
''COM RESTRIÇÃO DE LINKS A PÉ (p.m.k) K, Links "bi" com peso 0.9
'    Dim idp As Point_type
'    Dim thelink As Link_type
'    NSpread_From_Point = NStart
'    For i = 1 To P.NLinksDaqui
'        thelink = link(P.iLinkDaqui(i))
'        If (thelink.modes = "p" And thelink.tipo = 20) Or (thelink.modes <> "k" And thelink.modes <> "m" And thelink.modes <> "p" And thelink.isM2 <> 1 And thelink.tipo < 600) Then
'            idp = point(thelink.dp)
'            If idp.marker <> marker Then
'                point(thelink.dp).Distance = 1000000
'                point(thelink.dp).marker = marker
'            End If
'            If thelink.modes = "bi" Or thelink.modes = "bip" Then
'                thelink.Extension = thelink.Extension * CorreRate
'            End If
'            CandiDist = P.Distance + thelink.Extension
'            If CandiDist < point(thelink.dp).Distance Then
'                point(thelink.dp).Distance = CandiDist
'                point(thelink.dp).LastLink = P.iLinkDaqui(i)
'                NSpread_From_Point = NSpread_From_Point + 1
'                LinkListSPread(NSpread_From_Point) = thelink.dp
'            End If
'        End If
'    Next i
'End Function
'Sub Register_RoutesSTOPS_in_Network()
''Resets routesstops in nodes run all routes first TRIP to fill it. (supposedly all trips have the same pattern)
''Review route stop and extension to each stop and total route extension
''Calculate extension based on link extension
''if uses a non-existing link, adds Euclidian distance between those points with NO WARNINGS
''if a route STOPs twice on the same point, register it twice (and distance related to each pass)
''routines that want to check it, just need to see if next route is the same as all the passages of
''each route are always together in a point,STOP or link
'Dim ipoint As Long
'Dim iroute As Integer
'For i = 1 To Npoints
'    point(i).NStopRoutes = 0
'Next i
'For iroute = 1 To nRoutes
'    route(iroute).HasPara = True
'    ReDim route(iroute).para(route(iroute).Npoints)
'    For i = 1 To route(iroute).Npoints
'        route(iroute).para(i) = "#"
'    Next i
'    itrip = route(iroute).Departure(1)
'    laststopINroute = 0
'    For i = 1 To trip(itrip).Nstation
'        ipoint = trip(itrip).stop_point(i)
'        NPass = 0
'        Do
'            NPass = NPass + 1
'            iroutepos = Get_iroute_pointpos(ipoint, iroute, NPass)
'            nowstopinroute = point(ipoint).iRoutePosition(iroutepos)
'        Loop While iroutepos <> 0 And nowstopinroute < laststopINroute
'        If iroutepos = 0 Then
''            MsgBox "Route does not pass here, not at least after the last stop!"
''            Debug.Print "i==" & point(ipoint).Name & "   line==" & RoUtE(iroute).number
'            iroutepos = Get_iroute_pointpos(ipoint, iroute)
'            nowstopinroute = point(ipoint).iRoutePosition(iroutepos)
'        Else
'            laststopINroute = nowstopinroute
'        End If
'        If iroutepos <> 0 Then
'            route(iroute).para(nowstopinroute) = "+"
'            point(ipoint).NStopRoutes = point(ipoint).NStopRoutes + 1
'            ReDim Preserve point(ipoint).iStopRoute(point(ipoint).NStopRoutes)
'            ReDim Preserve point(ipoint).iStopPosition(point(ipoint).NStopRoutes)
'            ReDim Preserve point(ipoint).iStopDistance(point(ipoint).NStopRoutes)
'            point(ipoint).iStopRoute(point(ipoint).NStopRoutes) = iroute
'            point(ipoint).iStopPosition(point(ipoint).NStopRoutes) = i
'            point(ipoint).iStopDistance(point(ipoint).NStopRoutes) = point(ipoint).iRouteDistance(iroutepos)
'        End If
'    Next i
'Next iroute
'End Sub

