Attribute VB_Name = "BasicNetwork"
Public Const PI As Double = 3.14159265358979
Type Point_type
    X As Double
    Y As Double
    IsM2 As Integer '0=is not EMME, 1=is centroid, 2=is not centroid
    t1 As Single '0 or 1 or 10
    t2 As Single '0 or 1.2 or 2.5
    t3 As Single '0 or 50
    name As String '5 numbers
    STname As String
    nlinksdaqui As Integer 'número de links que começam nesse nó
    ilinkdaqui() As Long 'os links que começam nesse nó
    nlinkspraka As Integer 'número de links que acabam no nó
    ilinkpraka() As Long 'os links que acabam no nó
    NRoutes As Integer 'número de linhas que passam no nó
    iroute() As Integer 'número das linhas que passam no nó
    tipo As String
'    nroutegrupos As Integer 'quantidade de tipo de linhas que passam no nó
'    iroutegrupos() As Integer 'os tipos de linha que passam no nó
    NextNodeX As Long 'nó com próxima coordenada X maior ou igual a deste nó
    NextNodeY As Long 'nó com próxima coordenada Y maior ou igual a deste nó
    Marker As Single 'marcador para as funções que usam public marker, sempre crescente
    Corredor As Long 'número do corredor a que pertence, alternativamente
                      'em MarcaCorredoresPointsandLinks 0 se não pertence a corredor,1 se pertence
    DistânciadoCorredor As Integer 'abaixo de 600 metros... a distância de um ponto de corredor
    DistânciadaRedeBRT As Integer 'distância da rede link tipo 8, corredor de ônibus em BH
    DistânciadaRedeMetrô As Integer 'distância da rede link tipo 10, metrô em BH
    NewP As Integer 'aponta para o novo ponto no corredor
    Estação As Long '1=parada, 0=não para
    DistânciadaRede As Single 'distância do ponto até algum corredor
    DistânciadaEstação As Single 'distância do ponto no corredor até a estação
    Closer As Long 'ponto no caminho da estação, para guiar linhas
    Toexpand As Boolean
    distance As Single 'para caminho mínimo NLinksinShortestPath
    time As Single 'para caminho mínimo NLinksinShortestPath
    bikelanedistance As Single 'para caminho mínimo NLinksinShortestPath
    LastLink As Long 'para caminho mínimo NLinksinShortestPath
    DELETA As Boolean 'para refazer conectores
    Limite As Integer 'para indicar a que perímetro pertence
End Type
Type Link_type
    op As Long
    dp As Long
    IsM2 As Integer '0=is not, 1=is to centroid, 2=is not to centroid
    Extension As Single
    time As Single
    Auxiliar As Single
    modes As String
    Lanes As Integer
    tipo As Integer
    vdf As Integer
    t1 As Single
    t2 As Single
    t3 As Single
    Corredor As Integer
    Troncal As Single 'usado para contar linhas
    NRoutes As Integer 'número de linhas que passam no link
    iroute() As Integer 'número das linhas que passam no link
    centro As Boolean 'é link da área central
    NRoutesOferta As Integer
    iRouteOferta() As Integer
    DemandaSemCobertura As Single  'para ajuste especial de headways
    Freq(-2 To 0) As Single
    NRou(-2 To 0) As Single
    Limite As Integer 'para indicar a que perímetro pertence
End Type
Type Route_type
    number As String
    mode As String
    vehicle As Integer
    headway As Single
    speed As Single
    name As String
    path As String
    lay As String
    t1 As Single
    t2 As Single
    t3 As Single
    Npoints As Integer
    ipoint() As Long
    NLinks As Integer 'em tese seria sempre Npoints -1, mas eventualmente se escolhe pular alguns (os virtuais)
    ilink() As Long
    HasPara As Boolean
    para() As String ' < board, > allight, + board and alight , # non-stop
    HasDwt As Boolean
    dwt() As Single  'dwt must mark stop time
    HasTtf As Boolean
    ttf() As Integer
    HasTtfL As Boolean
    ttfL() As Integer
    HasTtfT As Boolean
    ttfT() As Integer
    HasUs1 As Boolean
    Us1() As Integer
    HasUs2 As Boolean
    Us2() As Integer
    HasUs3 As Boolean
    Us3() As Integer
    HasVoltr As Boolean
    voltr() As Single
    linkvoltr() As Single
    MVehicle As Single
    MVehCapacity As Single
    MLength As Single
    MFleet As Single
    MTime As Single
    MBoard As Single
    MPaxHour As Single
    MPaxKm As Single
    MAvgLoad As Single
    MMaxLoad As Single
    MMaxVol As Single
    Extension As Single
    ExtensionInCorridor As Single
    ReturnRoute As Integer
    NewHeadway As Single
    DistAntes As Single
    DistDurante As Single
    DistDepois As Single
    KmAntes As Single
    KmDurante As Single
    KmForaDurante As Single
    KmDepois As Single
    FirstStop As Integer
    LastStop As Integer
    ToPrint As Integer
    checked As Boolean
    changed As Boolean
    selected As Boolean
    uptopax As Single  'para cálculo de novo headway
    overkm As Single   'para cálculo de novo headway
    DELETA As Boolean 'para only changes
End Type
Public Npoints As Long
Public Point() As Point_type
Public NLinks As Long
Public Link() As Link_type
Public NRoutes As Integer
Public Route() As Route_type

'Para Marcar mudanças
Public OldNpoints As Long
Public OldNLinks As Long
Public OldNRoutes As Integer

'for metro fare integration
Public NEstaçõeS As Integer
Public Estação(30) As Long 'stations
Public EstaçãoIntegra(30, 3) As Boolean 'true if station integrates with route of group j

Public PontodoCorredor(20, 150) As Long
Public NovosPontos(2000) As Point_type
Public NNovosPontos As Integer
Public MetroLine(9) As Long

Const FirstZona As Integer = 6649
Const LastZona As Integer = 7652


Public Const MINY As Single = -30.243 'menor coordinada Y
Public Const MAXY As Single = -29.639 'maior coordinada Y
Public Const MINX As Single = -51.403 'menor coordenada X
Public Const MAXX As Single = -50.772 'maior coordinada X
Public Const MAXNUMNOMES As Long = 1000000 'maior número Y
Public Const XSCREENYDIVIDE As Single = 0.0005 'distância para dividir a tela em X para NODE FINDER
Public Const YSCREENYDIVIDE As Single = 0.0005 'distância para dividir a tela em Y para NODE FINDER
Public Const DIM_XPOINTFINDER As Long = (MAXX - MINX) / XSCREENYDIVIDE
Public Const DIM_YPOINTFINDER As Long = (MAXY - MINY) / YSCREENYDIVIDE
Public Const MAXLINKLIST As Long = 30000
Public PointList(MAXLINKLIST) As Long
Public LinkList(MAXLINKLIST) As Long
Public LinkListSPread(MAXLINKLIST) As Long 'para ShortestPath: Dijkstra não precisa
'As funções YFinder e XFinder, transformam a coordenada do ponto de partida de busca para Y e X PointFinder
Public XPointFinder(DIM_XPOINTFINDER) As Long
Public YPointFinder(DIM_YPOINTFINDER) As Long
Public CorreRate As Single
Public Marker As Long 'marker... always grows
Public PointNamed(MAXNUMNOMES) As Long
Public Contapontosexpandidos As Integer
Public icentro As Long
Public VehCap(30) As Integer
Public LimitedeBh As Route_type
Public LimiteRodoAnel As Route_type
Public LimiteContorno As Route_type
Public LimiteHiperCentro As Route_type
Sub ResetNetWork()
ResetNodes
ResetLinks
End Sub
Sub ResetNodes()
For i = 0 To MAXNUMNOMES
    PointNamed(i) = 0
Next i
For i = 0 To DIM_XPOINTFINDER
    XPointFinder(i) = 0
Next i
For i = 0 To DIM_YPOINTFINDER
    YPointFinder(i) = 0
Next i
Npoints = 0
ReDim Point(Npoints)
End Sub
Sub ResetLinks()
NLinks = 0
For i = 1 To Npoints
    Point(i).nlinksdaqui = 0
    Point(i).nlinkspraka = 0
Next i
ReDim Links(NLinks)
End Sub
Function XFinder(X As Double) As Long
    XFinder = (X - MINX) / XSCREENYDIVIDE
    If XFinder < 0 Then XFinder = 0
    If XFinder > DIM_XPOINTFINDER Then XFinder = DIM_XPOINTFINDER
End Function
Function YFinder(Y As Double) As Long
    YFinder = Int((Y - MINY) / XSCREENYDIVIDE)
    If YFinder < 0 Then YFinder = 0
    If YFinder > DIM_YPOINTFINDER Then YFinder = DIM_YPOINTFINDER
End Function
Function NPointsListed(ByVal XLow As Double, ByVal YLow As Double, Optional ByVal XHigh As Double = -1, Optional ByVal YHigh As Double = -1) As Long
'Lista pontos no quadrilatero formado por XLow,XHigh,YLow,YHigh) em Public PointList,
'considerando as larguras XSCREENYDIVIDE e YSCREENYDIVIDE
'Se XHigh ou YHigh não são fornecidos (=-1), so o valor em questão interessa
    NPointsListed = 0
    If XHigh = -1 Then XHigh = XLow
    If YHigh = -1 Then YHigh = YLow
    If XLow > XHigh Then 'coloca X em ordem
        mem = XHigh
        XHigh = XLow
        XLow = XHigh
    End If
    If YLow > YHigh Then 'coloca Y em ordem
        mem = YHigh
        YHigh = YLow
        YLow = YHigh
    End If
    Marker = Marker + 1
    For i = XFinder(XLow) To XFinder(XHigh)
        NextNode = XPointFinder(i)
        While NextNode <> 0
            Point(NextNode).Marker = Marker
            NextNode = Point(NextNode).NextNodeX
        Wend
    Next i
    For i = YFinder(YLow) To YFinder(YHigh)
        NextNode = YPointFinder(i)
        While NextNode <> 0
            If Point(NextNode).Marker = Marker Then
                NPointsListed = NPointsListed + 1
                PointList(NPointsListed) = NextNode
            End If
            NextNode = Point(NextNode).NextNodeY
        Wend
    Next i
End Function
Function NPointsListedInARadius(Center As Point_type, radius As Single) As Long
    Dim XLow As Double
    Dim XHigh As Double
    Dim YLow As Double
    Dim YHigh As Double
    NPointsListedInARadius = 0
    XLow = Center.X - radius
    XHigh = Center.X + radius
    YLow = Center.Y - radius
    YHigh = Center.Y + radius
    If XLow < 0 Then XLow = 0
    If XHigh > MAXX Then XHigh = MAXX
    If YLow < 0 Then YLow = 0
    If YHigh > MAXY Then YHigh = MAXY
'    For i = XLow To XHigh
    For i = XFinder(XLow) To XFinder(XHigh)
        NextNode = XPointFinder(i)
        While NextNode <> 0
            Point(NextNode).Marker = Marker
            NextNode = Point(NextNode).NextNodeX
        Wend
    Next i
'    For i = YLow To YHigh
    For i = YFinder(YLow) To YFinder(YHigh)
        NextNode = YPointFinder(i)
        While NextNode <> 0
            If Point(NextNode).Marker = Marker Then
                dista = Point_Distance(Center, Point(NextNode))
                If dista <= radius And dista >= 0 Then
                    NPointsListedInARadius = NPointsListedInARadius + 1
                    PointList(NPointsListedInARadius) = NextNode
                End If
            End If
            NextNode = Point(NextNode).NextNodeY
        Wend
    Next i
End Function
Function NPointsListedInARadiusX(Center As Point_type, radius As Single, Optional StartFrom As Integer = 0) As Long
    'Startfrom pode ser utilizado para adicionar mais pontos a uma lista já existente
    If StarFrom = 0 Then Marker = Marker + 1
    NPointsListedInARadiusX = StarFrom
    XLow = Center.X - radius
    XHigh = Center.X + radius
    YLow = Center.Y - radius
    YHigh = Center.Y + radius
    If XLow < 0 Then XLow = 0
    If XHigh > MAXX Then XHigh = MAXX
    If YLow < 0 Then YLow = 0
    If YHigh > MAXY Then YHigh = MAXY
    
    For i = XFinder(XLow) To XFinder(XHigh)
        NextNode = XPointFinder(i)
        While NextNode <> 0
            Point(NextNode).Marker = Marker
            NextNode = Point(NextNode).NextNodeX
        Wend
    Next i
    For i = YFinder(YLow) To YFinder(YHigh)
        NextNode = YPointFinder(i)
        While NextNode <> 0
            If Point(NextNode).Marker = Marker Then
                dista = Point_Distance(Center, Point(NextNode))
                If dista <= radius And dista >= 0 Then
                    NPointsListedInARadiusX = NPointsListedInARadiusX + 1
                    PointList(NPointsListedInARadiusX) = NextNode
                End If
            End If
            NextNode = Point(NextNode).NextNodeY
        Wend
    Next i
End Function
Function Is_Inside_Perimeter(Point_ent As Point_type, Perimeter_ent As Route_type) As Boolean
    Dim Distant_Point As Point_type
    Distant_Point.X = 1
    Distant_Point.X = 1
    For iper = 1 To Perimeter_ent.Npoints
        iper2 = iper + 1
        If iper2 > Perimeter_ent.Npoints Then iper2 = 1
        cruzo = Cross(Point_ent, Distant_Point, Point(Perimeter_ent.ipoint(iper)), Point(Perimeter_ent.ipoint(iper2))).Marker
        Is_Inside_Perimeter = Is_Inside_Perimeter Xor cruzo
    Next iper
End Function
Function Cross(A1 As Point_type, A2 As Point_type, B1 As Point_type, B2 As Point_type, Optional IsInfinite As Boolean = False) As Point_type
    ' If the segment A1->A2 crosses B1->B2 the node Cross.Marker will be minus one (which is boolean for true)
    ' (else)If there is no intersection it will be marked 0
    ' And .X and .Y indicates the point location
    ' If IsInfinite is used, we assume the segments are part of a infinite line,
    ' there fore Cross.Marker = 0 tells us that they are parallel lines.
    ' and Cross.Marker = 1 tells us that they are the same line.
    Dim AngA As Single
    Dim AngB As Single
    Dim AngAB As Single
    Dim Avertical As Boolean
    Dim Bvertical As Boolean
    Cross.Marker = -1
    Avertical = (A1.X = A2.X)
    Bvertical = (B1.X = B2.X)
    If Avertical And Bvertical Then
        Cross.Marker = 0
        If A1.X = B1.X Then Cross.Marker = 1
    ElseIf Avertical Then
        Cross.X = A1.X
        AngB = (B2.Y - B1.Y) / (B2.X - B1.X)
        Cross.Y = B1.Y + AngB * (Cross.X - B1.X)
    ElseIf Bvertical Then
        Cross.X = B1.X
        AngA = (A2.Y - A1.Y) / (A2.X - A1.X)
        Cross.Y = A1.Y + AngA * (Cross.X - A1.X)
    Else
        AngA = (A2.Y - A1.Y) / (A2.X - A1.X)
        AngB = (B2.Y - B1.Y) / (B2.X - B1.X)
        If AngA = AngB Then
            Cross.Marker = 0
            AngAB = (A2.Y - B1.Y) / (A2.X - B1.X)
            If AngAB = AngA Then Cross.Marker = 1
        Else
            Cross.X = (B1.Y - A1.Y - B1.X * AngB + A1.X * AngA) / (AngA - AngB)
            Cross.Y = A1.Y + AngA * (Cross.X - A1.X)
        End If
    End If
    If Cross.Marker = -1 And Not IsInfinite Then
       'Ext A/B: True if Cross is on the extreme of A and/or B
        If Avertical Then
            extA = Cross.Y = A1.Y Or Cross.Y = A2.Y
            crossa = Cross.Y > A1.Y Xor Cross.Y > A2.Y
        Else
            extA = Cross.X = A1.X Or Cross.X = A2.X
            crossa = Cross.X > A1.X Xor Cross.X > A2.X
        End If
        touchA = extA Or crossa
        If Bvertical Then
            extb = Cross.Y = B1.Y Or Cross.Y = B2.Y
            crossb = Cross.Y > B1.Y Xor Cross.Y > B2.Y
        Else
            extb = Cross.X = B1.X Or Cross.X = B2.X
            crossb = Cross.X > B1.X Xor Cross.X > B2.X
        End If
        touchB = extb Or crossb
        If Not (touchA And touchB) Then Cross.Marker = 0
    End If
End Function
Function Point_to_Link_Distance(ip As Long, ilink As Long) As Single
    Dim il As Link_type
    Dim P As Point_type
    Dim PA As Point_type
    Dim xis As Point_type
    ' P ->PA = segmento ortogonal
    P = Point(ip)
    il = Link(ilink)
    a = Point(il.dp).Y - Point(il.op).Y
    b = Point(il.op).X - Point(il.dp).X
    PA.X = P.X + a
    PA.Y = P.Y + b
    xis = Cross(Point(il.op), Point(il.dp), P, PA, True)
    If (xis.X > Point(il.op).X And xis.X < Point(il.dp).X) Or (xis.X < Point(il.op).X And xis.X > Point(il.dp).X) Then
        Point_to_Link_Distance = Point_Distance(P, xis)
    Else
        Point_to_Link_Distance = Point_Distance(P, Point(il.dp))
        pdis = Point_Distance(P, Point(il.op))
        If pdis < Point_to_Link_Distance Then Point_to_Link_Distance = pdis
    End If
' se fosse infinito:   Point_to_Link_Distance = Abs((A * P.X + B * P.Y - A * Point(il.OP).X - B * Point(il.OP).Y)) / (A ^ 2 * B ^ 2) ^ 0.5
End Function
Function Get_Point(Eastings As Double, Northings As Double, Optional CreatePoint As Boolean = True, Optional IsM2 As Integer = 0, Optional name As String = "") As Long
    'Return the node number, if not If the node does not exist, it will be created unless create is passed as false
    Get_Point = 0
    For i = 1 To NPointsListed(Eastings, Northings)
        If Point(PointList(i)).X = Eastings And Point(PointList(i)).Y = Northings Then
            Get_Point = PointList(i)
            Exit Function
        End If
    Next i
    If CreatePoint Then Get_Point = Add_Point(Eastings, Northings, IsM2, name)
End Function
Function Add_Point(ByVal Eastings As Double, ByVal Northings As Double, Optional IsM2 As Integer = 0, Optional name As String = "") As Long
    'Add_Point will return the number of the MPoint added
    'It also prepare insert the Mpoint in the finder
    Dim Chega As Boolean
    Dim Last_Node As Long
    Dim Next_Node As Long
    Add_Point = 0
    Npoints = Npoints + 1 'aumenta
    ReDim Preserve Point(Npoints)
    Point(Npoints).X = Eastings
    Point(Npoints).IsM2 = IsM2
    Point(Npoints).nlinkspraka = 0
    Point(Npoints).nlinksdaqui = 0
    Point(Npoints).name = name
    Point(Npoints).STname = name
    PointNamed(name) = Npoints
    ReDim Point(Npoints).iroute(0)
    'Find .NextNodeX
    Chega = False
    Last_Node = 0
    IntE = XFinder(Eastings)
    Next_Node = XPointFinder(IntE)
    If Next_Node = 0 Then Chega = True
    While Not Chega
        If Eastings <= Point(Next_Node).X Then
            Last_Node = Next_Node
            Next_Node = Point(Last_Node).NextNodeX
            If Next_Node = 0 Then Chega = True
        Else
            Chega = True
        End If
    Wend
    'Found .NextNodeX
    Point(Npoints).NextNodeX = Next_Node
    'Last_Node (on X) now will point to the node being added
    If Last_Node = 0 Then
        XPointFinder(IntE) = Npoints
    Else
        Point(Last_Node).NextNodeX = Npoints
    End If
    ' REPEAT FOR Y
    Point(Npoints).Y = Northings
    'Find .NextNodeY
    Chega = False
    Last_Node = 0
    IntN = YFinder(Northings)
    Next_Node = YPointFinder(IntN)
    If Next_Node = 0 Then Chega = True
    While Not Chega
        If Northings <= Point(Next_Node).Y Then
            Last_Node = Next_Node
            Next_Node = Point(Last_Node).NextNodeY
            If Next_Node = 0 Then Chega = True
        Else
            Chega = True
        End If
    Wend
    'Found .NextNodeY
    Point(Npoints).NextNodeY = Next_Node
    'Last_Node (on Y) now will point to the node being added
    If Last_Node = 0 Then
        YPointFinder(IntN) = Npoints
    Else
        Point(Last_Node).NextNodeY = Npoints
    End If
    Add_Point = Npoints
End Function
Function Get_Link(PO As Long, PD As Long) As Long
    For i = 1 To Point(PO).nlinksdaqui
        If Link(Point(PO).ilinkdaqui(i)).dp = PD Then
            Get_Link = Point(PO).ilinkdaqui(i)
            Exit Function
        End If
    Next i
End Function
Function Add_Link(DummyLink As Link_type) As Long
    NLinks = NLinks + 1
    ReDim Preserve Link(NLinks)
    Link(NLinks) = DummyLink
    'insert linkdaqui in origin point
    op = Link(NLinks).op
    Point(op).nlinksdaqui = Point(op).nlinksdaqui + 1
    ReDim Preserve Point(op).ilinkdaqui(Point(op).nlinksdaqui)
    Point(op).ilinkdaqui(Point(op).nlinksdaqui) = NLinks
    'insert linkpraka in destiny point
    dp = Link(NLinks).dp
    Point(dp).nlinkspraka = Point(dp).nlinkspraka + 1
    ReDim Preserve Point(dp).ilinkpraka(Point(dp).nlinkspraka)
    Point(dp).ilinkpraka(Point(dp).nlinkspraka) = NLinks
    Link(Nlink).NRoutes = 0
    ReDim Link(NLinks).iroute(0)
    Add_Link = NLinks
End Function
Function Delete_Link(il As Long) As Boolean
    Delete_Link = Tira_Link(Link(il).op, il) And Tira_Link(Link(il).dp, il)
    Link(il).tipo = 1000
End Function
Function Tira_Link(ipoint As Long, ilink As Long) As Boolean
    For i = 1 To Point(ipoint).nlinksdaqui
        Tira_Link = False
        If Point(ipoint).ilinkdaqui(i) = ilink Then
            For j = i + 1 To Point(ipoint).nlinksdaqui
                Point(ipoint).ilinkdaqui(j - 1) = Point(ipoint).ilinkdaqui(j)
            Next j
            Point(ipoint).nlinksdaqui = Point(ipoint).nlinksdaqui - 1
            Tira_Link = True
            Exit Function
        End If
    Next i
    For i = 1 To Point(ipoint).nlinkspraka
        If Point(ipoint).ilinkpraka(i) = ilink Then
            For j = i + 1 To Point(ipoint).nlinkspraka
                Point(ipoint).ilinkpraka(j - 1) = Point(ipoint).ilinkpraka(j)
            Next j
            Point(ipoint).nlinkspraka = Point(ipoint).nlinkspraka - 1
            Tira_Link = True
            Exit Function
        End If
    Next i
End Function
Function GetAngletoSplitPoint(ipoint As Integer) As Single
ReDim NumtoSort(100)
For i = 1 To Point(ipoint).nlinksdaqui
    NumtoSort(i) = Angle(Point(ipoint), Point(Link(Point(ipoint).ilinkdaqui(i)).dp))
Next i
For i = 1 To Point(ipoint).nlinkspraka
    NumtoSort(i + Point(ipoint).nlinksdaqui) = Angle(Point(ipoint), Point(Link(Point(ipoint).ilinkpraka(i)).op))
Next i
NtoSort = Point(ipoint).nlinkspraka + Point(ipoint).nlinksdaqui
NumHeapSort
GreaterAngle = NumtoSort(1) + (360 - NumtoSort(NtoSort))
GetAngletoSplitPoint = NumtoSort(NtoSort) + GreaterAngle / 2
If GetAngletoSplitPoint > 360 Then GetAngletoSplitPoint = GetAngletoSplitPoint - 360
For i = 2 To NtoSort
    If NumtoSort(i) - NumtoSort(i - 1) > GreaterAngle Then
        GreaterAngle = NumtoSort(i) - NumtoSort(i - 1)
        GetAngletoSplitPoint = NumtoSort(i - 1) + GreaterAngle / 2
    End If
Next i
End Function
Function GetRestrictedAngletoSplitPoint(ipoint As Long, AngFrom As Single, AngTo As Single) As Single
ReDim NumtoSort(100)
Dim ANg As Single
nn = 0
For i = 1 To Point(ipoint).nlinksdaqui
    ANg = Angle(Point(ipoint), Point(Link(Point(ipoint).ilinkdaqui(i)).dp))
    If IsAngBetween(ANg, AngFrom, AngTo) And Point(Link(Point(ipoint).ilinkdaqui(i)).dp).name < 10000 Then
        nn = nn + 1
        NumtoSort(nn) = ANg
    End If
Next i
For i = 1 To Point(ipoint).nlinkspraka
    ANg = Angle(Point(ipoint), Point(Link(Point(ipoint).ilinkpraka(i)).op)) And Point(Link(Point(ipoint).ilinkpraka(i)).op).name < 10000
    If IsAngBetween(ANg, AngFrom, AngTo) Then
        nn = nn + 1
        NumtoSort(nn) = ANg
    End If
    NumtoSort(i + Point(ipoint).nlinksdaqui) = Angle(Point(ipoint), Point(Link(Point(ipoint).ilinkpraka(i)).op))
Next i
nn = nn + 1
NumtoSort(nn) = AngFrom
nn = nn + 1
NumtoSort(nn) = AngTo
NtoSort = nn 'public parameter to NumHeapSort
NumHeapSort
NumtoSort(0) = NumtoSort(NtoSort) - 360
GreaterAngle = 0
For i = 1 To NtoSort
    If NumtoSort(i) - NumtoSort(i - 1) > GreaterAngle And NumtoSort(i) <> AngFrom Then
        GreaterAngle = NumtoSort(i) - NumtoSort(i - 1)
        GetRestrictedAngletoSplitPoint = NumtoSort(i - 1) + GreaterAngle / 2
    End If
Next i
End Function
Function IsAngBetween(ANg As Single, AngFrom As Single, AngTo As Single) As Boolean
IsAngBetween = False
If AngFrom >= AngTo Then
    If ANg > AngFrom Or ANg < AngTo Then IsAngBetween = True
Else
    If ANg < AngTo And ANg > AngFrom Then IsAngBetween = True
End If
End Function
Function PointInsertPosition(BasePoint As Point_type, CentralAngle As Single, RelX As Integer, RelY As Integer, Optional ModDist As Integer = 25) As Point_type
'This function returns the point in position for the insertion of the new point
'Central Angle is the direction the tree of points is being mounted in degrees
'RelX and RelY are the relative position in the tree, ModDist is the space between points (- to the right,+ to the left)
PointInsertPosition.X = Int(BasePoint.X + RelY * ModDist * Cos(CentralAngle * PI / 180) + RelX * ModDist * Sin(CentralAngle * PI / 180) / 2)
PointInsertPosition.Y = Int(BasePoint.Y + RelY * ModDist * Sin(CentralAngle * PI / 180) - RelX * ModDist * Cos(CentralAngle * PI / 180) / 2)
End Function
Function Angle(pos1 As Point_type, pos2 As Point_type) As Single
distx = pos2.X - pos1.X
disty = pos2.Y - pos1.Y
Angle = 0
If distx = 0 And disty = 0 Then Angle = 10000: Exit Function
If distx = 0 And disty < 0 Then Angle = 270: Exit Function
If distx = 0 And disty > 0 Then Angle = 90: Exit Function
If distx > 0 And distx = 0 Then Angle = 0: Exit Function
If distx < 0 And disty = 0 Then Angle = 180: Exit Function
If distx > 0 And disty > 0 Then Angle = (Atn(disty / distx)) * 180 / 3.14159265: Exit Function
If distx > 0 And disty < 0 Then Angle = 360 + ((Atn(disty / distx)) * 180 / 3.14159265): Exit Function
If distx < 0 Then Angle = 180 + ((Atn(disty / distx)) * 180 / 3.14159265): Exit Function
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
Function Point_Diference(P1 As Point_type, P2 As Point_type) As Point_type
Point_Diference.X = P1.X - P2.X
Point_Diference.Y = P1.Y - P2.Y
End Function
Function Point_Summ(P1 As Point_type, P2 As Point_type) As Point_type
Point_Summ.X = P1.X + P2.X
Point_Summ.Y = P1.Y + P2.Y
End Function
Function Point_Distance(P1 As Point_type, P2 As Point_type) As Single
Point_Distance = ((P1.X - P2.X) ^ 2 + (P1.Y - P2.Y) ^ 2) ^ 0.5
End Function
Function distância(iponto As Long, jponto As Long)
    distância = ((Point(iponto).X - Point(jponto).X) ^ 2 + (Point(iponto).Y - Point(jponto).Y) ^ 2) ^ 0.5
End Function
Function NCentroidsListedInARadius(Center As Point_type, radius As Single) As Long
    Dim XLow As Double
    Dim XHigh As Double
    Dim YLow As Double
    Dim YHigh As Double
    NCentroidsListedInARadius = 0
    XLow = Center.X - radius
    XHigh = Center.X + radius
    YLow = Center.Y - radius
    YHigh = Center.Y + radius
    If XLow < 0 Then XLow = 0
    If XHigh > MAXX Then XHigh = MAXX
    If YLow < 0 Then YLow = 0
    If YHigh > MAXY Then YHigh = MAXY
'    For i = XLow To XHigh
    For i = XFinder(XLow) To XFinder(XHigh)
        NextNode = XPointFinder(i)
        While NextNode <> 0
            Point(NextNode).Marker = Marker
            NextNode = Point(NextNode).NextNodeX
        Wend
    Next i
'    For i = YLow To YHigh
    For i = YFinder(YLow) To YFinder(YHigh)
        NextNode = YPointFinder(i)
        While NextNode <> 0
            If Point(NextNode).Marker = Marker And Point(NextNode).IsM2 = 1 Then
                dista = Point_Distance(Center, Point(NextNode))
                If dista <= radius And dista >= 0 Then
                    NCentroidsListedInARadius = NCentroidsListedInARadius + 1
                    PointList(NCentroidsListedInARadius) = NextNode
                End If
            End If
            NextNode = Point(NextNode).NextNodeY
        Wend
    Next i
End Function
Function Get_Straighter(jlink As Long) As Long
    Dim ang1 As Single
    Dim ang2 As Single
    Get_Straighter = 0
    menordif = 90
    ang1 = Angle(Link(jlink).dp, Link(jlink).op)
    For i = 1 To Point(Link(jlink).op).nlinkspraka
        klink = Point(Link(jlink).op).ilinkpraka(i)
        ang2 = Angle(Link(klink).dp, Link(klink).op)
        difang = DifAngle(ang1, ang2)
        If difang < menordif Then
            Get_Straighter = klink
            menordif = difang
        End If
    Next i
End Function
Function NewRefredPoint(ipoint As Long, NorS As String, distance As Integer) As Long
Dim Dpoint As Point_type
Dim Dlink As Link_type
Dim ANg As Single
Dim dista As Single
    ANg = -1
    If NorS = "Leste" Then ANg = 0
    If NorS = "Norte" Then ANg = 90
    If ANg = -1 Then MsgBox "Fail to learn angle"
    Dpoint = PointInsertPosition(Point(ipoint), ANg, 0, 1, distance)
    NewRefredPoint = Add_Point(Dpoint.X, Dpoint.Y, 2, Get_First_Avaiable_Name(7500))
End Function
Function Get_First_Avaiable_Name(Optional bottom As Integer = 0, Optional top As Integer = 9999) As String
Get_First_Avaiable_Name = ""
For i = bottom To top
    If PointNamed(i) = 0 Then
        Get_First_Avaiable_Name = i
        Exit For
    End If
Next i
If Get_First_Avaiable_Name = "" Then MsgBox "Could not get name!"
End Function
Function NLinksInShortestPathDijkstra(ByVal OPoint As Long, ByVal Dpoint As Long, Optional Arelinked As Boolean, Optional NetWorkDistance As Single) As Integer
    Dim NLast As Integer
    Dim BestPoint As Long
    Dim NLinkList As Integer
    NetWorkDistance = 0
    Marker = Marker + 1
    LinkList(1) = OPoint
    LinkList(2) = 0
    BestPoint = OPoint
    Point(OPoint).distance = 0
    Point(OPoint).Marker = Marker
    conta = 0
    Do While BestPoint <> Dpoint And BestPoint <> 0
        BestPoint = Spread_From_Point(OPoint)
        conta = conta + 1
    Loop
    If BestPoint = 0 Then
        Arelinked = False
        bestdist = (MAXX - MINX) + (MAXY - MINY)
        For i = 1 To Npoints
            If Point(i).Marker = Marker Then
                dist = Point_Distance(Point(Dpoint), Point(i))
                If dist < bestdist Then
                    BestPoint = i
                    bestdist = dist
                End If
            End If
        Next i
    Else
        Arelinked = True
    End If
    NLinksInShortestPathDijkstra = 0
    NetWorkDistance = 0
    Do While BestPoint <> OPoint And BestPoint <> 0
        NLinksInShortestPathDijkstra = NLinksInShortestPathDijkstra + 1
        LinkList(NLinksInShortestPathDijkstra) = Point(BestPoint).LastLink
        NetWorkDistance = NetWorkDistance + Link(Point(BestPoint).LastLink).Extension
        BestPoint = Link(Point(BestPoint).LastLink).op
    Loop
    For i = 1 To Int(NLinksInShortestPathDijkstra / 2)
        memo = LinkList(i)
        LinkList(i) = LinkList(NLinksInShortestPathDijkstra - i + 1)
        LinkList(NLinksInShortestPathDijkstra - i + 1) = memo
    Next i
End Function
Function NLinksInShortestPathDijkstraTO(ByVal TOPoint As Long, ByVal FROMpoint As Long, Optional Arelinked As Boolean, Optional NetWorkDistance As Single) As Integer
    Dim NLast As Integer
    Dim BestPoint As Long
    Dim NLinkList As Integer
    NetWorkDistance = 0
    Marker = Marker + 1
    LinkList(1) = TOPoint
    LinkList(2) = 0
    BestPoint = TOPoint
    Point(TOPoint).distance = 0
    Point(TOPoint).Marker = Marker
    conta = 0
    Do While BestPoint <> FROMpoint And BestPoint <> 0
        BestPoint = Spread_TO_Point(TOPoint)
        conta = conta + 1
    Loop
    If BestPoint = 0 Then
        Arelinked = False
        bestdist = (MAXX - MINX) + (MAXY - MINY)
        For i = 1 To Npoints
            If Point(i).Marker = Marker Then
                dist = Point_Distance(Point(FROMpoint), Point(i))
                If dist < bestdist Then
                    BestPoint = i
                    bestdist = dist
                End If
            End If
        Next i
    Else
        Arelinked = True
    End If
    NLinksInShortestPathDijkstraTO = 0
    NetWorkDistance = 0
    Do While BestPoint <> TOPoint And BestPoint <> 0
        NLinksInShortestPathDijkstraTO = NLinksInShortestPathDijkstraTO + 1
        LinkList(NLinksInShortestPathDijkstraTO) = Point(BestPoint).LastLink
        NetWorkDistance = NetWorkDistance + Link(Point(BestPoint).LastLink).Extension
        BestPoint = Link(Point(BestPoint).LastLink).dp
    Loop
End Function
Function Spread_From_Point(ipoint As Long) As Long
'Returns the point that was closed in the previous round... expands from there... and marks that point with -marker
'COM RESTRIÇÃO DE LINKS A PÉ (p.m.k) K, Links "bi" com peso CorreRate
'Presume que a extensão dos links é sempre positiva e que correrate é positiva
    Dim P As Point_type  'origin point
    Dim SP As Point_type 'spread point
    Dim idp As Point_type 'variable destiny point
    Dim thelink As Link_type
    P = Point(ipoint)
    Spread_From_Point = 0
    Bestdistance = 1000000
    NListed = 0
    Do While Point(LinkList(NListed + 1)).Marker = P.Marker
        If Point(LinkList(NListed + 1)).distance <= Bestdistance Then
            Bestdistance = Point(LinkList(NListed + 1)).distance
            Spread_From_Point = LinkList(NListed + 1)
            Position = NListed + 1
        End If
        NListed = NListed + 1
    Loop
    If Spread_From_Point = 0 Then Exit Function
    SP = Point(Spread_From_Point)
    LinkList(Position) = LinkList(NListed)
    LinkList(NListed) = 0
    NListed = NListed - 1
    For i = 1 To SP.nlinksdaqui
        thelink = Link(SP.ilinkdaqui(i))
        If (thelink.modes = "p" And thelink.tipo Mod 1000 = 20) Or (thelink.modes <> "k" And thelink.modes <> "m" And thelink.modes <> "p" And thelink.IsM2 <> 1 And thelink.tipo Mod 1000 < 600 And thelink.tipo Mod 1000 <> 19) Then
            idp = Point(thelink.dp)
            If idp.Marker <> P.Marker Then
                Point(thelink.dp).distance = 1000000
                Point(thelink.dp).Marker = Marker
                NListed = NListed + 1
                LinkList(NListed) = thelink.dp
                ipos = 1
            End If
            If thelink.tipo Mod 100 = 8 Then thelink.Extension = thelink.Extension * CorreRate 'tipo8 = corredor de ônibus
            candidist = SP.distance + thelink.Extension
            If candidist < Point(thelink.dp).distance Then
                Point(thelink.dp).distance = candidist
                Point(thelink.dp).LastLink = SP.ilinkdaqui(i)
            End If
        End If
    Next i
    LinkList(NListed + 1) = 0
End Function
Function Spread_TO_Point(ipoint As Long) As Long
'Returns the point that was closed in the previous round... expands from there... and marks that point with -marker
'COM RESTRIÇÃO DE LINKS A PÉ (p.m.k) K, Links "bi" com peso CorreRate
'Presume que a extensão dos links é sempre positiva e que correrate é positiva
    Dim P As Point_type  'origin point
    Dim SP As Point_type 'spread point
    Dim ido As Point_type 'variable destiny point
    Dim thelink As Link_type
    P = Point(ipoint)
    Spread_TO_Point = 0
    Bestdistance = 1000000
    NListed = 0
    Do While Point(LinkList(NListed + 1)).Marker = P.Marker
        If Point(LinkList(NListed + 1)).distance <= Bestdistance Then
            Bestdistance = Point(LinkList(NListed + 1)).distance
            Spread_TO_Point = LinkList(NListed + 1)
            Position = NListed + 1
        End If
        NListed = NListed + 1
    Loop
    If Spread_TO_Point = 0 Then Exit Function
    SP = Point(Spread_TO_Point)
    LinkList(Position) = LinkList(NListed)
    LinkList(NListed) = 0
    NListed = NListed - 1
    For i = 1 To SP.nlinkspraka
        thelink = Link(SP.ilinkpraka(i))
        If (thelink.modes = "p" And thelink.tipo Mod 1000 = 20) Or (thelink.modes <> "k" And thelink.modes <> "m" And thelink.modes <> "p" And thelink.IsM2 <> 1 And thelink.tipo Mod 1000 < 600 And thelink.tipo Mod 1000 <> 19) Then
            ido = Point(thelink.op)
            If ido.Marker <> P.Marker Then
                Point(thelink.op).distance = 1000000
                Point(thelink.op).Marker = Marker
                NListed = NListed + 1
                LinkList(NListed) = thelink.op
                ipos = 1
            End If
            If thelink.tipo Mod 100 = 8 Then thelink.Extension = thelink.Extension * CorreRate 'tipo8 = corredor de ônibus
            candidist = SP.distance + thelink.Extension
            If candidist < Point(thelink.op).distance Then
                Point(thelink.op).distance = candidist
                Point(thelink.op).LastLink = SP.ilinkpraka(i)
            End If
        End If
    Next i
    LinkList(NListed + 1) = 0
End Function
