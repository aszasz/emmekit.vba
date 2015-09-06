Attribute VB_Name = "Mod_A4_Sort"
'To use this module;
'Redim AlfatoSort or NumtoSort (or PointList) to the size you want
'Insert the values you want to sort on the vector AlfatoSort or NumtoSort (or PointList)
'Call AlfaHeapSort or NumHeapSort and the vector will be sorted
'till the parameter NtoSort... NtoSort is the Length of the vetor will be sorted
Public NtoSort As Long
Public AlfaToSort() As String
Public NumToSort() As Single
Sub TripHeapSortbyTime(irg As Integer)
    'Ajusta Route_Group(iRG).trip() com base em Trip().arrival_time(1)... hora no primeiro ponto
    Dim i As Integer
    Dim Triptemp As Long
    For i = RouteGroup(irg).Ntrips / 2 To 1 Step -1
        Call TripAdjust(i, irg, RouteGroup(irg).Ntrips) '  into heap
    Next i
    For i = RouteGroup(irg).Ntrips - 1 To 1 Step -1
        Triptemp = RouteGroup(irg).itrip(i + 1)
        RouteGroup(irg).itrip(i + 1) = RouteGroup(irg).itrip(1)
        RouteGroup(irg).itrip(1) = Triptemp
        Call TripAdjust(1, irg, i)
    Next i
End Sub
Sub TripAdjust(i As Integer, irg As Integer, N As Integer)
    Dim j As Long
    Dim TripAux As Long
    TripAux = RouteGroup(irg).itrip(i)
    j = 2 * i
    While j <= N
        If j < N Then
            If trip(RouteGroup(irg).itrip(j)).arrival_time(1) < trip(RouteGroup(irg).itrip(j + 1)).arrival_time(1) Then
                j = j + 1
            End If
        End If
        If trip(TripAux).arrival_time(1) >= trip(RouteGroup(irg).itrip(j)).arrival_time(1) Then
           Exit Sub
        End If
        RouteGroup(irg).itrip(Int(j / 2)) = RouteGroup(irg).itrip(j)
        RouteGroup(irg).itrip(j) = TripAux
        j = 2 * j
    Wend
End Sub
Sub LinkHeapSortbyDistance(ReferencePoint As Long, Optional IsLatLong As Boolean = False)
    'Ajusta LinkList em ordem de distância do ReferencePoint
    Dim i As Long
    Dim Linktemp As Long
    For i = NtoSort / 2 To 1 Step -1
        Call LinkAdjust(i, NtoSort, ReferencePoint, IsLatLong) ' Convert PointList into heap
    Next i
    
    For i = NtoSort - 1 To 1 Step -1
        Linktemp = LinkList(i + 1)
        LinkList(i + 1) = LinkList(1)
        LinkList(1) = Linktemp
        Call LinkAdjust(1, i, ReferencePoint, IsLatLong)
    Next i
End Sub
Sub LinkAdjust(i As Long, N As Long, ReferencePoint As Long, IsLatLong As Boolean)
    Dim j As Long
    Dim LinkAux As Long
    LinkAux = LinkList(i)
    j = 2 * i
    While j <= N
        If j < N Then
            If Point_to_Link_Distance(ReferencePoint, LinkList(j), IsLatLong) < Point_to_Link_Distance(ReferencePoint, LinkList(j + 1), IsLatLong) Then
                j = j + 1
            End If
        End If
        If Point_to_Link_Distance(ReferencePoint, LinkAux, IsLatLong) >= Point_to_Link_Distance(ReferencePoint, LinkList(j), IsLatLong) Then
           Exit Sub
        End If
        LinkList(Int(j / 2)) = LinkList(j)
        LinkList(j) = LinkAux
        j = 2 * j
    Wend
'     PointList(Int(j / 2)) = PointAux
End Sub

Sub AlfaHeapSort()
    Dim i As Long
    Dim alfatempo As String
    For i = NtoSort / 2 To 1 Step -1
        Call Adjust(i, NtoSort) ' Convert AlfatoSort into heap
    Next i
    
    For i = NtoSort - 1 To 1 Step -1
        alfatempo = AlfaToSort(i + 1)
        AlfaToSort(i + 1) = AlfaToSort(1)
        AlfaToSort(1) = alfatempo
        Call Adjust(1, i)
    Next i
End Sub
Sub Adjust(i As Long, N As Long)
    Dim j As Long
    Dim AlfaAux As String
    AlfaAux = AlfaToSort(i)
    j = 2 * i
    While j <= N
        If j < N Then
        If AlfaToSort(j) < AlfaToSort(j + 1) Then
           j = j + 1
        End If
        End If
        If AlfaAux >= AlfaToSort(j) Then
           Exit Sub
        End If
        AlfaToSort(Int(j / 2)) = AlfaToSort(j)
        AlfaToSort(j) = AlfaAux
        j = 2 * j
     Wend
'     alfatosort(Int(j / 2)) = alfaaux
End Sub
Sub PointHeapSortbyDistance(ReferencePoint As Long, Optional IsLatLong As Boolean = False)
    'Ajusta PointList em ordem de distância do ReferencePoint
    Dim i As Long
    Dim Pointtemp As Long
    For i = NtoSort / 2 To 1 Step -1
        Call PointAdjust(i, NtoSort, ReferencePoint, IsLatLong) ' Convert PointList into heap
    Next i
    
    For i = NtoSort - 1 To 1 Step -1
        Pointtemp = PointList(i + 1)
        PointList(i + 1) = PointList(1)
        PointList(1) = Pointtemp
        Call PointAdjust(1, i, ReferencePoint, IsLatLong)
    Next i
End Sub
Sub PointAdjust(i As Long, N As Long, ReferencePoint As Long, IsLatLong As Boolean)
    Dim j As Long
    Dim PointAux As Long
    PointAux = PointList(i)
    j = 2 * i
    While j <= N
        If j < N Then
            If Distância(PointList(j), ReferencePoint, IsLatLong) < Distância(PointList(j + 1), ReferencePoint, IsLatLong) Then
                j = j + 1
            End If
        End If
        If Distância(PointAux, ReferencePoint, IsLatLong) >= Distância(PointList(j), ReferencePoint, IsLatLong) Then
           Exit Sub
        End If
        PointList(Int(j / 2)) = PointList(j)
        PointList(j) = PointAux
        j = 2 * j
    Wend
'     PointList(Int(j / 2)) = PointAux
End Sub
Sub NumHeapSort()
    Dim i As Long
    Dim NumTemp As Single
    For i = NtoSort / 2 To 1 Step -1
        Call NumAdjust(i, NtoSort) ' Convert NumtoSort into heap
    Next i
    For i = NtoSort - 1 To 1 Step -1
        NumTemp = NumToSort(i + 1)
        NumToSort(i + 1) = NumToSort(1)
        NumToSort(1) = NumTemp
        Call NumAdjust(1, i)
    Next i
End Sub
Sub NumAdjust(i As Long, N As Long)
    'Adjust makes sure that I move i to a place where it has two children    smaller than himself
    Dim j As Long
    Dim NumAux As Single
    NumAux = NumToSort(i) 'save it in the memory... it may be moved to j position
                          ' among one of its sons
    j = 2 * i             ' j is the position i is considered to be moved to
    While j <= N
        If j < N Then
            If NumToSort(j) < NumToSort(j + 1) Then
               j = j + 1
            End If
        End If  'j is the position of the largest of current i sons
        If NumAux >= NumToSort(j) Then
           Exit Sub
        End If
        NumToSort(Int(j / 2)) = NumToSort(j) 'switch i to j, j to i, switch position with the largest of your sons
        NumToSort(j) = NumAux
        j = 2 * j
     Wend
End Sub
Sub TesteHeapSort()
ReDim AlfaToSort(100)
    For i = 4 To 39
        AlfaToSort(i - 3) = Cells(i, 5)
    Next i
    NtoSort = 40
    AlfaHeapSort
    For i = 4 To 50
        Cells(i, 10) = AlfaToSort(i - 3)
    Next i
End Sub
Sub Teste()
ReDim NumToSort(9)
    For i = 71 To 79
        NumToSort(i - 70) = Cells(i, 5)
    Next i
    NtoSort = 9
    NumHeapSort
    For i = 71 To 79
        Cells(i, 6) = NumToSort(i - 70)
    Next i

End Sub
Sub STATUS()
Static icol
If icol = 0 Then icol = 6
icol = icol + 1
    For i = 71 To 79
        Cells(i, icol) = NumToSort(i - 70)
    Next i
End Sub
