Attribute VB_Name = "Mod_A3_TextTools"
Global Splitword(1000) As String
Global SplitwordB(1000) As String
Function CountWords(nome As String) As Integer
    'Returns the number of words from nome, and let them be in global Splitword
    ' WilliamsString Module has a SplitA function that returns a Variant, that is a String vector
    Dim resto As String
        resto = Trim(nome)
        For j = 1 To 100
            prime = FirstSpace(resto)
            If prime = 1 Then Exit For
            Splitword(j) = Left(resto, prime - 1)
            resto = Trim(Mid(resto, prime + 1))
        Next j
        CountWords = j - 1
End Function
Function CountWordsB(nome As String) As Integer
    'Usa global SplitwordB, for when you want to use what is in Splitword
    Dim resto As String
        resto = Trim(nome)
        For j = 1 To 100
            prime = FirstSpace(resto)
            If prime = 1 Then Exit For
            SplitwordB(j) = Left(resto, prime - 1)
            resto = Trim(Mid(resto, prime + 1))
        Next j
        CountWordsB = j - 1
End Function
Function FirstSpace(nome As String) As Integer
        For j = 1 To Len(nome)
            letra = Mid(nome, j, 1)
            If letra = " " Then Exit For
        Next j
        FirstSpace = j
End Function
Function FirstChar(nome As String) As Integer
        For j = 1 To Len(nome)
            letra = Mid(nome, j, 1)
            If letra <> " " Then Exit For
        Next j
        FirstChar = j
End Function
Function CountStrings(nome As String, seq As String) As Integer
    'Usa global Splitword
    Dim resto As String
    resto = Trim(nome)
    comp = Len(seq)
    For j = 1 To 100
        prime = FirstString(resto, seq)
        Splitword(j) = Trim(Left(resto, prime - 1))
        resto = Mid(resto, prime + comp)
        If resto = "" Then Exit For
     Next j
     CountStrings = j
     Splitword(j + 1) = ""
End Function
Function CountStringsB(nome As String, seq As String) As Integer
    'Usa global SplitwordB
    Dim resto As String
    resto = Trim(nome)
    comp = Len(seq)
    For j = 1 To 100
        prime = FirstString(resto, seq)
        SplitwordB(j) = Trim(Left(resto, prime - 1))
        resto = Mid(resto, prime + comp)
        If resto = "" Then Exit For
     Next j
     CountStringsB = j
End Function
Function FirstString(nome As String, seq As String) As Integer
    comp = Len(seq)
    For j = 1 To Len(nome)
        letra = Mid(nome, j, comp)
        If letra = seq Then Exit For
    Next j
    FirstString = j
End Function
Function CountStringsSpecial(nome As String) As Integer
    'Usa global Splitword
    Dim resto As String
    resto = Trim(nome)
    For j = 1 To 100
        prime = FirstStringSpecial(resto, "-")
        If prime = 1 Then Exit For
        Splitword(j) = Trim(Left(resto, prime - 1))
        resto = Trim(Mid(resto, prime + 1))
     Next j
     CountStringsSpecial = j - 1
End Function
Function FirstStringSpecial(nome As String, seq As String) As Integer
    For j = 1 To Len(nome)
        letra = Mid(nome, j, 1)
        If letra = "/" Or letra = "-" Then Exit For
    Next j
    FirstStringSpecial = j
End Function
'Return all the text before the first non-capital-letter
Function BeforeFirstMin(nome As String) As String
    For i = 1 To Len(nome)
        letra = Mid(nome, i, 1)
        If Asc(letra) >= Asc("a") And Asc(letra) <= Asc("z") Then
            BeforeFirstMin = Left(nome, i - 1): Exit Function
        End If
    Next i
    BeforeFirstMin = nome
End Function
Function ClearName(nome As String) As String
    ClearName = ""
    Candi = Trim(Format(nome, ">"))
    letra = " "
    For j = 1 To Len(Candi)
        letrant = letra
        letra = Mid(Candi, j, 1)
        If letra = "Ã" Then letra = "A"
        If letra = "À" Then letra = "A"
        If letra = "Á" Then letra = "A"
        If letra = "Â" Then letra = "A"
        If letra = "È" Then letra = "E"
        If letra = "É" Then letra = "E"
        If letra = "Ê" Then letra = "E"
        If letra = "Î" Then letra = "E"
        If letra = "Í" Then letra = "I"
        If letra = "Ì" Then letra = "I"
        If letra = "Õ" Then letra = "O"
        If letra = "Ó" Then letra = "O"
        If letra = "Ò" Then letra = "O"
        If letra = "Ô" Then letra = "O"
        If letra = "Ù" Then letra = "O"
        If letra = "Ú" Then letra = "U"
        If letra = "Û" Then letra = "U"
        If letra = "Ç" Then letra = "C"
        If letra = "Ñ" Then letra = "N"
        If letra = "Ý" Then letra = "Y"
        If letra = "." Or letra = "," Or letra = ";" Or letra = ">" Or letra = "*" Or letra = "+" Or letra = "&" Or letra = "<" Or letra = "-" Or letra = """" Or letra = "(" Or letra = ")" Or letra = "[" Or letra = "]" Or letra = "|" Or letra = "/" Then letra = " "
        If letra = "." Then letra = " "
        If letra = "´" Or letra = "`" Or letra = "'" Or letra = """" Then letra = ""
'        If IsNumeric(letra) Or letra <> letrant Then ClearName = ClearName & letra
        If Not (letra = " " And letrant = " ") Then ClearName = ClearName & letra
        If letra = "" Then letra = letrant
    Next j
End Function
Function ReplaceANDtoPLUS(nome As String) As String
    ReplaceANDtoPLUS = ""
    Candi = Trim(Format(nome, ">"))
    letra = " "
    For j = 1 To Len(Candi)
        letrant = letra
        letra = Mid(Candi, j, 1)
        If letra = "&" Then letra = "+"
        If Not (letra = " " And letrant = " ") Then ReplaceANDtoPLUS = ReplaceANDtoPLUS & letra
        If letra = "" Then letra = letrant
    Next j
End Function
Function ReverseafterPlus(nome As String) As String
    If CountStrings(nome, "+") = 2 Then
        ReverseafterPlus = Splitword(2) & " + " & Splitword(1)
    Else
        ReverseafterPlus = nome
    End If
End Function
Function ClearSigns(nome As String) As String
    ClearSigns = ""
    Candi = Trim(Format(nome, ">"))
    letra = " "
    For j = 1 To Len(Candi)
        letrant = letra
        letra = Mid(Candi, j, 1)
        If letra = "." Or letra = "," Or letra = ";" Or letra = ">" Or letra = "*" Or letra = "+" Or letra = "&" Or letra = "<" Or letra = "-" Or letra = """" Or letra = "(" Or letra = ")" Or letra = "[" Or letra = "]" Or letra = "|" Or letra = "/" Then letra = " "
        If letra = "." Then letra = " "
        If letra = "´" Or letra = "`" Or letra = "'" Or letra = """" Then letra = ""
'        If IsNumeric(letra) Or letra <> letrant Then ClearSigns = ClearSigns & letra
        If Not (letra = " " And letrant = " ") Then ClearSigns = ClearSigns & letra
        If letra = "" Then letra = letrant
    Next j
End Function
Function CutRepeatedLetters(nome As String) As String
    CutRepeatedLetters = ""
    Candi = Trim(Format(nome, ">"))
    N = Len(Candi)
    letra = " "
    For j = 1 To N
        letrant = letra
        letra = Mid(Candi, j, 1)
        parletra = Mid(Candi, j, 2)
        triletra = Mid(Candi, j, 3)
        If letra = "Y" Then letra = "I"
        If parletra = "OE" Then
            CutRepeatedLetters = CutRepeatedLetters & "U"
            j = j + 1
            letra = "U"
        ElseIf parletra = "DH" Then
            CutRepeatedLetters = CutRepeatedLetters & "D"
            j = j + 1
        ElseIf triletra = "AHM" Then
            CutRepeatedLetters = CutRepeatedLetters & "ACHM"
            j = j + 2
            letra = "M"
        ElseIf IsNumeric(letra) Or letra <> letrant Then
            CutRepeatedLetters = CutRepeatedLetters & letra
        End If
    Next j
End Function
Function JoinSingleLetterWords(nome As String) As String
    N = CountWords(nome)
    JoinSingleLetterWords = Splitword(1)
    For i = 1 + 1 To N
        If Len(Splitword(i)) = 1 And Len(Splitword(i - 1)) = 1 Then
            JoinSingleLetterWords = JoinSingleLetterWords & Splitword(i)
        Else
            JoinSingleLetterWords = JoinSingleLetterWords & " " & Splitword(i)
        End If
    Next i
End Function
Function ClearParenthesis(nome As String, Optional inside As String) As String
    isBetWeen = False
    ClearParenthesis = ""
    inside = ""
    N = Len(nome)
    For j = 1 To N
        letra = Mid(nome, j, 1)
        If letra = "(" Then isBetWeen = True
        If Not isBetWeen Then
            ClearParenthesis = ClearParenthesis & letra
        Else
            inside = inside & letra
        End If
        If letra = ")" And isBetWeen Then
            isBetWeen = False
            docut = False
            If Right(ClearParenthesis, 1) = " " Then
                If j = N Then
                    docut = True
                Else
                   If Mid(nome, j + 1, 1) = " " Then docut = True
                End If
                If docut Then ClearParenthesis = Left(ClearParenthesis, Len(ClearParenthesis) - 1)
            Else
                If j < N Then
                   If Mid(nome, j + 1, 1) <> " " Then ClearParenthesis = ClearParenthesis & " "
                End If
            End If
        End If
    Next j
End Function
'Indonesian for NORTH EAST WEST SOUTH = UTARA BARAT TIMUR SELATAN
Function RemoveNEWS(nome As String, Optional inside As String) As String
    N = CountWords(nome)
    RemoveNEWS = ""
    For i = 1 To N
        If Splitword(i) = "BARAT" Then Splitword(i) = ""
        If Splitword(i) = "UTARA" Then Splitword(i) = ""
        If Splitword(i) = "TIMUR" Then Splitword(i) = ""
        If Splitword(i) = "SELATAN" Then Splitword(i) = ""
    Next i
    RemoveNEWS = Splitword(1)
    If N = 0 Then RemoveNEWS = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then
            RemoveNEWS = RemoveNEWS & Splitword(i)
        Else
            RemoveNEWS = RemoveNEWS & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
End Function
'Mostly prepositions
Function ClearRelativeToWords(nome As String, Optional inside As String) As String
    N = CountWords(nome)
    For i = 1 To N
        If Splitword(i) = "DEKAT" Then Exit For
        If Splitword(i) = "SEKITAR" Then Exit For
        If Splitword(i) = "DEPAN" Then Exit For
        If Splitword(i) = "BELAKANG" Then Exit For
        If Splitword(i) = "SEBELUM" Then Exit For
        If Splitword(i) = "SIMPANG" Then Exit For
        If Splitword(i) = "SAMPING" Then Exit For
        If Splitword(i) = "ATAU" Then Exit For
        If Splitword(i) = "DARI" Then Exit For
        If Splitword(i) = "KE" Then Exit For
        If Splitword(i) = "DKT" Then Exit For
        If Splitword(i) = "BLKG" Then Exit For
        If Splitword(i) = "BLK" Then Exit For
        If Splitword(i) = "SEBELAH" Then Exit For
        If Splitword(i) = "SIMPANGAN" Then Exit For
        If Splitword(i) = "SEBERANG" Then Exit For
        If Splitword(i) = "KIRI" And Splitword(i + 1) <> "INDAH" Then Exit For
        If Splitword(i) = "KANAN" And Splitword(i - 1) <> "KRISTEN" Then Exit For
    Next i
    ClearRelativeToWords = ""
    inside = ""
    For itill = 1 To i - 1
        If Len(ClearRelativeToWords) = 0 Then
            ClearRelativeToWords = ClearRelativeToWords & Splitword(itill)
        Else
            ClearRelativeToWords = ClearRelativeToWords & " " & Splitword(itill)
        End If
    Next itill
    For itill = i + 1 To N
        If Len(inside) = 0 Then
            inside = inside & Splitword(itill)
        Else
            inside = inside & " " & Splitword(itill)
        End If
    Next itill
End Function
Function ClearRelativeToBars(nome As String, Optional inside As String) As String
    'Usa global Splitword
    For j = 1 To Len(nome)
        If Mid(nome, j, 1) = "/" Or Mid(nome, j, 1) = "-" Then
            If j > 2 Then
                If Mid(nome, j - 2, 2) <> "AR" And Mid(nome, j - 2, 2) <> "AL" And Mid(nome, j - 2, 2) <> "AT" Then
                    ClearRelativeToBars = Trim(Left(nome, j - 1))
                    inside = Trim(Mid(nome, j + 1))
                    Exit Function
                End If
            Else
                ClearRelativeToBars = Trim(Left(nome, j - 1))
                inside = Trim(Mid(nome, j + 1))
                Exit Function
            End If
        End If
     Next j
     ClearRelativeToBars = nome
     inside = ""
End Function
Function CutTitles(nome As String) As String
    N = CountWords(nome)
    CutTitles = ""
    For i = 1 To N
        If Splitword(i) = "HAJI" Then Splitword(i) = ""
        If Splitword(i) = "DR" Then Splitword(i) = ""
        If Splitword(i) = "LET" Then Splitword(i) = ""
        If Splitword(i) = "PROF" Then Splitword(i) = ""
        If Splitword(i) = "JEND" Then Splitword(i) = ""
        If Splitword(i) = "MAYOR" Then Splitword(i) = ""
        If Splitword(i) = "KAPTEN" Then Splitword(i) = ""
        If Splitword(i) = "PANGERAN" Then Splitword(i) = ""
        If Splitword(i) = "TAMAN" Then Splitword(i) = ""
        If Splitword(i) = "BLOK" Then Splitword(i) = ""
    Next i
    CutTitles = Splitword(1)
    If N = 0 Then CutTitles = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If (LEN2 = 1 And LEN1 = 1) Or LEN1 = 0 Then
            CutTitles = CutTitles & Splitword(i)
        Else
            CutTitles = CutTitles & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
'    If Len(CutTitles) < 2 Then CutTitles = ""
End Function
Function ClearNameComma(nome As String) As String
    comma = 0
    ClearNameComma = ""
    Candi = Trim(Format(nome, ">"))
    letra = " "
    For j = 1 To Len(Candi)
        letrant = letra
        letra = Mid(Candi, j, 1)
        If letra = "Ã" Then letra = "A"
        If letra = "À" Then letra = "A"
        If letra = "Á" Then letra = "A"
        If letra = "Â" Then letra = "A"
        If letra = "È" Then letra = "E"
        If letra = "É" Then letra = "E"
        If letra = "Ê" Then letra = "E"
        If letra = "Î" Then letra = "E"
        If letra = "Í" Then letra = "I"
        If letra = "Ì" Then letra = "I"
        If letra = "Õ" Then letra = "O"
        If letra = "Ó" Then letra = "O"
        If letra = "Ò" Then letra = "O"
        If letra = "Ô" Then letra = "O"
        If letra = "Ù" Then letra = "O"
        If letra = "Ú" Then letra = "U"
        If letra = "Û" Then letra = "U"
        If letra = "Ç" Then letra = "C"
        If letra = "Ñ" Then letra = "N"
        If letra = "Ý" Then letra = "Y"
'        If letra = "." Or letra = "," Or letra = ";" Or letra = ">" Or letra = "*" Or letra = "+" Or letra = "&" Or letra = "<" Or letra = "-" Or letra = """" Or letra = "(" Or letra = ")" Or letra = "[" Or letra = "]" Or letra = "|" Or letra = "/" Then letra = " "
        If letra = "." Then letra = " "
        If letra = "," Then
            If comma <> 0 And Len(ClearNameComma) > comma + 1 Then ClearNameComma = Trim(Mid(ClearNameComma, comma + 1)) & " " & Trim(Left(ClearNameComma, comma))
            comma = Len(ClearNameComma)
            letra = " "
            letrant = ""
        End If
        If letra = "´" Or letra = "`" Or letra = "'" Or letra = """" Then letra = ""
'        If IsNumeric(letra) Or letra <> letrant Then ClearName = ClearName & letra
        If Not (letra = " " And letrant = " ") Then ClearNameComma = ClearNameComma & letra
        If letra = "" Then letra = letrant
    Next j
    If comma <> 0 Then
        If Len(ClearNameComma) > comma + 1 Then ClearNameComma = Trim(Mid(ClearNameComma, comma + 1)) & " " & Trim(Left(ClearNameComma, comma))
    End If
End Function
Function ReplaceWOrd(NametoChange As String, arg1 As String, arg2 As String) As String
    N = CountWords(NametoChange)
    ReplaceWOrd = ""
    For i = 1 To N
        If Splitword(i) = arg1 Then Splitword(i) = arg2
    Next i
    ReplaceWOrd = Splitword(1)
    If N = 0 Then ReplaceWOrd = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then '(LEN2 = 1 And LEN1 = 1) Or
            ReplaceWOrd = ReplaceWOrd & Splitword(i)
        Else
            ReplaceWOrd = ReplaceWOrd & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
'    If Len(CutTitles) < 2 Then CutTitles = ""
End Function
Function ReplaceRomans(NametoChange As String) As String
    N = CountWords(NametoChange)
    ReplaceRomans = ""
    For i = 1 To N
        If Splitword(i) = "I" Then Splitword(i) = "1"
        If Splitword(i) = "II" Then Splitword(i) = "2"
        If Splitword(i) = "III" Then Splitword(i) = "3"
        If Splitword(i) = "IV" Then Splitword(i) = "4"
        If Splitword(i) = "V" Then Splitword(i) = "5"
        If Splitword(i) = "VI" Then Splitword(i) = "6"
        If Splitword(i) = "VII" Then Splitword(i) = "7"
        If Splitword(i) = "VIII" Then Splitword(i) = "8"
        If Splitword(i) = "IX" Then Splitword(i) = "9"
        If Splitword(i) = "X" Then Splitword(i) = "10"
        If Splitword(i) = "XI" Then Splitword(i) = "11"
    Next i
    ReplaceRomans = Splitword(1)
    If N = 0 Then ReplaceRomans = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then '(LEN2 = 1 And LEN1 = 1) Or
            ReplaceRomans = ReplaceRomans & Splitword(i)
        Else
            ReplaceRomans = ReplaceRomans & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
End Function
Function ReplaceSiglas(NametoChange As String) As String
    N = CountWords(NametoChange)
    ReplaceSiglas = ""
    For i = 1 To N
'        If Splitword(i) = "RW" Then Splitword(i) = "RAWA"
'        If Splitword(i) = "RWA" Then Splitword(i) = "RAWA"
'        If Splitword(i) = "RAW" Then Splitword(i) = "RAWA"
'        If Splitword(i) = "PD" Then Splitword(i) = "PONDOK"
'        If Splitword(i) = "PDK" Then Splitword(i) = "PONDOK"
'        If Splitword(i) = "PON" Then Splitword(i) = "PONDOK"
'        If Splitword(i) = "PONDK" Then Splitword(i) = "PONDOK"
'        If Splitword(i) = "PS" Then Splitword(i) = "PASAR"
'        If Splitword(i) = "PAS" Then Splitword(i) = "PASAR"
'        If Splitword(i) = "PSR" Then Splitword(i) = "PASAR"
'        If Splitword(i) = "PASR" Then Splitword(i) = "PASAR"
'        If Splitword(i) = "PSAR" Then Splitword(i) = "PASAR"
'        If Splitword(i) = "TN" Then Splitword(i) = "TANAH"
'        If Splitword(i) = "TM" Then Splitword(i) = "TAMAN"
'        If Splitword(i) = "TAM" Then Splitword(i) = "TAMAN"
'        If Splitword(i) = "TAMN" Then Splitword(i) = "TAMAN"
'        If Splitword(i) = "TMN" Then Splitword(i) = "TAMAN"
'        If Splitword(i) = "TJ" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "TANJ" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "TANJG" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "TJNG" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "TNJG" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "TNJUNG" Then Splitword(i) = "TANJUNG"
'        If Splitword(i) = "SEL" Then Splitword(i) = "SELATAN"
'        If Splitword(i) = "SLTN" Then Splitword(i) = "SELATAN"
'        If Splitword(i) = "SELTN" Then Splitword(i) = "SELATAN"
'        If Splitword(i) = "BRT" Then Splitword(i) = "BARAT"
'        If Splitword(i) = "BART" Then Splitword(i) = "BARAT"
'        If Splitword(i) = "BRAT" Then Splitword(i) = "BARAT"
'        If Splitword(i) = "BAR" Then Splitword(i) = "BARAT"
'        If Splitword(i) = "UT" Then Splitword(i) = "UTARA"
'        If Splitword(i) = "UTA" Then Splitword(i) = "UTARA"
'        If Splitword(i) = "UTR" Then Splitword(i) = "UTARA"
'        If Splitword(i) = "UTAR" Then Splitword(i) = "UTARA"
'        If Splitword(i) = "PUS" Then Splitword(i) = "PUSAT"
'        If Splitword(i) = "PST" Then Splitword(i) = "PUSAT"
'        If Splitword(i) = "TIM" Then Splitword(i) = "TIMUR"
'        If Splitword(i) = "TIMR" Then Splitword(i) = "TIMUR"
'        If Splitword(i) = "TMR" Then Splitword(i) = "TIMUR"
'        If Splitword(i) = "KP" Then Splitword(i) = "KAMPUNG"
'        If Splitword(i) = "KPG" Then Splitword(i) = "KAMPUNG"
'        If Splitword(i) = "KAMP" Then Splitword(i) = "KAMPUNG"
'        If Splitword(i) = "KAMPG" Then Splitword(i) = "KAMPUNG"
'        If Splitword(i) = "KLP" Then Splitword(i) = "KELAPA"


'If Splitword(i) = "KEC " Then Splitword(i) = ""
'If Splitword(i) = "RT  " Then Splitword(i) = ""
'If Splitword(i) = "RW" Then Splitword(i) = ""
'If Splitword(i) = "UNIV" Then Splitword(i) = "UNIVERSITAS"
'If Splitword(i) = "TAMRIN" Then Splitword(i) = "THAMRIN"
'If Splitword(i) = "TERM" Then Splitword(i) = "TERMINAL"
'If Splitword(i) = "STATSIUN" Then Splitword(i) = "STASIUN"
'If Splitword(i) = "SMKN" Then Splitword(i) = "SMK"
'If Splitword(i) = "SIAFIAH" Then Splitword(i) = "SIAFIA"
'If Splitword(i) = "SERONG" Then Splitword(i) = "SERPONG"
'If Splitword(i) = "RSCM" Then Splitword(i) = "RS CMS"
'If Splitword(i) = "PROIEK" Then Splitword(i) = "PRIOK"
'If Splitword(i) = "POM" Then Splitword(i) = "POMPA"
'If Splitword(i) = "PRANCIS" Then Splitword(i) = "PERANCIS"
'If Splitword(i) = "HOLAND" Then Splitword(i) = "NETHERLANDS"
'If Splitword(i) = "MANSUR" Then Splitword(i) = "MANSIUR"
'If Splitword(i) = "MAIJEN" Then Splitword(i) = "MAIOR JENDRAL"
'If Splitword(i) = "MAIJEND" Then Splitword(i) = "MAIOR JENDRAL"
'If Splitword(i) = "LETJEN" Then Splitword(i) = "LETNAM JENDRAL"
'If Splitword(i) = "LETJEND" Then Splitword(i) = "LETNAM JENDRAL"
'If Splitword(i) = "LET" Then Splitword(i) = "LETNAM"
If Splitword(i) = "LATUMENTEN" Then Splitword(i) = "LATUMETEN"
'If Splitword(i) = "LAKS" Then Splitword(i) = "LAKSAMANA"
'If Splitword(i) = "KOMP" Then Splitword(i) = "KOMPLEK"
'If Splitword(i) = "KOMPL" Then Splitword(i) = "KOMPLEK"
'If Splitword(i) = "KOMPLEKS" Then Splitword(i) = "KOMPLEK"
'If Splitword(i) = "PERUMAHAN" Then Splitword(i) = "KOMPLEK"
'If Splitword(i) = "KALIMALANG" Then Splitword(i) = "KALI MALANG"
'If Splitword(i) = "JENGOT" Then Splitword(i) = "JENDRAL GATOT SUBROTO"
'If Splitword(i) = "JEND" Then Splitword(i) = "JENDRAL"
'If Splitword(i) = "JEMB" Then Splitword(i) = "JEMBATAN"
'If Splitword(i) = "JL" Then Splitword(i) = "JALAN"
'If Splitword(i) = "JLN" Then Splitword(i) = "JALAN"
'If Splitword(i) = "JAKPUS" Then Splitword(i) = "JAKARTA PUSAT"
'If Splitword(i) = "JAKBAR" Then Splitword(i) = "JAKARTA BARAT"
'If Splitword(i) = "IR" Then Splitword(i) = "INSINIUR"
'If Splitword(i) = "IMPRES" Then Splitword(i) = "INPRES"
'If Splitword(i) = "GDN" Then Splitword(i) = "GEDUNG"
'If Splitword(i) = "GED" Then Splitword(i) = "GEDUNG"
'If Splitword(i) = "DIPONOGORO" Then Splitword(i) = "DIPONEGORO"
'If Splitword(i) = "CITEREUP" Then Splitword(i) = "CITEUREUP"
'If Splitword(i) = "CITAIEM" Then Splitword(i) = "CITAIAM"
'If Splitword(i) = "CILEDUK" Then Splitword(i) = "CILEDUG"
'If Splitword(i) = "CASABLANGCA" Then Splitword(i) = "CASABLANCA"
'If Splitword(i) = "CAROLOUS" Then Splitword(i) = "CAROLUS"
'If Splitword(i) = "CAREFOURE" Then Splitword(i) = "CAREFOUR"
'If Splitword(i) = "AT-TAQWA" Then Splitword(i) = "ATAQWA"
'If Splitword(i) = "ASAFIAH" Then Splitword(i) = "AS SIAFIA"
'If Splitword(i) = "ASKES" Then Splitword(i) = "AKSES"
'If Splitword(i) = "XVI" Then Splitword(i) = "16"
'If Splitword(i) = "XI" Then Splitword(i) = "11"
'If Splitword(i) = "100M" Then Splitword(i) = ""
'If Splitword(i) = "10M" Then Splitword(i) = ""
'If Splitword(i) = "10NO" Then Splitword(i) = ""
'If Splitword(i) = "10RT" Then Splitword(i) = ""
'If Splitword(i) = "15M" Then Splitword(i) = ""
'If Splitword(i) = "200M" Then Splitword(i) = ""
'If Splitword(i) = "ZUSUKI" Then Splitword(i) = "SUZUKI"
'If Splitword(i) = "5SUNTER" Then Splitword(i) = "SUNTER"
'If Splitword(i) = "21KEMAIORAN" Then Splitword(i) = "KEMAIORAN"
'If Splitword(i) = "2KEBAIORAN" Then Splitword(i) = "KEBAIORAN"
'If Splitword(i) = "3CENGKARENG" Then Splitword(i) = "CENGKARENG"
'If Splitword(i) = "6CAWANG" Then Splitword(i) = "CAWANG"
'If Splitword(i) = "1DAHLIA" Then Splitword(i) = "DAHLIA"
'If Splitword(i) = "1CENGKARENG" Then Splitword(i) = "CENGKARENG"
'If Splitword(i) = "JAL" Then Splitword(i) = "JALAN"
'If Splitword(i) = "KLAPA" Then Splitword(i) = "KELAPA"
'If Splitword(i) = "JEN" Then Splitword(i) = "JENDRAL"
'If Splitword(i) = "DEPO" Then Splitword(i) = "DEPOK"
'If Splitword(i) = "SERATUS" Then Splitword(i) = "100"
'If Splitword(i) = "AL-" Then Splitword(i) = "AL"
'If Splitword(i) = "APARTEMENT" Then Splitword(i) = "APARTEMEN"
'If Splitword(i) = "BADUNG" Then Splitword(i) = "BANDUNG"
'If Splitword(i) = "BOUGENVIL" Then Splitword(i) = "BOUGENVILE"
'If Splitword(i) = "BRIGJEN" Then Splitword(i) = "BRIGADIR JENDRAL"
'If Splitword(i) = "CIKULIR" Then Splitword(i) = "CIPULIR"
'If Splitword(i) = "CIRENDE" Then Splitword(i) = "CIRENDEU"
'If Splitword(i) = "DEPO" Then Splitword(i) = "DEPOK"
'If Splitword(i) = "DOKTER" Then Splitword(i) = "DR"
'If Splitword(i) = "ISTIQLAL" Then Splitword(i) = "ISTIQLALA"
'If Splitword(i) = "JAKSEL" Then Splitword(i) = "JAKARTA SELATAN"
'If Splitword(i) = "KALIMATI" Then Splitword(i) = "KALI MATI"
If Splitword(i) = "LETNAM" Then Splitword(i) = "LETNAN"
'If Splitword(i) = "PALSIGUNUNG" Then Splitword(i) = "PALSI GUNUNG"
'If Splitword(i) = "SUPARMAN" Then Splitword(i) = "S PARMAN"
'If Splitword(i) = "TANJUNGPRIUK" Then Splitword(i) = "TANJUNG PRIOK"
'If Splitword(i) = "TELUGONG" Then Splitword(i) = "TELUK GONG"
''If Splitword(i) = "SMAN"

    Next i
    ReplaceSiglas = Splitword(1)
    If N = 0 Then ReplaceSiglas = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then '(LEN2 = 1 And LEN1 = 1) Or
            ReplaceSiglas = ReplaceSiglas & Splitword(i)
        Else
            ReplaceSiglas = ReplaceSiglas & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
End Function
Function ReplaceMD(NametoChange As String) As String
    N = CountWords(NametoChange)
    ReplaceMD = ""
    For i = 1 To N
        If Splitword(i) = "CIR" Then Splitword(i) = "CIRCLE"
        If Splitword(i) = "BLDG" Then Splitword(i) = "BUILDING"
        If Splitword(i) = "BD" Then Splitword(i) = "BUILDING"
        If Splitword(i) = "BDG" Then Splitword(i) = "BUILDING"
        If Splitword(i) = "NORTH" Then Splitword(i) = ""
        If Splitword(i) = "EAST" Then Splitword(i) = ""
        If Splitword(i) = "WEST" Then Splitword(i) = ""
        If Splitword(i) = "SOUTH" Then Splitword(i) = ""
        If Splitword(i) = "N" Then Splitword(i) = ""
        If Splitword(i) = "E" Then Splitword(i) = ""
        If Splitword(i) = "W" Then Splitword(i) = ""
        If Splitword(i) = "S" Then Splitword(i) = ""
        If Splitword(i) = "CTR" Then Splitword(i) = "CENTER"
        If Splitword(i) = "COURT" Then Splitword(i) = "CENTER"
        If Splitword(i) = "CT" Then Splitword(i) = "CENTER"
        If Splitword(i) = "LA" Then Splitword(i) = "LANE"
        If Splitword(i) = "LN" Then Splitword(i) = "LANE"
        If Splitword(i) = ";" Then Splitword(i) = "+"
        If Splitword(i) = "TERR" Then Splitword(i) = "TE"
        If Splitword(i) = "TER" Then Splitword(i) = "TE"
        If Splitword(i) = "DRIVE" Then Splitword(i) = "DR"
        If Splitword(i) = "ROAD" Then Splitword(i) = "RD"
        If Splitword(i) = "PI" Then Splitword(i) = "PIKE"
        If Splitword(i) = "PK" Then Splitword(i) = "PIKE"
        If Splitword(i) = "AV" Then Splitword(i) = "AVE"
        If Splitword(i) = "AVENUE" Then Splitword(i) = "AVE"
        If Splitword(i) = "AVNUE" Then Splitword(i) = "AVE"
        If Splitword(i) = "PKWY" Then Splitword(i) = "PIKEWAY"
        If Splitword(i) = "PW" Then Splitword(i) = "PIKEWAY"
        If Splitword(i) = "PKW" Then Splitword(i) = "PIKEWAY"
        If Splitword(i) = "WY" Then Splitword(i) = "WAY"
        If Splitword(i) = "STREET" Then Splitword(i) = "ST"
        If Splitword(i) = "BOULEVARD" Then Splitword(i) = "BLVD"
        If Splitword(i) = "BVD" Then Splitword(i) = "BLVD"
        If Splitword(i) = "BV" Then Splitword(i) = "BLVD"
    Next i
    ReplaceMD = Splitword(1)
    If N = 0 Then ReplaceMD = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then '(LEN2 = 1 And LEN1 = 1) Or
            ReplaceMD = ReplaceMD & Splitword(i)
        Else
            ReplaceMD = ReplaceMD & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
End Function
Function CutNoRtRw(NametoChange As String) As String
    N = CountWords(NametoChange)
    CutNoRtRw = ""
    For i = 1 To N
        If Splitword(i) = "NO" Or Splitword(i) = "RT" Or Splitword(i) = "RW" Then
            N = i - 1
            Exit For
        End If
    Next i
    CutNoRtRw = Splitword(1)
    If N = 0 Then CutNoRtRw = ""
    LEN1 = Len(Splitword(1))
    For i = 2 To N
        LEN2 = Len(Splitword(i))
        If LEN1 = 0 Then '(LEN2 = 1 And LEN1 = 1) Or
            CutNoRtRw = CutNoRtRw & Splitword(i)
        Else
            CutNoRtRw = CutNoRtRw & " " & Splitword(i)
        End If
        LEN1 = LEN2
    Next i
End Function
Function ReplaceComma(nome As String) As String
    ReplaceComma = ""
    For j = 1 To Len(nome)
        If Mid(nome, j, 1) = "," Then
            ReplaceComma = ReplaceComma & "."
        Else
            ReplaceComma = ReplaceComma & Mid(nome, j, 1)
        End If
    Next j
End Function
Sub EmergencyLixo()
For i = 4 To 1540
'    If Cells(i, 9) <> "" Then Cells(i, 10) = Mid(Cells(i, 10), Len(Cells(i, 9)) + 2)
    A = CountStrings(Cells(i, 19), "/")
'    If IsNumeric(Splitword(1)) Then
        Cells(i, 17) = Splitword(1)
        Cells(i, 18) = Splitword(2)
'    End If
Next i
End Sub
