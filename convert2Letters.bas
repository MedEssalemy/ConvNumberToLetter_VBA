Attribute VB_Name = "convert2Letters"
'-------------------------------------------------------------------------
' Number to Text Converter
' Author: Mohamed Essalemy (mohamed.essalemy@gmail.com)
' Purpose: Converts numbers to their text representation with currency support
'-------------------------------------------------------------------------

Option Explicit

'-------------------------------------------------------------------------
'
'   Devise      =0   Aucune
'               =1   Euro             €
'               =2   Dollar           $
'               =3   Dirham marocaine DH
'
'   Langue      =0   Français
'               =1   Anglais
'
'   Casse       =0   Minuscule
'               =1   Majuscule en début de phrase
'               =2   Majuscule
'               =3   Majuscule en début de chaque mot
'
'   ZeroCent    =0   Ne mentionne pas les cents s'ils sont égal à 0
'               =1   Mentionne toujours les cents
'
'-------------------------------------------------------------------------
' Conversion limitée à 999 999 999 999 999 ou 9 999 999 999 999,99
' si le nombre contient plus de 2 décimales,il est arrondit à 3 décimales
'-------------------------------------------------------------------------

Function ConvNumberLetter(Nombre As Double, Optional Devise As Long = 0, _
                                 Optional Langue As Long = 0, _
                                 Optional Casse As Long = 0, _
                                 Optional ZeroCent As Long = 0) As String
    Dim dblEnt As Variant, byDec As Long
    Dim bNegatif As Boolean
    Dim strDev As String, strCentimes As String

    If Nombre < 0 Then
        bNegatif = True
        Nombre = Abs(Nombre)
    End If

    dblEnt = Int(Nombre)
    byDec = CInt((Nombre - dblEnt) * 100)
    If byDec = 0 Then
        If dblEnt > 999999999999999# Then
            ConvNumberLetter = "#TropGrand"
            Exit Function
        End If
    Else
        If dblEnt > 9999999999999.99 Then
            ConvNumberLetter = "#TropGrand"
            Exit Function
        End If
    End If

    Select Case Devise
        Case 0
            If byDec > 0 Then strDev = " virgule "
        Case 1
            strDev = " Euro"
            If dblEnt >= 1000000 And Right$(dblEnt, 6) = "000000" Then strDev = " d'Euro"
            If byDec > 0 Then strCentimes = strCentimes & " Cent"
            If byDec > 1 Then strCentimes = strCentimes & "s"

        Case 2
            strDev = " Dollar"
            If byDec > 0 Then strCentimes = strCentimes & " Cent"

        Case 3
            strDev = " Dirham"
            If dblEnt >= 1000000 And Right$(dblEnt, 6) = "000000" Then strDev = " dirhams,"
            If byDec > 0 Then strCentimes = strCentimes & " Centime"
            If byDec > 1 Then strCentimes = strCentimes & "s"
    End Select

    If dblEnt > 1 And Devise <> 0 Then strDev = strDev & "s,"
    strDev = strDev & " "
    If dblEnt = 0 Then
        ConvNumberLetter = "z" & Chr(233) & "ro " & strDev
    Else
        ConvNumberLetter = ConvNumEnt(CDbl(dblEnt), Langue) & strDev
    End If

    If byDec = 0 Then
        If Devise <> 0 Then
            If ZeroCent = 1 Then
                Select Case Devise
                    Case 0 To 2
                        ConvNumberLetter = ConvNumberLetter & "z" & Chr(233) & "ro " & "Cent"
                    Case 3
                        ConvNumberLetter = ConvNumberLetter & "z" & Chr(233) & "ro " & "Centimes"
                End Select
            End If
        End If
    Else
        If Devise = 0 Then
            ConvNumberLetter = ConvNumberLetter & ConvNumCent(byDec, Langue) & strCentimes
        Else
            ConvNumberLetter = ConvNumberLetter & ConvNumCent(byDec, Langue) & strCentimes
        End If
    End If

    ConvNumberLetter = Replace(ConvNumberLetter, "  ", " ")
    If Left(ConvNumberLetter, 1) = " " Then ConvNumberLetter = _
       Right(ConvNumberLetter, Len(ConvNumberLetter) - 1)
    If Right(ConvNumberLetter, 1) = " " Then ConvNumberLetter = _
       Left(ConvNumberLetter, Len(ConvNumberLetter) - 1)

    Select Case Casse
        Case 0
            ConvNumberLetter = LCase$(ConvNumberLetter)
        Case 1
            ConvNumberLetter = UCase$(Left$(ConvNumberLetter, 1)) & LCase(Right(ConvNumberLetter, Len(ConvNumberLetter) - 1))
        Case 2
            ConvNumberLetter = UCase$(ConvNumberLetter)
        Case 3
            ConvNumberLetter = Application.WorksheetFunction.Proper(ConvNumberLetter)
            If Devise = 3 Then _
               ConvNumberLetter = Replace(ConvNumberLetter, "€Uros", "€uros", , , vbTextCompare)
    End Select
End Function

Private Function ConvNumEnt(Nombre As Double, Langue As Long)
    Dim iTmp As Variant, dblReste As Double
    Dim strTmp As String
    Dim iCent As Long, iMille As Long, iMillion As Long
    Dim iMilliard As Long, iBillion As Long

    iTmp = Nombre - (Int(Nombre / 1000) * 1000)
    iCent = CInt(iTmp)
    ConvNumEnt = Nz(ConvNumCent(iCent, Langue))
    dblReste = Int(Nombre / 1000)
    If iTmp = 0 And dblReste = 0 Then Exit Function
    iTmp = dblReste - (Int(dblReste / 1000) * 1000)
    If iTmp = 0 And dblReste = 0 Then Exit Function
    iMille = CInt(iTmp)
    strTmp = ConvNumCent(iMille, Langue)

    Select Case iTmp
        Case 0
        Case 1
            strTmp = " mille "
        Case Else
            strTmp = strTmp & " mille "
    End Select

    If iMille = 0 And iCent > 0 Then ConvNumEnt = "et " & ConvNumEnt
    ConvNumEnt = Nz(strTmp) & ConvNumEnt
    dblReste = Int(dblReste / 1000)
    iTmp = dblReste - (Int(dblReste / 1000) * 1000)
    If iTmp = 0 And dblReste = 0 Then Exit Function
    iMillion = CInt(iTmp)
    strTmp = ConvNumCent(iMillion, Langue)

    Select Case iTmp
        Case 0
        Case 1
            strTmp = strTmp & " million "
        Case Else
            strTmp = strTmp & " millions "
    End Select

    If iMille = 1 Then ConvNumEnt = "et " & ConvNumEnt
    ConvNumEnt = Nz(strTmp) & ConvNumEnt
    dblReste = Int(dblReste / 1000)
    iTmp = dblReste - (Int(dblReste / 1000) * 1000)
    If iTmp = 0 And dblReste = 0 Then Exit Function
    iMilliard = CInt(iTmp)
    strTmp = ConvNumCent(iMilliard, Langue)

    Select Case iTmp
        Case 0
        Case 1
            strTmp = strTmp & " milliard "
        Case Else
            strTmp = strTmp & " milliards "
    End Select

    If iMillion = 1 Then ConvNumEnt = "et " & ConvNumEnt
    ConvNumEnt = Nz(strTmp) & ConvNumEnt
    dblReste = Int(dblReste / 1000)
    iTmp = dblReste - (Int(dblReste / 1000) * 1000)
    If iTmp = 0 And dblReste = 0 Then Exit Function
    iBillion = CInt(iTmp)
    strTmp = ConvNumCent(iBillion, Langue)

    Select Case iTmp
        Case 0
        Case 1
            strTmp = strTmp & " billion "
        Case Else
            strTmp = strTmp & " billions "
    End Select

    If iMilliard = 1 Then ConvNumEnt = "et " & ConvNumEnt
    ConvNumEnt = Nz(strTmp) & ConvNumEnt
End Function

Private Function ConvNumDizaine(Nombre As Long, Langue As Long, bDec As Boolean) As String
Dim TabUnit As Variant, TabDiz As Variant
Dim byUnit As Long, byDiz As Long
Dim strLiaison As String

    If bDec Then
        TabDiz = Array("zéro", "", "vingt", "trente", "quarante", "cinquante", _
                       "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    Else
        TabDiz = Array("", "", "vingt", "trente", "quarante", "cinquante", _
                       "soixante", "soixante", "quatre-vingt", "quatre-vingt")
    End If

    If Nombre = 0 Then
        TabUnit = Array("zéro")
    Else
        TabUnit = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", _
                        "huit", "neuf", "dix", "onze", "douze", "treize", "quatorze", "quinze", _
                        "seize", "dix-sept", "dix-huit", "dix-neuf")
    End If

    If Langue = 1 Then
        TabDiz(7) = "seventeen"
        TabDiz(8) = "eighteen"
        TabDiz(9) = "nineteen"
    End If

    byDiz = Int(Nombre / 10)
    byUnit = Nombre - (byDiz * 10)
    strLiaison = "-"
    If byUnit = 1 Then strLiaison = " et "

    Select Case byDiz
        Case 0
            strLiaison = " "
        Case 1
            byUnit = byUnit + 10
            strLiaison = ""
        Case 7
            If Langue = 0 Then byUnit = byUnit + 10
        Case 8
            If Langue <> 2 Then strLiaison = "-"
        Case 9
            If Langue = 0 Then
                byUnit = byUnit + 10
                strLiaison = "-"
            End If
    End Select

    ConvNumDizaine = TabDiz(byDiz)
    If TabUnit(byUnit) <> "" Then
        ConvNumDizaine = ConvNumDizaine & strLiaison & TabUnit(byUnit)
    Else
        ConvNumDizaine = ConvNumDizaine
    End If
End Function

Private Function ConvNumCent(Nombre As Long, Langue As Long) As String
Dim TabUnit As Variant
Dim byCent As Long, byReste As Long
Dim strReste As String

    TabUnit = Array("", "un", "deux", "trois", "quatre", "cinq", "six", "sept", _
                    "huit", "neuf", "dix")
    byCent = Int(Nombre / 100)
    byReste = Nombre - (byCent * 100)
    strReste = ConvNumDizaine(byReste, Langue, False)

    Select Case byCent
        Case 0
            ConvNumCent = strReste
        Case 1
            If byReste = 0 Then
                ConvNumCent = "cent"
            Else
                ConvNumCent = "cent " & strReste
            End If
        Case Else
            If byReste = 0 Then
                ConvNumCent = TabUnit(byCent) & " cents"
            Else
                ConvNumCent = TabUnit(byCent) & " cent " & strReste
            End If
    End Select
End Function

Private Function Nz(strNb As String) As String
    If strNb <> " zéro" Then Nz = strNb
End Function

