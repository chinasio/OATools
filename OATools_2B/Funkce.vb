Imports System.Math
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry

Module Funkce
    Public Function Vektory_normala(ByVal sArray As Array) As Array
        'calculate xyz components of vector from angular orientation of normal to plane
        Dim Alfa, Fi As Double
        Alfa = (sArray(0) + 180) Mod 360

        Fi = 90 - sArray(1)

        'xyz components of vetor
        Dim x, y, z As Double
        x = 0
        y = 0
        z = 0

        Const PI = 3.141592
        x = (Cos(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
        y = (Sin(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
        z = Sin(Fi * PI / 180)

        'MsgBox "X je: " & X & "Y je: " & Y & "Z je: " & Z
        'return
        Dim kArray(2) As Double
        kArray(0) = x
        kArray(1) = y
        kArray(2) = z
        Vektory_normala = kArray


    End Function


    Public Function Vektory_lineace(ByVal lArray As Array) As Array
        'calculate xyz components of vector from angular orientation of fall line
        Dim Alfa, Fi As Double
        Alfa = lArray(0)
        Fi = lArray(1)

        'xyz components of vetor
        Dim x, y, z As Double
        x = 0
        y = 0
        z = 0

        Const PI = 3.141592
        x = (Cos(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
        y = (Sin(Alfa * PI / 180)) * (Cos(Fi * PI / 180))
        z = Sin(Fi * PI / 180)

        'MsgBox "X je: " & X & "Y je: " & Y & "Z je: " & Z
        'return values
        Dim mArray(2) As Double
        mArray(0) = x
        mArray(1) = y
        mArray(2) = z
        Vektory_lineace = mArray

    End Function
    Public Function Statistics(ByVal sArray As Array) As Array
        'calculate eigenvalues
        Dim m11, m12, m13, m22, m23, m33, a, b, c, Lam1, Lam2, Lam3, CosPsi, Psi, ni, p, q, N As Double
        'Dim Lam, E, f, G, Xprumer, Yprumer, Zprumer As Double
        'Dim Alfa, Fi, Fr As Double

        'components of orientation matrix
        m11 = sArray(0)
        m22 = sArray(1)
        m33 = sArray(2)
        m12 = sArray(3)
        m13 = sArray(4)
        m23 = sArray(5)

        Const PI = 3.141592654


        'MsgBox m11 + m22 + m33
        'sum of matrix components at diagonal equals number of tectonic orientations in case of unweighted matrix, in case of weighted matrix this equals the sum of weights.
        a = -(m11 + m22 + m33)
        b = m11 * m22 + m11 * m33 + m22 * m33 - (m12 ^ 2) - (m13 ^ 2) - (m23 ^ 2)
        c = (m12 ^ 2) * m33 + (m13 ^ 2) * m22 + (m23 ^ 2) * m11 - m11 * m22 * m33 - 2 * m12 * m13 * m23

        'If Not a = 0 Then
        'Lam1, Lam2, Lam3 are eigenvalues
        p = (b - ((a ^ 2) / 3)) / 3
        q = 1 / 2 * ((2 / 27) * (a ^ 3) - (a * b) / 3 + c)
        CosPsi = -q / (Sqrt(Abs(p) ^ 3))
        'Psi = Cos(CosPsi) není arcCos, převod na arcTan
        'ZASEKNE SE, COSPSI JE 1, PAK ODMOCNINA Z 0!!!, příp odm. záp.čísla, přidávám Abs():
        'Psi = Atn(Sqr(1 / (CosPsi ^ 2) - 1))
        Psi = Atan(Sqrt(Abs(1 / (CosPsi ^ 2) - 1)))
        ni = 2 * Sqrt(Abs(p))
        'eigenvalues
        Lam1 = ni * Cos(Psi / 3) - a / 3
        Lam2 = -ni * Cos((Psi - PI) / 3) - a / 3
        Lam3 = -ni * Cos((Psi + PI) / 3) - a / 3
        'number of tectonic features (size of the dataset)
        N = Lam1 + Lam2 + Lam3
        'find the highest and lowest eigenvalue
        Dim MyArray(2) As Double
        MyArray(0) = Lam1
        MyArray(1) = Lam2
        MyArray(2) = Lam3
        Array.Sort(MyArray)

        Lam1 = MyArray(2) / N 'the highest value
        Lam2 = MyArray(1) / N
        Lam3 = MyArray(0) / N 'the lowest value

        'the shape parameter (tvar rozdělení)
        Dim Gama As Double
        'ln x = log x/log e
        Dim e As Double
        e = 2.718281828
        Gama = (Log(Lam1 / Lam2) / Log(e)) / (Log(Lam2 / Lam3) / Log(e))
        If Gama > 999 Then Gama = 999

        'the strength parameter (míra přednostní orientace)
        Dim ks As Double
        ks = Log(Lam1 / Lam3) / Log(e)
        If ks > 999 Then ks = 999
        'return values
        Dim gArray(4) As Double
        gArray(0) = Lam1
        gArray(1) = Lam2
        gArray(2) = Lam3
        gArray(3) = Gama
        gArray(4) = ks
        Statistics = gArray

    End Function
    Public Function Char_cisla_osa(ByVal sArray As Array) As Array
        'calculate eigenvalues, using the lowest one, calculate eigenvector, convert to angles (dipdir,dip)
        Dim m11, m12, m13, m22, m23, m33, a, b, c, Lam1, Lam2, Lam3, CosPsi, Psi, ni, p, q As Double
        Dim Lam, E, f, G, Xprumer, Yprumer, Zprumer As Double
        Dim Alfa, Fi, Fr As Double

        'components of orientation matrix
        m11 = sArray(0)
        m22 = sArray(1)
        m33 = sArray(2)
        m12 = sArray(3)
        m13 = sArray(4)
        m23 = sArray(5)

        Const PI = 3.141592654


        'MsgBox m11 + m22 + m33
        'sum of matrix components at diagonal equals number of tectonic orientations in case of unweighted matrix, in case of weighted matrix this equals the sum of weights.
        a = -(m11 + m22 + m33)
        b = m11 * m22 + m11 * m33 + m22 * m33 - (m12 ^ 2) - (m13 ^ 2) - (m23 ^ 2)
        c = (m12 ^ 2) * m33 + (m13 ^ 2) * m22 + (m23 ^ 2) * m11 - m11 * m22 * m33 - 2 * m12 * m13 * m23

        'If Not a = 0 Then
        'Lam1, Lam2, Lam3 are eigenvalues
        p = (b - ((a ^ 2) / 3)) / 3
        q = 1 / 2 * ((2 / 27) * (a ^ 3) - (a * b) / 3 + c)
        CosPsi = -q / (Sqrt(Abs(p) ^ 3))
        'Psi = Cos(CosPsi) není arcCos, převod na arcTan
        'ZASEKNE SE, COSPSI JE 1, PAK ODMOCNINA Z 0!!!, příp odm. záp.čísla, přidávám Abs():
        'Psi = Atn(Sqr(1 / (CosPsi ^ 2) - 1))
        Psi = Atan(Sqrt(Abs(1 / (CosPsi ^ 2) - 1)))
        ni = 2 * Sqrt(Abs(p))
        'eigenvalues
        Lam1 = ni * Cos(Psi / 3) - a / 3
        Lam2 = -ni * Cos((Psi - PI) / 3) - a / 3
        Lam3 = -ni * Cos((Psi + PI) / 3) - a / 3

        'find the highest and lowest eigenvalue
        Dim MyArray(2) As Double
        MyArray(0) = Lam1
        MyArray(1) = Lam2
        MyArray(2) = Lam3
        Array.Sort(MyArray)

        Lam1 = MyArray(2)
        Lam2 = MyArray(1)
        Lam3 = MyArray(0)

        Lam = Lam3
        'MsgBox("Lam1 je: " & Lam1 & " , Lam2 je: " & Lam2 & " , Lam3 je: " & Lam3)
        'MsgBox(Lam1 + Lam2 + Lam3)

        'platí:
        'Dim bh As Double
        'bh = Lam ^ 3 + A * Lam ^ 2 + B * Lam + C
        'sum of eigenvalues should be equal to number of tectonic orientations, or to sum of weights if it is weighted.

        'eigenvectors (Xprumer, Yprumer and Zprumer)
        E = m11 - Lam
        f = (m22 - Lam) * E - (m12 ^ 2)
        G = E * m23 - m13 * m12
        Xprumer = m12 * G - m13 * f
        Yprumer = -G * E
        Zprumer = Abs(E * f)

        'convert vector orientation of normal to plane to orientation in anlges of fall line
        If Xprumer + Yprumer = 0 Then
            Fi = 90
            Alfa = 0
        Else

            'Fr je Fi v rad (Fr is the dip in radians]
            Fr = Atan(Zprumer / Sqrt(Xprumer * Xprumer + Yprumer * Yprumer) + 0.001)
            Fi = (Fr * 180) / PI
            Alfa = (((2 - Sign(Yprumer) - Sign(Xprumer * Yprumer)) * PI / 2) * 180 / PI) + (Atan(Yprumer / Xprumer) * 180 / PI)
            Alfa = Round(Alfa)
            Fi = Round(Fi)

        End If
        'return values
        Dim pArray(1) As Double
        pArray(0) = Alfa
        pArray(1) = Fi
        Char_cisla_osa = pArray

    End Function

    Public Function Char_cisla_prum_smer(ByVal sArray As Array) As Array
        'calculate eigenvalues, using the highest one, calculate eigenvector, convert to angles (dipdir,dip)
        Dim m11, m12, m13, m22, m23, m33, a, b, c, Lam1, Lam2, Lam3, CosPsi, Psi, ni, p, q As Double
        Dim Lam, E, f, G, Xprumer, Yprumer, Zprumer As Double
        Dim Alfa, Fi, Fr As Double

        'components of orientation matrix
        m11 = sArray(0)
        m22 = sArray(1)
        m33 = sArray(2)
        m12 = sArray(3)
        m13 = sArray(4)
        m23 = sArray(5)

        Const PI = 3.141592654

        'MsgBox m11 + m22 + m33
        'sum of matrix components at diagonal equals number of tectonic orientations in case of unweighted matrix, in case of weighted matrix this equals the sum of weights.
        'na uhlopříčce je součet složek matice rovný počtu měření při nevážené matici, u vážené součtu vah
        a = -(m11 + m22 + m33)
        b = m11 * m22 + m11 * m33 + m22 * m33 - (m12 ^ 2) - (m13 ^ 2) - (m23 ^ 2)
        c = (m12 ^ 2) * m33 + (m13 ^ 2) * m22 + (m23 ^ 2) * m11 - m11 * m22 * m33 - 2 * m12 * m13 * m23

        'If Not a = 0 Then
        'eigenvalues (Lam1, Lam2, Lam3)
        p = (b - ((a ^ 2) / 3)) / 3
        q = 1 / 2 * ((2 / 27) * (a ^ 3) - (a * b) / 3 + c)
        CosPsi = -q / (Sqrt(Abs(p) ^ 3))
        'Psi = Cos(CosPsi) není arcCos, převod na arcTan
        'ZASERE SE, COSPSI JE 1, PAK ODMOCNINA Z 0!!!, příp odm. záp.čísla, přidávám Abs():
        'Psi = Atn(Sqr(1 / (CosPsi ^ 2) - 1))
        Psi = Atan(Sqrt(Abs(1 / (CosPsi ^ 2) - 1)))
        ni = 2 * Sqrt(Abs(p))
        'eigenvalues
        Lam1 = ni * Cos(Psi / 3) - a / 3
        Lam2 = -ni * Cos((Psi - PI) / 3) - a / 3
        Lam3 = -ni * Cos((Psi + PI) / 3) - a / 3

        'find the highest and lowest eigenvalue
        Dim MyArray(2) As Double
        MyArray(0) = Lam1
        MyArray(1) = Lam2
        MyArray(2) = Lam3
        Array.Sort(MyArray)

        Lam1 = MyArray(2)
        Lam2 = MyArray(1)
        Lam3 = MyArray(0)

        Lam = Lam1
        'MsgBox("Lam1 je: " & Lam1 & " , Lam2 je: " & Lam2 & " , Lam3 je: " & Lam3)
        'MsgBox(Lam1 + Lam2 + Lam3)
        'platí:
        'Dim bh As Double
        'bh = Lam ^ 3 + A * Lam ^ 2 + B * Lam + C
        'sum of eigenvalues should be equal to number of tectonic orientations, or to sum of weights if it is weighted.
        'eigenvector
        E = m11 - Lam
        f = (m22 - Lam) * E - (m12 ^ 2)
        G = E * m23 - m13 * m12
        Xprumer = m12 * G - m13 * f
        Yprumer = -G * E
        Zprumer = Abs(E * f)

        'convert vector orientation of normal to plane to orientation in anlges of fall line
        If Xprumer + Yprumer = 0 Then
            Fi = 90
            Alfa = 0
        Else
            'Fr je Fi v rad
            Fr = Atan(Zprumer / Sqrt(Xprumer * Xprumer + Yprumer * Yprumer) + 0.001)
            Fi = (Fr * 180) / PI
            Alfa = (((2 - Sign(Yprumer) - Sign(Xprumer * Yprumer)) * PI / 2) * 180 / PI) + (Atan(Yprumer / Xprumer) * 180 / PI)
            Alfa = Round(Alfa)
            Fi = Round(Fi)

        End If
        'return values
        Dim pArray(1) As Double
        pArray(0) = Alfa
        pArray(1) = Fi
        Char_cisla_prum_smer = pArray

    End Function

    Public Function Get_Azimuth(ByVal pline As ILine)
        Dim pVector As IVector3D
        Dim dAzimuth As Double
        pVector = New Vector3D
        pVector.ConstructDifference(pline.ToPoint, pline.FromPoint)
        dAzimuth = pVector.Azimuth
        Get_Azimuth = dAzimuth
    End Function

    Public Function PtConstructAngleDistance(ByVal pPoint As IPoint, ByVal dAngle As Long, ByVal dDist As Long) As IPoint
        Const PI = 3.14159265358979
        Dim dAngleRad As Double
        Dim pCPoint As IConstructPoint
        pCPoint = New Point

        dAngleRad = dAngle * 2 * PI / 360 'Convert the angle degrees to radians
        'MsgBox ""
        pCPoint.ConstructAngleDistance(pPoint, dAngleRad, dDist)
        PtConstructAngleDistance = pCPoint
    End Function


End Module
