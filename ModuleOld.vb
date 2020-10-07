
'Imports System.Math
'' В0 – давление атмосферного воздуха
'' R – тяга двигателя
''GT – массовый расход топлива
''nвд – частота вращения ротора высокого давления
'' nнд – частота вращения ротора низкого давления
''t*ТНД – температура газов за турбиной низкого давления
''Δ РКВД – избыточное давление воздуха за компрессором
'' Р*Вх  - полное давление воздуха на входе в двигатель
''ПкΣ – суммарная степень повышения давления в компрессоре 
''Н – барометрическая высота, соответствующая давлению атмосферного воздуха

'Module ModuleOld
'    Private ik As Integer

'    Private Hu As Double
'    Private a1IPmin As Double
'    Private a1IPmax As Double
'    Private a1gradmin As Double
'    Private a1gradmax As Double
'    Private a2IPmin As Double
'    Private a2IPmax As Double
'    Private a2gradmin As Double
'    Private a2gradmax As Double
'    Private DkpIPmin As Double
'    Private DkpIPmax As Double
'    Private DkpMMmin As Double
'    Private DkpMMmax As Double
'    Private bGtoplTDR As Double
'    Private aGtoplTDR As Double
'    'Private iVAR As Integer
'    Private Kvvod1, PRIZNAK As Integer

'    Private U, kk As Integer
'    Private ij, I, J As Integer
'    Private Kvv, CD As Integer
'    Private KKK, Akn1, AkPiB, AkGB, Akm, Akn2, _
'    AkPiKS, AkGT0, AkT3, AkT4, AkR, _
'    AkCR, _
'    Fkn1, FkPiB, FkGB, Fkm, Fkn2, _
'    FkPiKS, FkGT0, FkT3, FkT4, FkR, _
'    FkCR, FkaS, FkPiKVD, FkGTF As Double

'    Private KKK1, Akn1fi, AkPiBfi, AkGBfi, Akmfi, Akn2fi, _
'    AkPiKSfi, AkGT0fi, AkT3fi, AkT4fi, AkRfi, _
'    AkCRfi, _
'    Fkn1fi, FkPiBfi, FkGBfi, Fkmfi, Fkn2fi, _
'    FkPiKSfi, FkGT0fi, FkT3fi, FkT4fi, FkRfi, _
'    FkCRfi, FkPiKVDfi, FkGTFfi As Double

'    'Private Kvv, CD As Integer
'    'Private ij As Integer
'    'Private Kvvod1, PRIZNAK As Integer
'    Private phyzPER, T3fizTDR, aN1, fi As Double
'    Private nvvod, jk, Nkt, jj, N0, N, L1, jjj, ii As Integer
'    Private eps, alRUD, B0, Pb, tbc, Rizm, freq, Gtopl, FcaTVD, t300fiz, u447, T4, T4ogr, Hbc, Bmc, P2aKVD, P2bKVD, P2CT, P6b, P6CT, P4b, P4CT, t6, del1, del2, del3, PpNas1, dH2O1 As Double
'    Private alfa1, alfa2, Dpc, Tb, Ph, Pbaz, PvxCTm, Pvxbm, Xptm, Tptm, dH2O, GBphyzB, ppB, dk, sigVX, P6CTb, P6bb, P6bbs, T6b, PiKND, PiKNDs, t06b, ethaKND, PRptm, T1ptm, TTTad, T0ptm, aLkptmK, aKND, aLkptm, ethaKND1, ethaKNDs, P2b, T2b, PiKVD As Double
'    Private t62b, ethaKVD, dk2, ethaKVD1, PiKsum, t02b, Tbb, ethaKsum, aKsum, ethaKsum1, GtoplTDR, Thhot, T3fizmem, T4B, P4CTb, P4bb, PiT, Rfiz, alfafiz, CR, GBPHYZI, GBPHYZca, uMem, PiTMem, Tgfors, Rfors, Fca, T3captm, GBCAptm As Double
'    Private Skc1, GBCA, T3f As Double
'    Private G, AG, M0, R0, R, AR, T1, T2, Q, DI, PI, KT, KC, TKP, PKP, AKP, FKP, M, M1, B, D, V As Double 'A,

'    Private AJ As Double
'    Private FYZGG, q10, PHYZGB As Double    ' PHYZ(27,j)=GBphyzB1=PHYZGB 
'    Private FYZGG1, GTp1, qTp, QT As Double  ' GTp1=qTp*GBphyzB1
'    Private AU, BU, DF, T3 As Double
'    Private T0, TT, ENTH, hh, aIg
'    Private kPER, CkN1, CkPiB, CkGB, Ckm, CkN2, CkPiKS, CkGT0, CkT3, CkT4, CkR, CkCR, CkPiKVD, CkGTF, CkaS As Double

'    Private koeffPER, n1cay, PiBcay, GBcay, mcay, Pikcay, Rcay, CRcay, alfaScay, n2cay, PiTcay, T3cay, T4cay As Double
'    Private Kpr, ABKR, ABKGT, ABKTG As Double

'    Private CP400, I9 As Double
'    Private PHYZ(45, 100), PRIV(40, 100), PER(21, 100), VZ(32), A(19), aKperes(15, 100), bKperesfi(15, 100) As Double
'    Private Izdelie, N1f, N2f, ISren As Double
'    Private Stend, Rejim As Integer



'    Sub Main()

'        '              6 февраля 2006 г Введён ключ:  iVAR=1 или 2 или 3 (см пояснение в блокноте) 
'        '         17 Декабря 2005 обобщение 1-й и 2-й задач:
'        '                     см далее п/п ABCD 

'        ' New  KENTAVR: with Wet Air  27.05.2009

'        'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV    РАБОЧАЯ   VVVVVVVVVVVVVVVVVVVVVVVРАБОЧАЯ
'        'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV

'        '         =========  M A I N    P R O G R A M   ===========

'        '	dimension PHYZ(44,100),PRIV(40,100),PER(21,100),VZ(31),A(19),      ' A(18)
'        'Dim PHYZ(45, 100), PRIV(40, 100), PER(21, 100), VZ(32), A(19), aKperes(15, 100), bKperesfi(15, 100) As Double
'        'Dim Izdelie, N1f, N2f, ISren As Double
'        'Dim Stend, Rejim As Integer

'        Const IJK As Integer = 100    ' предельное количество вариантов 
'        'C       - - - - - - - - - - - - - - - - - -

'        Dim nvvod As Integer = 31

'        '      print *,'before tarir'
'        Call TARIROVKA(IJK, A)
'        '     * Stend,Izdelie,Hu,NO,a1IPmin,a1IPmax,a1gradmin,
'        '     * a1gradmax,a2IPmin,a2IPmax,a2gradmin,a2gradmax,DkpIPmin,DkpIPmax,
'        '     * DkpMMmin,DkpMMmax,bGtoplTDR,aGtoplTDR)
'        'print *,'tarir'

'        Call VVOD(IJK, A, VZ, PHYZ, PRIV, PER, aKperes, bKperesfi)

'        '	      SUBROUTINE VVOD(ijk,A,VZ,PHYZ,PRIV,PER,cKper,cKperfi)
'        '     * Stend,Izdelie,Hu,NO,a1IPmin,a1IPmax,a1gradmin,
'        '     * a1gradmax,a2IPmin,a2IPmax,a2gradmin,a2gradmax,DkpIPmin,DkpIPmax,
'        '     * DkpMMmin,DkpMMmax,bGtoplTDR,aGtoplTDR,VZ)
'        'Stop
'    End Sub

'    Private Sub OPEN(ByVal int As Integer, ByVal strFile As String)

'    End Sub
'    Private Sub READ(ByVal int As Integer) ', ByVal arr1 As Integer, ByVal arr2 As Integer, ByVal arr3 As Integer)

'    End Sub

'    Private Sub READ(ByVal int As Integer, ByVal Val As Double) ', ByVal arr1 As Integer, ByVal arr2 As Integer, ByVal arr3 As Integer)

'    End Sub
'    Private Sub READ(ByVal int As Integer, ByVal Val(,) As Double) ', ByVal arr1 As Integer, ByVal arr2 As Integer, ByVal arr3 As Integer)

'    End Sub
'    '       MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'    '       MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

'    'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'    'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'    '      ATTENTION:  ВСЕ МАССИВЫ PHYZ(43,ijk),PRIV(40,ijk),PER(21,ijk) ВО ВСЕХ П/П ДОЛЖНЫ
'    '      ИМЕТЬ ОДИНАКОВЫЕ РАЗМЕРНОСТИ СООТВЕТСТВЕННО''' 

'    '       --------П Р О Д О Л Ж Е Н И Е   23 СЕНТ 2004------
'    '                        Т А Р И Р О В К А 
'    Private Sub TARIROVKA(ByVal IJK As Integer, ByVal A() As Double)
'        '     * Stend,Izdelie,Hu,NO,a1IPmin,a1IPmax,a1gradmin,
'        '     * a1gradmax,a2IPmin,a2IPmax,a2gradmin,a2gradmax,DkpIPmin,DkpIPmax,
'        '     * DkpMMmin,DkpMMmax,bGtoplTDR,aGtoplTDR)
'        '      Читаем из файла Tarirovka данные Учётной Карты:
'        Dim Izdelie As Double
'        Dim Stend As Integer ', NO
'        'Dim A(19) As Double         '  A(18)
'        Dim TAR(3, 19) As String ' * 10, 
'        Dim BINIT(32, IJK) As Double

'        '	COMMON/TARIR/Stend,Izdelie,Hu,NO,a1IPmin,a1IPmax,a1gradmin,
'        '     * a1gradmax,a2IPmin,a2IPmax,a2gradmin,a2gradmax,DkpIPmin,DkpIPmax,
'        '     * DkpMMmin,DkpMMmax,aGtoplTDR,bGtoplTDR
'        '		COMMON/FIZPRIV/alfa1,alfa2,Ph,N1fiz,N2,P6CTb,P6bb,
'        '     * T6bb,PiKND,ethaKND,P2b,T2b,PiKVD,ethaKVD,PiKsum,ethaKsum,
'        '     * GBPHYZs,m,T4B,P4st,PiT,Kv,Rf,alfafiz,CR,njuOHL,Tgfors,mfors
'        Call OPEN(7, "A init\Tarirovka.TXT") '- файл тарировочных данных соотв 34 стенду (8 вариантов)

'        For K = 1 To 19  '- 18-число строк 
'            'READ(7)
'            READ(TAR(1, K)) : READ(TAR(2, K)) : READ(TAR(3, K))    '- СТРОКА ' чтение из 10 файла исх данн в своб формате
'        Next

'        ''              READ(AINIT(4,K),*) tb1   
'        '              -ЧИТАЕМ ИЗ ПОДСТРОКИ TAR(3,K) ЗНАЧЕНИЕ a(k) Матрица TAR(3,16)*10 -символьная
'        For K = 1 To 19  '- 19-число строк
'            READ(TAR(3, K)) : READ(A(K))
'        Next

'        Stend = A(1)

'        Izdelie = A(2)
'        Hu = A(3)
'        'NO = A(4)
'        a1IPmin = A(5)
'        a1IPmax = A(6)
'        a1gradmin = A(7)
'        a1gradmax = A(8)
'        a2IPmin = A(9)
'        a2IPmax = A(10)
'        a2gradmin = A(11)
'        a2gradmax = A(12)
'        DkpIPmin = A(13)
'        DkpIPmax = A(14)
'        DkpMMmin = A(15)
'        DkpMMmax = A(16)
'        bGtoplTDR = A(17)
'        aGtoplTDR = A(18)
'        'iVAR = A(19)
'    End Sub

'    'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT
'    'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT


'    '     VVVVVVVVVVVVVVVVVVVVVVVVVVVV РАБОЧАЯ  VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
'    '               VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
'    '                         ===========      В В О Д      ==============

'    Private Sub VVOD(ByVal IJK As Integer, ByVal A() As Double, ByVal VZ() As Double, ByVal PHYZ(,) As Double, ByVal PRIV(,) As Double, ByVal PER(,) As Double, ByVal cKper(,) As Double, ByVal cKperfi(,) As Double)
'        Dim b(200) As Double ', N1f, N2f, Izdelie, mfizmem, Kv, njuOHL, mfors As Double
'        Dim Akn1(38, 22), AkPiB(38, 22), AkGB(38, 22), Akm(38, 22), Akn2(38, 22), AkPiKS(38, 22), AkGT0(38, 22), AkT3(38, 22), AkT4(38, 22), AkR(38, 22), AkCR(38, 22), Fkn1(38, 22), FkPiB(38, 22), FkGB(38, 22), Fkm(38, 22), Fkn2(38, 22), FkPiKS(38, 22), FkGT0(38, 22), FkT3(38, 22), FkT4(38, 22), FkR(38, 22), FkCR(38, 22), FkaS(38, 22), FkPiKVD(38, 22), FkGTF(38, 22), _
'       Akn1fi(12, 20), AkPiBfi(12, 20), AkGBfi(12, 20), Akmfi(12, 20), _
'       Akn2fi(12, 20), _
'       AkPiKSfi(12, 20), AkGT0fi(12, 20), AkT3fi(12, 20), AkT4fi(12, 20), _
'       AkRfi(12, 20), _
'       AkCRfi(12, 20), _
'       Fkn1fi(12, 20), FkPiBfi(12, 20), FkGBfi(12, 20), Fkmfi(12, 20), _
'       Fkn2fi(12, 20), _
'       FkPiKSfi(12, 20), FkGT0fi(12, 20), FkT3fi(12, 20), FkT4fi(12, 20), _
'       FkRfi(12, 20), _
'       FkCRfi(12, 20), FkPiKVDfi(12, 20), FkGTFfi(12, 20) As Double


'        'Dim Stend As Integer    'Rejim,Stend    
'        'Dim ISren As Integer
'        ' 28-ЧИСЛО ВАРИАНТОВ
'        Dim AINIT(32, IJK), BINIT(32, IJK), CD, CE, CF, CINIT(32, 2), DINIT(32, IJK), PRIZNAK As String
'        PRIZNAK = Nothing
'        '      common/k14/Kperes
'        '	common/Kvvod/CD,CE,CF
'        'common/Kvvod1/PRIZNAK

'        Dim VV(32, IJK) As Double
'        'Dim VV(32, IJK), PHYZ(45, IJK), PRIV(40, IJK), PER(21, IJK), _
'        'cKper(15, IJK), cKperfi(15, IJK), _
'        'VZ(32), A(19) As Double

'        '     common/U/kk
'        'common/BOUND/ij
'        '     common/Kvv/CD
'        'COMMON/KKK/Akn1,AkPiB,AkGB,Akm,Akn2,
'        '    *AkPiKS,AkGT0,AkT3,AkT4,AkR,
'        '    *AkCR,
'        '    *Fkn1,FkPiB,FkGB,Fkm,Fkn2,
'        '    *FkPiKS,FkGT0,FkT3,FkT4,FkR,
'        '    *FkCR,FkaS,FkPiKVD,FkGTF

'        'COMMON/KKK1/Akn1fi,AkPiBfi,AkGBfi,Akmfi,Akn2fi,
'        '    *AkPiKSfi,AkGT0fi,AkT3fi,AkT4fi,AkRfi,
'        '    *AkCRfi,
'        '    *Fkn1fi,FkPiBfi,FkGBfi,Fkmfi,Fkn2fi,
'        '    *FkPiKSfi,FkGT0fi,FkT3fi,FkT4fi,FkRfi,
'        '    *FkCRfi,FkPiKVDfi,FkGTFfi

'        '	  common/k14/Kper    ?????
'        '-------------------------------------- Вставка
'        nvvod = 32

'        '--------------------------------------
'        OPEN(229, "'A init\VARIANT.TXT")


'        If (jk > 0) Then GoTo 421 ' ключ однократного открытия 6 файла и записи в него строки j  Nkt Rejim ...

'        '	OPEN(6,FILE='B_txt\VVOD1.txt')
'        '--------------------------------------
'        OPEN(77, "B rez\ex_VVOD.xls")
'        '--------------------------------------

'        '	WRITE(6,33)
'        '  33  FORMAT('      j     Nkt  Rejim  alRUD  B0         Pb  
'        '     *       tb      Rizm        N1f      N2f        freq   Gtopl     Fc
'        '     *aTVD   t300fiz    u447        T4      T4ogr     Hbc          Bmc    
'        '     *   P2aKVD     P2bKVD     P2CT      P6b      P6CT        P4b        
'        '     *P4CT    T6         del1    del2    del3     ISren')  
'        '--------------------------------------------------------------
'        '       Запись в файл PROBA vvod.XLS (с промежуточным знаком табуляции)
'        '        Write(77, 7779)
'        '7779:   Format("j", "	", "Nkt", "	", "Rejim", "	", "alRUD", "	", "B0", "	", _
'        '         "Pb", "	", "tb", "	", "Rizm", "	", "N1f", "	", "N2f", "	", "freq", _
'        '         "	", _
'        '         "Gtopl", "	", "FcaTVD", "	", "t300fiz", "	", "u447", "	", "T4", _
'        '         "	", "T4ogr", "	", "Hbc", "	", "Bmc", "	", "P2aKVD", "	", "P2bKVD", _
'        '         "	", "P2CT", "	", "P6b", "	", "P6CT", "	", "P4b", "	", "P4CT", "	", _
'        '         "T6", "	", _
'        '         "del1", "	", "del2", "	", "del3", "	", "ISren", "	", "Fi,%", "	")
'        '--------------------------------------------------------------- 
'        jk = jk + 1


'        '***********************15.05.2009**********
'        OPEN(10, "A init\Akn1.TXT")
'        OPEN(11, "A init\AkPiB.TXT")
'        OPEN(12, "A init\AkGB.TXT")
'        OPEN(13, "A init\Akm.TXT")
'        OPEN(14, "A init\Akn2.TXT")
'        OPEN(15, "A init\AkPiKS.TXT")
'        OPEN(16, "A init\AkGT0.TXT")
'        OPEN(17, "A init\AkT3.TXT")
'        OPEN(18, "A init\AkT4.TXT")
'        OPEN(19, "A init\AkR.TXT")
'        OPEN(20, "A init\AkCR.TXT")


'        OPEN(21, "A init\Fkn1.TXT")
'        OPEN(22, "A init\FkPiB.TXT")
'        OPEN(23, "A init\FkGB.TXT")

'        OPEN(24, "A init\Fkm.TXT")
'        OPEN(25, "A init\Fkn2.TXT")
'        OPEN(26, "A init\FkPiKS.TXT")

'        OPEN(27, "A init\FkGT0.TXT")
'        OPEN(28, "A init\FkT3.TXT")
'        OPEN(29, "A init\FkT4.TXT")

'        OPEN(30, "A init\FkR.TXT")
'        OPEN(31, "A init\FkCR.TXT")
'        OPEN(32, "A init\FkaS.TXT")
'        OPEN(33, "A init\FkPiKVD.TXT")
'        OPEN(34, "A init\FkGTF.TXT")


'        READ(10, Akn1)
'        READ(11, AkPiB)
'        READ(12, AkGB)
'        READ(13, Akm)
'        READ(14, Akn2)
'        READ(15, AkPiKS)
'        READ(16, AkGT0)
'        READ(17, AkT3)
'        READ(18, AkT4)
'        READ(19, AkR)
'        READ(20, AkCR)



'        READ(21, Fkn1)
'        READ(22, FkPiB)
'        READ(23, FkGB)
'        READ(24, Fkm)
'        READ(25, Fkn2)
'        READ(26, FkPiKS)
'        READ(27, FkGT0)
'        READ(28, FkT3)
'        READ(29, FkT4)

'        READ(30, FkR)
'        READ(31, FkCR)
'        READ(32, FkaS)
'        READ(33, FkPiKVD)
'        READ(34, FkGTF)
'        'PRINT *,'END OF INTRODUCTION'


'        ' коэффициенты пересч, учитывающие относительную влажн. воздуха
'        OPEN(110, "A init\Akn1fi.TXT")

'        OPEN(111, "A init\AkPiBfi.TXT")
'        OPEN(112, "A init\AkGBfi.TXT")
'        OPEN(113, "A init\Akmfi.TXT")
'        OPEN(114, "A init\Akn2fi.TXT")
'        OPEN(115, "A init\AkPiKSfi.TXT")

'        OPEN(116, "A init\AkGT0fi.TXT")
'        OPEN(117, "A init\AkT3fi.TXT")
'        OPEN(118, "A init\AkT4fi.TXT")

'        OPEN(119, "A init\AkRfi.TXT")
'        OPEN(120, "A init\AkCRfi.TXT")


'        OPEN(121, "A init\Fkn1fi.TXT")
'        OPEN(122, "A init\FkPiBfi.TXT")
'        OPEN(123, "A init\FkGBfi.TXT")

'        OPEN(124, "A init\Fkmfi.TXT")
'        OPEN(125, "A init\Fkn2fi.TXT")
'        OPEN(126, "A init\FkPiKSfi.TXT")

'        OPEN(127, "A init\FkGT0fi.TXT")
'        OPEN(128, "A init\FkT3fi.TXT")
'        OPEN(129, "A init\FkT4fi.TXT")

'        OPEN(130, "A init\FkRfi.TXT")
'        OPEN(131, "A init\FkCRfi.TXT")
'        '	  OPEN(132,FILE="A init\FkaSfi.TXT")
'        OPEN(133, "A init\FkPiKVDfi.TXT")
'        OPEN(134, "A init\FkGTFfi.TXT")



'        READ(110, Akn1fi)
'        READ(111, AkPiBfi)
'        READ(112, AkGBfi)
'        READ(113, Akmfi)
'        READ(114, Akn2fi)
'        READ(115, AkPiKSfi)
'        READ(116, AkGT0fi)
'        READ(117, AkT3fi)
'        READ(118, AkT4fi)
'        READ(119, AkRfi)
'        READ(120, AkCRfi)



'        READ(121, Fkn1fi)
'        READ(122, FkPiBfi)
'        READ(123, FkGBfi)
'        READ(124, Fkmfi)
'        READ(125, Fkn2fi)
'        READ(126, FkPiKSfi)
'        READ(127, FkGT0fi)
'        READ(128, FkT3fi)
'        READ(129, FkT4fi)
'        READ(130, FkRfi)
'        READ(131, FkCRfi)

'        '	  READ(132,FkaSfi
'        READ(133, FkPiKVDfi)
'        READ(134, FkGTFfi)

'        '*********************15.05.2009***********
'        Dim I, J, K, t, L As Integer
'421:    'Continue Do
'        For K = 1 To IJK  '- ijk-число строк   ' предельное количество вариантов ijk=100

'            'READ(229,*),(AINIT(I,K),I=1,32)
'            READ(AINIT(I, K))
'            '           			  '- СТРОКА ' чтение из 29 файла исх данн в своб формате

'            If (AINIT(1, K) = "0" AndAlso AINIT(2, K) = "0") Then GoTo 1

'            'CCC                      =================================>>>>>>>> to_ 1 >>>>>>>>>
'            J = K

'            CD = AINIT(1, K)  '-это РПТ, Б, или УБ
'            PRIZNAK = CD
'            CE = AINIT(3, K)  '-это индекс режима
'            CF = AINIT(4, K)  '-это значение альфа РУД 

'            BINIT(1, J) = AINIT(1, J)
'            '
'            'READ(AINIT(2,j),*) b(j) ' ЧТЕНИЕ СИМВОЛЬНОЙ КОНСТАНТЫ-ПРИСВОЕНИЕ ЕЁ ЗНАЧЕНИЯ ЧИСЛОВОЙ ЯЧЕЙКЕ 
'            READ(AINIT(2, J)) : READ(b(J))
'            Nkt = b(J)

'            'CCC   --------------------------- =======>>>>>>>>>>>>

'            For I = 2 To nvvod
'                'READ(AINIT(i,K),*) b(i)
'                READ(AINIT(I, K)) : READ(b(I))
'            Next

'            VV(1, K) = K
'            VZ(1) = K

'            For I = 2 To nvvod
'                VZ(I) = b(I)
'                VV(I, K) = b(I)
'            Next
'            '      ------------------------------------------------------------------ 
'            'CCC            CALL PHYZPARAMETERS(IJK,VZ,PHYZ,PRIV,PER)

'            '      ------------------------------------------------------------------ 
'        Next

'        'print *,'after 299'
'1:      'Continue Do

'        '      ++++++++++++++++++++++++++++++++++++++++++++++++++++
'        'CCC          ========================================================================
'        K = K - 1    ' - учёт нулевой строки, на которой происходит окончание расчётов
'        '         Запись в каждый из файлов 6,25,9,7 в два приёма: в 1-й строка идентификаторов, во 2-й матрицы результатов
'        '           DO 221 JK=1,K
'        '        WRITE(6,7771)  (VV(I,JK),I=1,32)           ' -запись в файл (6,FILE='B rez\VVOD.TXT')
'        ' 7771 FORMAT(4f5.0,28f10.4)  ' несмотря на то, что массив VV имеет 94 строки
'        ' 221  CONTINUE

'        '                                 Запись в файл   ex_VVOD.xls --
'        For jk As Integer = 1 To K
'            'Write(77, 7778)(VV(I, JK), I = 1, 32)
'            '7778 FORMAT(4(f5.0,'	 '),28(F10.4,'	'))
'        Next
'        '        ----------------------------------------------------

'        'CCC     ========================================================================
'        'XXXXXXXXXXXXX                     K=K-1
'        'CCC     ========================================================================

'        For J = 1 To K  ' - j-номер контрольной точки

'            'CCC            =================================>>>>>>>> to _ 1 >>>>>>>>>

'            CD = AINIT(1, J)  '-это РПТ, Б, УБ, Др или Мг
'            CE = AINIT(3, J)  '-это индекс режима
'            CF = AINIT(4, J)  '-это значение альфа РУД 
'            If (CD = "M") Then Debug.Print("M", J)
'            If (CD = "F") Then Debug.Print("F", J)

'            'CCC   --------------------------- =======>>>>>>>>>>>>

'            Nkt = b(J)

'            VZ(1) = J

'            For I = 2 To nvvod
'                VZ(I) = VV(I, J)
'            Next
'            '      ------------------------------------------------------------------ 
'            'subroutine KperMF(ijk,PRIZNAK,t,aN1,Kper)
'            Debug.Print("BEF CALL bKperMF(PRIZNAK,j,ijk,t,aN1,cKper,VZ)")
'            '    		read(*,*)
'            t = VZ(7)    'tb
'            aN1 = VZ(9)  ' N1f
'            Debug.Print("VZ(7) VZ(9)", t, aN1)
'            '      read(*,*)
'            '	ijk=100
'            Debug.Print("----1----IJK", IJK)
'            Call bKperMF(PRIZNAK, J, IJK, t, aN1, cKper, VZ)
'            Debug.Print("----2----IJK", IJK)
'            Debug.Print("AFTER CALL bKperMF(PRIZNAK,j,ijk,t,aN1,cKper,VZ)")  ' определение k пересчёта
'            '    		read(*,*)
'            fi = VZ(32)
'            Debug.Print("----3----IJK", IJK)
'            Call cKperMFfi(PRIZNAK, J, IJK, t, fi, cKperfi, VZ)   ' определение k пересчёта с учётом относит влажности воздуха
'            Debug.Print("----4----IJK", IJK)
'            Debug.Print("BEF CALL PHYZPARAMETERS")
'            '	ijk=100
'            '    		read(*,*)
'            '	print *,'j,ijk,A',j,ijk,A
'            '	read(*,*)
'            '	print *,'VZ',VZ
'            '		read(*,*)
'            '	print *,'PHYZ',PHYZ
'            '		read(*,*)
'            Call PHYZPARAMETERS(J, IJK, A, VZ, PHYZ)   'j-номер контрольной точки(варианта);ijk-разм-ть массивов
'            Debug.Print("AFTER CALL PHYZPARAMETERS")
'            '    		read(*,*)
'            'iVAR = A(19)
'            Call PRIVEDENIE(J, IJK, PHYZ, PRIV)  'iVAR, iVAR=1 или 2 или 3
'            Debug.Print("AFTER CALL PRIVEDENIE")
'            '   		read(*,*)
'            '      print *,'j into VVOD',j                
'            Call PERESCHET(J, IJK, PHYZ, PER, cKper, cKperfi)
'            Debug.Print("'---2---j,PER',(PER(i,j),i=1,21)")
'            '       read(*,*)

'            '     ЗАПИСЬ ПАРАМЕТРОВ (КОЭФ-ТЫ ПЕРЕСЧ И ПРИВЕД) ПРОИСХОДИТ ПОСЛЕ ОКОНЧАНИЯ РАСЧЕТА КОНТРОЛЬНОЙ ТОЧКИ
'            '      ------------------------------------------------------------------ 
'        Next

'        'CCC        ========================================================================
'        '----------------------------------------------------



'        '       ЗАПИСЬ ПАРАМЕТРОВ PHYZ,PER,PRIV ПРОИСХОДИТ ПОСЛЕ ОКОНЧАНИЯ РАСЧЕТА ВСЕХ ВАР-ТОВ(КОНТР ТОЧЕК)
'        '      WRITE(325,7772) ((PHYZ(i,j),i=1,44),j=1,K)
'        ' 7772 FORMAT(44G12.5)       '2f5.0,41f12.5)
'        '---------------------------------------------------
'        '       Запись в файл      ex_PHYZICAL.XLS
'        Write(PHYZ(I, J)) ',i=1,45),j=1,K)
'        '7700 FORMAT(45(G12.6,'	'))
'        '----------------------------------------------------
'        I = 0
'        J = 0
'        L = 0

'        '      WRITE(9,7773) ((PER(i,j),i=1,21),j=1,K) ' ij передан из SUBROUTINE PERESCHET(IJK):common /BOUND/ij
'        ' 7773 FORMAT(G12.3,G12.2,2G12.4,G12.5,G12.3,G12.5,11G12.4,F12.3, ' G12.5e3
'        '     * 2G12.4)
'        '       print *,'---2a--j,PER',((PER(i,l),i=1,21),l=1,7)  
'        '       read(*,*)
'        '         Запись в файл      ex_PERESCHET.XLS
'        Write(PER(I, J)) ',i=1,21),j=1,K)  ' OPEN(79,FILE='B rez\ex_PERESCHET.xls') зап строки идент
'        '7701 format(21(G12.5,'	'))
'        Debug.Print("'-3-j,PER,(PER(i,j),i=1,21")
'        '	 read(*,*)  
'        '-----------------------------------------------
'        '      WRITE(87,7774) ((PRIV(i,j),i=1,40),j=1,K)
'        ' 7774 FORMAT(6G12.4,G12.5,G12.4,G12.5,4G12.4,F12.3,23G12.4,3G15.4)
'        '
'        '---------------------------------------------------
'        '         Запись в файл      ex_PRIVEDENIE.XLS
'        Write(PRIV(I, J)) ',i=1,40),j=1,K)
'        '7703 FORMAT(40(G12.5,'	')) ' (6G12.4,G12.5,G12.4,G12.5,4G12.4,F12.3,23G12.4,G15.4)
'        Write(cKper(I, J)) ',i=1,15),j=1,K)
'        '1772 FORMAT(15(G12.5,'	'))       '2f5.0,41f12.5)


'        Write(cKperfi(I, J)) ',i=1,15),j=1,K)
'        '3772 FORMAT(15(G12.5,'	'))       '2f5.0,41f12.5)
'        '        +++++++++++++++++++++++++++++++++++++++++++++++++++++

'        '      close (9)
'        '      close (7)
'        '      close (15)

'        '      call printper(K)
'        '	call printpr(K)
'        '	call koeff per pr(K)
'    End Sub

'    'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV    РАБОЧАЯ   VVVVVVVVVVVVVVVVVV    РАБОЧАЯ
'    'VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV

'    'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
'    '                  РАСЧЕТ  Ф И З И Ч Е С К И Х  ПАРАМЕТРОВ

'    Private Sub PHYZPARAMETERS(ByVal j As Integer, ByVal IJK As Integer, ByVal A() As Double, ByVal VZ() As Double, ByVal PHYZ(,) As Double)    ',   '  массив VZ(см файл Variant)
'        'Dim Izdelie, N1fiz, N2fiz, kKND6, mmem, M0, KT, KC As Double
'        Dim N1f, N2f, kKND, kKVD, kKsum, Kv, njuOHL, mfors, mfiz, mfizmem, KOHL, n2pr As Double ', kKVD6
'        Dim Stend, ISren As Integer      'Rejim,Stend
'        Dim PRIV(40, IJK), PER(21, IJK), Tvozd(10), PpNas(10), F(4), Aptm(6, 11), Cptm(6) As Double 'PHYZ(45, IJK),VZ(32), A(19),
'        'ReDim_PHYZ(45, IJK)
'        'ReDim_VZ(32)
'        'ReDim_A(19)

'        Dim BINIT(32, j), CLet(45) As String 'CE, CF,, PRIZNAK CD,
'        '     	common/Kvvod/CD,CE,CF
'        '     common/Kvv/CD
'        'common /BOUND/ij
'        '	common/Kvvod1/PRIZNAK
'        '       '		common/phyzTVD/Gtopl
'        '   common/phyzPER/T3fizTDR
'        Aptm = New Double(,) {{0.2521923, 0.2079764, 0.1047056, 0.4489375, 0.2083632, 0.3070881}, _
'        {-0.1186612, 1.211806, 0.4234367, -0.1088401, -0.0112279, 0.2230734}, _
'        {0.3360775, -1.464097, -0.3953561, 0.4027652, 0.2235868, -0.4909}, _
'        {-0.3073812, 1.291195, 0.2249471, -0.2638393, -0.2732668, 0.5321652}, _
'        {0.1382207, -0.6385396, -0.07729786, 0.07993751, 0.1461334, -0.2756533}, _
'        {-0.03090246, 0.1574277, 0.0146217, -0.0115716, -0.03687021, 0.6851526}, _
'        {0.002745383, -0.01518199, -0.001166819, 0.0006241951, 0.003584204, -0.06596988}, _
'        {0.0, 0.0, 0.0, 0.0, 0.0, 0.0}, _
'        {0.0, 0.0, 0.0, 0.0, 0.0, 0.0}, _
'        {0.0, 0.0, 0.0, 0.0, 0.0, 0.0}, _
'        {0.85, 0.15, 0.0, 0.0, 10331.0, 0}}

'        Tvozd = New Double() {228.15, 253.15, 263.15, 273.15, 283.15, 288.15, 293.15, 303.15, 313.15, 323.15}
'        PpNas = New Double() {127.968, 127.968, 287.795, 610.381, 1227.693, 1704.907, 2332.75, 4238.94, 7371.49, 12330.25}

'        Debug.Print("into FIZ")
'        '	read(*,*)
'        '-------------------------------------------------------
'        '      N0=0  ИЛИ -999  -УСЛОВНОЕ ЧИСЛО (Метка отсутствия записи параметра) 
'        'CCC      N0=0  См.   COMMON/TARIR/Stend,Izdelie,Hu,NO,a1IPmin,a1IPmax,a1gradmin,
'        '-------------------------------------------------------
'        'If (jj > 0) Then GoTo 321 ' -обход оператора откр файла и однокр записи начиная со 2-го варианта
'        '	OPEN(325,FILE='B_txt\PHYZICAL1.txt') '- файл, в котором будет нах-ся матрица физ значений
'        '-------------------------------------------------------
'        If jj = 0 Then
'            OPEN(75, "B rez\PHYZICAL.xls")
'            OPEN(65, "'B rez\KV.xls")
'            '-------------------------------------------------------
'            OPEN(185, "B rez\Mistakes.TXT")
'            Debug.Print("(1)='Nкт'")
'            Debug.Print("(2)='Режим'")
'            Debug.Print("(3)='а1'")
'            Debug.Print("(4)='а2'")
'            Debug.Print("(5)='Ph'")
'            Debug.Print("(6)='Pb'")
'            Debug.Print("(7)='n1физ'")
'            Debug.Print("(8)='n2физ'")
'            Debug.Print("(9)='tb'")
'            Debug.Print("(10)='P*вх\КНД'")
'            Debug.Print("(11)='Р6СТ'")
'            Debug.Print("(12)='Р*6'")
'            Debug.Print("(13)='Р*6\s'")
'            Debug.Print("(14)='T*6'")
'            Debug.Print("(15)='Pi*КНД'")
'            Debug.Print("(16)='Pi*КНД\s'")
'            Debug.Print("(17)='etaКНД' ")
'            Debug.Print("(18)='etaКНДs'  ")
'            Debug.Print("(19)='Р2СТ'")
'            Debug.Print("(20)='Р*2'")
'            Debug.Print("(21)='Т*2'")
'            Debug.Print("(22)='РiКВД'")
'            Debug.Print("(23)='etaКВД'")
'            Debug.Print("(24)='PiKs'")
'            Debug.Print("(25)='etaKs'")
'            Debug.Print("(26)='T3физ'")
'            Debug.Print("(27)='T3физТДР' ")
'            Debug.Print("(28)='Т*4' ")
'            Debug.Print("(29)='Р4СТ'")
'            Debug.Print("(30)='Р*4'")
'            Debug.Print("(31)='PiT'")
'            Debug.Print("(32)='GBPHYZI'")
'            Debug.Print("(33)='GBфизB'")
'            Debug.Print("(34)='Rфиз'")
'            Debug.Print("(35)='Gт'")
'            Debug.Print("(36)='GтТДР'")
'            Debug.Print("(37)='альф_физ' ")
'            Debug.Print("(38)='CR' ")
'            Debug.Print("(39)='M'")
'            Debug.Print("(40)='Dpc'")
'            Debug.Print("(41)='Fca'  ")      ''Hu-топл'
'            Debug.Print("(42)='u447'")
'            Debug.Print("(43)='T4огр'")
'            Debug.Print("(44)='al_RUD.'")
'            Debug.Print("(45)='Fi,% отн вл'")


'            '----------------------------------------
'            Debug.Write("CLet(j5), J5 = 1, 45")
'            '235  format(45(G12.1,'	'))
'            '----------------------------------------   
'            jj = jj + 1
'        End If

'        '321:    'Continue Do
'        '---------------------------------------- 
'        eps = 0.00001   '   (машинный ноль)

'        j = VZ(1)
'        Nkt = VZ(2)
'        Rejim = VZ(3)
'        alRUD = VZ(4)
'        B0 = VZ(5)
'        Pb = VZ(6) 'измеренный
'        tbc = VZ(7)
'        Rizm = VZ(8)
'        N1f = VZ(9)
'        N2f = VZ(10)
'        freq = VZ(11)
'        Gtopl = VZ(12)
'        '	print *,' begin PHYZ:Gtopl',Gtopl
'        FcaTVD = VZ(13)
'        t300fiz = VZ(14)
'        u447 = VZ(15) + 273.15
'        T4 = VZ(16)       '+273.15
'        T4ogr = VZ(17)       '+273.15
'        Hbc = VZ(18)
'        Bmc = VZ(19)
'        P2aKVD = VZ(20)
'        P2bKVD = VZ(21)
'        '       ------ 21 12 2005   --------- ' Избавление от нулевого замера
'        If (P2aKVD < eps) Then
'            P2aKVD = P2bKVD
'            VZ(20) = P2aKVD
'        End If

'        If (P2bKVD < eps) Then
'            P2bKVD = P2aKVD
'            VZ(21) = P2bKVD
'        End If

'        P2CT = VZ(22)
'        '      --------21 12 2005    -------- ' Избавление от нулевого замера  
'        P6b = VZ(23)
'        P6CT = VZ(24)
'        P4b = VZ(25)
'        P4CT = VZ(26)
'        t6 = VZ(27)
'        del1 = VZ(28)
'        del2 = VZ(29)
'        del3 = VZ(30)
'        ISren = VZ(31)
'        fi = VZ(32)
'        '---------------------------------------- 

'        '1	Стенд	19
'        '2	Изделие	99004
'        '3	Hu	10331.     ! 10300.
'        '4	n0	0
'        '5	ИПа1мин	3.0
'        '6	ИПа1макс	106.0
'        '7	а1градмин	-30.0
'        '8	а1градмакс	0
'        '9	ИПа2мин		5
'        '10	ИПа2макс	117.0
'        '11	а2градмин	-25.0
'        '12	а2градмакс	3.0
'        '13	Дкрделмин	48.0
'        '14	Дкрделмакс	102.0
'        '15	Дкрминмм	620.8
'        '16	Дкрмаксмм	868.4
'        '17      aGtoplTDR	349.33 !5648.1	!-151.9                !23.17
'        '18	bGtoplTDR	21.454 !2.8119	!23.17                 !151.9
'        '19	iVAR	1	3	2    

'        Stend = A(1)
'        a1IPmin = A(5)
'        a1IPmax = A(6)
'        a1gradmin = A(7)
'        a1gradmax = A(8)
'        a2IPmin = A(9)
'        a2IPmax = A(10)
'        a2gradmin = A(11)
'        a2gradmax = A(12)
'        DkpIPmin = A(13)
'        DkpIPmax = A(14)
'        DkpMMmin = A(15)
'        DkpMMmax = A(16)
'        Hu = A(3)

'        ''''	CD=PRIZNAK
'        '---------------------------------------- 
'        '      WRITE(125,7706) K,Rejim,CD
'        ' 7706 format('PHYZICAL:K indexRej,Rejim:	',3G12.5) 
'        '      print *,'Rejim',Rejim
'        '	print *,CD
'        '      if(Rejim = 7.) goto 400
'        '      if(Rejim >= 3.2 andalso Rejim <= 6.) goto 401
'        '      if(Rejim = 3 andalso (CD = 'РПТ' orelse CD = 'Б' orelse CD = 'УБ'))
'        '     *goto 402
'        '      if(Rejim = 3 andalso (CD = 'MFC' orelse CD = 'MPIT'))
'        '     *goto 402
'        '	if(Rejim = 2.0 andalso (CD = 'Б' orelse CD = 'УБ')) Kv=1.018   '   Б 
'        '      if(Rejim = 1.0 andalso (CD = 'FALF'))    ' РПТ   ???'''''
'        '     *Kv=1.0145  '  РПТ ПФ
'        '       goto 405
'        ' 400	Kv=1.012
'        '      goto 405
'        ' 401  Kv=1.017   
'        '      goto 405
'        ' 402  Kv=1.018
'        ' 405  continue
'        '      continue
'        '      alfa1=0.
'        '	alfa2=0.
'        '	Dpc=0.
'        '      Kv=1.015
'        '      print *,'Kv',Kv
'        '	read(*,*)


'        '         --------  ВЛАЖНОСТЬ ВОЗДУХА -----05,06,2009-------
'        '  СМ. Е\М2\99М2\"влажность 99М2 2009".xls -исх данные по влажности воздуха
'        '           1. существует зависимость Рп. насыщ.(Тн,К)
'        '-------------------------------------------------------------------------------
'        '  Т,К    	253.15	263.15	273.1	283.15	    288.15	    293.15	303.15	313.15	323.15
'        'рП.Н,[Па] 	127.968	287.795	610.381	1227.693	1704.907	2332.75	4238.94	7371.49	12330.25
'        '------------------------------------------------------------------------------------
'        '	Data Тvozd/253.15,263.15,273.1,283.15,288.15,293.15,303.15,313.15,
'        '     *323.15/,
'        '     *PpNas/127.968,287.795,610.381,1227.693,1704.907,2332.75,4238.94,
'        '     *7371.49,12330.25/
'        '           2. по формуле d=0.622*PpNas*(Fi/100)/(101325-Ppnas*(Fi/100))
'        '       d -влагосодержание воздуха (по РТМ это количество килограммов водяного пара, приходящегося на 1 кг сух воздуха)
'        F(1) = tbc + 273.15
'        Call SPLINEЛинейный(9, Tvozd, PpNas, F, 222) '&222)
'        '          call SPLINE1(19,X,Y,F,&23)
'        PpNas1 = F(2)     ' давление насыщенных паров воды
'        dH2O1 = 0.622 * PpNas1 * fi / 100 / (101325 - PpNas1 * fi / 100) ' влагосодержание воздуха 

'        Debug.Print(" *,'F(1),Fi,PpNas1,dH2O',F(1),Fi,PpNas1,dH2O1")
'        '	  read(*,*)
'        '         --------  ВЛАЖНОСТЬ ВОЗДУХА ------05,06,2009---


'        '-----------------19.05.2009------------см описание в листе 99M2\B rez\analize.xls
'        If (alRUD >= 0 AndAlso alRUD <= 12) Then Kv = 1.0 + 0.011 / 12 * alRUD
'        If (alRUD > 12 AndAlso alRUD <= 30) Then Kv = 1.007 + 0.006 / 18 * alRUD
'        If (alRUD > 30 AndAlso alRUD <= 67) Then Kv = 1.019 - 0.002 * 67 / 37 + 0.002 / 37 * alRUD
'        If (alRUD > 67 AndAlso alRUD <= 78) Then Kv = 1.018 + 0.001 * 78 / 11 - 0.001 / 11 * alRUD
'        If (alRUD > 78 AndAlso alRUD <= 105) Then Kv = 1.015 + 0.003 * 105 / 27 - 0.003 / 27 * alRUD
'        If (alRUD > 105 AndAlso alRUD <= 116) Then Kv = 1.013 + 0.002 * 116 / 11 - 0.002 / 11 * alRUD
'        '      print *,'alRUD,Kv',alRUD,Kv
'        '	end do
'        '      read(*,*)
'        Debug.Write("alRUD,Kv")
'        '237  format(2(G12.5,'	'))
'        '      write(65,*) Kv,alRUD ',Kv)
'        '-----------------19.05.2009------------

'        If (del1 <> N0) Then alfa1 = a1gradmax + (a1gradmax - a1gradmin) * (del1 - a1IPmax) / (a1IPmax - a1IPmin)

'        If (del2 <> N0) Then alfa2 = a2gradmax + (a2gradmax - a2gradmin) * (del2 - a2IPmax) / (a2IPmax - a2IPmin)

'        If (del3 <> N0) Then Dpc = (-DkpIPmin + del3) * (-DkpMMmin + DkpMMmax) / (DkpIPmax - DkpIPmin) + DkpMMmin


'        Tb = tbc + 273.15


'        Ph = B0 / 735.6
'        Pbaz = Ph
'        '	print *,'a(1)',a(1)' -5-
'        'ввести признак базового давления
'        If (Stend = 25 OrElse Stend = 41 OrElse Stend = 13 OrElse Stend = 15) Then
'            '
'        End If
'        ' -  базовым явл давл в кабине В0
'        If (Stend = 19 OrElse Stend = 34) Then
'            ' - базовым явл давл в боксе: Pb    
'            Pbaz = Pb / 735.6    ' Рbaz- базовое                              ' -6-
'        End If

'        Debug.Print("PHYZICAL: stend number...do not EXSIST', Stend")

'        PvxCTm = Pbaz - Math.Abs(Hbc / 735.6)    ' - статическое давл в мерн сеч М
'        Pvxbm = Pbaz - Math.Abs(Bmc / 735.6)     ' - полное давл в мерн сеч М

'        '        CALL PARAMRMK(A,VZ,GBphyzB,ppB)
'        Call ПараметрыРасходомерногоКоллектора(Stend, Tb, PvxCTm, Pvxbm, GBphyzB, ppB)  ' ' параметры сухого воздуха 
'        Debug.Print("'dry: GBphyzB',GBphyzB")

'        Xptm = 0 ' для чистого воздуха
'        Tptm = Tb / 1000
'        dH2O = dH2O1
'        Call CPSptm(Cptm, Tptm, Xptm, Aptm, dH2O)
'        dk = Cptm(4) ' коэф-т адиабаты влажного воздуха
'        Debug.Print("'Cptm,dk,Tb',Cptm,dk,Tb")
'        '	read(*,*)   
'        'расчёт нового dk и заново вычислить - Виктор
'        Call RMKd(Stend, Tb, PvxCTm, Pvxbm, GBphyzB, ppB, dk) ' параметры воздуха с учётом влажности
'        Debug.Print("'wet: GBphyzB',GBphyzB")

'        sigVX = ppB / Pbaz
'        '       if(P6CT = N0) goto 68
'        If (P6CT < eps) Then
'            P6CTb = 0   '-если отсутствует замер P6CT ' P6CTb- стат давл за КНД
'        Else
'            P6CTb = Pbaz + P6CT ' P6CTb-истинное значение стат давления
'            '	P6CT -значение стат давления, измеренное диффер датчиком
'        End If
'        If (P6b < eps) Then
'            ' перестраховка, т.к. P6b=P6* замеряется всегда
'            P6bb = 0    '-если отсутствует замер P6b                     ' -13-
'        Else
'            P6bb = Pbaz + P6b    ' P6bb- полное давл за КНД
'        End If
'        'P6bbs = 0   '? не вычисляем значение Р*6 по знач стат давления 
'        If (t6 < eps) Then
'            ' t6 -Тстатич за КНД           ' -14-
'            T6b = 0                                                    ' -15-
'        Else
'            T6b = 273.15 + t6 ' это Т*6 за КНД
'        End If

'        PiKND = P6bb / ppB              '     (4.2.1)

'        'PiKNDs = 0  '? не вычисляем значение PiKND по знач стат давления    ' -16-
'        '      --------    Параметры КНД:   ------------
'        '	t6b=T6  ' -см 27-й столбец (Ввод ИД): Т6 это t*6,C
'        t06b = 0.5 * (tbc + t6)    ' t*6,C

'        If (t06b > 500) Then
'            Debug.Print(" TEMPERATURE t0-6 MORE THEN 500C' t06b")
'            '908  FORMAT(' TEMPERATURE t0-6 MORE THEN 500C',F6.1) ' АВАРИЙНАЯ ПЕЧАТЬ
'            Stop
'        End If


'        'If (t06b <= 50) Then
'        '    kKND = 1.4                     ' (4.2.3)
'        'ElseIf (t06b > 50 AndAlso t06b < 500) Then
'        '    kKND = 1.4035 - 5.05 * 0.00001 * t06b - 8.5 * 0.00000001 * t06b * t06b    ' (4.2.3)
'        'End If

'        kKND = Kadiabat(t06b)

'        If (T6b = Tb OrElse t6 < eps) Then
'            ethaKND = 0
'        Else
'            '      print*,'ethaKND,Tb,PiKND,kKND,T6b,Tb',ethaKND,Tb,PiKND,kKND,T6b,Tb
'            '	read(*,*)
'            ethaKND = Tb * (PiKND ^ ((kKND - 1) / kKND) - 1) / (T6b - Tb)    ' -17-  (4.2.2)
'            Debug.Print("'dry air:ethaKND,T06b,PiKND,kKND,T6b,Tb,dH2O',ethaKND,T06b,PiKND,kKND,T6b,Tb,dH2O")
'            '	read(*,*)

'            '   ------  кпд КНД--через энтальпии    --------------------------       '

'            PRptm = PiKND                                                        '
'            T1ptm = Tb                                                           '
'            Xptm = 0.0 ' для воздуха весовая доля топлива =0                     '
'            Debug.Print("'PRptm,T1ptm,Xptm',PRptm,T1ptm,Xptm  ")                      '                                              
'            '
'            TTTad = TPIptm(PRptm, T1ptm, Xptm, Aptm, Cptm, dH2O)  ' адиабатическая температура    '
'            Debug.Print("'P T M: TTTad=TPIptm(PRptm,T1ptm,Xptm,Aptm,Cptm,dH2O)', TTTad,PRptm,T1ptm")                                         '

'            '      READ(*,*)                                                          '
'            '                                                                        '
'            '  работа сжатия 1 кг воздуха в КНД равна разности энтальпий на вых и вх в КНД:
'            T0ptm = Tb / 1000    '  Tb                                            '  
'            Tptm = T6b / 1000   '  T6                                            '
'            Xptm = 0.0
'            Debug.Print("'T0ptm,Tptm,Xptm',T0ptm,Tptm,Xptm  ")                        '
'            aLkptmK = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                          '
'            aKND = aLkptmK                                                      ' 
'            '       aLkptmK=aLkptmK*4.1868                                           '
'            Debug.Print("'aLkptmK,DIptm,T0ptm,Tptm,Xptm', aLkptmK,T0ptm,Tptm,Xptm")
'            '                                                                        '
'            ' располагаемая ад. работа турбины 1 кг газа:                            '
'            T0ptm = Tb / 1000                                                    '
'            Tptm = TTTad / 1000                                                  '
'            Xptm = 0.0                                                           '
'            aLkptm = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                           '

'            ethaKND1 = aLkptm / aKND                ' кпд КНД                      '
'            Debug.Print("'ethaKND1=aLkptm/aKND',ethaKND1,aLkptm,aKND")
'            '	read(*,*)                                                          '
'            Debug.Print("'aLkptm,T0ptm,Tptm,Xptm',aLkptm,T0ptm,Tptm,Xptm")
'            ethaKND = Tb * (PiKND ^ ((kKND - 1) / kKND) - 1) / (T6b - Tb)                     '
'            Debug.Print("'wet air:ethaKND,T06b,PiKND,kKND,T6b,Tb,dH2O',ethaKND,T06b, GOOD REZULT 'PiKND,kKND,T6b,Tb,dH2O")
'            '   ------  кпд КНД--через энтальпии  -----------------                  '
'            '	read(*,*) 
'            '   ------  Учёт влажности воздуха k=k(d)-------	  
'        End If

'        'ethaKNDs = 0 '? не вычисляем значение ethaKNDs по знач стат давления 
'        If (P2CT < eps) Then
'            P2CT = 0
'        Else
'            P2CT = P2CT + Pbaz
'        End If
'        P2b = Pbaz + (P2aKVD + P2bKVD) * 0.99 / 2  ' -полное давл за КВД
'        '      print *,'into PHYZPARAM:P2b,Pbaz,P2aKVD,P2bKVD',
'        '     *P2b,Pbaz,P2aKVD,P2bKVD
'        '      read (*,*)
'        T2b = 273.15 + t300fiz ' температура перед КС
'        '------------------------------------------------------------
'        '       --------     Параметры КВД:  ------------
'        If (P6bb < eps) Then
'            '  P6bb -определено выше
'            PiKVD = 0
'        Else
'            PiKVD = P2b / P6bb
'        End If
'        t62b = 0.5 * (t300fiz + t6)   ' -t300fiz=t2b-за КВД ,t6b- за КНД  
'        If (t62b > 500) Then
'            Debug.Print("' TEMPERATURE t6-2 MORE THEN 500C' t62b")
'            '909  FORMAT(' TEMPERATURE t6-2 MORE THEN 500C',F6.1) ' АВАРИЙНАЯ ПЕЧАТЬ
'        End If

'        'If (t62b <= 50) Then
'        '    kKVD = 1.4
'        'ElseIf (t62b > 50 AndAlso t62b < 500) Then
'        '    kKVD = 1.4035 - 5.05 * 0.00001 * t62b - 8.5 * 0.00000001 * t62b * t62b
'        '    kKVD6 = (kKVD - 1) / kKVD
'        'End If

'        kKVD = Kadiabat(t62b)

'        T6b = 273.15 + t6

'        If (t6 < eps) Then
'            ethaKVD = 0
'            T6b = 0.0
'        Else
'            ethaKVD = T6b * (PiKVD ^ ((kKVD - 1) / kKVD) - 1) / (T2b - T6b)   ' -22-
'            Debug.Print("'dry air:ethaKVD,T62b,PiKVD,kKVD,T2b,T6b,dH2O',ethaKVD,T62b,PiKVD,kKVD,T2b,T6b,dH2O")
'            '	read(*,*)

'            '   ------  Учёт влажности воздуха k=k(d)--КВД--
'            Tptm = (t62b + 273.15) / 1000
'            Xptm = 0 ' для чистого воздуха
'            Call CPSptm(Cptm, Tptm, Xptm, Aptm, dH2O)
'            dk2 = Cptm(4)
'            kKVD = dk2
'            Debug.Print("'Cptm,dk1,Tb',Cptm,dk1,Tb")
'            ethaKVD = T6b * (PiKVD ^ ((kKVD - 1) / kKVD) - 1) / (T2b - T6b)   ' -22-
'            Debug.Print("'wet air:ethaKVD,T62b,PiKVD,kKVD,T2b,T6b,dH2O',ethaKVD,T62b,PiKVD,kKVD,T2b,T6b,dH2O")
'            '	read(*,*) 

'            '   ------  кпд КВД--через энтальпии    ---------------------------'

'            PRptm = PiKVD                                                        '
'            T1ptm = T6b                                                          '
'            Xptm = 0.0  ' для воздуха весовая доля топлива =0                    '
'            Debug.Print("'PRptm,T1ptm,Xptm',PRptm,T1ptm,Xptm'")
'            '                                                                        '
'            TTTad = TPIptm(PRptm, T1ptm, Xptm, Aptm, Cptm, dH2O)  ' адиабатическая температура    '
'            Debug.Print("'P T M: TTTad=TPIptm(PRptm,T1ptm,Xptm,Aptm,Cptm,dH2O)',TTTad,PRptm,T1ptm ")

'            '      READ(*,*)                                                          '
'            '                                                                        '
'            '  работа сжатия 1 кг воздуха в КВД равна разности энтальпий на вых и вх в КВД:
'            T0ptm = T6b / 1000    '  Tb                                            '  
'            Tptm = T2b / 1000    '  T6                                            '
'            Xptm = 0.0
'            Debug.Print("'T0ptm,Tptm,Xptm',T0ptm,Tptm,Xptm'")
'            aLkptmK = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                          '
'            aKND = aLkptmK                                                      ' 
'            '       aLkptmK=aLkptmK*4.1868                                           '
'            Debug.Print("'aLkptmK,DIptm,T0ptm,Tptm,Xptm', aLkptmK,T0ptm,Tptm,Xptm")
'            '                                                                        '
'            ' располагаемая ад. работа турбины 1 кг газа:                            '
'            T0ptm = T6b / 1000                                                 '
'            Tptm = TTTad / 1000                                                 '
'            Xptm = 0.0                                                           '
'            aLkptm = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                           '

'            ethaKVD1 = aLkptm / aKND                ' кпд КНД                      '
'            Debug.Print("'ethaKVD1=aLkptm/aKND',ethaKVD1,aLkptm,aKND")  ' GOOD REZULT '  2
'            '	read(*,*)                                                          '
'            '   ------  Учёт влажности воздуха k=k(d)-------	  
'        End If

'        '---------------------------------------------------------------
'        '       --------     Параметры Кsum:  ------------	 
'        PiKsum = P2b / ppB
'        t02b = 0.5 * (tbc + t300fiz)
'        If (t02b > 500) Then
'            Debug.Print("' TEMPERATURE t0-2 MORE THEN 500C' t02b")
'            '1909  FORMAT(' TEMPERATURE t0-2 MORE THEN 500C',F6.1) ' АВАРИЙНАЯ ПЕЧАТЬ
'            Stop
'        End If

'        'If (t62b <= 50) Then
'        '    kKVD = 1.4
'        'ElseIf (t02b > 50 AndAlso t02b < 500) Then
'        '    kKsum = 1.4035 - 5.05 * 0.00001 * t02b - 8.5 * 0.00000001 * t02b * t02b
'        'End If
'        ' исправлено kKsum и t62b
'        kKsum = Kadiabat(t02b)

'        Tbb = tbc + 273.15
'        ethaKsum = Tbb * (PiKsum ^ ((kKsum - 1) / kKsum) - 1) / (T2b - Tbb)  ' -24-    (4.4.2) tbb см метку 92
'        Debug.Print("'ethaKsum=Tbb*(PiKsum**((kKsum-1)/kKsum)-1)/(t2b-tbb)',ethaKsum,Tbb,PiKsum,kKsum,kKsum,t2b,tbb")
'        '   ------  сумм кпд компрессора--через энтальпии    --------------'

'        PRptm = PiKsum                                                       '
'        T1ptm = Tb                                                           '
'        Xptm = 0.0  ' для воздуха весовая доля топлива =0                    '
'        Debug.Print("'PRptm,T1ptm,Xptm',PRptm,T1ptm,Xptm ")                       '
'        '                                                                        '
'        TTTad = TPIptm(PRptm, T1ptm, Xptm, Aptm, Cptm, dH2O)  ' адиабатическая температура    '
'        Debug.Print("'P T M: TTTad=TPIptm(PRptm,T1ptm,Xptm,Aptm,Cptm,dH2O)',TTTad,PRptm,T1ptm")                                                  '

'        '      READ(*,*)                                                          '
'        '                                                                        '
'        '  работа сжатия 1 кг воздуха в КВД равна разности энтальпий на вых и вх в компрессор:
'        T0ptm = Tb / 1000    '  Tb                                            '  
'        Tptm = T2b / 1000    '  T6                                           '
'        Xptm = 0.0
'        Debug.Print("'T0ptm,Tptm,Xptm',T0ptm,Tptm,Xptm")                          '
'        aLkptmK = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                               '
'        aKsum = aLkptmK                                                     ' 
'        '       aLkptmK=aLkptmK*4.1868                                           '
'        Debug.Print("'aKsum,DIptm,T0ptm,Tptm,Xptm',   aKsum,T0ptm,Tptm,Xptm")
'        '                                                                        '
'        ' располагаемая ад. работа турбины 1 кг газа:                            '
'        T0ptm = Tb / 1000                                                     '
'        Tptm = TTTad / 1000                                                   '
'        Xptm = 0.0                                                           '
'        aLkptm = DIptm(T0ptm, Tptm, Xptm, Aptm, dH2O)                           '

'        ethaKsum1 = aLkptm / aKsum                ' кпд КНД                    '
'        Debug.Print("'ethaKsum1=aLkptm/aKsum',ethaKsum1,aLkptm,aKsum") ' GOOD REZULT ' 3
'        '	read(*,*)                                                          '
'        '   ------  Учёт влажности воздуха k=k(d)-------	  


'        '------------------------      T3fizTDR        -------------------

'        '  --------------------   8.06.2009г   --------------
'        'if(freq-N0)300,300,301 ' частота датчика ТДР

'        If freq <= N0 Then
'            GtoplTDR = 0
'            T3fizTDR = 0
'        Else
'            ' 301  GtoplTDR=aGtoplTDR+bGtoplTDR*freq     '   -151.90+23.17*N TDR     ' 7 Sept 2005  

'            GtoplTDR = A(17) + A(18) * freq
'            'GtoplTDR = bGtoplTDR + aGtoplTDR * freq
'            'сделать расчет перевода литров в килограмм
'            Debug.Print("'GtoplTDR=a(17)+a(18)*freq',GtoplTDR,a(17),a(18),freq")
'            If GtoplTDR < eps Then
'                GtoplTDR = 0
'                T3fizTDR = 0
'                ''      if(GtoplTDR < eps) goto 303

'                '        Gtopl=Gtopl1    '   возвр знач Gtopl из промежут ячейки
'                '' 303  continue
'                '      print *,'GtoplTDR,Gtopl,Gtopl1',GtoplTDR,Gtopl,Gtopl1
'                '	read(*,*)
'            Else
'                '      Gtopl1=Gtopl           
'                ' 	Gtopl=GtoplTDR   ' - режим Мах, Дросс, Мг:

'                '		call PARAM TVD(Rejim,T2b,GtoplTDR,FcaTVD,P2b,A(3),      '9 Nov 2005   ' GtoplTDR 
'                '     *qT1,GBPHYZca,Thhot,ISren)
'                '     print *,'GtoplTDR',GtoplTDR
'                '	Gtopl=GtoplTDR
'                '	call PARAM TVD(A,VZ,GtoplTDR,GBPHYZca,Thhot)

'                Call TVD(Rejim, ISren, Hu, FcaTVD, P2b, T2b, GtoplTDR, GBPHYZca, Thhot)
'                T3fizTDR = Thhot   '  -режим Мах 
'                Debug.Print("'T3fizTDR,GtoplTDR,'T3fizTDR,GtoplTDR")
'            End If
'        End If

'        '        ---------вычисление Т*3 физ--------


'        'print *,'---Rejim',Rejim

'        Debug.Print("'+++CD','	',CD")
'        'If (CD = "F") Then GoTo 297 'print *,'F'

'        If (CD = "M") Then

'            Debug.Print("'===M'")
'            '      goto 297
'            '	read(*,*)
'            '------------------------------------------------------------- 
'            '                       ^^^^^^^^^^^^^^^^^^^
'            '                       Максимал, Дросс, Мг:
'            '
'            '                       ^^^^^^^^^^^^^^^^^^^
'            '                  Gtopl1=Gtopl  ' запомин в промеж яч: Gtopl1 значение Gtopl
'            '              call PARAM TVD  ' -определение т-ры газа перед ТВД в кр сеч СА ТВД 
'            '			                   и расх воздуха через СА ТВД. 
'            '                                Для ПФ и Ф вышеупомянутые параметры ,а также т(степ двухконт)
'            '                                вычисляются по формулам (4.6.5) и (4.6.6)
'            '       Gtopl=Gtopl1
'            '       print *,'Gtopl',Gtopl

'            '      call PARAM TVD(Rejim,T2b,Gtopl,FcaTVD,P2b,A(3),             '9 Nov 2005   ' Gtopl 
'            '     *qT1,GBPHYZca,Thhot,ISren)

'            '''''''''''''''''''''''''''''''''''''''''
'            '         call PARAM TVD(Rejim,T2b,GtoplTDR,FcaTVD,P2b,a(3),      '9 Nov 2005   ' GtoplTDR 
'            '     *qT1,GBPHYZ1,Thhot,ISren)
'            '''''''''''''''''''''''''''''''''''''''''''''
'            Gtopl = VZ(12)

'            '	print *,'Gtopl=VZ(12)',Gtopl
'            '      call PARAM TVD(A,VZ,Gtopl,GBPHYZca,Thhot)
'            ''       call PARAM TVD(Rejim,ISren,Hu,FcaTVD,P2b,T2b,Gtopl,GBPHYZca,Thhot)
'            '	print *,'Rejim,ISren,Hu,FcaTVD,P2b,T2b,Gtopl,GBPHYZca,Thhot,njuOHL'
'            '      print *,Rejim,ISren,Hu,FcaTVD,P2b,T2b,Gtopl,GBPHYZca,Thhot,njuOHL
'            '	read(*,*)
'            '*****************************************************************************
'            Call TVD(Rejim, ISren, Hu, FcaTVD, P2b, T2b, Gtopl, GBPHYZca, Thhot)
'            Debug.Print("'L.M.:Gtopl,GBPHYZca,Thhot',Gtopl,GBPHYZca,Thhot  ")
'            Call TVD1(Rejim, ISren, Hu, FcaTVD, P2b, T2b, Gtopl, GBCA, T3f)
'            Debug.Print("'Jankin:Gtopl,GBCA,T3f',Gtopl,GBCA,T3f")
'            Call GGPARAM_TVD_ptm(Rejim, ISren, Skc1, Hu, FcaTVD, P2b, T2b, Gtopl, GBCAptm, T3captm, Xptm, Aptm, Cptm, dH2O)
'            Debug.Print("'PTM:Gtopl,GBCA,T3f',Gtopl,GBCA,T3f ")
'            '*****************************************************************************

'            'SUBROUTINE PARAM TVD(Rejim,Hu,FcaTVD,P2b,T2b,Gtopl,GBCA,T3f)
'            T3fizmem = Thhot      ' - режим Мах, Дросс, Мг:
'            Debug.Print("'--Gtopl,T3fizmem',Gtopl,T3fizmem")
'            '-------------------- вычисление Пи турбины и КПД турбины---------------
'            'Continue Do
'            T4B = T4 + 273.15 '-температура газов за ТНД
'            '------------------------
'            '      if(P4CT-N0)1301,1300,1301
'            ' 1300	P4CT=0.
'            '      goto 1302
'            ' 1301     P4CT=Pbaz+P4CT
'            ' 1302 continue 
'            '------------------------

'            If (P4CT < eps) Then
'                P4CTb = 0
'            Else
'                P4CTb = Pbaz + P4CT
'            End If
'            P4bb = Pbaz + P4b
'            PiT = P2b * 0.95 / P4bb
'            Rfiz = 1.01 * Rizm * Kv
'            Debug.Print("'Rfiz,Rizm,Kv',Rfiz,Rizm,Kv")
'            '	read(*,*)
'            Debug.Write("'Max: Rejim,CD,Kv,Rfiz,Rizm:	' Rejim,CD,Kv,Rfiz,Rizm")
'            '7708 format('Max: Rejim,CD,Kv,Rfiz,Rizm:	',5G12.5)
'            '      print *,'max:alfafiz,GBphyzB,Gtopl',alfafiz,GBphyzB,Gtopl
'            '	read(*,*)
'            alfafiz = 3600 * GBphyzB / Gtopl / 14.95
'            CR = Gtopl / Rfiz                             '  ----   CR
'            If (Gtopl < eps) Then
'                Gtopl = 0.0
'                T3fizmem = 0.0
'                CR = 0.0
'                alfafiz = 0.0
'            End If



'            '      print *,'Gtopl',Gtopl
'            '	read(*,*)
'            '----------------------------------------------------------------
'            '      ОЦЕНКА СТЕПЕНИ ДВУХКОНТУРНОСТИ И ПАРАМЕТРОВ ТУРБИНЫ:   (njuOHL) 

'            '----------  охлаждение СА ТВД  -----------
'            '                  kOHL
'            n2pr = N2f * Math.Sqrt(288.15 / Tb)
'            '   -------   New WIEW:  ISren=0(отсутств сигнала), ISren=1(частич охл), ISren=5(полное охл)  -----------


'            '   3 Октября 2005г
'            '      njuOHL=0.88               '-njuOHL- (НЮ)коэфф отн расх возд на охл СА ТВД  
'            '     if(ISren > 1 orelse n2 > 90 orelse n2pr < 80.) goto 147
'            '     njuOHL=0.94
'            ' 147  continue
'            '

'            KOHL = ISren
'            '-njuOHL- (НЮ)коэфф отн расх возд на охл СА ТВД
'            If (KOHL = 0) Then njuOHL = 0.94
'            If (KOHL >= 1 AndAlso N2f > 91.5) Then njuOHL = 0.88 '  n2fiz > 0.90
'            If (KOHL >= 1 AndAlso n2pr < 80.0) Then njuOHL = 0.89 '  n2fiz < 0.80

'            '      print *,'kOHL,njuOHL',kOHL,njuOHL
'            '      GBPHYZs=GBPHYZ1/njuOHL   ' здесь GBPHYZ1 -это GвСА (4,6,1); GBPHYZ1 получ из п/п /PAR TVD
'            GBPHYZI = GBPHYZca / njuOHL   ' здесь GBPHYZca -это GвСА (4,6,1); GBPHYZ1 получ из п/п /PAR TVD

'            '      print *,'not FORS: GBPHYZI,GBPHYZca,njuOHL,Thhot',
'            '     *	GBPHYZI,GBPHYZca,njuOHL,Thhot
'            '      if(Rejim = 1 orelse Rejim = 2) goto 183    ' для Форсажа T*Г и mф вычисляются по другим формулам
'            '      mfizmem=(GBphyzB-GBPHYZs)/GBPHYZs     ' GBphyzB  получ из п/п  PARRMK
'            mfizmem = (GBphyzB - GBPHYZI) / GBPHYZI ' GBphyzB  получ из п/п  PARRMK, GBPHYZI- расх возд через 1-й конт

'            '183:        'Continue Do
'            PiT = P2b * 0.95 / P4bb '  ?????       ' 28 столбец Сержа
'            '	if(Rejim = 1 orelse Rejim = 2) goto 83    
'            uMem = u447         ' -запоминание значения u447 для М режима
'            PiTMem = PiT         ' -запоминание значения PiT для М режима
'            '83:
'            T4ogr = T4ogr + 273.15   ' -температура ограничения 855С

'            'End If

'            '-------------проверка: на наличие форсажа (т.е. анализ значения Rejim)
'            'If (Rejim > 2) Then GoTo 6789 '184  ' до 19,03,2005 было 1, а не 2
'            'Для М1
'            '1-ПФ115
'            '2-МФ78
'            '3-М
'            '4-КР
'            'Для М2
'            '1-F
'            '3-M

'            '297:    'Continue Do

'        ElseIf (CD = "F") Then

'            Gtopl = VZ(12)
'            '	print *,'print  FORS: Gtopl',Gtopl
'            '           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^  
'            T4B = T4 + 273.15
'            'if(P4CT-N0)2301,2300,2301

'            If P4CT = N0 Then
'                P4CTb = 0
'            Else
'                P4CTb = Pbaz + P4CT
'            End If
'            P4bb = Pbaz + P4b
'            PiT = P2b * 0.95 / P4bb
'            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'            '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'            Tgfors = T3fizmem + 1.4 * (u447 - uMem)     ' -T*3fiz   -25-' -запоминание значения u447 для М режима
'            '	print *,'FORS: GBPHYZI,GBPHYZca,njuOHL,Tgfors',
'            '     *	GBPHYZI,GBPHYZca,njuOHL,Tgfors
'            Debug.Print("'Tgfors,T3fizmem,u447,uMem',Tgfors,T3fizmem,u447,uMem")
'            '	read(*,*)
'            '1778:       'Continue Do
'            mfors = mfizmem + 0.23 * (PiT - PiTMem)              '           -38-
'            Rfors = 1.01 * Rizm * Kv
'            Rfiz = 1.01 * Rizm * Kv

'            Debug.Write("'Fors: Rejim,CD,Kv,Rizm,Rfiz:	' Rejim,CD,Kv,Rizm,Rfiz")
'            '7709 format('Fors: Rejim,CD,Kv,Rizm,Rfiz:	',5G12.5)
'            '           -33-
'            '      print *,'fors:alfafiz,GBphyzB,Gtopl',alfafiz,GBphyzB,Gtopl
'            '	read(*,*)
'            alfafiz = 3600 * GBphyzB / Gtopl / 14.95   '           -36-
'            CR = Gtopl / Rfors                     '           -37- 
'            If (Gtopl < eps) Then
'                Gtopl = 0.0
'                T3fizmem = 0.0
'                Tgfors = 0.0
'                CR = 0.0
'                alfafiz = 0.0
'            End If
'            '      print *,'++Gtopl,T3fizmem,CR,alfafiz',Gtopl,T3fizmem,CR,alfafiz
'            u447 = u447                    '           -41-  
'            T4ogr = T4ogr + 273.15
'            mfiz = mfors

'            '                              	 Tgfors=Thhot+1.4*(u447-uM) 

'            '      stop
'            '2779:       GoTo 6790 ' -выход из форс режима

'            '            Fca = FcaTVD
'            '6790:       'Continue Do
'            '          6 September 2005
'            '      Запоминание(memorization) параметров на текущем шаге
'            T3fizmem = Tgfors
'            mfizmem = mfors
'            uMem = u447
'            PiTMem = PiT
'        End If


'        '6789:   'Continue Do
'        '      read(*,*)
'        '////////////////////////////////////////////////////////////////////////
'        PHYZ(1, j) = Nkt
'        PHYZ(2, j) = Rejim
'        PHYZ(3, j) = alfa1
'        PHYZ(4, j) = alfa2
'        PHYZ(5, j) = Ph
'        PHYZ(6, j) = Pbaz
'        PHYZ(7, j) = N1f
'        PHYZ(8, j) = N2f    '
'        PHYZ(9, j) = Tb ' tb           	          
'        PHYZ(10, j) = ppB '	 P*вх\КНД 
'        PHYZ(11, j) = P6CTb ' Р6СТ   
'        PHYZ(12, j) = P6bb '	 Р*6 
'        'PHYZ(13, j) = P6bbs ' 	 Р*6\s      
'        PHYZ(14, j) = T6b
'        PHYZ(15, j) = PiKND
'        'PHYZ(16, j) = PiKNDs
'        PHYZ(17, j) = ethaKND 'КПД кнд *
'        'PHYZ(18, j) = ethaKNDs
'        PHYZ(19, j) = P2CT
'        PHYZ(20, j) = P2b
'        PHYZ(21, j) = T2b
'        PHYZ(22, j) = PiKVD
'        PHYZ(23, j) = ethaKVD
'        PHYZ(24, j) = PiKsum
'        PHYZ(25, j) = ethaKsum
'        PHYZ(26, j) = T3captm   'T3fizmem
'        PHYZ(27, j) = T3fizTDR
'        PHYZ(28, j) = T4B

'        PHYZ(29, j) = P4CTb
'        PHYZ(30, j) = P4bb

'        PHYZ(31, j) = PiT
'        PHYZ(32, j) = GBCAptm  'GBPHYZI
'        PHYZ(33, j) = GBphyzB
'        If (Rfiz < eps) Then CR = 0.0
'        PHYZ(34, j) = Rfiz
'        PHYZ(35, j) = Gtopl
'        PHYZ(36, j) = GtoplTDR
'        PHYZ(37, j) = alfafiz
'        PHYZ(38, j) = CR
'        PHYZ(39, j) = mfizmem
'        PHYZ(40, j) = Dpc
'        PHYZ(41, j) = VZ(13)       'A(3)  '      Hu
'        PHYZ(42, j) = u447
'        PHYZ(43, j) = T4ogr
'        '      PHYZ(44,j)=sigVX
'        PHYZ(44, j) = alRUD
'        PHYZ(45, j) = fi
'        '      	lki=lki+1

'        '      CALL PRIVEDENIE(IJK,VZ,PHYZ,PRIV,GBPHYZI)                       

'        '	CALL PERESCHET(IJK,VZ,PHYZ,PER) 
'        'Return
'        '222:    Debug.Print("'spline Wet Air: dH2O=',dH2O ")
'    End Sub

'    'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF



'    '       --------    РТМ   23.03.2009        --------------------НАЧАЛО-----/////// 

'    '                     ---------------- Р Т М  ---------------------


'    '                         PPPPPP     TTTTTTTTT    M     M
'    '                         P     P        T        MM   MM
'    '                         P     P        T        M M M M
'    '                         PPPPPP         T        M  M  M
'    '                         P              T        M     M
'    '                         P              T        M     M
'    '                         P              T        M     M



'    '----------              CPS(C,T,X,A)             ------------   ----НАЧАЛО------

'    ' ОПРЕДЕЛЕНИЕ Т.ЁМК(С), ГАЗ. ПОСТ.И ПОКАЗ. АДИАБАТЫ ПО ЗАД. Т-РЕ Т И КОЭФФ. ИЗБ. ВОЗДУХА АЛЬФА(Х=1/АЛЬФА)
'    '  	C(1) - Cp    C(2) -R    C(4) - k     C(6) - L0    

'    Private Sub CPSptm(ByVal Cptm() As Double, ByVal Tptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal dH2O As Double)
'        'ReDim_Aptm(6, 11)
'        Dim P(10) As Double
'        'ReDim_Cptm(6)

'        '  	C(1) - Cp    C(2) -R    C(4) - k     C(6) - L0    
'        Cptm(6) = 28.966 / 0.2095 * (Aptm(1, 11) / 12.01 + Aptm(2, 11) / 4.032 - Aptm(3, 11) / 32.0)
'        Aptm(4, 11) = dH2O
'        D = Aptm(4, 11)
'        '	print *,'D',D
'        '	read(*,*)
'        QT = (1.0 - D) * Xptm / Cptm(6)
'        M = 6
'        N = M + 1



'        If (QT < 0.00001 AndAlso D < 0.00001) Then
'            '2:
'            Cptm(3) = 29.27
'            For I As Integer = 1 To N
'                '		print *,'////i,i,A(1,I),D,A(4,I),P(I)',i,A(1,I),D,A(4,I)
'                P(I) = Aptm(1, I)    '  AIR KOEFFICIENTS 
'            Next
'            '      read(*,*)
'            'GoTo 10
'        Else
'            If QT < 0.00001 Then
'                '14:
'                Cptm(3) = 29.27 * (1.0 + 0.60779 * D)
'                For I As Integer = 1 To N
'                    P(I) = Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D
'                Next
'                '      print *,'---i,A(1,I),D,A(4,I),P(I)',i,A(1,I),D,A(4,I)
'                'GoTo 10
'            End If
'            '15:
'            'IF (Aptm(2,11)-0.999)1,5,5   ' 0.15
'            If Aptm(2, 11) < 0.999 Then
'                'GoTo 1
'                '1:
'                If (Math.Abs(Aptm(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(Aptm(2, 11) - 0.15) < 0.001 AndAlso D < 0.000001) Then
'                    'GoTo 3
'                    '3:
'                    Cptm(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
'                    For I As Integer = 1 To N
'                        P(I) = (Aptm(2, I) * QT + Aptm(1, I)) / (1.0 + QT)
'                    Next
'                    'GoTo 10
'                Else
'                    Cptm(3) = 29.27 * (1.0 + 28.966 * (Aptm(2, 11) / 4.032 + Aptm(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
'                    For I As Integer = 1 To N
'                        AU = (44.01 * Aptm(3, I) - 32.0 * Aptm(5, I)) / 12.01
'                        BU = (9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008
'                        P(I) = ((Aptm(1, 11) * AU + Aptm(2, 11) * BU + Aptm(3, 11) * Aptm(5, I)) * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Debug.Print("'****i,Aptm(1,I)',i,Aptm(1,I)")
'                    Next
'                    'GoTo 10
'                End If
'            Else
'                'GoTo 5
'                '5  IF(Xptm-1.0)7,8,8
'                If Xptm < 1.0 Then
'                    '7:
'                    Cptm(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
'                    For I As Integer = 1 To N
'                        P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                    Next
'                    'GoTo 10
'                Else
'                    '8:
'                    Cptm(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
'                    For I As Integer = 1 To N
'                        P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * (1.0 - D) / Cptm(6) + D * Aptm(4, I) + (1.0 - D) * Aptm(1, I) + (QT - (1.0 - D) / Cptm(6)) * Aptm(6, I)) / (1.0 + QT)
'                    Next
'                End If
'            End If
'        End If


'        '10:
'        Tptm = Tptm + 0.001
'        Cptm(1) = P(M + 1)

'        For I As Integer = 1 To M ',1
'            Cptm(1) = Cptm(1) * Tptm + P(M + 1 - I)
'        Next
'        Cptm(4) = Cptm(1) / (Cptm(1) - Cptm(3) / 426.9)
'        Tptm = Tptm + 1000
'        '	RETURN 
'    End Sub
'    '----------              CPS(C,T,X,A)       ---------------------КОНЕЦ------





'    '---------------     FUNCTION DI(T0,T,X,A)  ( P400(1))     -------     --НАЧАЛО------

'    ' ОПРЕДЕЛЕНИЕ ПРИРАЩЕНИЯ ЭНТАЛЬПИИ ОТ T0 ДО T ПРИ ЗАД. АЛЬФА(Х=1/АЛЬФА)
'    Private Function DIptm(ByVal T0ptm As Double, ByVal Tptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal dH2O As Double)
'        'subroutine(DI(T0, T, X, A, D7))
'        Dim L0, I0, I1 As Double
'        Dim P(10) As Double
'        'ReDim_Aptm(6, 11)
'        L0 = 28.966 / 0.2095 * (Aptm(1, 11) / 12.01 + Aptm(2, 11) / 4.032 - Aptm(3, 11) / 32.0)
'        '	 PRINT *,'++++++++++++++'

'        '      PRINT *,'L0,Aptm(1,11),Aptm(2,11),Aptm(3,11)',
'        '     *L0,Aptm(1,11),Aptm(2,11),Aptm(3,11)

'        '	 PRINT *,'++++++++++++++'
'        Aptm(4, 11) = dH2O
'        D = Aptm(4, 11)       ' V.G.
'        QT = (1.0 - D) * Xptm / L0
'        M = 6
'        N = M + 1
'        If (QT < 0.00001 AndAlso D < 0.00001) Then
'            '2:
'            For I As Integer = 1 To N
'                P(I) = Aptm(1, I)
'            Next
'            'GoTo 50
'        Else
'            'if(QT-0.00001) 14,15,15
'            If QT < 0.00001 Then
'                '14:
'                For I As Integer = 1 To N
'                    P(I) = Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D   '   KOZEL''''
'                Next
'                'GoTo 50
'            Else
'                '15:  IF (Aptm(2,11)-0.999)10,20,20
'                If Aptm(2, 11) < 0.999 Then
'                    '10:
'                    If (Math.Abs(Aptm(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(Aptm(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
'                        For I As Integer = 1 To N
'                            P(I) = (Aptm(2, I) * QT + Aptm(1, I)) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        For I As Integer = 1 To N
'                            AU = (44.01 * Aptm(3, I) - 32.0 * Aptm(5, I)) / 12.01
'                            BU = (9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008
'                            P(I) = ((Aptm(1, 11) * AU + Aptm(2, 11) * BU + Aptm(3, 11) * Aptm(5, I)) * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    End If
'                Else
'                    '20:  IF(Xptm-1.0)70,80,80
'                    If Xptm < 1.0 Then
'                        '70:
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        '80:
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * (1.0 - D) / L0 + D * Aptm(4, I) * (1.0 - D) * Aptm(1, I) + (QT - (1.0 - D) / L0) * Aptm(6, I)) / (1.0 + QT)
'                        Next
'                    End If
'                End If
'            End If
'        End If

'        '50:
'        T0ptm = T0ptm + 0.001
'        Tptm = Tptm + 0.001
'        I0 = P(M + 1) / (M + 1)
'        I1 = I0
'        For I As Integer = 1 To M
'            I0 = I0 * T0ptm + P(M + 1 - I) / (M + 1 - I)
'            I1 = I1 * Tptm + P(M + 1 - I) / (M + 1 - I)
'        Next
'        DIptm = (I1 * Tptm - I0 * T0ptm) * 1000.0
'        'DI7 = (I1 * T - I0 * T0) * 1000.0
'        Tptm = Tptm + 1000.0
'        T0ptm = T0ptm + 1000.0
'        'Return
'    End Function

'    '----------              FUNCTION DI(T0,T,X,A)   ( P400(1))          -----КОНЕЦ-----




'    '----------              FUNCTION TI(DI,T0,X,A,C) ( P400(2))            -----НАЧАЛО------

'    '   ОПРЕДЕЛЕНИЕ Т-РЫ Т ПО НАЧ. Т-РЕ Т0, ПРИРАЩЕНИЮ ЭНТАЛЬПИИ DI И АЛЬФА(Х=1/АЛЬФА) 
'    Private Function TIptm(ByVal DIptm As Double, ByVal T0ptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal Cptm() As Double, ByVal dH2O As Double)
'        Dim I0, I1 As Double
'        Dim P(10) As Double
'        'ReDim_Cptm(6)
'        'ReDim_Aptm(6, 11)

'        Cptm(6) = 28.966 / 0.2095 * (Aptm(1, 11) / 12.01 + Aptm(2, 11) / 4.032 - Aptm(3, 11) / 32.0)
'        Aptm(4, 11) = dH2O
'        D = Aptm(4, 11)
'        QT = (1.0 - D) * Xptm / Cptm(6)
'        M = 6
'        N = M + 1
'        If (QT < 0.00001 AndAlso D < 0.00001) Then
'            Cptm(3) = 29.27
'            For I As Integer = 1 To N
'                P(I) = Aptm(1, I)
'            Next
'            'GoTo 50
'        Else
'            'if(QT-0.00001) 14,15,15
'            If QT < 0.00001 Then
'                '14:
'                Cptm(3) = 29.27 * (1.0 + 0.60779 * D)
'                For I As Integer = 1 To N
'                    P(I) = Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D
'                Next
'                'GoTo 50
'            Else
'                '15  IF (Aptm(2,11)-0.999)10,20,20
'                If Aptm(2, 11) < 0.999 Then
'                    '10:
'                    If (Math.Abs(Aptm(1, 11) - 0.85) < 0.001 AndAlso Math.Abs(Aptm(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
'                        '                        GoTo 30
'                        '30:
'                        Cptm(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = (Aptm(2, I) * QT + Aptm(1, I)) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        Cptm(3) = 29.27 * (1.0 + 28.966 * (Aptm(2, 11) / 4.032 + Aptm(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            AU = (44.01 * Aptm(3, I) - 32.0 * Aptm(5, I)) / 12.01
'                            BU = (9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008
'                            P(I) = ((Aptm(1, 11) * AU + Aptm(2, 11) * BU + Aptm(3, 11) * Aptm(5, I)) * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    End If
'                Else
'                    '20:  IF(Xptm-1.0)70,80,80
'                    If Xptm < 1.0 Then
'                        '70:
'                        Cptm(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        '80:
'                        Cptm(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * (1.0 - D) / Cptm(6) + D * Aptm(4, I) * (1.0 - D) * Aptm(1, I) + (QT - (1.0 - D) / Cptm(6)) * Aptm(6, I)) / (1.0 + QT)
'                        Next
'                    End If
'                End If
'            End If
'        End If

'        '50:
'        T0ptm = T0ptm * 0.001
'        T1 = T0ptm + DIptm / 0.25 * 0.001
'        I0 = P(M + 1) / (M + 1)
'        For I As Integer = 1 To M
'            I0 = I0 * T0ptm + P(M + 1 - I) / (M + 1 - I)
'        Next
'        I0 = I0 * T0ptm * 1000.0

'        T2 = T1
'        Do
'            T1 = T2
'            '120:
'            Cptm(1) = P(M + 1)
'            I1 = P(M + 1) / (M + 1)
'            For I As Integer = 1 To M
'                Cptm(1) = Cptm(1) * T1 + P(M + 1 - I)
'                I1 = I1 * T1 + P(M + 1 - I) / (M + 1 - I)
'            Next
'            I1 = I1 * T1 * 1000.0
'            T2 = T1 - (I1 - I0 - DIptm) / (1000.0 * Cptm(1))

'            'If (Abs(T2 - T1) <= 0.000001) Then GoTo 140
'            'T1 = T2
'            'ij = ij + 1
'            '	print *,'ij',ij
'            'GoTo 120
'        Loop Until Abs(T2 - T1) <= 0.000001

'        '140:
'        Cptm(4) = Cptm(1) / (Cptm(1) - Cptm(3) / 426.9)
'        TIptm = T2 * 1000.0
'        T0ptm = T0ptm * 1000.0
'        'Return
'    End Function

'    '----------              FUNCTION TI(DI,T0,X,A,C)             ------------КОНЕЦ-----



'    '----------   FUNCTION PIT(T1,T2,X,A)   Р400(4): Т1,Т2,qT ||PI=P2/P1       ---------НАЧАЛО------

'    '   ОПРЕДЕЛЕНИЕ ОТНОШЕНИЯ КОНЕЧ. ДАВЛ. Р2 В АДИАБАТ. ПРОЦ. СЖ. ИЛИ РАСШИР. К НАЧ. Р1 ПО ЗАД. ЗНАЧ.
'    '   НАЧАЛЬНОЙ Т1 И КОНЕЧНОЙ Т2 ТЕМПЕРАТУР ПРИ ИЗВЕСТНОМ альфа
'    Private Function PITptm(ByVal T1ptm As Double, ByVal T2ptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double)
'        Dim L0, I1, I2 As Double
'        Dim P(10) As Double
'        'ReDim_Aptm(6, 11)
'        L0 = 28.966 / 0.2095 * (Aptm(1, 11) / 12.01 + Aptm(2, 11) / 4.032 - Aptm(3, 11) / 32.0)
'        D = Aptm(4, 11)
'        QT = (1.0 - D) * Xptm / L0
'        M = 6
'        N = M + 1
'        L1 = M - 1
'        If (QT < 0.00001 AndAlso D < 0.00001) Then
'            '            GoTo 2
'            '2:
'            R = 29.27
'            For I As Integer = 1 To N
'                P(I) = Aptm(1, I)
'            Next
'            'GoTo 50
'        Else
'            'if(QT-0.00001) 14,15,15
'            If QT < 0.00001 Then
'                '14:
'                R = 29.27 * (1.0 + 0.60779 * D)
'                For I As Integer = 1 To N
'                    P(I) = Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D
'                Next
'                'GoTo 50
'            Else
'                '15  IF (Aptm(2,11)-0.999)10,20,20
'                If Aptm(2, 11) < 0.999 Then
'                    '10:
'                    If (Abs(Aptm(1, 11) - 0.85) < 0.001 AndAlso Abs(Aptm(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
'                        '                        GoTo 30
'                        '30:
'                        R = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = (Aptm(2, I) * QT + Aptm(1, I)) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        R = 29.27 * (1.0 + 28.966 * (Aptm(2, 11) / 4.032 + Aptm(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            AU = (44.01 * Aptm(3, I) - 32.0 * Aptm(5, I)) / 12.01
'                            BU = (9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008
'                            P(I) = ((Aptm(1, 11) * AU + Aptm(2, 11) * BU + Aptm(3, 11) * Aptm(5, I)) * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    End If
'                Else
'                    '20  IF(Xptm-1.0)70,80,80
'                    If Xptm < 1.0 Then
'                        '70:
'                        R = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        '80:
'                        R = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * (1.0 - D) / L0 + D * Aptm(4, I) * (1.0 - D) * Aptm(1, I) + (QT - (1.0 - D) / L0) * Aptm(6, I)) / (1.0 + QT)
'                        Next
'                    End If
'                End If
'            End If
'        End If

'        '50:
'        T1ptm = T1ptm * 0.001
'        T2ptm = T2ptm * 0.001
'        I1 = P(M + 1) / M
'        I2 = I1
'        For I As Integer = 1 To L1
'            I1 = I1 * T1ptm + P(M + 1 - I) / (M - I)
'            I2 = I2 * T2ptm + P(M + 1 - I) / (M - I)
'        Next
'        I1 = I1 * T1ptm
'        I2 = I2 * T2ptm
'        'PITptm=EXP((I2-I1+P(1)*ALOG(T2ptm/T1ptm))*426.9/R)
'        PITptm = Exp((I2 - I1 + P(1) * Math.Log(T2ptm / T1ptm)) * 426.9 / R)
'        T1ptm = T1ptm * 1000.0
'        T2ptm = T2ptm * 1000.0
'        'Return
'    End Function

'    '-----------------------       FUNCTION PIT(T1,T2,X,A)     --------------------КОНЕЦ-----      



'    '----------     FUNCTION TPI(T1,T2,X,A)   Р400(7): Т1,PI=P2/P1,qT || Т2       ----НАЧАЛО-----

'    '  ОПРЕДЕЛЕНИЕ КОНЕЧНОЙ Т-РЫ Т2 В АДИАБАТ. ПРОЦЕССЕ ПО ЗАДАННЫМ ЗНАЧЕНИЯМ НАЧ. Т-РЫ Т1 И ОТНОШЕНИЮ
'    '  КОНЕЧНОГО ДАВЛЕНИЯ К НАЧАЛЬНОМУ Р2/Р1 ПРИ ИЗВЕСТНОМ АЛЬФА (Х=1/АЛЬФА)
'    'TPIptm(PRptm, T1ptm, Xptm, Aptm, Cptm)

'    Private Function TPIptm(ByVal PRptm As Double, ByVal T1ptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal Cptm() As Double, ByVal dH2O As Double)
'        Dim I1, I2 As Double 'L0,
'        Dim P(10) As Double
'        'ReDim_Cptm(6)
'        'ReDim_Aptm(6, 11)
'        Cptm(6) = 28.966 / 0.2095 * (Aptm(1, 11) / 12.01 + Aptm(2, 11) / 4.032 - Aptm(3, 11) / 32.0)

'        Aptm(4, 11) = dH2O    '  V.G.
'        D = Aptm(4, 11)
'        QT = (1.0 - D) * Xptm / Cptm(6)
'        M = 6
'        N = M + 1
'        L1 = M - 1
'        Debug.Print("'PRptm,T1ptm,',PRptm,T1ptm")
'        If (QT < 0.00001 AndAlso D < 0.00001) Then
'            '            GoTo 2
'            '2:
'            Cptm(3) = 29.27
'            For I As Integer = 1 To N
'                P(I) = Aptm(1, I)
'            Next
'            'GoTo 50
'        Else
'            'if(QT-0.00001) 14,15,15
'            If QT < 0.00001 Then
'                '14:
'                Cptm(3) = 29.27 * (1.0 + 0.60779 * D)
'                For I As Integer = 1 To N
'                    P(I) = Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D
'                Next
'                'GoTo 50
'            Else
'                '15  IF (Aptm(2,11)-0.999)10,20,20
'                If Aptm(2, 11) < 0.999 Then
'                    '10:
'                    If (Abs(Aptm(1, 11) - 0.85) < 0.001 AndAlso Abs(Aptm(2, 11) - 0.15) < 0.001 AndAlso D < 0.00001) Then
'                        '                        GoTo 30
'                        '30:
'                        Cptm(3) = 29.27 * (1.0 + 1.07757 * QT) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = (Aptm(2, I) * QT + Aptm(1, I)) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        Cptm(3) = 29.27 * (1.0 + 28.966 * (Aptm(2, 11) / 4.032 + Aptm(3, 11) / 32.0) * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            AU = (44.01 * Aptm(3, I) - 32.0 * Aptm(5, I)) / 12.01
'                            BU = (9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008
'                            P(I) = ((Aptm(1, 11) * AU + Aptm(2, 11) * BU + Aptm(3, 11) * Aptm(5, I)) * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    End If
'                Else
'                    '20  IF(Xptm-1.0)70,80,80
'                    If Xptm < 1.0 Then
'                        '70:
'                        Cptm(3) = 29.27 * (1.0 + 7.18378 * QT + 0.60779 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * QT + Aptm(1, I) * (1.0 - D) + Aptm(4, I) * D) / (1.0 + QT)
'                        Next
'                        'GoTo 50
'                    Else
'                        '80:
'                        Cptm(3) = 29.27 * (0.7905 + 14.368 * QT + 0.8173 * D) / (1.0 + QT)
'                        For I As Integer = 1 To N
'                            P(I) = ((9.008 * Aptm(4, I) - 8.0 * Aptm(5, I)) / 1.008 * (1.0 - D) / Cptm(6) + D * Aptm(4, I) * (1.0 - D) * Aptm(1, I) + (QT - (1.0 - D) / Cptm(6)) * Aptm(6, I)) / (1.0 + QT)
'                        Next
'                    End If
'                End If
'            End If
'        End If
'        '50:
'        T1ptm = T1ptm * 0.001
'        T2 = T1ptm * PRptm ^ 0.25
'        Debug.Print("'T2=T1ptm*PRptm**0.25',T2,T1ptm,PRptm,(PRptm^0.25)")
'        I1 = P(M + 1) / M
'        For I As Integer = 1 To L1
'            I1 = I1 * T1ptm + P(M + 1 - I) / (M - I)
'        Next
'        I1 = I1 * T1ptm

'        T3 = T2
'        Do
'            T2 = T3
'            '120:
'            DF = P(M + 1)
'            I2 = P(M + 1) / M
'            For I As Integer = 1 To L1
'                I2 = I2 * T2 + P(M + 1 - I) / (M - I)
'                DF = DF * T2 + P(M + 1 - I)
'            Next

'            DF = DF + P(1) / T2
'            'I2 = I2 * T2 - I1 + P(1) * ALOG(T2 / T1ptm)
'            'T3 = T2 - (I2 - Cptm(3) * ALOG(PRptm) / 426.9) / DF

'            I2 = I2 * T2 - I1 + P(1) * Math.Log(T2 / T1ptm)
'            T3 = T2 - (I2 - Cptm(3) * Math.Log(PRptm) / 426.9) / DF

'            'If (Abs(T3 - T2) <= 0.000001) Then GoTo 140
'            'T2 = T3
'            'GoTo 120
'        Loop Until Abs(T3 - T2) <= 0.000001
'        '140:
'        Cptm(1) = DF * T3
'        Cptm(4) = Cptm(1) / (Cptm(1) - Cptm(3) / 426.9)
'        TPIptm = T3 * 1000.0
'        T1ptm = T1ptm * 1000.0
'        'Return
'    End Function

'    '---------------------FUNCTION TPI(T1,T2,X,A)           --------------------КОНЕЦ-----



'    '  -----------------V.G. hand made   --- аналог Р403 --


'    '  Янкин: по Р403 можно определить qTp=Q9 весовую долю топлива подогрева по замерам tвоздОСК и т-ре рабо-
'    '  чего тела в РМК1, а также при заданной величине HU*ethaKC 

'    ''''        qTp=Q9

'    '   ********      И Т А К :   с помощью п/п  Р Т М  *************
'    '                                                               *

'    '      стандарт по Янкину: T0ptm=293.15/1000. - температура отсчёта 
'    '      TTptm=343.15/1000.  Т*2 за КВД


'    Private Sub P4033(ByVal Tvh1 As Double, ByVal ttt As Double, ByVal GBphyzB1 As Double, ByVal GBphyzBsum As Double, ByVal Cptm() As Double, ByVal Tptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal qTp As Double, ByVal Tvh As Double)  ' qTp - результат (весовая доля топл. подогрева)
'        'ReDim_Aptm(6, 11)
'        'ReDim_Cptm(6)

'        T0 = 293.15 / 1000.0
'        TT = Tvh1 / 1000.0 '  343.15/1000.
'        Debug.Print("'ttt',ttt")
'        Debug.Print("'DIptm(T0,TT,Xptm,Aptm)',T0,TT,Xptm")
'        ENTH = DIptm(T0, TT, Xptm, Aptm, dH2O)   ' ЭНТАЛЬПИЯ ВОЗДУХА ЗА ОСК
'        Debug.Print(",'ENTH,T0,TT,',ENTH,T0,TT ")

'        hh = 0.1
'        qTp = hh
'        Debug.Print("'tt',tt")
'        '	read(*,*)
'        '                 И Т Е Р А Ц И Я:
'        qTp = 0

'        Do
'            Do
'                qTp = qTp + hh
'                '3340:
'                Xptm = qTp * 14.92914 ' qTp*14.92914  '14.95    ' X=1/alfa=0.0657  ' alfa=15.22
'                ' 	print *,'GTp,qTp=GTp/GBphyzBsum,X=qTp*14.92914',
'                '     *GTp,qTp,GBphyzBsum,Xptm

'                'GTp1 = qTp * GBphyzB1
'                GTp1 = qTp * GBphyzB1 / (1 + qTp)
'                aIg = (ENTH * (GBphyzB1 - GTp1) + GTp1 * 10388.0 * 0.95) / GBphyzB1      ' энтальпия продуктов сгорания 
'                '                                                             	Янк ур энергии (41) на стр 22

'                '	print*,'into P4033:aIg=(ENTH*(GBphyzB1-GTp1)+GTp1*10388.*.95)
'                '     */GBphyzB1,qTp',aIg,ENTH,GBphyzB1,GTp1,qTp   ',AI(3)
'                T0ptm = 293.15

'                '      	PRINT *, 'aIg,T0,X',aIg,T0,X
'                T0 = 293.15   '/1000.
'                'TT = Tvh1
'                '      read(*,*)

'                Tvh = TIptm(aIg, T0, Xptm, Aptm, Cptm, dH2O)  ' РТМ: опред конеч т-ры по зад. начальной т-ре , приращ. энтальпии
'                'и(X = 1 / alfa)
'                '	print *,'      ++++++  Tvh,aIg,T0,tt',Tvh,aIg,T0,tt
'                J = J + 1
'                I = I + 1
'                'If ((Tvh - ttt) > 0) Then GoTo 3399
'                'qTp = qTp + hh
'                ''	print *,'   &&&& Tvh,ttt+273.15 &&&&&&&&&',Tvh,(ttt+273.15)
'                ''	read(*,*)
'                'GoTo 3340
'            Loop Until Tvh - ttt > 0

'            '3399:
'            qTp = qTp - hh
'            'If (hh < 0.0000001) Then GoTo 3341
'            If hh < 0.0000001 Then Exit Do

'            hh = hh / 10.0   ' 2 (i=48,j=38)    10 (i=33,j=28)
'            I = I + 1
'            'GoTo 3340
'        Loop

'        '3341:   Debug.Print(" 'into P4033_VIVAT: Tvh,qTp,i,j',Tvh,qTp,i,j    '   Tvh -температура рабочего тела на входе в КВД")
'        '	read(*,*)
'    End Sub

'    ''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'    ''     аналог Р403: поиск Т*3 по заданным Т*2, Hu*etha, qT(Xptm=qTp*14.92914)
'    'Sub P403T(ByVal Tvh1 As Double, ByVal ttt As Double, ByVal GBphyzB1 As Double, ByVal GBphyzBsum As Double, ByVal Cptm() As Double, ByVal Tptm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal qTp As Double, ByVal Tvh As Double)  ' qTp - результат (весовая доля топл. подогрева)
'    '    'ReDim_Aptm(6, 11)
'    '    'ReDim_Cptm(6)

'    '    T0 = 293.15 / 1000.0
'    '    TT = Tvh1 / 1000.0 '  343.15/1000.
'    '    Xptm = 0.0
'    '    Debug.Print("'into P403T:  Tvh1,ttt,GBphyzB1,GBphyzBsum,Cptm,Tptm,Xptm,qTp,Tvh',Tvh1,ttt,GBphyzB1,GBphyzBsum,Cptm,Tptm,Xptm,qTp,Tvh")
'    '    ENTH = DIptm(T0, TT, Xptm, Aptm, dH2O)   ' ЭНТАЛЬПИЯ ВОЗДУХА на входе в КС
'    '    Debug.Print("'into P403T:ENTH,T0,TT,',ENTH,T0,TT ")

'    '    Xptm = qTp * 14.92914 ' qTp*14.92914  '14.95    ' X=1/alfa=0.0657  ' alfa=15.22
'    '    ' 	print *,'GTp,qTp=GTp/GBphyzBsum,X=qTp*14.92914',
'    '    '     *GTp,qTp,GBphyzBsum,Xptm

'    '    'GTp1 = qTp * GBphyzB1
'    '    GTp1 = qTp * GBphyzB1 / (1 + qTp)
'    '    aIg = (ENTH * (GBphyzB1 - GTp1) + GTp1 * 10388.0 * 0.95) / GBphyzB1      ' энтальпия продуктов сгорания 
'    '    '                                                             	Янк ур энергии (41) на стр 22
'    '    Debug.Print(",'aIg=(ENTH*(GBphyzB1-GTp1)+GTp1*10388.*.95)/GBphyzB1',aIg,ENTH,GBphyzB1,GTp1   ',AI(3)")
'    '    T0ptm = 293.15

'    '    '      	PRINT *, 'aIg,T0,X',aIg,T0,X
'    '    T0 = 293.15   '/1000.
'    '    'TT = Tvh1
'    '    '      read(*,*)
'    '    Xptm = qTp * 14.92914
'    '    Tvh = TIptm(aIg, T0, Xptm, Aptm, Cptm, dH2O)  ' РТМ: опред конеч т-ры по зад. начальной т-ре , приращ. энтальпии
'    '    'и(X = 1 / alfa)
'    '    '	print *,'      ++++++  Tvh,aIg,T0,tt',Tvh,aIg,T0,tt

'    '    '	print *,'   &&&& Tvh,ttt+273.15 &&&&&&&&&',Tvh,(ttt+273.15)
'    '    '	read(*,*)

'    '    Debug.Print("'into P403T_ T-PA T*3: Tvh,Tvh1,qTp',Tvh,Tvh1,qTp    '   Tvh -температура рабочего тела на входе в КВД")

'    '    '	read(*,*)

'    'End Sub



'    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$                                 
'    '                               	          *
'    '   ********      И Т О Г :     *************


'    '  -----------------V.G. hand made   ----------------


'    '       --------    РТМ   23.03.2009        --------------------КОНЕЦ----///////    

'    '////////////////////////////////////////////////////////////////



'    'CALL(bKperMF(PRIZNAK, j, ijk, t, aN1, cKper))
'    '  -------    15.05.2009    KperMF-------

'    Private Sub bKperMF(ByVal PRIZNAK As String, ByVal j As Integer, ByVal ijk As Double, ByVal t As Double, ByVal aN1 As Double, ByVal cKper(,) As Double, ByVal VZ() As Double)
'        Dim Akn1(38, 22), AkPiB(38, 22), AkGB(38, 22), Akm(38, 22), Akn2(38, 22), _
'        AkPiKS(38, 22), AkGT0(38, 22), AkT3(38, 22), AkT4(38, 22), AkR(38, 22), _
'        AkCR(38, 22), _
'        Fkn1(38, 22), FkPiB(38, 22), FkGB(38, 22), Fkm(38, 22), Fkn2(38, 22), _
'        FkPiKS(38, 22), FkGT0(38, 22), FkT3(38, 22), FkT4(38, 22), FkR(38, 22), _
'        FkCR(38, 22), FkaS(38, 22), FkPiKVD(38, 22), FkGTF(38, 22) As Double
'        ReDim_cKper(15, ijk), VZ(32)

'        Dim aName(15) As String 'CD,

'        'common/kPER/CkN1,CkPiB,CkGB,Ckm,CkN2,CkPiKS,CkGT0,CkT3,CkT4,
'        '    *CkR,CkCR,CkPiKVD,CkGTF,CkaS
'        'COMMON/KKK/Akn1,AkPiB,AkGB,Akm,Akn2,
'        '    *AkPiKS,AkGT0,AkT3,AkT4,AkR,
'        '    *AkCR,
'        '    *Fkn1,FkPiB,FkGB,Fkm,Fkn2,
'        '    *FkPiKS,FkGT0,FkT3,FkT4,FkR,
'        '    *FkCR,FkaS,FkPiKVD,FkGTF

'        'common/Kvv/CD   ',CE,CF    ' ЗНАЧЕНИЯ BINIT ПЕРЕДАЮТСЯ ИЗ п/п VVOD
'        '       PRINT *,'bKperMF:J,IJK',J,IJK
'        '	READ(*,*)
'        '      print *,'CD',CD
'        'y = 123
'        '	print *,'y',y
'        '	read(*,*)
'        '       common/k14/Kper

'        '      print *,'AkGT0',AkGT0
'        '       print *,'AkGT0(1,2),AkGT0(2,1)',AkGT0(1,2),AkGT0(2,1)

'        ' 	t=-35.5  '40. '43. '-35. '-37. '-35. '13. '15.3 '13.     '20.
'        '	aN1=80.5 '98. '93.  '90.5
'        '      print *,'j,jj',j,jj
'        '	read(*,*)
'        'If (jj > 0) Then GoTo 321 ' -обход оператора откр файла и однокр записи начиная со 2-го варианта
'        If jj = 0 Then
'            OPEN(425, "'B rez\KKperesch.xls'")
'            'stop()
'            aName(1) = "Nкт"
'            aName(2) = "kn1"
'            aName(3) = "kPiB"
'            aName(4) = "kGB"
'            aName(5) = "km"
'            aName(6) = "kn2"
'            aName(7) = "kPiKS"
'            aName(8) = "kGT0"
'            aName(9) = "kT3"
'            aName(10) = "kT4"
'            aName(11) = "kR"
'            aName(12) = "kCR"
'            aName(13) = "FkaS"
'            aName(14) = "FkPiKVD"
'            aName(15) = "FkGTF"

'            Debug.Write("(aName(j1), J1 = 1, 15)")
'            '234  format(15(G12.1,'	'))
'            jj = jj + 1
'            '		print *,'PRIZNAK,j,ijk,t,aN1,cKper,VZ',
'            '     *PRIZNAK,j,ijk,t,aN1
'            '	 read(*,*)
'        End If

'        '321:
'        'If (CD = "F") Then GoTo 159
'        If CD = "F" Then
'            '159:
'            Debug.Print("'++++ F O R S ++++'")
'            '       read(*,*)
'            cKper(1, j) = VZ(2)
'            Call PERESCH(Fkn1, t, aN1, V)
'            cKper(2, j) = V
'            '	  print *,'Fkn1' 
'            Call PERESCH(FkPiB, t, aN1, V)
'            cKper(3, j) = V
'            '	  	  print *,'FkPiB' 
'            Call PERESCH(FkGB, t, aN1, V)
'            cKper(4, j) = V
'            '		  print *,'FkGB'
'            Call PERESCH(Fkm, t, aN1, V)
'            cKper(5, j) = V
'            '	      print *,'Fkm'
'            Call PERESCH(Fkn2, t, aN1, V)
'            cKper(6, j) = V
'            '            print *,'Fkn2'
'            Call PERESCH(FkPiKS, t, aN1, V)
'            cKper(7, j) = V
'            '             print *,'FkPiKS'
'            Call PERESCH(FkGT0, t, aN1, V)
'            cKper(8, j) = V
'            '		      print *,'FkGT0'
'            '	read(*,*)
'            Call PERESCH(FkT3, t, aN1, V)
'            cKper(9, j) = V
'            '	    print *,'FkT3' 
'            Call PERESCH(FkT4, t, aN1, V)
'            cKper(10, j) = V
'            '	      print *,'FkT4'
'            Call PERESCH(FkR, t, aN1, V)
'            cKper(11, j) = V
'            '	         print *,'FkR'
'            Call PERESCH(FkCR, t, aN1, V)
'            cKper(12, j) = V
'            '	            print *,'FkCR'
'            Call PERESCH(FkaS, t, aN1, V)
'            cKper(13, j) = V
'            '	            print *,'FkaS'
'            Call PERESCH(FkPiKVD, t, aN1, V)
'            cKper(14, j) = V
'            '	            print *,'FkPiKVD'
'            Call PERESCH(FkGTF, t, aN1, V)
'            cKper(15, j) = V
'            Debug.Print("'FkGTF'")
'            '        print *,'cKper',cKper
'        Else
'            Debug.Print("'==== M A X I M A L ===='")
'            '	read(*,*)
'            cKper(1, j) = VZ(2)
'            '	print *,'cKper(1,j)=VZ(2)',j,cKper(1,j),VZ(2)
'            Call PERESCH(Akn1, t, aN1, V)
'            cKper(2, j) = V
'            '        print *,'Akn1'
'            '	  print *,'cKper(1,j)=V',ijk,cKper(1,j),V 
'            Call PERESCH(AkPiB, t, aN1, V)
'            cKper(3, j) = V
'            '	  	  print *,'AkPiB' 
'            Call PERESCH(AkGB, t, aN1, V)
'            cKper(4, j) = V
'            '		  print *,'AkGB'
'            Call PERESCH(Akm, t, aN1, V)
'            cKper(5, j) = V
'            '	      print *,'Akm'
'            Call PERESCH(Akn2, t, aN1, V)
'            cKper(6, j) = V
'            '            print *,'Akn2'
'            '	 read(*,*) 
'            Call PERESCH(AkPiKS, t, aN1, V)
'            cKper(7, j) = V
'            '            print *,'AkPiKS'
'            Call PERESCH(AkGT0, t, aN1, V)
'            cKper(8, j) = V
'            '	            print *,'AkGT0'
'            Call PERESCH(AkT3, t, aN1, V)
'            cKper(9, j) = V
'            '	              print *,'AkT3' 
'            Call PERESCH(AkT4, t, aN1, V)
'            cKper(10, j) = V
'            '	            print *,'AkT4'
'            Call PERESCH(AkR, t, aN1, V)
'            cKper(11, j) = V
'            '	            print *,'AkR'
'            Call PERESCH(AkCR, t, aN1, V)
'            cKper(12, j) = V
'            '	      print *,'AkCR'
'            cKper(13, j) = 0
'            cKper(14, j) = 0
'            cKper(15, j) = 0

'            '	print *,'j,cKper',j,(cKper(i,j),i=1,15)
'            Debug.Print("'--------------M-(cKper(i,j),i=1,15)--------------'")
'            '	read(*,*)
'            '	print *,'-+- cKper',(cKper(i,100),i=1,14)
'            'GoTo 333

'            '      print *,'cKper',cKper
'            '       read(*,*) 
'        End If

'        '333:    'Continue 
'        Debug.Print("'---------------F (cKper(i,j),i=1,15)--------------'")
'        '	read(*,*)
'    End Sub

'    '******************************************************

'    Private Sub PERESCH(ByVal ABC(,) As Double, ByVal t As Double, ByVal aN1 As Double, ByVal V As Double)
'        Dim X(21), Y(21), F(3) As Double
'        ReDim_ABC(38, 22)
'        Dim aLeft, a1, b1, a2, b2, aakGT0 As Double

'        J = 0
'        '       t=-35. '13. '15.3 '13.     '20.
'        '      print *,'AkGT0',AkGT0
'        '       print *,'AkGT0(1,2),AkGT0(2,1)',AkGT0(1,2),AkGT0(2,1)
'        For I As Integer = 1 To 18
'            '	if((t+AkGT0(2*i,1)) > 0.)
'            If (t >= ABC(2 * I, 1)) AndAlso (t <= ABC(2 * (I + 1), 1)) Then 'попал в диапазон

'                '                GoTo 2
'                '2:
'                aLeft = ABC(2 * I, 1) 'запомнить левую температуру
'                J = J + 1  ' индекс "целостности" параметра t и попадание в Царьковский узел при j=2
'                ii = I
'            End If
'            'GoTo 3
'            '  2 	print *,'i,ABC(2*i,1),t',i,ABC(2*i,1),t
'            '3:          'Continue Do
'        Next

'        '	if  j=1 -дробная температура 
'        '    решение ищется между кривыми с номерами ii и ii+1

'        '	if j=2 ' целочисленное значение температуры
'        '    решение ищется на кривой с номером ii 

'        '       print *,'j,ii',j,ii
'        '      aN1=80.5 '90.5

'        'If (J = 2.0) Then GoTo 100
'        If J = 2 Then
'            '100:
'            For I As Integer = 1 To 21       ' целочисленное значение температуры
'                X(22 - I) = ABC(2 * ii - 1, I + 1)
'                Y(22 - I) = ABC(2 * ii, I + 1)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            F(1) = aN1
'            Call SPLINE1(21, X, Y, F) ',&23)
'            V = F(2)
'        Else
'            For I As Integer = 1 To 21
'                X(22 - I) = ABC(2 * ii - 1, I + 1)   ' аргументы только в порядке возрастания
'                Y(22 - I) = ABC(2 * ii, I + 1)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            F(1) = aN1
'            Call SPLINE1(21, X, Y, F) ',&22)
'            '	print *,'F',F
'            a1 = F(1)
'            b1 = F(2)
'            '      print *,'a1,b1',a1,b1
'            For I As Integer = 1 To 21
'                X(22 - I) = ABC(2 * ii + 1, I + 1)    ' аргументы только в порядке возрастания
'                Y(22 - I) = ABC(2 * ii + 2, I + 1)
'                'a3(i) = X(22 - i)
'                'b3(i) = Y(22 - i)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            F(1) = aN1
'            Call SPLINE1(21, X, Y, F) ', 22)
'            '	print *,'F',F
'            a2 = F(1)
'            b2 = F(2)
'            '      print *,'a2,b2',a2,b2
'            aakGT0 = b1 + (t - aLeft) * (b2 - b1) / 5.0
'            '	print *,'result k_perescheta',aakGT0
'            V = aakGT0
'        End If

'        'GoTo 200

'        '	print *,'+++++++ABC+++++++res k_peresch',F
'        '        Return
'        '23:     Debug.Print("'spline1 &&&'")
'        '22:     Debug.Print(",'spline1 ???'")
'        '200:    Debug.Print("----------ABC---------end of PROBLEM'")
'    End Sub

'    '  -------    15.05.2009    KperMF-------
'    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

'    'FFFFFFF(II)
'    'F(II)
'    'F(II)
'    'F(II)
'    'FFFFFF(II)
'    'F(II)
'    'F(II)
'    'F(II)
'    'F(II)




'    '              26.05.2009    KperMFfi
'    '      CALL cKperMF(PRIZNAK,j,ijk,t,aN1,cKper) поправочный коэф-т пересчёта, учитывающий влажность воздуха 
'    'call(cKperMFfi(PRIZNAK, j, ijk, t, fi, cKperfi, VZ))
'    '  -------    15.05.2009    cKperMF-------

'    Private Sub cKperMFfi(ByVal PRIZNAK As String, ByVal j As Integer, ByVal ijk As Double, ByVal t As Double, ByVal fi As Double, ByVal cKperfi(,) As Double, ByVal VZ() As Double)
'        Dim Akn1fi(12, 20), AkPiBfi(12, 20), AkGBfi(12, 20), Akmfi(12, 20), _
'        Akn2fi(12, 20), _
'        AkPiKSfi(12, 20), AkGT0fi(12, 20), AkT3fi(12, 20), AkT4fi(12, 20), _
'        AkRfi(12, 20), _
'        AkCRfi(12, 20), _
'        Fkn1fi(12, 20), FkPiBfi(12, 20), FkGBfi(12, 20), Fkmfi(12, 20), _
'        Fkn2fi(12, 20), _
'        FkPiKSfi(12, 20), FkGT0fi(12, 20), FkT3fi(12, 20), FkT4fi(12, 20), _
'        FkRfi(12, 20), _
'        FkCRfi(12, 20), FkaSfi(12, 20), FkPiKVDfi(12, 20), FkGTFfi(12, 20) As Double
'        ReDim_cKperfi(15, ijk), VZ(32)

'        Dim aNamefi(15) As String 'CD,

'        'common/kPER/CkN1,CkPiB,CkGB,Ckm,CkN2,CkPiKS,CkGT0,CkT3,CkT4,
'        '    *CkR,CkCR,CkPiKVD,CkGTF,CkaS
'        '     COMMON/KKK1/Akn1fi,AkPiBfi,AkGBfi,Akmfi,Akn2fi,
'        '    *AkPiKSfi,AkGT0fi,AkT3fi,AkT4fi,AkRfi,
'        '    *AkCRfi,
'        '    *Fkn1fi,FkPiBfi,FkGBfi,Fkmfi,Fkn2fi,
'        '    *FkPiKSfi,FkGT0fi,FkT3fi,FkT4fi,FkRfi,
'        '    *FkCRfi,FkPiKVDfi,FkGTFfi

'        'common/Kvv/CD   ',CE,CF    ' ЗНАЧЕНИЯ BINIT ПЕРЕДАЮТСЯ ИЗ п/п VVOD
'        '	PRINT *,'cKperMFfi:J,IJK',J,IJK
'        '	READ(*,*)
'        '       common/k14/Kper

'        '      print *,'AkGT0',AkGT0
'        '       print *,'AkGT0(1,2),AkGT0(2,1)',AkGT0(1,2),AkGT0(2,1)

'        ' 	t=-35.5  '40. '43. '-35. '-37. '-35. '13. '15.3 '13.     '20.
'        '	aN1=80.5 '98. '93.  '90.5
'        If (jjj = 0) Then
'            'GoTo 321 ' -обход оператора откр файла и однокр записи начиная со 2-го варианта
'            'OPEN(525,FILE='B rez\KKpereschfi.xls')

'            aNamefi(1) = "Nкт"
'            aNamefi(2) = "kn1fi"
'            aNamefi(3) = "kPiBfi"
'            aNamefi(4) = "kGBfi"
'            aNamefi(5) = "kmfi"
'            aNamefi(6) = "kn2fi"
'            aNamefi(7) = "kPiKSfi"
'            aNamefi(8) = "kGT0fi"
'            aNamefi(9) = "kT3fi"
'            aNamefi(10) = "kT4fi"
'            aNamefi(11) = "kRfi"
'            aNamefi(12) = "kCRfi"
'            aNamefi(13) = "FkaSfi"
'            aNamefi(14) = "FkPiKVDfi"
'            aNamefi(15) = "FkGTFfi"

'            Debug.Write("(aNamefi(j4), J4 = 1, 15)")
'            '234  format(15(G12.1,'	'))
'            jjj = jjj + 1
'        End If


'        If (CD = "F") Then
'            Debug.Print("'++++ F O R S  WET AIR ++++'")
'            '      read(*,*)
'            cKperfi(1, j) = VZ(2)
'            Call PERESCHfi(Fkn1fi, t, fi, V)
'            cKperfi(2, j) = V
'            '	  print *,'Fkn1fi' 
'            Call PERESCHfi(FkPiBfi, t, fi, V)
'            cKperfi(3, j) = V
'            '	  	  print *,'FkPiBfi' 
'            Call PERESCHfi(FkGBfi, t, fi, V)
'            cKperfi(4, j) = V
'            '		  print *,'FkGBfi'
'            Call PERESCHfi(Fkmfi, t, fi, V)
'            cKperfi(5, j) = V
'            '	      print *,'Fkmfi'
'            Call PERESCHfi(Fkn2fi, t, fi, V)
'            cKperfi(6, j) = V
'            '            print *,'Fkn2fi'
'            Call PERESCHfi(FkPiKSfi, t, fi, V)
'            cKperfi(7, j) = V
'            '             print *,'FkPiKSfi'
'            Call PERESCHfi(FkGT0fi, t, fi, V)
'            cKperfi(8, j) = V
'            '		      print *,'FkGT0fi'
'            '	read(*,*)

'            Call PERESCHfi(FkT3fi, t, fi, V)
'            cKperfi(9, j) = V
'            '	    print *,'FkT3fi' 
'            Call PERESCHfi(FkT4fi, t, fi, V)
'            cKperfi(10, j) = V
'            '	      print *,'FkT4fi'
'            Call PERESCHfi(FkRfi, t, fi, V)
'            cKperfi(11, j) = V
'            '	         print *,'FkRfi'
'            Call PERESCHfi(FkCRfi, t, fi, V)
'            cKperfi(12, j) = V
'            Debug.Print("'FkCRfi'")
'            'call(PERESCHfi(FkaSfi, t, fi, V))
'            'cKperfi(13, j) = V
'            '	            print *,'FkaSfi'
'            Call PERESCHfi(FkPiKVDfi, t, fi, V)
'            cKperfi(14, j) = V
'            '	            print *,'FkPiKVDfi'
'            Call PERESCHfi(FkGTFfi, t, fi, V)
'            cKperfi(15, j) = V
'            '	            print *,'FkGTFfi'
'            '        print *,'cKper',cKper
'        Else
'            Debug.Print("'==== M A X I M A L WET AIR===='")
'            '	read(*,*)
'            cKperfi(1, j) = VZ(2)
'            '	print *,'j,ijk,fi,cKperfi(1,j)=VZ(2)',j,ijk,fi,cKperfi(1,j),VZ(2)

'            'call(PERESCHfi(ABCD, t, fi, V))
'            '      print *,'--//ijk,Akn1fi',ijk,fi
'            '	read(*,*)
'            Call PERESCHfi(Akn1fi, t, fi, V)
'            cKperfi(2, j) = V

'            '	print *,'---\\\\ijk,Akn1fi',ijk,fi
'            '	read(*,*)

'            '	  print *,'cKperfi(1,j)=V',ijk,cKperfi(1,j),V 
'            Call PERESCHfi(AkPiBfi, t, fi, V)
'            cKperfi(3, j) = V
'            '	  	  print *,'++++/////ijk,AkPiBfi',ijk,AkPiBfi
'            '	PRINT *,'AVARIA'
'            '	READ(*,*)
'            Call PERESCHfi(AkGBfi, t, fi, V)
'            '	print *,''''''/////ijk,AkGBfi',ijk,AkGBfi
'            cKperfi(4, j) = V
'            '		  print *,'AkGBfi'
'            Call PERESCHfi(Akmfi, t, fi, V)
'            cKperfi(5, j) = V
'            '	      print *,'Akmfi'
'            Call PERESCHfi(Akn2fi, t, fi, V)
'            cKperfi(6, j) = V
'            '            print *,'Akn2fi'
'            '	 read(*,*) 
'            Call PERESCHfi(AkPiKSfi, t, fi, V)
'            cKperfi(7, j) = V
'            '            print *,'AkPiKSfi'
'            Call PERESCHfi(AkGT0fi, t, fi, V)
'            cKperfi(8, j) = V
'            '            print *,'AkGT0fi'
'            Call PERESCHfi(AkT3fi, t, fi, V)
'            cKperfi(9, j) = V
'            '	              print *,'AkT3' 
'            Call PERESCHfi(AkT4fi, t, fi, V)
'            cKperfi(10, j) = V
'            '	            print *,'-*-AkT4fi'
'            '       print *,'AkT4fi',AkT4fi

'            Call PERESCHfi(AkRfi, t, fi, V)
'            cKperfi(11, j) = V
'            '	            print *,'AkRfi'
'            Call PERESCHfi(AkCRfi, t, fi, V)
'            cKperfi(12, j) = V
'            '	            print *,'AkCRfi'
'            '	      print *,'AkCRfi'
'            cKperfi(13, j) = 0
'            cKperfi(14, j) = 0
'            cKperfi(15, j) = 0

'            '	print *,'j,cKperfi',j,(cKperfi(i,j),i=1,15)
'            '	print *,'---------------M-(cKperfi(i,j),i=1,15)--------------'
'            '	read(*,*)
'            '	print *,'-+- cKper',(cKper(i,100),i=1,14)
'        End If

'        '      print *,'cKper',cKper
'        '       read(*,*) 


'        Debug.Print("'---------------F-(cKperfi(i,j),i=1,15)--------------'")
'        '	read(*,*)
'    End Sub

'    '******************************************************

'    Private Sub PERESCHfi(ByVal ABCD(,) As Double, ByVal t As Double, ByVal fi As Double, ByVal V As Double)  ' fi - относит. влажность воздуха
'        Dim X(19), Y(19), f(3) As Double
'        'ReDim_ABCD(12, 20)
'        Dim aLeft, a1, b1, a2, b2, aakGT0 As Double

'        Debug.Print("'into PERESCHfi: ijk',ijk")
'        J = 0
'        '       fi=0. 20. 40. 60. 80. 100.t=-35. '13. '15.3 '13.     '20.
'        '      print *,'AkGT0',AkGT0
'        '       print *,'AkGT0(1,2),AkGT0(2,1)',AkGT0(1,2),AkGT0(2,1)
'        Debug.Print("'t,fi',t,fi")
'        For I As Integer = 1 To 5
'            '	if((t+AkGT0(2*i,1)) > 0.)
'            If (fi >= ABCD(2 * I, 1)) AndAlso (fi <= ABCD(2 * (I + 1), 1)) Then
'                '                GoTo 2
'                '2:
'                Debug.Print("'i,ABCD(2*i,1),t',i,ABCD(2*i,1),t")
'                aLeft = ABCD(2 * I, 1)
'                J = J + 1  ' индекс "целочисленности" параметра fi и попадание в Царьковский узел при j=2
'                ii = I
'            End If
'            '            GoTo 3
'            '3:          'Continue Do
'        Next
'        '	if  j=1 -дробная температура 
'        '    решение ищется между кривыми с номерами ii и ii+1

'        '	if j=2 ' целочисленное значение температуры
'        '    решение ищется на кривой с номером ii 

'        Debug.Print("'j,ii',j,ii")
'        '      aN1=80.5 '90.5

'        If (J = 2.0) Then
'            '            GoTo 100
'            '100:
'            For I As Integer = 1 To 19       ' целочисленное значение температуры
'                X(I) = ABCD(2 * ii - 1, I + 1)
'                Y(I) = ABCD(2 * ii, I + 1)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            f(1) = t
'            Call SPLINE1(19, X, Y, f) ',&23)
'            V = f(2)
'        Else


'            For I As Integer = 1 To 19
'                X(I) = ABCD(2 * ii - 1, I + 1)   ' аргументы только в порядке возрастания
'                Y(I) = ABCD(2 * ii, I + 1)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            f(1) = t
'            Call SPLINE1(19, X, Y, f) ',&22)
'            Debug.Print("'F',F")
'            a1 = f(1)
'            b1 = f(2)
'            '      print *,'a1,b1',a1,b1
'            For I As Integer = 1 To 19
'                X(I) = ABCD(2 * ii + 1, I + 1)    ' аргументы только в порядке возрастания
'                Y(I) = ABCD(2 * ii + 2, I + 1)
'                'a3(i) = X(22 - i)
'                'b3(i) = Y(22 - i)
'            Next
'            '	print *,'X',X
'            '		print *,'y',y
'            f(1) = t
'            Call SPLINE1(19, X, Y, f) ',&22)
'            '	print *,'F',F
'            a2 = f(1)
'            b2 = f(2)
'            '      print *,'a2,b2',a2,b2
'            aakGT0 = b1 + (fi - aLeft) * (b2 - b1) / 20.0  '5.
'            '	print *,'result k_perescheta',aakGT0
'            V = aakGT0
'        End If

'        '-------------------------------

'        '  100 do i=1,21       ' целочисленное значение температуры
'        'X(22 - i) = ABC(2 * ii - 1, i + 1)
'        'Y(22 - i) = ABC(2 * ii, i + 1)
'        '      end do 
'        '	print *,'X',X
'        '		print *,'y',y

'        '-------------------------------


'        '        Debug.Print("'+++++++ABCD++++++++res k_pereschfi',F")

'        '        Return
'        '23:     Debug.Print("'spline1fi &&&'")
'        '22:     Debug.Print("'spline1fi ???'")
'        '200:    Debug.Print("'--------ABCD-------end of PROBLEM res k_pereschfi'")
'    End Sub

'    '  -------    15.05.2009    KperMF-------



'    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    '*******************************************************






'    '      '*****************  SPLINE1   К У Б И Ч Н Ы Й ***************  
'    '  -------    B.Я.Калешин  -------
'    '               В Н И М А Н И Е:
'    '----------13 Мая 2004   значения аргумента Х д.б. расположены в порядке возрастания
'    '-----------3 Февраля 2005 значения аргумента X(I+1),X(I) не должны совпадать
'    Private Sub SPLINE1(ByVal N As Integer, ByVal X() As Double, ByVal Y() As Double, ByVal F() As Double)
'        Dim P As Integer
'        'dimension(F(3), Z(100), X(N), Y(N))
'        ReDim_F(4), X(N), Y(N)
'        Dim Z(100) As Double
'        Dim NK, T, C, H, NI, E As Integer
'        Dim A As Double

'        'if(50-N)15,16,16
'        '15:     Debug.Print("'   РАЗМЕРНОСТИ  МАСС.Х,Y  УМЕНЬШИТЬ ДО 50'")
'        '250  format(/'   РАЗМЕРНОСТИ  МАСС.Х,Y  УМЕНЬШИТЬ ДО 50')

'        If N > 50 Then
'            Stop
'            Exit Sub
'        End If

'        '16	if((F(1)-X(1))*(F(1)-X(N)))5,5,6

'        If (F(1) - X(1)) * (F(1) - X(N)) > 0 Then
'            '6:          Debug.Print("1H ,' SPLINE1: ARGUMENT ISren BEYOND OF LIMITS SEGMENT'/3H X=,E15.6,5H  XN=,E15.6,5H  XK=,E15.6,5H  YN=,E15.6,5H  YK=,E15.6")
'            '            '  9  format(1H ,43HSPLINE. АРГУМЕНТ ВЫХОДИТ ЗА ПРЕДЕЛЫ ОТРЕЗКА/
'            '            '9  format(1H ,' SPLINE1: ARGUMENT ISren BEYOND OF LIMITS SEGMENT'/3H X=,E15.6,5H  XN=,E15.6,5H  XK=,E15.6,5H  YN=,E15.6,5H  YK=,E15.6) 
'            '            Debug.Write("1H ,' SPLINE1: ARGUMENT ISren BEYOND OF LIMITS SEGMENT'/3H X=,E15.6,5H  XN=,E15.6,5H  XK=,E15.6,5H  YN=,E15.6,5H  YK=,E15.6 F(1),X(1),X(N),Y(1),Y(N)")
'            Stop
'            Exit Sub     '-это условие * ''' 10.11.2004
'        Else
'            '5:
'            Z(1) = 0
'            Z(N) = 0
'            Z(N + 1) = 0
'            Z(2 * N) = 0
'            NK = N - 1

'            For I As Integer = 2 To NK
'                H = X(I + 1) - X(I)
'                'if(H) 30,31,30
'                If (H) = 0 Then
'                    '31:
'                    Debug.Print("'SPLINE1(N,X,Y,F,*):the X(i+1) X(i)-points couldn not be agree',X(I+1),X(I)  ' СОВПАДЕНИЕ ТОЧЕК X(I+1),X(I)")
'                    Stop
'                Else
'                    '30:
'                    T = Y(I + 1) - Y(I)
'                    C = H / 6.0
'                    D = T / H
'                    H = X(I) - X(I - 1)
'                    T = Y(I) - Y(I - 1)
'                    A = H / 6.0
'                    B = 2 * (A + C)
'                    D = D - T / H
'                    E = A * Z(I - 1) + B
'                    Z(I) = -C / E
'                    NI = N + I
'                    Z(NI) = (D - Z(NI - 1) * A) / E
'                End If
'            Next

'            For ii As Integer = 1 To NK
'                I = N - ii
'                NI = N + I
'                Z(I) = Z(I) * Z(I + 1) + Z(NI)
'            Next

'            For I As Integer = 2 To N
'                'if(F(1)-X(I))10,10,3
'                '10:         P = I
'                'GoTo 11
'                If (F(1) - X(I)) <= 0 Then
'                    P = I
'                    Exit For
'                End If
'            Next

'            '11:
'            H = X(P) - X(P - 1)
'            T = Y(P) - Y(P - 1)
'            A = X(P) * (Y(P - 1) - Z(P - 1) * H * H / 6.0)
'            B = X(P - 1) * (Y(P) - Z(P) * H * H / 6.0)
'            C = T / H - H * (Z(P) - Z(P - 1)) / 6.0
'            D = (A - B) / H
'            E = F(1) - X(P - 1)
'            T = X(P) - F(1)
'            A = Z(P) * E
'            B = Z(P - 1) * T
'            ' ----------------  V.G.   9.02.2004 ---------
'            '        F(4)- ВТОРАЯ  производная ф-ии в  точке F(1):
'            F(4) = (A + B) / H
'            A = A * E
'            B = B * T
'            F(3) = (A - B) / (2.0 * H) + C
'            A = A * E
'            B = B * T
'            F(2) = (A + B) / (6.0 * H) + C * F(1) + D
'            'Return
'        End If
'    End Sub
'    '****************************   SPLINE   ************************




'    '*********************   SPLINE  Л И Н Е Й Н Ы Й     V G    *************
'    '               В Н И М А Н И Е:
'    '----------7 Dec 2004   значения аргумента Х д.б. расположены в порядке возрастания
'    '-----------3 Февраля 2005 значения аргумента X(I+1),X(I) не должны совпадать
'    Private Sub SPLINEЛинейный(ByVal N As Integer, ByVal X() As Double, ByVal Y() As Double, ByVal F() As Double, ByVal Obg As Double)
'        Dim P, H As Integer
'        'dimension(F(3), Z(100), X(N), Y(N))
'        'Dim F(2), Z(100), X(N), Y(N) As Double

'        '	if(50-N)15,16,16
'        '15:         Print(250)
'        ' 250  format(/' V G   РАЗМЕРНОСТИ  МАСС.Х,Y  УМЕНЬШИТЬ ДО 50')
'        '            Return 1
'        If (50 - N) < 0 Then
'            Debug.Print("' V G   РАЗМЕРНОСТИ  МАСС.Х,Y  УМЕНЬШИТЬ ДО 50'")
'            Exit Sub
'        End If


'        '16	if((F(1)-X(1))*(F(1)-X(N)))5,5,6
'        If (F(1) - X(1)) * (F(1) - X(N)) > 0 Then
'            Debug.Print("43HSPLINE. АРГУМЕНТ ВЫХОДИТ ЗА ПРЕДЕЛЫ ОТРЕЗКАX=,E15.6,5H  XN=,E15.6,5H  XK=,E15.6,5H  YN=,E15.6,5H  YK=,E15.6' F(1),X(1),X(N),Y(1),Y(N)")
'            Exit Sub
'        Else

'            '6:      Print(9, F(1), X(1), X(N), Y(1), Y(N))
'            '        '  9  format(1H ,43HSPLINE. АРГУМЕНТ ВЫХОДИТ ЗА ПРЕДЕЛЫ ОТРЕЗКА/
'            '   9  format(1H ,' SPLINE V G: ARGUMENT ISren BEYOND OF LIMITS SEGMENT'/3H X=,E15.6,5H  XN=,E15.6,5H  XK=,E15.6,5H  YN=,E15.6,5H  YK=,E15.6) 
'            '	WRITE (15,9) F(1),X(1),X(N),Y(1),Y(N)
'            '        Return 1      '-это условие * ''' 10.11.2004

'            '5:
'            For I As Integer = 2 To N
'                'if(F(1)-X(I))10,20,3 ' - определение интервала, которому принадлежит F(1)
'                If F(1) < X(I) Then
'                    P = I
'                    Exit For
'                ElseIf F(1) = X(I) Then
'                    P = I
'                    If (I = N) Then
'                        F(2) = Y(N)
'                        Exit Sub
'                    End If
'                    Exit For
'                End If
'            Next

'            '  11  PRINT *,'X,F(1)',X,f(1)
'            H = X(P) - X(P - 1)
'            'if(H) 30,31,30
'            If (H) = 0 Then
'                Debug.Print("'SPLINE(N,X,Y,F,*):the X(i+1) X(i)-points couldn not be agree',X(I+1),X(I)")
'                Stop
'            Else
'                F(2) = Y(P - 1) + (F(1) - X(P - 1)) * (Y(P) - Y(P - 1)) / (X(P) - X(P - 1))
'            End If

'            '30:     Continue Do
'            '        F(2) = Y(P - 1) + (F(1) - X(P - 1)) * (Y(P) - Y(P - 1)) / (X(P) - X(P - 1))
'            '12:     Continue Do
'            '        Return
'        End If

'    End Sub
'    '****************************   SPLINE   ************************






'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET
'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET

'    '       Приведение к САУ основных параметров изделия на Др режимах
'    '      j1 - номер варианта
'    Private Sub PRIVEDENIE(ByVal j1 As Integer, ByVal IJK As Integer, ByVal PHYZ(,) As Double, ByVal PRIV(,) As Double)   'ByVal iVAR As Integer, Tb,GBphyzB,ppB,Rejim,
'        Dim Index As Integer  ',Rejim
'        Dim n1pr, n2pr, N1f, mpr, n2prKVD As Double ', kKND, kKVD, kKs, njuOHL,N2f, KP2, mfors, mfizmem, Kv,, ISren
'        Dim AKR(7, 10), AKGT(7, 10), AKTG(7, 10), VZ(32) As Double
'        ReDim_PRIV(40, IJK), PHYZ(45, IJK)

'        Dim ALet(40), B As String
'        Dim _as, F, T3prTDR, Rpr, GTpr, CRpr, Gvspr, alfaspr, PiKNDpr, PiKspr, T3pr, T4pr, PiTpr, U447pr, P6bpr, P6pr, T6bpr As Double
'        Dim P2bpr, P2pr, P4bpr, P4pr, T4bpr, T2bpr, GprKVD, wKND, wKVD As Double

'        ' common/Kvv/CD
'        ' common/koeffPER/n1cay,
'        '*PiBcay,GBcay,mcay,
'        '*Pikcay,Rcay,CRcay,alfaScay,n2cay,PiTcay,T3cay,T4cay
'        'common/Kpr/ABKR,ABKGT,ABKTG
'        'If (ik > 0) Then GoTo 632 ' - ключ обхода участка открытие файла 7 и запись строки идентификаторов
'        If (ik = 0) Then
'            ''------------------------------------------------------
'            'OPEN(57,FILE='B rez\PRIVEDENIE.xls')
'            '------------------------------------------------------
'            B = "Nкт"
'            '	print *,B    'выводится на экран вместо кириллицы -абракадабра; оператор  write выводит в текстовый файл кириллицу правильно 
'            '++++++++++++++++++++

'            '             Р У С И Ф И К А Ц И Я     И Д Е Н Т И Ф И К А Т О Р О В  (15 ФЕВ 2005Г)
'            ALet(1) = " Nкт"
'            ALet(2) = " Режим"
'            ALet(3) = " n1пр"
'            ALet(4) = " n2пр"
'            ALet(5) = " а1"
'            ALet(6) = " а2"
'            ALet(7) = " Rпр"
'            ALet(8) = " CRпр"
'            ALet(9) = " Gтпр"
'            ALet(10) = " Gвsпр"
'            ALet(11) = " альфаsпр"
'            ALet(12) = " mпр"
'            ALet(13) = " PiКНДпр"
'            ALet(14) = " wКНДпр"
'            ALet(15) = " etaКНДпр"
'            ALet(16) = " PiКВДпр"
'            ALet(17) = " wКВДпр"
'            ALet(18) = " etaКВДпр"
'            ALet(19) = " PiКsпр"
'            ALet(20) = " T3пр"
'            ALet(21) = " T3прТДР"
'            ALet(22) = " T4пр"
'            ALet(23) = " PiTпр"
'            ALet(24) = " u447пр"
'            ALet(25) = " P*6пр"
'            ALet(26) = " P*6прСТ"
'            ALet(27) = " P6пр"
'            ALet(28) = " T*6пр"
'            ALet(29) = " P*2пр"
'            ALet(30) = " P2пр"
'            ALet(31) = " T*2пр"
'            ALet(32) = " P*4пр"
'            ALet(33) = " P4пр"
'            ALet(34) = " T*4пр"
'            ALet(35) = " O"
'            ALet(36) = " n2прКВД"
'            ALet(37) = " GпрКВД"
'            ALet(38) = " O"
'            ALet(39) = " 1.033\Pb"
'            ALet(40) = "SQRT(288.15\Tb)"

'            'write(7, 177)(Alet(j), j = 1, 40)
'            ' 177  format(39G12.2,G15.2)
'            '-------------------------------------------------------
'            Debug.Write("ALet(J), J = 1, 40")
'            '277  format(39(G12.2,'	'),G15.2,'	')

'            '-------------------------------------------------------
'            ik = ik + 1   ' - ключ обхода участка открытие файла 7 и запись строки идентификаторов
'        End If

'        '632:    'Continue 
'        'u = 288.15 / Tb
'        '       print *,'PRIVEDENIE'
'        '	print *,'j1,IJK,VZ',j1,IJK,VZ
'        GBPHYZI = PHYZ(32, j1)
'        U = 288.15 / PHYZ(9, j1)
'        '	print *,'PHYZ(9,j1),j1',PHYZ(9,j1),j1
'        _as = Sqrt(U)
'        '     n1pr=VZ(9)*as   
'        '	n2pr=VZ(10)*_as 
'        n1pr = PHYZ(7, j1) * _as
'        n2pr = PHYZ(8, j1) * _as
'        '	print *,'n1pr,VZ(9),as,n2pr,VZ(10)',n1pr,VZ(9),as,n2pr,VZ(10)       

'        '****************************************	

'        '      CALL ABCD(Index,N1f,n1pr,GBphyzB,Rfiz,IJK,PiKND,PiKsum,Tb,
'        '     *Rejim,alRUD,ppB)
'        N1f = PHYZ(7, j1)
'        'Rejim = VZ(3)
'        Rejim = PHYZ(2, j1)
'        '      CALL ABCD(Index,N1f,n1pr,PHYZ(33,IJK),PHYZ(34,IJK),IJK,
'        '     *PHYZ(15,IJK),PHYZ(24,IJK),
'        '     *PHYZ(9,IJK),Rejim,VZ(4),PHYZ(10,IJK))

'        'CALL(ABCD(Index, j1, ijk, VZ, PHYZ))
'        ''      CALL ABCD(iVAR,Index,j1,ijk,PHYZ) 
'        '****************************************	
'        Index = Index + 1
'        F = 1.033227 / PHYZ(10, j1)
'        If PHYZ(27, j1) <= 0.00001 Then
'            T3prTDR = 0
'        Else
'            T3prTDR = PHYZ(27, j1) * U '*ABKTG
'        End If

'        Rpr = PHYZ(34, j1) * F '*ABKR     ' ABKR - это KR  см. рис. 2
'        GTpr = PHYZ(35, j1) * F * _as  'ABKGT  ' ABKGT - это KGT  см. рис. 3
'        Debug.Print("'GTpr=PHYZ(35,j1)*f*as',GTpr,PHYZ(35,j1),f,as")
'        '	read(*,*)
'        CRpr = GTpr / Rpr
'        Gvspr = PHYZ(33, j1) * F / _as            ' PHYZ(33,j1)=GBphyzB
'        alfaspr = 3600 * Gvspr / GTpr / 14.95            'почему меньше 3 у Сержа ?
'        mpr = PHYZ(39, j1) 'дурь
'        'mpr = mfizmem
'        ' -------     параметры компрессора:     --------------
'        PiKNDpr = PHYZ(15, j1) 'дурь
'        PiKspr = PHYZ(24, j1) 'дурь
'        ' --------- Параметры газа перед турбиной (ТВД)  ----- 
'        '      ТДР-Турбинный датчик расхода - параметр используется не всегда.
'        ' --------- Параметры газа за турбиной (ТНД)  ----- 
'        T3pr = PHYZ(26, j1) * U 'ABKTG'T3fiz * u * ABKTG
'        PiTpr = PHYZ(31, j1)   'дурь   'PiT      ' 28 столбец Сержа
'        U447pr = 288.15 * PHYZ(42, j1) * ABKTG / PHYZ(9, j1)    'дурь   ' РЛ ТВД    ' -13-
'        P6bpr = PHYZ(12, j1) * F  ' -P6bb это P6*
'        P6pr = PHYZ(11, j1) * F           ' P6CT*f                 'P6b*f     '     
'        'PHYZ(11, j) = P6CTb

'        T6bpr = U * PHYZ(14, j1) 'u * T6b
'        P2bpr = F * PHYZ(20, j1) 'f * P2b
'        P2pr = F * PHYZ(19, j1) 'f * P2CT
'        P4bpr = F * PHYZ(30, j1) 'f * P4b
'        If (VZ(24) = 0) Then P6pr = 0 'if(P6CT = 0.) P6pr=0.  '    ??????????
'        P4pr = F * PHYZ(29, j1) 'f * P4CT

'        T4pr = PHYZ(28, j1) * U   'дурь   'T4B*u  ' 21 ' -за ТНД   ' -8-
'        T4bpr = U * PHYZ(28, j1) 'дурь'u * T4b
'        T4pr = PHYZ(28, j1) * 288.15 / PHYZ(9, j1) 'дурь'T4pr = T4B * 288.15 / Tb

'        T2bpr = 288.15 / PHYZ(9, j1) * PHYZ(21, j1)    ' 288.15/Tb*T2b
'        If (PHYZ(14, j1) = 0) Then 'GoTo 72 ' if(T6b = 0.) goto 72
'            n2prKVD = 0
'            GprKVD = 0
'        Else
'            n2prKVD = PHYZ(8, j1) * 1.234 * Sqrt(288.15 / PHYZ(14, j1)) 'n2prKVD = N2f * 1.234 * sqrt(288.15 / T6b)
'            '	GprKVD=GBPHYZs*1.033227/P6bb*sqrt(T6b/288.15)      ' см. (4.6.1) GBPHYZs=GBPHYZ1/ню
'            GprKVD = GBPHYZI * 1.033227 / PHYZ(12, j1) * Sqrt(PHYZ(14, j1) / 288.15)
'        End If

'        T3pr = PHYZ(26, j1) * U   '*ABKTG    '    T3pr=T3fiz*u*ABKTG
'        PiTpr = PHYZ(31, j1) 'PiT'дурь
'        U447pr = 288.15 * PHYZ(42, j1) / PHYZ(9, j1)    '*ABKTG/PHYZ(9,j1)     ' U447pr=288.15*U447*ABKTG/Tb
'        wKND = PHYZ(15, j1) / Gvspr 'wKND = PiKND / Gvspr
'        If (PHYZ(14, j1) <= 0) Then
'            wKVD = 0
'        Else
'            wKVD = PHYZ(22, j1) / (GBPHYZI * 1.033227 / PHYZ(12, j1) * Sqrt(PHYZ(14, j1) / 288.15))
'        End If
'        ' GoTo 162 ' if(T6b <= 0) goto 162
'        'wKVD = PiKVD / (GBPHYZs * 1.033227 / P6bb * sqrt(T6bb / 288.15))

'        '      if(ethaKND < 0) ethaKND=0.
'        If (PHYZ(17, j1) < 0) Then PHYZ(17, j1) = 0

'        '      if(ethaKVD < 0) ethaKVD=0.
'        If (PHYZ(23, j1) < 0) Then PHYZ(23, j1) = 0

'        PRIV(1, j1) = PHYZ(1, j1)
'        PRIV(2, j1) = PHYZ(2, j1)
'        PRIV(3, j1) = n1pr
'        PRIV(4, j1) = n2pr
'        PRIV(5, j1) = PHYZ(3, j1)
'        PRIV(6, j1) = PHYZ(4, j1)
'        PRIV(7, j1) = Rpr
'        PRIV(8, j1) = CRpr
'        PRIV(9, j1) = GTpr
'        PRIV(10, j1) = Gvspr
'        PRIV(11, j1) = alfaspr
'        PRIV(12, j1) = mpr
'        PRIV(13, j1) = PiKNDpr
'        PRIV(14, j1) = wKND
'        PRIV(15, j1) = PHYZ(17, j1)
'        PRIV(16, j1) = PHYZ(22, j1)
'        PRIV(17, j1) = wKVD
'        PRIV(18, j1) = PHYZ(23, j1) 'ethaKVD
'        PRIV(19, j1) = PiKspr
'        PRIV(20, j1) = T3pr
'        PRIV(21, j1) = T3prTDR
'        PRIV(22, j1) = T4pr
'        PRIV(23, j1) = PiTpr
'        PRIV(24, j1) = U447pr
'        PRIV(25, j1) = P6bpr
'        'PRIV(26, j1) = P6bprCT
'        PRIV(27, j1) = P6pr
'        PRIV(28, j1) = T6bpr
'        PRIV(29, j1) = P2bpr
'        PRIV(30, j1) = P2pr
'        PRIV(31, j1) = T2bpr
'        PRIV(32, j1) = P4bpr
'        PRIV(33, j1) = P4pr
'        PRIV(34, j1) = T4bpr
'        PRIV(35, j1) = 0.0
'        PRIV(36, j1) = n2prKVD
'        PRIV(37, j1) = GprKVD
'        PRIV(38, j1) = 0.0
'        PRIV(39, j1) = F
'        PRIV(40, j1) = _as
'        '	      PRINT*,'VIVAT : PRIVEDENIE ISren OVER     14 oct 2004 '
'    End Sub

'    '--прСАУ------прСАУ------прСАУ------прСАУ------прСАУ------прСАУ------прСАУ
'    '--прСАУ------прСАУ------прСАУ------прСАУ------прСАУ------прСАУ------прСАУ 




'    '&&&&&&&&&&&&&&&&&&&&&&&&&&
'    '-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK
'    '           19 Dec 2005

'    '                        (A,VZ,   GBphyzB, ppB)
'    'SUBROUTINE(PARAMRMK(Stend, Tb, PvxCTm, PvxPOLN, GBphyzB, GBprB, ppB))

'    'СУХОЙ(ВОЗДУХ)


'    Private Sub ПараметрыРасходомерногоКоллектора(ByVal Stend As Integer, ByVal Tb As Double, ByVal PvxCTm As Double, ByVal PvxPOLN As Double, ByVal GBphyzB As Double, ByVal ppB As Double)
'        'dimension(A(18), VZ(31))
'        'Dim Stend As Integer
'        Dim k, k1, k2, LamM, LamB As Double
'        Dim DM, DB, akkk, _as, PiM, PMcp, PiMcp, a1, c, E, qLamMcp, PB1, yLamB, PiLb, qLb, Pbtorm, SB, GBprB As Double

'        DM = 0.905
'        DB = 0.905

'        akkk = 1.0
'        U = 288.15 / Tb
'        _as = Sqrt(U)

'        '----      ГДФ ПИ в мерн сечении М:
'        PiM = PvxCTm / PvxPOLN  ' -(1.2) -нумерация формул см стр 12 ПРИЛОЖЕНИЯ 1 в файле 99М1ПМ
'        '----      для входного устр-ва коэфф k1 поля полного давления в сечении "М": 
'        k1 = -0.05019 + 3.3523 * PiM - 3.5814 * PiM * PiM + 1.2794 * PiM * PiM * PiM   '- (1.3)

'        'НомерВходногоУстройства		К1ПолнДавл	К2ПолнДавл	К3ПолнДавл	К4ПолнДавл	К1Потерь	К2Потерь	К3Потерь	К4Потерь	Дм	Дв
'        '6464-4203		                -0.050191	3.352285	-3.581382	1.27943	    -0.464944	4.628497	-4.947336	1.784405	905	905


'        '----      среднее значение полного давления в сечении "М":
'        PMcp = PvxPOLN * k1   '-(1.4)
'        '----      среднее значение ГДФ ПИ в сечении "М":
'        PiMcp = PvxCTm / PMcp   '-(1.4)
'        '	print *,'PiMcp',PiMcp
'        k = 1.4
'        a1 = (k - 1) / (k + 1)
'        c = (k - 1) / k
'        D = (k + 1) / 2
'        E = 1 / (k - 1)
'        B = D ^ E
'        '	v=(1-PiMcp**c)

'        '----------------------------------- 
'        LamM = Sqrt((1 - PiMcp ^ c) / a1)  ' -аналитич реш-ие 1-го ур-ия (1.5)  22.02.2005
'        qLamMcp = LamM * (D * (1 - a1 * LamM * LamM)) ^ E '- второе ур-ие (1.5)
'        '	print *,'LamM,qLamMcp',LamM,qLamMcp
'        '-----------------------------------
'        '  -коэффициент k2 пересчёта статич давл от сеч "М" к сеч "В":
'        '           19 мая 2005г: изменение в зависимости от № стенда: 
'        If (Stend = 34) Then
'            k2 = 0.5359 + 1.285 * PiM - 1.2048 * PiM * PiM + 0.3843 * PiM * PiM * PiM  ' для стенда номер 34 
'            PB1 = PvxCTm * k2   ' -стат давл в сечении В 
'            DM = 0.907
'            akkk = 1.015             '
'            '	  yLamB=qLamMcp*PMcp/PB1*.907*.907/.905/.905      '-(1.7)
'        Else
'            If (Stend <> 15) Then
'                DM = 0.924    ' 20,03,2006 для 15 стенда
'                DB = 0.924    ' 20,03,2006 для 15 стенда
'            End If
'            'Коэф-т пересч статич давления от сеч "М"к сеч "В"  
'            k2 = 0.2803 + 2.161 * PiM - 2.215 * PiM * PiM + 0.7738 * PiM * PiM * PiM  '-(1.6) -для стендов отличных от 34
'            PB1 = PvxCTm * k2   ' -стат давл в сечении В                '-(1.6)
'            '	 yLamB=qLamMcp*PMcp/PB1      '-(1.7)
'        End If


'        yLamB = qLamMcp * PMcp / PB1 * DM * DM / DB / DB     '-(1.7)  ' Учёт наличия конусности
'        '	print *,'Stend,DM,DB',Stend,DM,DB
'        LamB = (Sqrt(B * B + 4 * a1 * yLamB * yLamB) - B) / 2 / a1 / yLamB '-аналит реш-ие ур-ия yLamB(LamB)=q(LamB)/Pi(LamB) относи
'        '   тельно LamB, где правая часть представлена в виде Lam/(1-(k-1)/(k+1))*((k+1)/2)**(1/k-1)   22.02.2005       
'        PiLb = (1 - a1 * LamB * LamB) ^ (1 / c)    '- первое ур-ие (1.5)
'        qLb = LamB * (D * (1 - a1 * LamB * LamB)) ^ E '- второе ур-ие (1.5)
'        '            PiLb1= qLb/yLamB   '-контроль
'        Pbtorm = PB1 / PiLb
'        SB = Math.PI * DB * DB / 4.0                                             '-(1.8)
'        GBphyzB = 0.3964 * Pbtorm / Sqrt(Tb) * qLb * SB * 10000.0 * akkk              '-(1.9)
'        '	 print *,'GBphyzB,Pbtorm,Tb,qLb',GBphyzB,Pbtorm,Tb,qLb
'        '	read(*,*)
'        '\\\\\\\\\\\\\\\\\\\\\\\ Расчёт воздуха в сечении М  \\\\\\\\\\\\\\\\\\\\
'        '      PbtormM=PMcp                                                 '-(1.4)
'        '  	qLamMcpM=LamM*(d*(1-a*LamM*LamM))**e '- второе ур-ие (1.5)
'        '      GBphyzM=0.3964*PbtormM/sqrt(Tb)*qLamMcp*3.14159*.905*.905/4.0*
'        '     *10000.0 
'        '      print *,'GBphyzB,GBphyzM,Pbtorm,PbtormM,qLb,qLamMcpM,
'        '     *qLamMcp,bLam,aLam,
'        '     *PiLM,PiLB,PiLB1',
'        '     *GBphyzB,GBphyzM,Pbtorm,PbtormM,qLb,qLamMcpM,qLamMcp,LamB,LamM,
'        '     *PiMcp,PiLB,PiLB1
'        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'        GBprB = GBphyzB * 1.033227 / Pbtorm * Sqrt(Tb / 288.15)

'        ppB = PB1 / PiLb ' ppB это полное давление Р*_В
'        '	PRINT *,'GBphyzB,ppB',GBphyzB,ppB
'        '	READ(*,*)
'    End Sub

'    ' -RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-
'    ' -RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK



'    '                   8.06.2009г
'    '   -------  УЧЁТ ВЛАЖНОСТИ ВОЗДУХА   ----------

'    '                        (A,VZ,   GBphyzB, ppB)
'    'SUBROUTINE(PARAMRMK(Stend, Tb, PvxCTm, PvxPOLN, GBphyzB, GBprB, ppB))

'    Private Sub RMKd(ByVal Stend As Integer, ByVal Tb As Double, ByVal PvxCTm As Double, ByVal PvxPOLN As Double, ByVal GBphyzB As Double, ByVal ppB As Double, ByVal k As Double)
'        'dimension(A(18), VZ(31))
'        'Dim Stend As Integer
'        Dim k1, k2, LamM, LamB As Double 'k,
'        Dim DM, DB, akkk, _as, PiM, PMcp, PiMcp, a1, qLamMcp, c, E, PB1, PiLb, yLamB, qLb, Pbtorm, SB, ABCmkr, GBprB As Double
'        DM = 0.905
'        DB = 0.905

'        akkk = 1.0
'        U = 288.15 / Tb
'        _as = Sqrt(U)

'        '----      ГДФ ПИ в мерн сечении М:
'        PiM = PvxCTm / PvxPOLN  ' -(1.2) -нумерация формул см стр 12 ПРИЛОЖЕНИЯ 1 в файле 99М1ПМ
'        '----      для входного устр-ва коэфф k1 поля полного давления в сечении "М": 
'        k1 = -0.05019 + 3.3523 * PiM - 3.5814 * PiM * PiM + 1.2794 * PiM * PiM * PiM   '- (1.3)
'        '----      среднее значение полного давления в сечении "М":
'        PMcp = PvxPOLN * k1   '-(1.4)
'        '----      среднее значение ГДФ ПИ в сечении "М":
'        PiMcp = PvxCTm / PMcp   '-(1.4)
'        '	print *,'PiMcp',PiMcp
'        'k = 1.4
'        Debug.Print("into PARAMRMKd:k',k")
'        a1 = (k - 1) / (k + 1)
'        c = (k - 1) / k
'        D = (k + 1) / 2
'        E = 1 / (k - 1)
'        B = D ^ E
'        '	v=(1-PiMcp**c)

'        '----------------------------------- 
'        LamM = Sqrt((1 - PiMcp ^ c) / a1)  ' -аналитич реш-ие 1-го ур-ия (1.5)  22.02.2005
'        qLamMcp = LamM * (D * (1 - a1 * LamM * LamM)) ^ E '- второе ур-ие (1.5)
'        '	print *,'LamM,qLamMcp',LamM,qLamMcp
'        '-----------------------------------
'        '  -коэффициент k2 пересчёта статич давл от сеч "М" к сеч "В":
'        '           19 мая 2005г: изменение в зависимости от № стенда: 
'        If (Stend = 34) Then
'            k2 = 0.5359 + 1.285 * PiM - 1.2048 * PiM * PiM + 0.3843 * PiM * PiM * PiM  ' для стенда номер 34 
'            PB1 = PvxCTm * k2   ' -стат давл в сечении В 
'            DM = 0.907
'            akkk = 1.015             '
'            '	  yLamB=qLamMcp*PMcp/PB1*.907*.907/.905/.905      '-(1.7)
'        Else
'            If (Stend <> 15) Then
'                DM = 0.924    ' 20,03,2006 для 15 стенда
'                DB = 0.924    ' 20,03,2006 для 15 стенда
'            End If
'            'Коэф-т пересч статич давления от сеч "М"к сеч "В"  
'            k2 = 0.2803 + 2.161 * PiM - 2.215 * PiM * PiM + 0.7738 * PiM * PiM * PiM  '-(1.6) -для стендов отличных от 34
'            PB1 = PvxCTm * k2   ' -стат давл в сечении В                '-(1.6)
'            '	 yLamB=qLamMcp*PMcp/PB1      '-(1.7)
'        End If


'        yLamB = qLamMcp * PMcp / PB1 * DM * DM / DB / DB     '-(1.7)  ' Учёт наличия конусности
'        '	print *,'Stend,DM,DB',Stend,DM,DB
'        LamB = (Sqrt(B * B + 4 * a1 * yLamB * yLamB) - B) / 2 / a1 / yLamB '-аналит реш-ие ур-ия yLamB(LamB)=q(LamB)/Pi(LamB) относи
'        '   тельно LamB, где правая часть представлена в виде Lam/(1-(k-1)/(k+1))*((k+1)/2)**(1/k-1)   22.02.2005       
'        PiLb = (1 - a1 * LamB * LamB) ^ (1 / c)    '- первое ур-ие (1.5)
'        qLb = LamB * (D * (1 - a1 * LamB * LamB)) ^ E '- второе ур-ие (1.5)
'        '            PiLb1= qLb/yLamB   '-контроль
'        Pbtorm = PB1 / PiLb
'        SB = Math.PI * DB * DB / 4.0                                             '-(1.8)

'        ABCmkr = Sqrt(k * (2 / (k + 1)) ^ ((k + 1) / (k - 1)))
'        Debug.Print("'ABCmkr=sqrt(k*(2/(k+1)**((k+1)/(k-1)))',ABCmkr,k")
'        '	read(*,*)

'        GBphyzB = 0.3964 * Pbtorm / Sqrt(Tb) * qLb * SB * 10000.0 * akkk              '-(1.9)
'        '	 print *,'GBphyzB,Pbtorm,Tb,qLb',GBphyzB,Pbtorm,Tb,qLb
'        '	read(*,*)
'        '\\\\\\\\\\\\\\\\\\\\\\\ Расчёт воздуха в сечении М  \\\\\\\\\\\\\\\\\\\\
'        '      PbtormM=PMcp                                                 '-(1.4)
'        '  	qLamMcpM=LamM*(d*(1-a*LamM*LamM))**e '- второе ур-ие (1.5)
'        '      GBphyzM=0.3964*PbtormM/sqrt(Tb)*qLamMcp*3.14159*.905*.905/4.0*
'        '     *10000.0 
'        '      print *,'GBphyzB,GBphyzM,Pbtorm,PbtormM,qLb,qLamMcpM,
'        '     *qLamMcp,bLam,aLam,
'        '     *PiLM,PiLB,PiLB1',
'        '     *GBphyzB,GBphyzM,Pbtorm,PbtormM,qLb,qLamMcpM,qLamMcp,LamB,LamM,
'        '     *PiMcp,PiLB,PiLB1
'        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'        GBprB = GBphyzB * 1.033227 / Pbtorm * Sqrt(Tb / 288.15)

'        ppB = PB1 / PiLb ' ppB это полное давление Р*_В
'        '	PRINT *,'GBphyzB,ppB',GBphyzB,ppB
'        '	READ(*,*)
'    End Sub

'    ' -RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-RMK-

'    '   -------  УЧЁТ ВЛАЖНОСТИ ВОЗДУХА   ----------


'    '*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD 


'    Private Sub TVD(ByVal Rejim As Integer, ByVal ISren As Integer, ByVal Hu As Double, ByVal FcaTVD As Double, ByVal P2b As Double, ByVal T2b As Double, ByVal Gtopl As Double, ByRef GBCA As Double, ByRef T3f As Double)
'        Dim KsOHL, sigKC, dP, PP2, Tg0, dTg0, Tg, dTg1, GR, GPRTVD, GBPHYZ, qT1, Tg10, dTg10, Tg1, dTg7, Tg3, Y, Y1 As Double 'ISren, Thhot1
'        'Dim Stend As Integer
'        '  	print *,'Rejim,Hu,FcaTVD,P2b,T2b,Gtopl,GBCA,T3f',
'        '     *Rejim,Hu,FcaTVD,P2b,T2b,Gtopl,GBCA,T3f
'        '      read(*,*)
'        If (ISren = 1) Then
'            KsOHL = 1.0
'        End If
'        If (ISren = 0) Then
'            KsOHL = 1.083
'        End If

'        QT = Gtopl / 3600.0 / 60.0          'GBphyzB             '?????? ?? GT
'        '---------------------------------------------------------  
'        '        ?????? ???????????? ??????? ???????? ??
'        'Для М1
'        '1-ПФ115
'        '2-МФ78
'        '3-М
'        '4-КР
'        'Для М2
'        '1-F
'        '3-M
'        If (Rejim = 4) Then
'            '            GoTo 501
'            '501:
'            dP = (0.0224 + 0.07487 * P2b - 0.0006163 * P2b * P2b) * KsOHL
'            PP2 = P2b - dP
'            sigKC = 0.0995 + 0.905 * PP2 / P2b
'            ' 501  dP=(0.03292+0.07302*P2b-0.000554*P2b*P2b)*KsOHL  ' OLD
'            'PP2 = P2b - dP'OLD
'            'sigKC = 0.0852 + 0.9203 * PP2 / P2b'OLD
'        Else 'If (Rejim = 1 OrElse Rejim = 2 OrElse Rejim = 3) Then GoTo 500

'            '500:
'            sigKC = 0.95
'            'GoTo 505
'        End If

'        '505:
'        Tg0 = 40.347 + 0.928819 * T2b + 41830.4 * QT - 6.27948 * T2b * QT + 0.000029283 * T2b * T2b - 160994.0 * QT * QT
'        dTg0 = -0.00000003574 * Tg0 * Tg0 * Tg0 + 0.0001307 * Tg0 * Tg0 - 0.15396 * Tg0 + 58.27
'        Tg = Tg0 - dTg0
'        '	print *,'Tg0 dTg0 Tg',Tg0,dTg0,Tg
'        '--------------------------------------------------
'        '  ???????? ??  ???????? Hu ???????? ?? 10250:  
'        dTg1 = -0.353 + 0.0093184 * Tg - 0.0086214 * T2b
'        Tg = Tg + dTg1 * (Hu - 10250) / 100
'        '--------------------------------------------------
'        '      print *,'Tg0,dTg0,Tg,dTg1',Tg0,dTg0,Tg,dTg1
'        GR = 118.0
'        GPRTVD = GR * FcaTVD / 305.0

'        GBPHYZ = GPRTVD * sigKC * P2b / (1 + QT) / Sqrt(Tg)
'        qT1 = Gtopl / 3600.0 / GBPHYZ
'        Do
'            '50:
'            Tg10 = 40.347 + 0.928819 * T2b + 41830.4 * qT1 - 6.27948 * T2b * qT1 + 0.000029283 * T2b * T2b - 160994 * qT1 * qT1
'            dTg10 = -0.00000003574 * Tg10 * Tg10 * Tg10 + 0.0001307 * Tg10 * Tg10 - 0.15396 * Tg10 + 58.27
'            Tg1 = Tg10 - dTg10
'            dTg7 = -0.353 + 0.0093184 * Tg1 - 0.0086214 * T2b
'            '  ???????? ??  ???????? Hu ???????? ?? 10250:  
'            Tg3 = Tg1 + dTg7 * (Hu - 10250) / 100
'            'GBPHYZca = GPRTVD * sigKC * P2b / (1 + qT1) / sqrt(Tg3)
'            GBCA = GPRTVD * sigKC * P2b / (1 + qT1) / Sqrt(Tg3)
'            Y = Tg1 - Tg
'            Y1 = Abs(Y)
'            'if(abs(Tg1-Tg)-0.01) 60,60,49
'            If Abs(Tg1 - Tg) <= 0.01 Then Exit Do
'            '49:
'            qT1 = Gtopl / 3600.0 / GBCA     'GBPHYZca
'            Tg = Tg1
'            '	print *,'qT1,T2b,Gtopl,GBCA',qT1,T2b,Gtopl,GBCA
'            I = I + 1
'            'GoTo 50
'        Loop
'        '60:     'Continue 

'        '      print *,'Tg10,dTg10,Tg1,dTg7,Tg3',Tg10,dTg10,Tg1,dTg7,Tg3
'        T3f = Tg3 'Thhot = Tg3
'        '	print *,'T3f',T3f
'        '	read(*,*)
'        'Thhot1 = Tg1
'        '	print *,'sigKC,GBCA',sigKC,GBCA
'        '	print *,'-------------------------'
'    End Sub

'    '*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD
'    '*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD

'    '      29.05.2006  Нахождение Т*3 по Янкину с помощью Р400 и Р403 в "ЛОБ"'''
'    '                  Здесь же для значения коэф отбора NjuOHL=089 вычисл GBsI в I контуре
'    Private Sub TVD1(ByVal Rejim As Integer, ByVal ISren As Integer, ByVal Hu As Double, ByVal FcaTVD As Double, ByVal P2b As Double, ByVal T2b As Double, ByVal Gtopl As Double, ByRef GBCA As Double, ByRef T3f As Double)
'        Dim KsOHL, KH, sigKC, dP, PP2, TH0, RH, MH, VH, TH, GB, GR, AL, HU0, TG0, QTT, QT3, GPRTVD, Skc, GBPHYZ, AL1, QTT1, GB1, Y, Y1 As Double 'ISren,, Thhot1, Tg1
'        Dim KT, A(26), V(9), T(11) As Double 'M0,KC,
'        'common /A/G,AG,M0,R0,R,AR,T1,T2,Q,DI,PI,KT,KC,TKP,PKP,AKP,FKP,M,M1/B/A/D/V
'        'Dim Stend As Integer
'        '  	print *,'Rejim,Hu,FcaTVD,P2b,T2b,Gtopl,GBCA,T3f',
'        '     *Rejim,Hu,FcaTVD,P2b,T2b,Gtopl,GBCA,T3f
'        '      read(*,*)
'        A = New Double() {29.27, 0.0, 1.0775667, 0.0, 14.95, 10331, _
'                          0.24582008, -0.035655427, -0.04983327, 0.57930396, _
'                          -1.0122514, 0.86642789, -0.42615499, 0.12290562, _
'                          -0.019373617, 0.0012910064, 0.16192593, 1.434004, _
'                          -1.5125178, -0.049143299, 2.4095975, -2.9655719, _
'                          1.7457616, -0.56350004, 0.095991782, -0.0067651442}

'        If (ISren = 1) Then
'            KsOHL = 1.0
'        End If
'        If (ISren = 0) Then
'            KsOHL = 1.083
'        End If

'        QT = Gtopl / 3600.0 / 60.0          'GBphyzB             '?????? ?? GT
'        '---------------------------------------------------------  
'        '        ?????? ???????????? ??????? ???????? ??

'        If (Rejim = 4) Then
'            'GoTo 501
'            ' 501  dP=(0.03292+0.07302*P2b-0.000554*P2b*P2b)*KsOHL  ' OLD
'            'PP2 = P2b - dP'OLD
'            'sigKC = 0.0852 + 0.9203 * PP2 / P2b'OLD
'            '501:
'            dP = (0.0224 + 0.07487 * P2b - 0.0006163 * P2b * P2b) * KsOHL
'            PP2 = P2b - dP
'            sigKC = 0.0995 + 0.905 * PP2 / P2b
'        Else 'If (Rejim = 1 OrElse Rejim = 2 OrElse Rejim = 3) Then GoTo 500
'            '500:
'            sigKC = 0.95
'            'GoTo 505
'        End If

'        '505:
'        TH0 = T2b 'ijk
'        G = 9.81
'        AG = 8377.74
'        Q = 0
'        Call P400(9)
'        T1 = 785
'        Call P400(5)
'        KH = KT
'        RH = R
'        MH = 0.5
'        VH = MH * Sqrt(KH * G * A(1) * TH)
'        DI = VH * VH / AG
'        Call P400(2)
'        Call P400(4)
'        GB = 118
'        Debug.Print("'A(5)',A(5)")
'        Debug.Print("'Gtopl',Gtopl")
'        GR = 118
'        AL = GR / Gtopl / A(5) * 3600          'ВЫЧИСЛЕНИЕ AL ПО GT
'        'HZ = X(20)
'        'HU = HZ * A(6)


'        HU0 = 0.99 * Hu
'        Debug.Print(",HZ,Hu")
'        Debug.Print("'A',A")

'        Debug.Print("'TH0,HU0,A(5),AL',TH0,HU0,A(5),AL")
'        Call P403(TH0, HU0, 0, AL, QT3, TG0, Q, 1)
'        Debug.Print("     THE   B E G I N   OF SOLUTIONS'")
'        Debug.Print(" ijk, TH0, HU0, AL, QT3, TG0, Q")
'        '75  format(I5,6F12.5)
'        '          Р403-при S=1-ОПРЕДЕЛЕНИЕ температуры на выходе из камеры 
'        '          сгорания по заданным коэффициенту -AL и т-ре на вх в КС-Тн0
'        '-----------------------------------------------------------------------
'        Thhot = TG0
'        QTT = QT3
'        ''''	FcaTVD=303.07   '- 003 машина в ТБК дек 2004' 303.48
'        GPRTVD = GR * FcaTVD / 305.0  '303.96/305.  '303.32' 303.48
'        Debug.Print("'GPRTVD=GR*FcaTVD/305.',GPRTVD,GR,FcaTVD")
'        ''      GBPHYZ=GPRTVD*.95*P22(ijk)/(1+QTT)/sqrt(TG0)   '     ijk   25.05.2006
'        'as = P22(ijk) - (0.0224 + 0.07487 * P22(ijk) - 0.0006163 * P22(ijk) * P22(ijk))
'        '     * 	*1.083   '1.0 
'        '	Skc=0.0995+0.905*as/P22(ijk)
'        Skc = 0.95
'        Debug.Print("'GPRTVD,Skc,P2b,QTT,TG0',GPRTVD,Skc,P2b,QTT,TG0")
'        GBPHYZ = GPRTVD * Skc * P2b / (1 + QTT) / Sqrt(TG0)

'        Debug.Print("' GBPHYZ      GR      TH0      AL      THHOT   QTT     TH               QT3      MH   GBPHYZ'")
'        Debug.Print("F7.3,F9.2,F9.2,F7.3,F9.2,F7.4,F9.2,F7.4,F5.1,F7.2)',GBPHYZ,GR,TH0,AL,THHOT,QTT,TH,QT3,MH,GBPHYZ ")

'        AL1 = GBPHYZ / Gtopl / A(5) * 3600.0 ' ВЫЧИСЛЕНИЕ AL1 В ПРОЦЕССЕ ИТЕРАЦИЙ ПРИ РЕШ ЗАД ПО GT
'        '-------------------------------------------------------------------------
'        Do

'            '50:
'            Call P403(TH0, HU0, 0, AL1, QT3, TG0, Q, 1)
'            QTT1 = QT3
'            GBCA = GPRTVD * Skc * P2b / (1 + QTT1) / Sqrt(TG0)
'            GB1 = GBCA / 0.89 '      ijk
'            ''      print *,'TG0      GBPHYZ1      KH      RH      GBPHYZ  P22(ijk) '
'            '' 	print '(F9.2,F9.2,F7.4,F7.3,F9.2,F9.2)',TG0,GBPHYZ1,KH,RH,GBPHYZ,
'            ''     *P22(ijk)
'            '       print *,'TG0  GBPHYZ1  KH   RH  GBPHYZ  P2b QT3,Skc'
'            Debug.Print(" TG0, GB1, H, RH, GBCA, P2b, QT3, Skc)")
'            Debug.Print("'TG0,GB1,H,RH,GBCA,P2b,QT3,Skc',F9.2,F9.2,F7.4,F7.3,F9.2,F9.2,F9.6,F6.3)")

'            Y = TG0 - Thhot
'            Y1 = Abs(Y)
'            Debug.Print("'Y,      Y1'")
'            Debug.Print("( F9.2,F9.2)',Y,Y1")
'            'if(abs(TG0-THHOT)-0.1) 60,60,49
'            If Abs(TG0 - Thhot) <= 0.1 Then Exit Do
'            '49:
'            AL1 = GBCA / Gtopl / A(5) * 3600.0
'            Thhot = TG0
'            I = I + 1
'            'GoTo 50
'        Loop

'        '  60   PRINT 104, I,TH0,P2,Gtopl,qtopl,THOT,GBPHYZ  
'        '      WRITE(15,104)ijk,T2b,P2b,GTkc(ijk),QT3,THHOT,GBPHYZ1,Skc
'        '     *,FcaTVD
'        ' 104  FORMAT(I5,8F9.3)

'        '60:
'        T3f = Thhot 'Thhot = Tg3
'        '	print *,'T3f',T3f
'        '	read(*,*)
'        'Thhot1 = Tg1
'        Debug.Print("sigKC,GBCA',sigKC,GBCA")
'        Debug.Print("'-----------------------------------------'")
'        '	Hu=Hu/0.99   ' Восстановление, необходимое для корректной работы при повторном обращении
'    End Sub

'    '*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD
'    '*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD*TVD

'    ' &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'    '          ----------------     P T M   ------------------

'    '      24.03.2009  Нахождение Т*3 по P T M  
'    '                  
'    Private Sub GGPARAM_TVD_ptm(ByVal Rejim As Integer, ByVal ISren As Integer, ByVal Skc1 As Double, ByVal Hu As Double, ByVal FcaTVD As Double, ByVal P2b As Double, ByVal T2b As Double, ByVal Gtopl As Double, ByRef GBphyz As Double, ByRef T3captm As Double, ByVal Xptm As Double, ByVal Aptm(,) As Double, ByVal Cptm() As Double, ByVal dH2O As Double)

'        Dim Gts As Double 'ISren,KsOHL, KH,
'        Dim A(26), V(9), T(11) As Double 'M0, KT, KC,
'        'ReDim_Aptm(6, 11)
'        'ReDim_Cptm(6)
'        Dim TH0, GR, HU0, GPRTVD, TG0, TTptm, ENTH1ptm, aIg1ptm, TT0ptm As Double
'        'common /AJ/G,AG,M0,R0,R,AR,T1,T2,Q,DI,PI,KT,KC,TKP,PKP,AKP,FKP,M,M1/B/A/D/V
'        'common /FYZGG/ q10,PHYZGB     ' PHYZ(27,j)=GBphyzB1=PHYZGB 
'        'common/FYZGG1/GTp1,qTp   ' GTp1=qTp*GBphyzB1

'        TH0 = T2b
'        GR = 118.0  '118.2
'        HU0 = 0.99 * Hu

'        Debug.Print("'TH0,HU0,A(5),AL',TH0,HU0,A(5),AL")
'        'call(P403(TH0, HU0, 0, AL, QT3, TG0, Q, 1))

'        Gts = Gtopl / 3600.0  ' секундный расход топлива в ОКС 

'        ''	DO J=1,100
'        ''	T0ptm=293.15/1000.
'        ''	 TTptm=(293.15+(J-1)*10.)/1000.
'        ''	Xptm=qTp*14.92914 
'        ''           ENTH1ptm=DIptm(T0ptm,TTptm,Xptm,Aptm)   ' ЭНТАЛЬПИЯ рабочего тела на выходе из КВД
'        ''      PRINT *,'ENTH1ptm',(293.15+(J-1)*10.),ENTH1ptm
'        ''      write(475,335) T2,Q,DI5
'        '' 335  format(3(F15.9,'	'))
'        ''	END DO
'        Skc1 = 0.95
'        '--------------------------------------------------------
'        GPRTVD = GR * FcaTVD / 305.0  '303.96/305.  '303.32' 303.48
'        TG0 = 500.0
'        QT = Gts / GPRTVD
'        Do
'            '44:
'            GBphyz = GPRTVD * Skc1 * P2b / (1 + QT) / Sqrt(TG0)  ' на самом деле это не GBPHYZ, а " газ в горле СА"
'            Debug.Print("'GBPHYZ=GPRTVD*Skc1*P2b/(1+QT)/sqrt(TG0)',GBPHYZ,GPRTVD,Skc1,P2b,QT,sqrt(TG0),TG0")
'            '--------------------------------------------------------

'            '                  ---------------- Р Т М  --------------------- 

'            '  ----     решение задачи: ОСНОВНАЯ КАМЕРА СГОРАНИЯ -------

'            '       T0=Tvh/1000.
'            T0ptm = 293.15 / 1000.0
'            '      TTptm=(Tvhptm+273.15)/1000.
'            TTptm = T2b / 1000.0

'            Debug.Print("'into GGPARAM TVD ptm:qTp,q10',qTp,q10")
'            'q10 = qTp
'            '       Xptm=qTp*14.92914   '5
'            'Xptm = QT * 14.92914
'            Debug.Print("'T0ptm,TTptm,Xptm,QT,dH2O',T0ptm,TTptm,Xptm,QT,dH2O")
'            Debug.Print("' GGPARAM TVD ptm:DIptm(T0ptm,TTptm,Xptm,Aptm)',T0ptm,TTptm,Xptm")

'            T0ptm = 293.15 / 1000.0
'            TTptm = T2b / 1000.0
'            Xptm = 0.0    '   9.06.09 LOPUHH'''
'            Debug.Print("'ENTH1ptm=DIptm(T0ptm,TTptm,Xptm,Aptm,dH2O)',ENTH1ptm,T0ptm,TTptm,Xptm,Aptm,dH2O")
'            ENTH1ptm = DIptm(T0ptm, TTptm, Xptm, Aptm, dH2O)   ' энтальпия воздуха
'            aIg1ptm = (ENTH1ptm * GBphyz + Gts * HU0) / (GBphyz * (1 + QT))  ' ЭНТАЛЬПИЯ газа в горле СА
'            Debug.Print("'aIg1ptm=(ENTH1ptm*GBPHYZ+Gts*Hu0)/(GBPHYZ)',aIg1ptm,ENTH1ptm,GBPHYZ,Gts,Hu0")

'            Xptm = Gts / GBphyz * 14.92914 ' -Gts-GTp1)*14.92914     '14.95
'            Debug.Print("'Xptm=Gts/GBPHYZ',  '-Gts-GTp1)*14.92914',  ' 95,Xptm,Gts,GBPHYZ")

'            'GTp1 = qTp * GBphyzB1

'            TT0ptm = 293.15
'            T3captm = TIptm(aIg1ptm, TT0ptm, Xptm, Aptm, Cptm, dH2O)
'            Debug.Print("aIg1ptm,TT0ptm,Xptm',aIg1ptm,TT0ptm,Xptm")
'            Debug.Print("'T3captm',T3captm")

'            If Abs(T3captm - TG0) < 0.01 Then Exit Do
'            TG0 = T3captm
'            QT = Xptm / 14.92914
'            'GoTo 44
'        Loop

'        '55:
'        Debug.Print("'result TG0=T3captm: Gtopl',TG0,T3captm,Gtopl")
'        T3f = TG0
'        '	read(*,*)
'    End Sub

'    '

'    '          ----------------     P T M   ------------------

'    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&



'    'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP()


'    '---------       P400(S)       -------------'  12.02.03 
'    Private Sub P400(ByVal S As Integer)
'        Dim F, Z, NN, IW, K, L As Integer 'S
'        Dim E As Double
'        Dim M0, KT, KC, A(46), V(9), T(11) As Double
'        Dim IU As Double
'        '	common /A/G,AG,M0,R0,R,AR,T1,T2,Q,DI,PI,KT,KC,TKP,
'        '*PKP,AKP,FKP,M,M1/B/A/CP400/I9
'        '     N-порядок аппроксимирующего полинома ''' см стр 96 Янкина.
'        N = 9   '-ЯНКИН
'        '      N=8  '-ЦАРЬКОВ
'        T(4) = 1.0 + Q
'        IU = S
'        'реально вызываются 1,2,4,5,9

'        '901:goto (11,12,13,14,15,16,17,18,19),IU
'        Select Case IU
'            Case 1
'                '11:
'                K = 1
'                '51:
'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31
'            Case 2
'                '12:
'                T(1) = DI
'                ''''''''''''''      T(2)=100.   ''''''''' POPA-VOVA 24.07 2002
'                T2 = 100.0
'                K = 5
'                'print(1000, T(1))
'                ' 1000 format('12-51 T(1)=DI',F10.4)
'                'GoTo 51
'                '51:
'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31
'            Case 3
'                '13:
'                K = 1
'                '53:
'                R = R0 * (1 + M0 * Q) / T(4)
'                AR = R / 427
'                IW = K
'                GoTo 902
'            Case 4
'                F = 1
'                '54:
'                T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Log(T2 / T1)
'                Z = 8
'                L = 2
'                GoTo 31
'            Case 5
'                '15:
'                J = 1
'                'GoTo 55
'                Call Блок55(T, NN, A, K, IW)
'                GoTo 902
'            Case 6
'                '16:
'                'if(abs(T2-T1)-0.5)911,911,562
'                If Abs(T2 - T1) <= 0.5 Then
'                    '911:
'                    J = 11
'                    'GoTo 55
'                    Call Блок55(T, NN, A, K, IW)
'                    GoTo 902
'                Else
'                    '562:
'                    K = 3
'                    'GoTo 51
'                    '51:
'                    T(8) = 0
'                    Z = 7
'                    L = 1
'                    GoTo 31
'                End If
'            Case 7
'                '17:
'                T(11) = PI
'                T2 = 100
'                F = 9
'                'GoTo 54
'                '54:
'                T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Log(T2 / T1)
'                Z = 8
'                L = 2
'                GoTo 31
'            Case 8
'                '18:
'                T2 = 0.8 * T1
'                T(11) = T2
'                E = -1
'                T(10) = T1
'                '73:
'                T1 = T2
'                J = 12
'                'GoTo 55
'                Call Блок55(T, NN, A, K, IW)
'                GoTo 902
'            Case 9
'                '19:
'                'If (Q < 1.0 / A(5) OrElse A(2) = 0.0) Then GoTo 904
'                If (Q < 1.0 / A(5) OrElse A(2) = 0.0) Then
'                    M = 0
'                    R0 = A(1)
'                    M0 = A(3)
'                Else
'                    M = 2 * (N + 1)
'                    R0 = A(2)
'                    M0 = A(4)
'                End If

'                M1 = M + N + 1
'                'GoTo 999
'                Exit Sub
'        End Select


'522:    T2 = T(9) + (100.0 - T(9)) * (T(1) - T(8)) / (T(10) - T(8))
'        'if(abs(T2-T(9))-0.1)888,906,906
'        If Abs(T2 - T(9)) < 0.1 Then
'            '888	if(S-2)572,523,572
'            If S = 2 Then
'                '523:
'                DI = T(1)
'            Else
'                '572:
'                PI = T(11)
'            End If
'            Exit Sub
'        Else
'            '906:
'            T(9) = T2

'            'if(S-2)907,908,907
'            If S = 2 Then
'                '908:
'                K = 2
'                'GoTo 51
'                '51:
'                T(8) = 0
'                Z = 7
'                L = 1
'                'GoTo 31
'            Else
'                '907:
'                F = 2
'                'GoTo 54
'                '54:
'                T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Log(T2 / T1)
'                Z = 8
'                L = 2
'                'GoTo 31
'            End If
'        End If

'        'GoTo 31
'31:
'        T(5) = 1
'        T(6) = 1
'        T(7) = 0
'        T(2) = T2 / 1000
'        T(3) = T1 / 1000
'        NN = N + 7

'        For I As Integer = Z To NN
'            T(5) = T(5) * T(2)
'            T(6) = T(6) * T(3)
'            T(7) = T(7) + 1
'            ii = I + M
'            ij = I + M1
'            '903:
'            T(8) = T(8) + (T(5) - T(6)) * (A(ii) + Q * A(ij)) / (T(7) * T(4))
'        Next

'        'print(1001, T(8), I)
'        ' 1001 format('31- T(8)',F10.4,I5)
'        'if(L-1)33,32,33
'        If L = 1 Then
'            '32:
'            DI = T(8) * 1000
'            T(8) = DI
'            IW = K
'            GoTo 902
'        Else
'            '33	if(F-2)909,522,909
'            If F = 2 Then
'                GoTo 522
'            Else
'                '909:
'                K = 6
'                'GoTo 53
'                '53:
'                R = R0 * (1 + M0 * Q) / T(4)
'                AR = R / 427
'                IW = K
'                GoTo 902
'            End If
'        End If
'        Exit Sub



'        '902  goto (999,522,563,581,34,531,532,564,571,582,561,58),IW
'902:
'        Select Case IW
'            Case 1
'                'Exit Sub
'            Case 2
'                GoTo 522
'            Case 3
'                '563:
'                T(1) = DI / (T2 - T1)
'                K = 8
'                'GoTo 53
'                '53:
'                R = R0 * (1 + M0 * Q) / T(4)
'                AR = R / 427
'                IW = K
'                GoTo 902
'            Case 4
'                '581:
'                T(1) = DI + AR * KT * T2 / 2.0
'                'if(E)912,912,7
'                If E <= 0 Then
'                    '912:
'                    T(9) = T(1)
'                    T2 = T1
'                    E = 1
'                    'GoTo 73
'                    '73:
'                    T1 = T2
'                    J = 12
'                    'GoTo 55
'                    Call Блок55(T, NN, A, K, IW)
'                    GoTo 902
'                Else
'                    '7:
'                    TKP = (T2 * T(9) - T(11) * T(1)) / (T(9) - T(1))
'                    'if(abs(T2-TKP)-0.1)71,913,913
'                    If Abs(T2 - TKP) < 0.1 Then
'                        '71:
'                        T2 = TKP
'                        F = 10
'                        'GoTo 54
'                        '54:
'                        T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Log(T2 / T1)
'                        Z = 8
'                        L = 2
'                        GoTo 31
'                    Else
'                        '913:
'                        T2 = TKP
'                        'GoTo 73
'                        '73:
'                        T1 = T2
'                        J = 12
'                        'GoTo 55
'                        Call Блок55(T, NN, A, K, IW)
'                        GoTo 902
'                    End If
'                End If
'            Case 5
'                '34:
'                T(10) = T(8)
'                T(8) = 0.0
'                T(9) = T1
'                GoTo 522
'            Case 6
'                '531:
'                PI = Exp(T(8) / AR)
'                'print(987, PI)
'                ' 987  format('PI',F10.5)
'                IW = F
'                GoTo 902
'            Case 7
'                '532:
'                KT = T(3) / (T(3) - AR)
'                IW = J
'                GoTo 902
'            Case 8
'                '564:
'                KC = T(1) / (T(1) - AR)
'                'Exit Sub
'            Case 9
'                '571
'                T(1) = AR * Log(T(11))
'                T(10) = T(8)
'                T(8) = 0.0
'                T(9) = T1
'                GoTo 522
'            Case 10
'                '582:
'                PKP = PI
'                AKP = Sqrt(G * KT * R * TKP)
'                '999:            ' Continue Do
'            Case 11
'                '561:
'                KC = KT
'                'Exit Sub
'            Case 12
'                '58:
'                T1 = T(10)
'                K = 4
'                'GoTo 51
'                '51:
'                T(8) = 0
'                Z = 7
'                L = 1
'                GoTo 31
'        End Select

'        '571:    T(1) = AR * Log(T(11))
'        '34:     T(10) = T(8)
'        '        T(8) = 0.0
'        '        T(9) = T1


'        '13:     K = 1
'        '53:     R = R0 * (1 + M0 * Q) / T(4)
'        '        AR = R / 427
'        '        IW = K
'        '        GoTo 902
'        '14:     F = 1
'        '        '  54  PRINT 7654,T2,T1
'        '        ' 7654 FORMAT('--T2,T1===',2D12.4)
'        '54:     T(8) = ((A(M + 7) + Q * A(M1 + 7)) / T(4)) * Log(T2 / T1)
'        '        'UU = log(T2 / T1) - нат.лог
'        '        'VV = alog(T2 / T1) - нат.лог
'        '        '	WW=log10(T2/T1) - десятичный логарифм
'        '        'print(701, UU, VV, WW)
'        '        ' 701  format('UU,VV,WW',3F10.4)
'        '        'stop()
'        '        Z = 8
'        '        L = 2
'        '        GoTo 31
'        'Return
'    End Sub

'    Private Sub Блок55(ByVal T() As Double, ByVal NN As Double, ByVal A() As Double, ByVal K As Integer, ByVal IW As Integer)
'        '55:
'        Dim I As Integer
'        T(1) = T1 / 1000
'        T(2) = 1
'        T(3) = (A(M + 7) + Q * A(M1 + 7)) / T(4)
'        NN = N + 7

'        For I = 8 To NN
'            T(2) = T(1) * T(2)
'            ii = I + M
'            ij = I + M1
'            '910:
'            T(3) = T(3) + T(2) * (A(ii) + Q * A(ij)) / T(4)
'        Next
'        K = 7
'        'GoTo 53
'        '53:
'        R = R0 * (1 + M0 * Q) / T(4)
'        AR = R / 427
'        IW = K
'    End Sub


'    Private Sub P403(ByVal X1 As Double, ByVal X2 As Double, ByVal X3 As Double, ByVal X4 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal Y3 As Double, ByVal S As Integer)
'        '    Р403- тепловой расчет КС(топливо керосин или водород):
'        'Dim S As Integer
'        '	real*8 X1,X2,X3,X4,Y1,Y2,Y3,L,
'        '     *G,AG,M0,R0,R,AR,T1,T2,Q,DI,PI,KT,KC,TKP,
'        '     *PKP,AKP,FKP,
'        '     *A(46),V(9)
'        Dim A(26), V(9), L As Double 'M0, KT, KC,
'        Dim K As Integer

'        'common /A/G,AG,M0,R0,R,AR,T1,T2,Q,DI,PI,KT,KC,TKP,
'        '    *PKP,AKP,FKP,M,M1/B/A/D/V
'        '	 print *,'X1,X2,X3,X4',X1,X2,X3,X4
'        Q = X3
'        Call P400(9)
'        V(4) = 1.0 + Q
'        L = -1
'        'if(S)1,3,1
'        If S <> 0 Then
'            '1:
'            V(2) = X4
'            X4 = X1
'            V(3) = 200
'            L = 1
'        End If

'        '4:
'        Do
'            If S <> 0 Then X4 = X4 + V(3)
'            '3:
'            T1 = X1
'            T2 = X4
'            Call P400(1)
'            V(1) = DI
'            T1 = 293
'            V(5) = Q
'            K = M
'            Q = 0
'            M = M1
'            Call P400(1)
'            Q = V(5)
'            M = K
'            Y1 = V(4) * V(1) / (X2 - DI)
'            Y2 = 1 / ((Y1 + Q) * A(5))
'            '	print *,'V(4),V(1),X2,DI,Y1,Y2,Y1,Q,A(5)'
'            '     *	,V(4),V(1),X2,DI,Y1,Y2,Y1,Q,A(5)
'            'if(L)2,6,6
'            If L < 0 Then ' выполнится когда If S = 0 Then
'                '2:
'                Y3 = Q + Y1
'                Exit Do
'            Else
'                '6  if(V(3)-.1)5,7,7
'                If V(3) < 0.1 Then
'                    '5:
'                    Y2 = X4
'                    X4 = V(2)
'                    '2:
'                    Y3 = Q + Y1
'                    Exit Do
'                Else
'                    '7  if(Y2-V(2))8,4,4
'                    If (Y2 - V(2)) < 0 Then
'                        '8:
'                        X4 = X4 - V(3)
'                        V(3) = V(3) / 2
'                        '	print *,'V(3),Y2,V(2)',V(3),Y2,V(2)
'                    End If
'                    'GoTo 4
'                End If
'            End If
'        Loop

'        '5:          Y2 = X4
'        '            X4 = V(2)
'        '2:          Y3 = Q + Y1
'        '            Return
'    End Sub








'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET
'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET

'    '           Для KENTAVR:  ДОБАВОЧНЫЙ БЛОК subroutine Ralfa(j1,IJK,tb1,PHYZ,rkRper) 
'    '                         ДЛЯ РАСЧЁТА Rпер на режиме FALF (что вызвано семейством кривых
'    '                           rekR(alfa) для ряда значений alfa-альфа_топлива) 

'    '     Замечание: для режимов Др и Мал газ параметры пересчёта не определять'''
'    '     Это условие записано в SUBROUTINE PHYZPARAMETERS(IJK)
'    '       j1 - номер варианта 
'    '       Пересчёт к САУ основных параметров изделия на Мах и Форс режимах
'    Private Sub PERESCHET(ByVal j1 As Integer, ByVal IJK As Integer, ByVal PHYZ(,) As Double, ByVal PER(,) As Double, ByVal cKper(,) As Double, ByVal cKperfi(,) As Double)     'GBphyzB,ppB,
'        Dim ck(13), n1per, n2per, mper As Double ' N1f, N2f, mfizmem, N1fiz, N2fiz, mfors, Kv, njuOHL,, mcay,, n1pero, n2pero, mpero, n1cay, n2cay
'        'Dim CD, CE, CF As String
'        Dim BLet(21) As Double
'        '     common/Kvv/CD
'        'common/phyzPER/T3fizTDR
'        Dim VZ(32) As Double
'        'Dim PER(21, IJK), PHYZ(45, IJK), cKper(15, IJK), cKperfi(15, IJK) As Double
'        Dim im As Integer
'        Dim pq, sqr, bc, Rper, CRper, Gtper, T3per, T3perTDR, T4per, GBper, PiKNDper, PiKSper, alfaper, aPiKVD, aGTf As Double

'        'If (im > 0) Then GoTo 732 ' - ключ обхода участка открытие файла 9 и запись строки идентификаторов
'        If im = 0 Then

'            '------------------------------------------------
'            'OPEN(79,FILE='B rez\PERESCHET.xls')
'            '------------------------------------------------

'            BLet(1) = "Nкт"
'            BLet(2) = "Режим"
'            BLet(3) = "n1пер"
'            BLet(4) = "n2пер"
'            BLet(5) = "Rпер"
'            BLet(6) = "CRпер"
'            BLet(7) = "Gт_пер"
'            BLet(8) = "GВ_пер"
'            BLet(9) = "альфа_пер"
'            BLet(10) = "mпер"
'            BLet(11) = "PiКНДпер"
'            BLet(12) = "PiКsпер"
'            BLet(13) = "T3пер"
'            BLet(14) = "T4пер"
'            BLet(15) = "alfaSпер"
'            BLet(16) = "PiKVDпер"
'            BLet(17) = "GTFпер"

'            BLet(18) = "T3перТДР"  '""T4-5-Tогр"
'            BLet(19) = "dT"
'            BLet(20) = "Т4огр"
'            BLet(21) = "1.033\Pb"

'            '      write(9,177) (Blet(j), j=1,21)    '   OPEN(9,FILE='B rez\PERESCHET.TXT')
'            ' 177  format(21G12.2)
'            '---------------------------------------------
'            Debug.Write("Blet(J), J = 1, 21)")  '   OPEN(279,FILE='B rez\ex_PERESCHET.xls') зап строки идент
'            '5177  format(21(G12.2,'	'))
'            '-------------------------------------------
'            im = im + 1  ' - ключ обхода участка открытие файла 9 и запись строки идентификаторов
'        End If

'        '732:    'Continue Do
'        ''       print *,'j1,cKper(i,j1)',j1,(cKper(j1,1),j1=1,15)
'        '	print *,'j1,cKper',j1,(cKper(i,1),i=1,15)
'        '1	  2	   3	   4	 5	  6	        7	      8	        9	    10	    11	      12     13	       14	      15
'        'Nкт   kn1fi  kPiBfi  kGBfi  kmfi  kn2fi   kPiKSfi  kGT0fi    kT3fi    kT4fi   kRfi     kCRfi  FkaSfi  FkPiKVDfi FkGTFfi   

'        '      if(PHYZ(2,j1) = 4) goto 990
'        n1per = cKper(2, j1) * cKperfi(2, j1) * PHYZ(7, j1) 'n1cay * PHYZ(7, j1)'N1f
'        '	print *,'n1per,n1cay,VZ(9)',n1per,n1cay,VZ(9)
'        '      print *,'n1per=cKper(2,j1)*PHYZ(7,j1)',n1per,cKper(2,16),
'        '     *PHYZ(7,j1)
'        Debug.Print("'n1per=cKper(2,j1)*cKperfi(2,j1)*PHYZ(7,j1)',n1per,cKper(2,j1),cKperfi(2,j1),PHYZ(7,j1)")
'        '       read(*,*)
'        'Царьков()
'        pq = 288.15 / PHYZ(9, j1)
'        sqr = Sqrt(pq)
'        '	print *,'PHYZ(9,j1),pq,sqr',PHYZ(9,j1),pq,sqr
'        n2per = cKper(6, j1) * cKperfi(6, j1) * PHYZ(8, j1)  '*sqr  'rekN2*PHYZ(8,j1)*sqr      ' N2f
'        bc = 1.033227 / PHYZ(10, j1) 'ppB
'        'Rper = bc * Rcay * PHYZ(34, j1)'Rfiz
'        Rper = cKper(11, j1) * cKperfi(11, j1) * PHYZ(34, j1) * bc 'bc * Rfiz * rekR * PHYZ(34, j1)'Rfiz
'        '	print *,'Rper,bc,cKper(4,1),PHYZ(34,j1),j1',Rper,bc,cKper(4,16),
'        '       *PHYZ(34,j1),j1
'        '	read(*,*)
'        CRper = cKper(12, j1) * cKperfi(12, j1) * PHYZ(38, j1) 'CR
'        Gtper = PHYZ(35, j1) * cKper(8, j1) * cKperfi(8, j1) * bc   '*sqr*bc

'        T3per = cKper(9, j1) * cKperfi(9, j1) * PHYZ(26, j1) 'T3fizmem * pq
'        '	print *,'T3per,rekTG,PHYZ(26,j1)',T3per,rekTG,PHYZ(26,j1)
'        '	read(*,*)
'        T3perTDR = cKper(9, j1) * cKperfi(9, j1) * PHYZ(27, j1)
'        T4per = cKper(10, j1) * cKperfi(10, j1) * PHYZ(28, j1) 'T4B * pq
'        '      print *,'T4per,T4cay,PHYZ(28,j1),j1',T4per,T4cay,PHYZ(28,j1),
'        '     *	j1
'        GBper = cKper(4, j1) * cKperfi(4, j1) * PHYZ(33, j1) * bc 'bc * GBphyzB
'        '	print *,'GBper=bc*GBcay* PHYZ(33,j1)',
'        '     *GBper,bc,GBcay,PHYZ(33,j1),j1
'        '	 read(*,*)


'        mper = cKper(5, j1) * cKperfi(5, j1) * PHYZ(39, j1) 'mfizmem
'        PiKNDper = cKper(3, j1) * cKperfi(3, j1) * PHYZ(15, j1) 'PiKND
'        PiKSper = cKper(7, j1) * cKperfi(7, j1) * PHYZ(24, j1) 'PiKsum
'        'If (CD = "F") Then GoTo 88
'        If (CD = "F") Then
'            '88:
'            alfaper = cKper(13, j1) * PHYZ(37, j1)     ' *cKperfi(13,j1)*PHYZ(37,j1)
'            aPiKVD = cKper(14, j1) * cKperfi(14, j1) * PHYZ(22, j1)
'            aGTf = cKper(15, j1) * cKperfi(15, j1) * PHYZ(35, j1)
'        Else
'            alfaper = 0
'            aPiKVD = 0
'            aGTf = 0.0
'        End If
'        'GoTo 99
'        '	print *,'alfaper=cKper(13,j1)*PHYZ(37,j1)',alfaper,cKper(13,j1),
'        '     *PHYZ(37,j1)
'        '++++++++++++++++++++++++++++++++++++++++++++++++++
'        '  99     ij=ij+1
'        '99:
'        PER(1, j1) = PHYZ(1, j1) 'Nkt
'        PER(2, j1) = PHYZ(2, j1) 'Rejim
'        PER(3, j1) = n1per  'T2per    """
'        '"	print *,"PER(3,ij)",PER(3,ij)
'        PER(4, j1) = n2per  '  """  
'        PER(5, j1) = Rper
'        PER(6, j1) = CRper
'        PER(7, j1) = Gtper
'        PER(8, j1) = GBper


'        PER(9, j1) = alfaper 'Глюк
'        PER(15, j1) = alfaper 'Глюк

'        PER(10, j1) = mper
'        PER(11, j1) = PiKNDper
'        PER(12, j1) = PiKSper
'        PER(13, j1) = T3per
'        PER(14, j1) = T4per

'        PER(16, j1) = aPiKVD
'        PER(17, j1) = aGTf
'        '"PER(17, ij) = T3perTDR
'        Dim T45tOGR, delTT As Double
'        PER(18, j1) = T45tOGR
'        PER(19, j1) = delTT
'        PER(20, j1) = PHYZ(43, j1)
'        PER(21, j1) = bc


'        Debug.Print("'--1--ij,PER',(PER(i,j1),i=1,21) ")
'        '      IF(CD = 'FALF')  call Ralfa(j1,IJK,tb1,PHYZ,rekR) 
'        'F(2) = rkRper
'        'Rper = bc * rekR * PHYZ(34, j1)
'        '	print *,'Rper=bc*rekR*PHYZ(34,j1)',Rper,bc,rekR,PHYZ(34,j1) 
'        'PER(5, ij) = Rper
'        'PER(6, ij) = rekR
'    End Sub
'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET
'    '----PERESCHET----PERESCHET       П Е Р Е С Ч Ё Т       -PERESCHET----PERESCHET

'    Private Function Kadiabat(ByVal Tgrad As Double) As Double
'        Dim K As Double
'        If (Tgrad <= 50) Then
'            K = 1.4
'        ElseIf (Tgrad > 50 AndAlso Tgrad < 500) Then
'            K = 1.4035 - 5.05 * 0.00001 * Tgrad - 8.5 * 0.00000001 * Tgrad * Tgrad    ' (4.2.3)
'        End If
'        Return K
'    End Function

'End Module
