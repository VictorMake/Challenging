'Module TermoDinamik

'    'Type FormatParName
'    '  CharacterI() As String
'    '  FormatI() As Object
'    'End Type


'    'Type DataXl
'    '  NamePar As String
'    '  NameLocation As Range
'    '  ValLocation As Range
'    '  Val() As Single
'    '  QuantSample As Integer
'    '  FormatName As FormatParName
'    '  ErrData() As Boolean
'    'End Type


'#Region "Const"

'    Const AH1i1 As Double = 47.67039
'    Const AH1i2 As Double = 50.05635
'    Const AH1i3 As Double = 52.44334
'    Const AH1i4 As Double = 54.83153
'    Const AH1i5 As Double = 57.22099
'    Const AH1i6 As Double = 59.61192
'    Const AH1i7 As Double = 62.00439
'    Const AH1i8 As Double = 64.39851
'    Const AH1i9 As Double = 66.79445
'    Const AH1i10 As Double = 69.19234
'    Const AH1i11 As Double = 71.59227
'    Const AH1i12 As Double = 73.99432
'    Const AH1i13 As Double = 76.39873
'    Const AH1i14 As Double = 78.8055
'    Const AH1i15 As Double = 81.21484
'    Const AH1i16 As Double = 83.6268
'    Const AH1i17 As Double = 86.04156
'    Const AH1i18 As Double = 88.45918
'    Const AH1i19 As Double = 90.87985
'    Const AH1i20 As Double = 93.30356
'    Const AH1i21 As Double = 95.73053
'    Const AH1i22 As Double = 98.1608
'    Const AH1i23 As Double = 100.5945
'    Const AH1i24 As Double = 103.0318
'    Const AH1i25 As Double = 105.4726
'    Const AH1i26 As Double = 107.9172
'    Const AH1i27 As Double = 110.3656
'    Const AH1i28 As Double = 112.818
'    Const AH1i29 As Double = 115.2743
'    Const AH1i30 As Double = 117.7348
'    Const AH1i31 As Double = 120.1994
'    Const AH1i32 As Double = 122.6683
'    Const AH1i33 As Double = 125.1415
'    Const AH1i34 As Double = 127.6191
'    Const AH1i35 As Double = 130.1013
'    Const AH1i36 As Double = 132.588
'    Const AH1i37 As Double = 135.0793
'    Const AH1i38 As Double = 137.5753
'    Const AH1i39 As Double = 140.0761
'    Const AH1i40 As Double = 142.5817
'    Const AH1i41 As Double = 145.0921
'    Const AH1i42 As Double = 147.6075
'    Const AH1i43 As Double = 150.1278
'    Const AH1i44 As Double = 152.6532
'    Const AH1i45 As Double = 155.1836
'    Const AH1i46 As Double = 157.7191
'    Const AH1i47 As Double = 160.2597
'    Const AH1i48 As Double = 162.8056
'    Const AH1i49 As Double = 165.3566
'    Const AH1i50 As Double = 167.9128
'    Const AH1i51 As Double = 170.4743
'    Const AH1i52 As Double = 173.0411
'    Const AH1i53 As Double = 175.6132
'    Const AH1i54 As Double = 178.1905
'    Const AH1i55 As Double = 180.7733
'    Const AH1i56 As Double = 183.3613
'    Const AH1i57 As Double = 185.9547
'    Const AH1i58 As Double = 188.5536
'    Const AH1i59 As Double = 191.1576
'    Const AH1i60 As Double = 193.7671
'    Const AH1i61 As Double = 196.3819
'    Const AH1i62 As Double = 199.0021
'    Const AH1i63 As Double = 201.6276
'    Const AH1i64 As Double = 204.2584
'    Const AH1i65 As Double = 206.8946
'    Const AH1i66 As Double = 209.5361
'    Const AH1i67 As Double = 212.1829
'    Const AH1i68 As Double = 214.8349
'    Const AH1i69 As Double = 217.4922
'    Const AH1i70 As Double = 220.1548
'    Const AH1i71 As Double = 222.8225
'    Const AH1i72 As Double = 229.5144
'    Const AH1i73 As Double = 236.2382
'    Const AH1i74 As Double = 242.9933
'    Const AH1i75 As Double = 249.7793
'    Const AH1i76 As Double = 256.595


'    Const AH2i1 As Double = 263.4402
'    Const AH2i2 As Double = 270.3145
'    Const AH2i3 As Double = 277.2163
'    Const AH2i4 As Double = 284.1453
'    Const AH2i5 As Double = 291.1003
'    Const AH2i6 As Double = 298.0806
'    Const AH2i7 As Double = 305.0857
'    Const AH2i8 As Double = 312.114
'    Const AH2i9 As Double = 319.1653
'    Const AH2i10 As Double = 326.2383
'    Const AH2i11 As Double = 333.3323
'    Const AH2i12 As Double = 340.4465
'    Const AH2i13 As Double = 347.5798
'    Const AH2i14 As Double = 354.7317
'    Const AH2i15 As Double = 361.9011
'    Const AH2i16 As Double = 369.0874
'    Const AH2i17 As Double = 376.29
'    Const AH2i18 As Double = 383.5081
'    Const AH2i19 As Double = 390.741
'    Const AH2i20 As Double = 397.9878
'    Const AH2i21 As Double = 405.2483
'    Const AH2i22 As Double = 412.522
'    Const AH2i23 As Double = 419.8081
'    Const AH2i24 As Double = 427.1064
'    Const AH2i25 As Double = 434.4165
'    Const AH2i26 As Double = 441.7375
'    Const AH2i27 As Double = 449.0701
'    Const AH2i28 As Double = 456.4128
'    Const AH2i29 As Double = 463.7664
'    Const AH2i30 As Double = 471.1299
'    Const AH2i31 As Double = 478.5032
'    Const AH2i32 As Double = 485.887
'    Const AH2i33 As Double = 493.28
'    Const AH2i34 As Double = 500.6831
'    Const AH2i35 As Double = 508.0955
'    Const AH2i36 As Double = 515.5173
'    Const AH2i37 As Double = 522.949
'    Const AH2i38 As Double = 530.3901
'    Const AH2i39 As Double = 537.8408
'    Const AH2i40 As Double = 545.3008
'    Const AH2i41 As Double = 552.77
'    Const AH2i42 As Double = 560.2493
'    Const AH2i43 As Double = 567.738
'    Const AH2i44 As Double = 575.2361
'    Const AH2i45 As Double = 582.7434
'    Const AH2i46 As Double = 590.2605
'    Const AH2i47 As Double = 597.7874
'    Const AH2i48 As Double = 605.323
'    Const AH2i49 As Double = 612.8684
'    Const AH2i50 As Double = 620.4224
'    Const AH2i51 As Double = 627.9856
'    Const AH2i52 As Double = 635.5574
'    Const AH2i53 As Double = 643.1382
'    Const AH2i54 As Double = 650.7271
'    Const AH2i55 As Double = 658.3237

'    Const ACPH1i1 As Double = 0.2385955
'    Const ACPH1i2 As Double = 0.2386993
'    Const ACPH1i3 As Double = 0.2388183
'    Const ACPH1i4 As Double = 0.2389465
'    Const ACPH1i5 As Double = 0.239093
'    Const ACPH1i6 As Double = 0.2392471
'    Const ACPH1i7 As Double = 0.2394119
'    Const ACPH1i8 As Double = 0.2395935
'    Const ACPH1i9 As Double = 0.2397888
'    Const ACPH1i10 As Double = 0.2399933
'    Const ACPH1i11 As Double = 0.2402053
'    Const ACPH1i12 As Double = 0.2404404
'    Const ACPH1i13 As Double = 0.2406769
'    Const ACPH1i14 As Double = 0.2409347
'    Const ACPH1i15 As Double = 0.2411957
'    Const ACPH1i16 As Double = 0.2414764
'    Const ACPH1i17 As Double = 0.2417617
'    Const ACPH1i18 As Double = 0.2420669
'    Const ACPH1i19 As Double = 0.2423706
'    Const ACPH1i20 As Double = 0.2426971
'    Const ACPH1i21 As Double = 0.2430267
'    Const ACPH1i22 As Double = 0.2433715
'    Const ACPH1i23 As Double = 0.243724
'    Const ACPH1i24 As Double = 0.2440857
'    Const ACPH1i25 As Double = 0.2444625
'    Const ACPH1i26 As Double = 0.2448379
'    Const ACPH1i27 As Double = 0.2452377
'    Const ACPH1i28 As Double = 0.2456314
'    Const ACPH1i29 As Double = 0.246048
'    Const ACPH1i30 As Double = 0.2464645
'    Const ACPH1i31 As Double = 0.2468887
'    Const ACPH1i32 As Double = 0.2473175
'    Const ACPH1i33 As Double = 0.2477616
'    Const ACPH1i34 As Double = 0.2482178
'    Const ACPH1i35 As Double = 0.2486649
'    Const ACPH1i36 As Double = 0.2491287
'    Const ACPH1i37 As Double = 0.2496063
'    Const ACPH1i38 As Double = 0.2500778
'    Const ACPH1i39 As Double = 0.25056
'    Const ACPH1i40 As Double = 0.2510421
'    Const ACPH1i41 As Double = 0.251538
'    Const ACPH1i42 As Double = 0.2520325
'    Const ACPH1i43 As Double = 0.252539
'    Const ACPH1i44 As Double = 0.2530395
'    Const ACPH1i45 As Double = 0.2535522
'    Const ACPH1i46 As Double = 0.2540588
'    Const ACPH1i47 As Double = 0.2545837
'    Const ACPH1i48 As Double = 0.255101
'    Const ACPH1i49 As Double = 0.2556259
'    Const ACPH1i50 As Double = 0.2561523
'    Const ACPH1i51 As Double = 0.2566742
'    Const ACPH1i52 As Double = 0.2572128
'    Const ACPH1i53 As Double = 0.2577301
'    Const ACPH1i54 As Double = 0.2582733
'    Const ACPH1i55 As Double = 0.2588089
'    Const ACPH1i56 As Double = 0.2593399
'    Const ACPH1i57 As Double = 0.2598816
'    Const ACPH1i58 As Double = 0.260408
'    Const ACPH1i59 As Double = 0.260942
'    Const ACPH1i60 As Double = 0.2614853
'    Const ACPH1i61 As Double = 0.2620163
'    Const ACPH1i62 As Double = 0.2625534
'    Const ACPH1i63 As Double = 0.2630829
'    Const ACPH1i64 As Double = 0.2636185
'    Const ACPH1i65 As Double = 0.264151
'    Const ACPH1i66 As Double = 0.2646804
'    Const ACPH1i67 As Double = 0.2651978
'    Const ACPH1i68 As Double = 0.2657303
'    Const ACPH1i69 As Double = 0.2662537
'    Const ACPH1i70 As Double = 0.2667724
'    Const ACPH1i71 As Double = 0.2672912
'    Const ACPH1i72 As Double = 0.2689526
'    Const ACPH1i73 As Double = 0.2702069
'    Const ACPH1i74 As Double = 0.2714373
'    Const ACPH1i75 As Double = 0.2726275
'    Const ACPH1i76 As Double = 0.2738085

'    Const ACPH2i1 As Double = 0.2749707
'    Const ACPH2i2 As Double = 0.2760742
'    Const ACPH2i3 As Double = 0.2771582
'    Const ACPH2i4 As Double = 0.2782031
'    Const ACPH2i5 As Double = 0.279209
'    Const ACPH2i6 As Double = 0.2802051
'    Const ACPH2i7 As Double = 0.2811328
'    Const ACPH2i8 As Double = 0.2820507
'    Const ACPH2i9 As Double = 0.2829199
'    Const ACPH2i10 As Double = 0.2837597
'    Const ACPH2i11 As Double = 0.2845783
'    Const ACPH2i12 As Double = 0.285332
'    Const ACPH2i13 As Double = 0.2860742
'    Const ACPH2i14 As Double = 0.2867773
'    Const ACPH2i15 As Double = 0.2874511
'    Const ACPH2i16 As Double = 0.2881054
'    Const ACPH2i17 As Double = 0.2887207
'    Const ACPH2i18 As Double = 0.2893164
'    Const ACPH2i19 As Double = 0.289873
'    Const ACPH2i20 As Double = 0.2904199
'    Const ACPH2i21 As Double = 0.2909473
'    Const ACPH2i22 As Double = 0.2914453
'    Const ACPH2i23 As Double = 0.2919335
'    Const ACPH2i24 As Double = 0.2924023
'    Const ACPH2i25 As Double = 0.2928418
'    Const ACPH2i26 As Double = 0.2933007
'    Const ACPH2i27 As Double = 0.2937109
'    Const ACPH2i28 As Double = 0.2941406
'    Const ACPH2i29 As Double = 0.294541
'    Const ACPH2i30 As Double = 0.2949316
'    Const ACPH2i31 As Double = 0.2953515
'    Const ACPH2i32 As Double = 0.2957226
'    Const ACPH2i33 As Double = 0.296123
'    Const ACPH2i34 As Double = 0.2964941
'    Const ACPH2i35 As Double = 0.296875
'    Const ACPH2i36 As Double = 0.2972656
'    Const ACPH2i37 As Double = 0.2976465
'    Const ACPH2i38 As Double = 0.2980273
'    Const ACPH2i39 As Double = 0.2983984
'    Const ACPH2i40 As Double = 0.2987695
'    Const ACPH2i41 As Double = 0.2991699
'    Const ACPH2i42 As Double = 0.2995508
'    Const ACPH2i43 As Double = 0.2999219
'    Const ACPH2i44 As Double = 0.300293
'    Const ACPH2i45 As Double = 0.3006836
'    Const ACPH2i46 As Double = 0.3010742
'    Const ACPH2i47 As Double = 0.3014258
'    Const ACPH2i48 As Double = 0.3018164
'    Const ACPH2i49 As Double = 0.3021582
'    Const ACPH2i50 As Double = 0.3025293
'    Const ACPH2i51 As Double = 0.302871
'    Const ACPH2i52 As Double = 0.3032324
'    Const ACPH2i53 As Double = 0.3035547
'    Const ACPH2i54 As Double = 0.3038672
'    Const ACPH2i55 As Double = 0.3041699


'    Const AS1i1 As Double = 1.500759
'    Const AS1i2 As Double = 1.5124
'    Const AS1i3 As Double = 1.523504
'    Const AS1i4 As Double = 1.53412
'    Const AS1i5 As Double = 1.544289
'    Const AS1i6 As Double = 1.554049
'    Const AS1i7 As Double = 1.563433
'    Const AS1i8 As Double = 1.572469
'    Const AS1i9 As Double = 1.581182
'    Const AS1i10 As Double = 1.589597
'    Const AS1i11 As Double = 1.597733
'    Const AS1i12 As Double = 1.60569
'    Const AS1i13 As Double = 1.613243
'    Const AS1i14 As Double = 1.620648
'    Const AS1i15 As Double = 1.627841
'    Const AS1i16 As Double = 1.634832
'    Const AS1i17 As Double = 1.641636
'    Const AS1i18 As Double = 1.648259
'    Const AS1i19 As Double = 1.654715
'    Const AS1i20 As Double = 1.661011
'    Const AS1i21 As Double = 1.667154
'    Const AS1i22 As Double = 1.673156
'    Const AS1i23 As Double = 1.679021
'    Const AS1i24 As Double = 1.684756
'    Const AS1i25 As Double = 1.690368
'    Const AS1i26 As Double = 1.69586
'    Const AS1i27 As Double = 1.701242
'    Const AS1i28 As Double = 1.706516
'    Const AS1i29 As Double = 1.711688
'    Const AS1i30 As Double = 1.716761
'    Const AS1i31 As Double = 1.72174
'    Const AS1i32 As Double = 1.726629
'    Const AS1i33 As Double = 1.731432
'    Const AS1i34 As Double = 1.736151
'    Const AS1i35 As Double = 1.74079
'    Const AS1i36 As Double = 1.745353
'    Const AS1i37 As Double = 1.749843
'    Const AS1i38 As Double = 1.754261
'    Const AS1i39 As Double = 1.758609
'    Const AS1i40 As Double = 1.762892
'    Const AS1i41 As Double = 1.767113
'    Const AS1i42 As Double = 1.77127
'    Const AS1i43 As Double = 1.775368
'    Const AS1i44 As Double = 1.779408
'    Const AS1i45 As Double = 1.783393
'    Const AS1i46 As Double = 1.787325
'    Const AS1i47 As Double = 1.791203
'    Const AS1i48 As Double = 1.795032
'    Const AS1i49 As Double = 1.798812
'    Const AS1i50 As Double = 1.802543
'    Const AS1i51 As Double = 1.806229
'    Const AS1i52 As Double = 1.80987
'    Const AS1i53 As Double = 1.813467
'    Const AS1i54 As Double = 1.817022
'    Const AS1i55 As Double = 1.820537
'    Const AS1i56 As Double = 1.82401
'    Const AS1i57 As Double = 1.827445
'    Const AS1i58 As Double = 1.830842
'    Const AS1i59 As Double = 1.834203
'    Const AS1i60 As Double = 1.837526
'    Const AS1i61 As Double = 1.840816
'    Const AS1i62 As Double = 1.84407
'    Const AS1i63 As Double = 1.847293
'    Const AS1i64 As Double = 1.850481
'    Const AS1i65 As Double = 1.853638
'    Const AS1i66 As Double = 1.856764
'    Const AS1i67 As Double = 1.859859
'    Const AS1i68 As Double = 1.862926
'    Const AS1i69 As Double = 1.865963
'    Const AS1i70 As Double = 1.868971
'    Const AS1i71 As Double = 1.871952
'    Const AS1i72 As Double = 1.879286
'    Const AS1i73 As Double = 1.886458
'    Const AS1i74 As Double = 1.893477
'    Const AS1i75 As Double = 1.90035
'    Const AS1i76 As Double = 1.907081

'    Const AS2i1 As Double = 1.913679
'    Const AS2i2 As Double = 1.92015
'    Const AS2i3 As Double = 1.926497
'    Const AS2i4 As Double = 1.932725
'    Const AS2i5 As Double = 1.938839
'    Const AS2i6 As Double = 1.944844
'    Const AS2i7 As Double = 1.950743
'    Const AS2i8 As Double = 1.95654
'    Const AS2i9 As Double = 1.962238
'    Const AS2i10 As Double = 1.967841
'    Const AS2i11 As Double = 1.973351
'    Const AS2i12 As Double = 1.978771
'    Const AS2i13 As Double = 1.984104
'    Const AS2i14 As Double = 1.989354
'    Const AS2i15 As Double = 1.994521
'    Const AS2i16 As Double = 1.999609
'    Const AS2i17 As Double = 2.004619
'    Const AS2i18 As Double = 2.009555
'    Const AS2i19 As Double = 2.014418
'    Const AS2i20 As Double = 2.019209
'    Const AS2i21 As Double = 2.023932
'    Const AS2i22 As Double = 2.028586
'    Const AS2i23 As Double = 2.033176
'    Const AS2i24 As Double = 2.037703
'    Const AS2i25 As Double = 2.042167
'    Const AS2i26 As Double = 2.04657
'    Const AS2i27 As Double = 2.050916
'    Const AS2i28 As Double = 2.055203
'    Const AS2i29 As Double = 2.059436
'    Const AS2i30 As Double = 2.063614
'    Const AS2i31 As Double = 2.067739
'    Const AS2i32 As Double = 2.071813
'    Const AS2i33 As Double = 2.075836
'    Const AS2i34 As Double = 2.079811
'    Const AS2i35 As Double = 2.083738
'    Const AS2i36 As Double = 2.087619
'    Const AS2i37 As Double = 2.091455
'    Const AS2i38 As Double = 2.095246
'    Const AS2i39 As Double = 2.098995
'    Const AS2i40 As Double = 2.102702
'    Const AS2i41 As Double = 2.106368
'    Const AS2i42 As Double = 2.109995
'    Const AS2i43 As Double = 2.113582
'    Const AS2i44 As Double = 2.117131
'    Const AS2i45 As Double = 2.120644
'    Const AS2i46 As Double = 2.12412
'    Const AS2i47 As Double = 2.12756
'    Const AS2i48 As Double = 2.130967
'    Const AS2i49 As Double = 2.134338
'    Const AS2i50 As Double = 2.137678
'    Const AS2i51 As Double = 2.140984
'    Const AS2i52 As Double = 2.144258
'    Const AS2i53 As Double = 2.147501
'    Const AS2i54 As Double = 2.150714
'    Const AS2i55 As Double = 2.153895


'    Const ACPS1i1 As Double = 0.001164055
'    Const ACPS1i2 As Double = 0.001110458
'    Const ACPS1i3 As Double = 0.001061535
'    Const ACPS1i4 As Double = 0.001016903
'    Const ACPS1i5 As Double = 0.0009760857
'    Const ACPS1i6 As Double = 0.0009383201
'    Const ACPS1i7 As Double = 0.0009036064
'    Const ACPS1i8 As Double = 0.0008712767
'    Const ACPS1i9 As Double = 0.0008415221
'    Const ACPS1i10 As Double = 0.0008135794
'    Const ACPS1i11 As Double = 0.0007876395
'    Const ACPS1i12 As Double = 0.0007634163
'    Const ACPS1i13 As Double = 0.0007405281
'    Const ACPS1i14 As Double = 0.0007192611
'    Const ACPS1i15 As Double = 0.0006991385
'    Const ACPS1i16 As Double = 0.0006803512
'    Const ACPS1i17 As Double = 0.0006623268
'    Const ACPS1i18 As Double = 0.0006455421
'    Const ACPS1i19 As Double = 0.0006296157
'    Const ACPS1i20 As Double = 0.0006143569
'    Const ACPS1i21 As Double = 0.0006001471
'    Const ACPS1i22 As Double = 0.0005865097
'    Const ACPS1i23 As Double = 0.0005735396
'    Const ACPS1i24 As Double = 0.0005611419
'    Const ACPS1i25 As Double = 0.0005492209
'    Const ACPS1i26 As Double = 0.0005382537
'    Const ACPS1i27 As Double = 0.0005273819
'    Const ACPS1i28 As Double = 0.0005171774
'    Const ACPS1i29 As Double = 0.0005072593
'    Const ACPS1i30 As Double = 0.0004979132
'    Const ACPS1i31 As Double = 0.0004889488
'    Const ACPS1i32 As Double = 0.0004802702
'    Const ACPS1i33 As Double = 0.0004718779
'    Const ACPS1i34 As Double = 0.0004639626
'    Const ACPS1i35 As Double = 0.0004562377
'    Const ACPS1i36 As Double = 0.0004489897
'    Const ACPS1i37 As Double = 0.0004418371
'    Const ACPS1i38 As Double = 0.00043478
'    Const ACPS1i39 As Double = 0.000428295
'    Const ACPS1i40 As Double = 0.0004220961
'    Const ACPS1i41 As Double = 0.0004157065
'    Const ACPS1i42 As Double = 0.0004097938
'    Const ACPS1i43 As Double = 0.0004040718
'    Const ACPS1i44 As Double = 0.0003984449
'    Const ACPS1i45 As Double = 0.0003931997
'    Const ACPS1i46 As Double = 0.0003878595
'    Const ACPS1i47 As Double = 0.0003828048
'    Const ACPS1i48 As Double = 0.0003780364
'    Const ACPS1i49 As Double = 0.0003730773
'    Const ACPS1i50 As Double = 0.0003685951
'    Const ACPS1i51 As Double = 0.0003641127
'    Const ACPS1i52 As Double = 0.0003597259
'    Const ACPS1i53 As Double = 0.0003555296
'    Const ACPS1i54 As Double = 0.000351429
'    Const ACPS1i55 As Double = 0.0003473281
'    Const ACPS1i56 As Double = 0.0003435134
'    Const ACPS1i57 As Double = 0.0003396987
'    Const ACPS1i58 As Double = 0.0003360747
'    Const ACPS1i59 As Double = 0.0003323555
'    Const ACPS1i60 As Double = 0.0003289222
'    Const ACPS1i61 As Double = 0.0003254889
'    Const ACPS1i62 As Double = 0.0003222465
'    Const ACPS1i63 As Double = 0.0003188131
'    Const ACPS1i64 As Double = 0.0003156662
'    Const ACPS1i65 As Double = 0.0003126143
'    Const ACPS1i66 As Double = 0.0003095625
'    Const ACPS1i67 As Double = 0.0003067015
'    Const ACPS1i68 As Double = 0.0003036498
'    Const ACPS1i69 As Double = 0.0003007888
'    Const ACPS1i70 As Double = 0.0002981185
'    Const ACPS1i71 As Double = 0.0002953529
'    Const ACPS1i72 As Double = 0.0002869032
'    Const ACPS1i73 As Double = 0.0002807616
'    Const ACPS1i74 As Double = 0.0002748871
'    Const ACPS1i75 As Double = 0.0002692412
'    Const ACPS1i76 As Double = 0.0002639387

'    Const ACPS2i1 As Double = 0.0002588271
'    Const ACPS2i2 As Double = 0.0002538681
'    Const ACPS2i3 As Double = 0.0002491379
'    Const ACPS2i4 As Double = 0.0002445602
'    Const ACPS2i5 As Double = 0.0002402115
'    Const ACPS2i6 As Double = 0.000235939
'    Const ACPS2i7 As Double = 0.0002318954
'    Const ACPS2i8 As Double = 0.0002279282
'    Const ACPS2i9 As Double = 0.0002241135
'    Const ACPS2i10 As Double = 0.000220375
'    Const ACPS2i11 As Double = 0.0002168274
'    Const ACPS2i12 As Double = 0.0002133179
'    Const ACPS2i13 As Double = 0.0002099991
'    Const ACPS2i14 As Double = 0.0002066803
'    Const ACPS2i15 As Double = 0.0002035141
'    Const ACPS2i16 As Double = 0.000200386
'    Const ACPS2i17 As Double = 0.0001974487
'    Const ACPS2i18 As Double = 0.0001945114
'    Const ACPS2i19 As Double = 0.0001916504
'    Const ACPS2i20 As Double = 0.0001889038
'    Const ACPS2i21 As Double = 0.0001861954
'    Const ACPS2i22 As Double = 0.0001836014
'    Const ACPS2i23 As Double = 0.0001810455
'    Const ACPS2i24 As Double = 0.000178566
'    Const ACPS2i25 As Double = 0.0001761246
'    Const ACPS2i26 As Double = 0.0001738358
'    Const ACPS2i27 As Double = 0.0001715088
'    Const ACPS2i28 As Double = 0.0001692963
'    Const ACPS2i29 As Double = 0.0001671219
'    Const ACPS2i30 As Double = 0.0001650238
'    Const ACPS2i31 As Double = 0.0001629257
'    Const ACPS2i32 As Double = 0.0001609421
'    Const ACPS2i33 As Double = 0.0001589966
'    Const ACPS2i34 As Double = 0.0001570892
'    Const ACPS2i35 As Double = 0.00015522
'    Const ACPS2i36 As Double = 0.0001534271
'    Const ACPS2i37 As Double = 0.0001516724
'    Const ACPS2i38 As Double = 0.0001499557
'    Const ACPS2i39 As Double = 0.0001482773
'    Const ACPS2i40 As Double = 0.000146637
'    Const ACPS2i41 As Double = 0.0001450729
'    Const ACPS2i42 As Double = 0.0001434708
'    Const ACPS2i43 As Double = 0.000141983
'    Const ACPS2i44 As Double = 0.0001404953
'    Const ACPS2i45 As Double = 0.0001390457
'    Const ACPS2i46 As Double = 0.0001375961
'    Const ACPS2i47 As Double = 0.0001362991
'    Const ACPS2i48 As Double = 0.0001348495
'    Const ACPS2i49 As Double = 0.0001335907
'    Const ACPS2i50 As Double = 0.0001322174
'    Const ACPS2i51 As Double = 0.0001309967
'    Const ACPS2i52 As Double = 0.0001296997
'    Const ACPS2i53 As Double = 0.0001285171
'    Const ACPS2i54 As Double = 0.0001272583
'    Const ACPS2i55 As Double = 0.0001260757

'    Const FH1i1 As Double = 63.74435
'    Const FH1i2 As Double = 68.25485
'    Const FH1i3 As Double = 72.81036
'    Const FH1i4 As Double = 77.40707
'    Const FH1i5 As Double = 82.04727
'    Const FH1i6 As Double = 86.72943
'    Const FH1i7 As Double = 91.45355
'    Const FH1i8 As Double = 96.22269
'    Const FH1i9 As Double = 101.0315
'    Const FH1i10 As Double = 105.883
'    Const FH1i11 As Double = 110.778
'    Const FH1i12 As Double = 115.7181
'    Const FH1i13 As Double = 120.697
'    Const FH1i14 As Double = 125.7248
'    Const FH1i15 As Double = 130.7877
'    Const FH1i16 As Double = 135.8955
'    Const FH1i17 As Double = 141.0461
'    Const FH1i18 As Double = 146.2395
'    Const FH1i19 As Double = 151.4717
'    Const FH1i20 As Double = 156.7467
'    Const FH1i21 As Double = 162.0644
'    Const FH1i22 As Double = 167.4255
'    Const FH1i23 As Double = 172.8256
'    Const FH1i24 As Double = 178.2692
'    Const FH1i25 As Double = 183.7532
'    Const FH1i26 As Double = 189.2769
'    Const FH1i27 As Double = 194.8479
'    Const FH1i28 As Double = 200.4517
'    Const FH1i29 As Double = 206.102
'    Const FH1i30 As Double = 211.7882
'    Const FH1i31 As Double = 217.5209
'    Const FH1i32 As Double = 223.2903
'    Const FH1i33 As Double = 229.1016
'    Const FH1i34 As Double = 234.9533
'    Const FH1i35 As Double = 240.8432
'    Const FH1i36 As Double = 246.7743
'    Const FH1i37 As Double = 252.7458
'    Const FH1i38 As Double = 258.7561
'    Const FH1i39 As Double = 264.804
'    Const FH1i40 As Double = 270.8931
'    Const FH1i41 As Double = 277.0217
'    Const FH1i42 As Double = 283.187
'    Const FH1i43 As Double = 289.3921
'    Const FH1i44 As Double = 295.635
'    Const FH1i45 As Double = 301.9165
'    Const FH1i46 As Double = 308.2336
'    Const FH1i47 As Double = 314.5979
'    Const FH1i48 As Double = 320.989
'    Const FH1i49 As Double = 327.426
'    Const FH1i50 As Double = 333.8943
'    Const FH1i51 As Double = 340.4006
'    Const FH1i52 As Double = 346.9473
'    Const FH1i53 As Double = 353.5254
'    Const FH1i54 As Double = 360.1484
'    Const FH1i55 As Double = 366.8013
'    Const FH1i56 As Double = 373.49
'    Const FH1i57 As Double = 380.2175
'    Const FH1i58 As Double = 386.9758
'    Const FH1i59 As Double = 393.7766
'    Const FH1i60 As Double = 400.6125
'    Const FH1i61 As Double = 407.4768
'    Const FH1i62 As Double = 414.3828
'    Const FH1i63 As Double = 421.3179
'    Const FH1i64 As Double = 428.2959
'    Const FH1i65 As Double = 435.2981
'    Const FH1i66 As Double = 442.3408
'    Const FH1i67 As Double = 449.4131
'    Const FH1i68 As Double = 456.5247
'    Const FH1i69 As Double = 463.6687
'    Const FH1i70 As Double = 470.845
'    Const FH1i71 As Double = 478.0518
'    Const FH1i72 As Double = 496.218
'    Const FH1i73 As Double = 514.5881
'    Const FH1i74 As Double = 533.1558
'    Const FH1i75 As Double = 551.905
'    Const FH1i76 As Double = 570.874

'    Const FH2i1 As Double = 590.0269
'    Const FH2i2 As Double = 609.3506
'    Const FH2i3 As Double = 628.8574
'    Const FH2i4 As Double = 648.5596
'    Const FH2i5 As Double = 668.4326
'    Const FH2i6 As Double = 688.4888
'    Const FH2i7 As Double = 708.6914
'    Const FH2i8 As Double = 729.0771
'    Const FH2i9 As Double = 749.6338
'    Const FH2i10 As Double = 770.3369
'    Const FH2i11 As Double = 791.1865
'    Const FH2i12 As Double = 812.207
'    Const FH2i13 As Double = 833.374
'    Const FH2i14 As Double = 854.6753
'    Const FH2i15 As Double = 876.1475
'    Const FH2i16 As Double = 897.7417
'    Const FH2i17 As Double = 919.4702
'    Const FH2i18 As Double = 941.333
'    Const FH2i19 As Double = 963.3301
'    Const FH2i20 As Double = 985.4492
'    Const FH2i21 As Double = 1007.69
'    Const FH2i22 As Double = 1030.066
'    Const FH2i23 As Double = 1052.551
'    Const FH2i24 As Double = 1075.146
'    Const FH2i25 As Double = 1097.852
'    Const FH2i26 As Double = 1120.667
'    Const FH2i27 As Double = 1143.579
'    Const FH2i28 As Double = 1166.589
'    Const FH2i29 As Double = 1189.709
'    Const FH2i30 As Double = 1212.915
'    Const FH2i31 As Double = 1236.218
'    Const FH2i32 As Double = 1259.595
'    Const FH2i33 As Double = 1283.069
'    Const FH2i34 As Double = 1306.604
'    Const FH2i35 As Double = 1330.225
'    Const FH2i36 As Double = 1353.931
'    Const FH2i37 As Double = 1377.71
'    Const FH2i38 As Double = 1401.538
'    Const FH2i39 As Double = 1425.464
'    Const FH2i40 As Double = 1449.426
'    Const FH2i41 As Double = 1473.474
'    Const FH2i42 As Double = 1497.583
'    Const FH2i43 As Double = 1521.716
'    Const FH2i44 As Double = 1545.911
'    Const FH2i45 As Double = 1570.166
'    Const FH2i46 As Double = 1594.446
'    Const FH2i47 As Double = 1618.811
'    Const FH2i48 As Double = 1643.188
'    Const FH2i49 As Double = 1667.615
'    Const FH2i50 As Double = 1692.078
'    Const FH2i51 As Double = 1716.589
'    Const FH2i52 As Double = 1741.138
'    Const FH2i53 As Double = 1765.698
'    Const FH2i54 As Double = 1790.308
'    Const FH2i55 As Double = 1814.978

'    Const FCPH1i1 As Double = 0.4510498
'    Const FCPH1i2 As Double = 0.4555511
'    Const FCPH1i3 As Double = 0.459671
'    Const FCPH1i4 As Double = 0.4640198
'    Const FCPH1i5 As Double = 0.4682159
'    Const FCPH1i6 As Double = 0.4724121
'    Const FCPH1i7 As Double = 0.4769135
'    Const FCPH1i8 As Double = 0.4808807
'    Const FCPH1i9 As Double = 0.4851532
'    Const FCPH1i10 As Double = 0.489502
'    Const FCPH1i11 As Double = 0.4940033
'    Const FCPH1i12 As Double = 0.4978943
'    Const FCPH1i13 As Double = 0.5027771
'    Const FCPH1i14 As Double = 0.5062866
'    Const FCPH1i15 As Double = 0.510788
'    Const FCPH1i16 As Double = 0.5150604
'    Const FCPH1i17 As Double = 0.5193329
'    Const FCPH1i18 As Double = 0.5232239
'    Const FCPH1i19 As Double = 0.5274963
'    Const FCPH1i20 As Double = 0.5317688
'    Const FCPH1i21 As Double = 0.5361176
'    Const FCPH1i22 As Double = 0.5400085
'    Const FCPH1i23 As Double = 0.5443573
'    Const FCPH1i24 As Double = 0.5484009
'    Const FCPH1i25 As Double = 0.5523682
'    Const FCPH1i26 As Double = 0.5570984
'    Const FCPH1i27 As Double = 0.560379
'    Const FCPH1i28 As Double = 0.565033
'    Const FCPH1i29 As Double = 0.5686188
'    Const FCPH1i30 As Double = 0.5732727
'    Const FCPH1i31 As Double = 0.5769348
'    Const FCPH1i32 As Double = 0.581131
'    Const FCPH1i33 As Double = 0.5851746
'    Const FCPH1i34 As Double = 0.5889893
'    Const FCPH1i35 As Double = 0.5931091
'    Const FCPH1i36 As Double = 0.5971527
'    Const FCPH1i37 As Double = 0.6010284
'    Const FCPH1i38 As Double = 0.6047851
'    Const FCPH1i39 As Double = 0.6089111
'    Const FCPH1i40 As Double = 0.6128662
'    Const FCPH1i41 As Double = 0.6165283
'    Const FCPH1i42 As Double = 0.6205078
'    Const FCPH1i43 As Double = 0.624292
'    Const FCPH1i44 As Double = 0.6281494
'    Const FCPH1i45 As Double = 0.6317139
'    Const FCPH1i46 As Double = 0.6364257
'    Const FCPH1i47 As Double = 0.6391113
'    Const FCPH1i48 As Double = 0.6437011
'    Const FCPH1i49 As Double = 0.6468261
'    Const FCPH1i50 As Double = 0.6506348
'    Const FCPH1i51 As Double = 0.6546631
'    Const FCPH1i52 As Double = 0.6578125
'    Const FCPH1i53 As Double = 0.6623046
'    Const FCPH1i54 As Double = 0.6652832
'    Const FCPH1i55 As Double = 0.6688721
'    Const FCPH1i56 As Double = 0.6727539
'    Const FCPH1i57 As Double = 0.6758301
'    Const FCPH1i58 As Double = 0.6800781
'    Const FCPH1i59 As Double = 0.6835938
'    Const FCPH1i60 As Double = 0.6864257
'    Const FCPH1i61 As Double = 0.6906006
'    Const FCPH1i62 As Double = 0.6935058
'    Const FCPH1i63 As Double = 0.6978027
'    Const FCPH1i64 As Double = 0.7002197
'    Const FCPH1i65 As Double = 0.7042724
'    Const FCPH1i66 As Double = 0.7072265
'    Const FCPH1i67 As Double = 0.7111572
'    Const FCPH1i68 As Double = 0.7144043
'    Const FCPH1i69 As Double = 0.7176269
'    Const FCPH1i70 As Double = 0.7206787
'    Const FCPH1i71 As Double = 0.7240967
'    Const FCPH1i72 As Double = 0.7348046
'    Const FCPH1i73 As Double = 0.742705
'    Const FCPH1i74 As Double = 0.7499707
'    Const FCPH1i75 As Double = 0.7587597
'    Const FCPH1i76 As Double = 0.7661133

'    Const FCPH2i1 As Double = 0.7729492
'    Const FCPH2i2 As Double = 0.7802734
'    Const FCPH2i3 As Double = 0.7880859
'    Const FCPH2i4 As Double = 0.7949219
'    Const FCPH2i5 As Double = 0.8022461
'    Const FCPH2i6 As Double = 0.8081055
'    Const FCPH2i7 As Double = 0.8154297
'    Const FCPH2i8 As Double = 0.8222656
'    Const FCPH2i9 As Double = 0.828125
'    Const FCPH2i10 As Double = 0.8339844
'    Const FCPH2i11 As Double = 0.8408203
'    Const FCPH2i12 As Double = 0.8466797
'    Const FCPH2i13 As Double = 0.8520508
'    Const FCPH2i14 As Double = 0.8588867
'    Const FCPH2i15 As Double = 0.8637695
'    Const FCPH2i16 As Double = 0.8691406
'    Const FCPH2i17 As Double = 0.8745117
'    Const FCPH2i18 As Double = 0.8798828
'    Const FCPH2i19 As Double = 0.8847656
'    Const FCPH2i20 As Double = 0.8896484
'    Const FCPH2i21 As Double = 0.8950195
'    Const FCPH2i22 As Double = 0.8994141
'    Const FCPH2i23 As Double = 0.9038086
'    Const FCPH2i24 As Double = 0.9082031
'    Const FCPH2i25 As Double = 0.9125977
'    Const FCPH2i26 As Double = 0.9165039
'    Const FCPH2i27 As Double = 0.9204102
'    Const FCPH2i28 As Double = 0.9248047
'    Const FCPH2i29 As Double = 0.9282227
'    Const FCPH2i30 As Double = 0.9321289
'    Const FCPH2i31 As Double = 0.9350586
'    Const FCPH2i32 As Double = 0.9389648
'    Const FCPH2i33 As Double = 0.9414063
'    Const FCPH2i34 As Double = 0.9448242
'    Const FCPH2i35 As Double = 0.9482422
'    Const FCPH2i36 As Double = 0.9511719
'    Const FCPH2i37 As Double = 0.953125
'    Const FCPH2i38 As Double = 0.9570313
'    Const FCPH2i39 As Double = 0.9584961
'    Const FCPH2i40 As Double = 0.9619141
'    Const FCPH2i41 As Double = 0.9643555
'    Const FCPH2i42 As Double = 0.965332
'    Const FCPH2i43 As Double = 0.9677734
'    Const FCPH2i44 As Double = 0.9702148
'    Const FCPH2i45 As Double = 0.9711914
'    Const FCPH2i46 As Double = 0.9746094
'    Const FCPH2i47 As Double = 0.9750977
'    Const FCPH2i48 As Double = 0.9770508
'    Const FCPH2i49 As Double = 0.9785156
'    Const FCPH2i50 As Double = 0.9804688
'    Const FCPH2i51 As Double = 0.9819336
'    Const FCPH2i52 As Double = 0.9824219
'    Const FCPH2i53 As Double = 0.984375
'    Const FCPH2i54 As Double = 0.9868164
'    Const FCPH2i55 As Double = 0.9868144

'    Const FS1i1 As Double = 1.494265
'    Const FS1i2 As Double = 1.516294
'    Const FS1i3 As Double = 1.537466
'    Const FS1i4 As Double = 1.557922
'    Const FS1i5 As Double = 1.577711
'    Const FS1i6 As Double = 1.596785
'    Const FS1i7 As Double = 1.615334
'    Const FS1i8 As Double = 1.63331
'    Const FS1i9 As Double = 1.65081
'    Const FS1i10 As Double = 1.667786
'    Const FS1i11 As Double = 1.684427
'    Const FS1i12 As Double = 1.700592
'    Const FS1i13 As Double = 1.716328
'    Const FS1i14 As Double = 1.731873
'    Const FS1i15 As Double = 1.746988
'    Const FS1i16 As Double = 1.761818
'    Const FS1i17 As Double = 1.776266
'    Const FS1i18 As Double = 1.790571
'    Const FS1i19 As Double = 1.804447
'    Const FS1i20 As Double = 1.81818
'    Const FS1i21 As Double = 1.831675
'    Const FS1i22 As Double = 1.844931
'    Const FS1i23 As Double = 1.857901
'    Const FS1i24 As Double = 1.87068
'    Const FS1i25 As Double = 1.883316
'    Const FS1i26 As Double = 1.895809
'    Const FS1i27 As Double = 1.907969
'    Const FS1i28 As Double = 1.920033
'    Const FS1i29 As Double = 1.931906
'    Const FS1i30 As Double = 1.943684
'    Const FS1i31 As Double = 1.955223
'    Const FS1i32 As Double = 1.966619
'    Const FS1i33 As Double = 1.977921
'    Const FS1i34 As Double = 1.989079
'    Const FS1i35 As Double = 2.000093
'    Const FS1i36 As Double = 2.011061
'    Const FS1i37 As Double = 2.021694
'    Const FS1i38 As Double = 2.03228
'    Const FS1i39 As Double = 2.042913
'    Const FS1i40 As Double = 2.053356
'    Const FS1i41 As Double = 2.063608
'    Const FS1i42 As Double = 2.073765
'    Const FS1i43 As Double = 2.083921
'    Const FS1i44 As Double = 2.093887
'    Const FS1i45 As Double = 2.103806
'    Const FS1i46 As Double = 2.113581
'    Const FS1i47 As Double = 2.12326
'    Const FS1i48 As Double = 2.132893
'    Const FS1i49 As Double = 2.142382
'    Const FS1i50 As Double = 2.151871
'    Const FS1i51 As Double = 2.161264
'    Const FS1i52 As Double = 2.170515
'    Const FS1i53 As Double = 2.17967
'    Const FS1i54 As Double = 2.188778
'    Const FS1i55 As Double = 2.197838
'    Const FS1i56 As Double = 2.20685
'    Const FS1i57 As Double = 2.215815
'    Const FS1i58 As Double = 2.224636
'    Const FS1i59 As Double = 2.233362
'    Const FS1i60 As Double = 2.242088
'    Const FS1i61 As Double = 2.250767
'    Const FS1i62 As Double = 2.259302
'    Const FS1i63 As Double = 2.26779
'    Const FS1i64 As Double = 2.276278
'    Const FS1i65 As Double = 2.28467
'    Const FS1i66 As Double = 2.293015
'    Const FS1i67 As Double = 2.301311
'    Const FS1i68 As Double = 2.309465
'    Const FS1i69 As Double = 2.317667
'    Const FS1i70 As Double = 2.325821
'    Const FS1i71 As Double = 2.333832
'    Const FS1i72 As Double = 2.353716
'    Const FS1i73 As Double = 2.373362
'    Const FS1i74 As Double = 2.392626
'    Const FS1i75 As Double = 2.411604
'    Const FS1i76 As Double = 2.430344

'    Const FS2i1 As Double = 2.448845
'    Const FS2i2 As Double = 2.466965
'    Const FS2i3 As Double = 2.484941
'    Const FS2i4 As Double = 2.502632
'    Const FS2i5 As Double = 2.52018
'    Const FS2i6 As Double = 2.537394
'    Const FS2i7 As Double = 2.554417
'    Const FS2i8 As Double = 2.571201
'    Const FS2i9 As Double = 2.587795
'    Const FS2i10 As Double = 2.604151
'    Const FS2i11 As Double = 2.620459
'    Const FS2i12 As Double = 2.636385
'    Const FS2i13 As Double = 2.652264
'    Const FS2i14 As Double = 2.667856
'    Const FS2i15 As Double = 2.683353
'    Const FS2i16 As Double = 2.69866
'    Const FS2i17 As Double = 2.713823
'    Const FS2i18 As Double = 2.728748
'    Const FS2i19 As Double = 2.74353
'    Const FS2i20 As Double = 2.758121
'    Const FS2i21 As Double = 2.77257
'    Const FS2i22 As Double = 2.78697
'    Const FS2i23 As Double = 2.801037
'    Const FS2i24 As Double = 2.815104
'    Const FS2i25 As Double = 2.828932
'    Const FS2i26 As Double = 2.842712
'    Const FS2i27 As Double = 2.856255
'    Const FS2i28 As Double = 2.869701
'    Const FS2i29 As Double = 2.883005
'    Const FS2i30 As Double = 2.896166
'    Const FS2i31 As Double = 2.909136
'    Const FS2i32 As Double = 2.922106
'    Const FS2i33 As Double = 2.934885
'    Const FS2i34 As Double = 2.947521
'    Const FS2i35 As Double = 2.960014
'    Const FS2i36 As Double = 2.97246
'    Const FS2i37 As Double = 2.984715
'    Const FS2i38 As Double = 2.996874
'    Const FS2i39 As Double = 3.00889
'    Const FS2i40 As Double = 3.020811
'    Const FS2i41 As Double = 3.032541
'    Const FS2i42 As Double = 3.044176
'    Const FS2i43 As Double = 3.055859
'    Const FS2i44 As Double = 3.067303
'    Const FS2i45 As Double = 3.078651
'    Const FS2i46 As Double = 3.089905
'    Const FS2i47 As Double = 3.101015
'    Const FS2i48 As Double = 3.111982
'    Const FS2i49 As Double = 3.12295
'    Const FS2i50 As Double = 3.133678
'    Const FS2i51 As Double = 3.144503
'    Const FS2i52 As Double = 3.155041
'    Const FS2i53 As Double = 3.165579
'    Const FS2i54 As Double = 3.176022
'    Const FS2i55 As Double = 3.186321

'    Const FCPS1i1 As Double = 0.002202988
'    Const FCPS1i2 As Double = 0.002117157
'    Const FCPS1i3 As Double = 0.002045631
'    Const FCPS1i4 As Double = 0.001978874
'    Const FCPS1i5 As Double = 0.001907349
'    Const FCPS1i6 As Double = 0.001854897
'    Const FCPS1i7 As Double = 0.001797676
'    Const FCPS1i8 As Double = 0.001749992
'    Const FCPS1i9 As Double = 0.00169754
'    Const FCPS1i10 As Double = 0.001664162
'    Const FCPS1i11 As Double = 0.001616478
'    Const FCPS1i12 As Double = 0.001573563
'    Const FCPS1i13 As Double = 0.001554489
'    Const FCPS1i14 As Double = 0.001511574
'    Const FCPS1i15 As Double = 0.001482964
'    Const FCPS1i16 As Double = 0.001444817
'    Const FCPS1i17 As Double = 0.001430511
'    Const FCPS1i18 As Double = 0.001387596
'    Const FCPS1i19 As Double = 0.001373291
'    Const FCPS1i20 As Double = 0.001349449
'    Const FCPS1i21 As Double = 0.001325607
'    Const FCPS1i22 As Double = 0.001296997
'    Const FCPS1i23 As Double = 0.001277924
'    Const FCPS1i24 As Double = 0.001263618
'    Const FCPS1i25 As Double = 0.001249313
'    Const FCPS1i26 As Double = 0.001215935
'    Const FCPS1i27 As Double = 0.001206398
'    Const FCPS1i28 As Double = 0.001187325
'    Const FCPS1i29 As Double = 0.001177788
'    Const FCPS1i30 As Double = 0.001153946
'    Const FCPS1i31 As Double = 0.001139641
'    Const FCPS1i32 As Double = 0.001130104
'    Const FCPS1i33 As Double = 0.001115799
'    Const FCPS1i34 As Double = 0.001101494
'    Const FCPS1i35 As Double = 0.001096725
'    Const FCPS1i36 As Double = 0.001063347
'    Const FCPS1i37 As Double = 0.001058578
'    Const FCPS1i38 As Double = 0.001063347
'    Const FCPS1i39 As Double = 0.001044273
'    Const FCPS1i40 As Double = 0.0010252
'    Const FCPS1i41 As Double = 0.001015663
'    Const FCPS1i42 As Double = 0.001015663
'    Const FCPS1i43 As Double = 0.0009965897
'    Const FCPS1i44 As Double = 0.0009918213
'    Const FCPS1i45 As Double = 0.0009775162
'    Const FCPS1i46 As Double = 0.0009679794
'    Const FCPS1i47 As Double = 0.0009632111
'    Const FCPS1i48 As Double = 0.0009489059
'    Const FCPS1i49 As Double = 0.0009489059
'    Const FCPS1i50 As Double = 0.0009393692
'    Const FCPS1i51 As Double = 0.0009250641
'    Const FCPS1i52 As Double = 0.0009155273
'    Const FCPS1i53 As Double = 0.000910759
'    Const FCPS1i54 As Double = 0.0009059906
'    Const FCPS1i55 As Double = 0.0009012222
'    Const FCPS1i56 As Double = 0.0008964539
'    Const FCPS1i57 As Double = 0.0008821487
'    Const FCPS1i58 As Double = 0.000872612
'    Const FCPS1i59 As Double = 0.000872612
'    Const FCPS1i60 As Double = 0.0008678436
'    Const FCPS1i61 As Double = 0.0008535385
'    Const FCPS1i62 As Double = 0.0008487701
'    Const FCPS1i63 As Double = 0.0008487701
'    Const FCPS1i64 As Double = 0.0008392334
'    Const FCPS1i65 As Double = 0.000834465
'    Const FCPS1i66 As Double = 0.0008296967
'    Const FCPS1i67 As Double = 0.0008153915
'    Const FCPS1i68 As Double = 0.0008201599
'    Const FCPS1i69 As Double = 0.0008153915
'    Const FCPS1i70 As Double = 0.0008010864
'    Const FCPS1i71 As Double = 0.0008010864
'    Const FCPS1i72 As Double = 0.0007858276
'    Const FCPS1i73 As Double = 0.0007705688
'    Const FCPS1i74 As Double = 0.0007591248
'    Const FCPS1i75 As Double = 0.000749588
'    Const FCPS1i76 As Double = 0.0007400513

'    Const FCPS2i1 As Double = 0.0007247925
'    Const FCPS2i2 As Double = 0.0007190704
'    Const FCPS2i3 As Double = 0.0007076263
'    Const FCPS2i4 As Double = 0.0007019043
'    Const FCPS2i5 As Double = 0.0006885529
'    Const FCPS2i6 As Double = 0.0006809235
'    Const FCPS2i7 As Double = 0.0006713867
'    Const FCPS2i8 As Double = 0.0006637573
'    Const FCPS2i9 As Double = 0.0006542206
'    Const FCPS2i10 As Double = 0.0006523132
'    Const FCPS2i11 As Double = 0.0006370544
'    Const FCPS2i12 As Double = 0.0006351471
'    Const FCPS2i13 As Double = 0.000623703
'    Const FCPS2i14 As Double = 0.0006198883
'    Const FCPS2i15 As Double = 0.0006122589
'    Const FCPS2i16 As Double = 0.0006065369
'    Const FCPS2i17 As Double = 0.0005970001
'    Const FCPS2i18 As Double = 0.0005912781
'    Const FCPS2i19 As Double = 0.0005836487
'    Const FCPS2i20 As Double = 0.0005779266
'    Const FCPS2i21 As Double = 0.0005760193
'    Const FCPS2i22 As Double = 0.0005626678
'    Const FCPS2i23 As Double = 0.0005626678
'    Const FCPS2i24 As Double = 0.0005531311
'    Const FCPS2i25 As Double = 0.0005512238
'    Const FCPS2i26 As Double = 0.000541687
'    Const FCPS2i27 As Double = 0.0005378723
'    Const FCPS2i28 As Double = 0.0005321503
'    Const FCPS2i29 As Double = 0.0005264282
'    Const FCPS2i30 As Double = 0.0005187988
'    Const FCPS2i31 As Double = 0.0005187988
'    Const FCPS2i32 As Double = 0.0005111694
'    Const FCPS2i33 As Double = 0.0005054474
'    Const FCPS2i34 As Double = 0.0004997253
'    Const FCPS2i35 As Double = 0.000497818
'    Const FCPS2i36 As Double = 0.0004901886
'    Const FCPS2i37 As Double = 0.0004863739
'    Const FCPS2i38 As Double = 0.0004806519
'    Const FCPS2i39 As Double = 0.0004768372
'    Const FCPS2i40 As Double = 0.0004692078
'    Const FCPS2i41 As Double = 0.0004653931
'    Const FCPS2i42 As Double = 0.0004673004
'    Const FCPS2i43 As Double = 0.0004577637
'    Const FCPS2i44 As Double = 0.000453949
'    Const FCPS2i45 As Double = 0.0004501343
'    Const FCPS2i46 As Double = 0.0004444122
'    Const FCPS2i47 As Double = 0.0004386902
'    Const FCPS2i48 As Double = 0.0004386902
'    Const FCPS2i49 As Double = 0.0004291534
'    Const FCPS2i50 As Double = 0.0004329681
'    Const FCPS2i51 As Double = 0.000421524
'    Const FCPS2i52 As Double = 0.000421524
'    Const FCPS2i53 As Double = 0.0004177094
'    Const FCPS2i54 As Double = 0.0004119873
'    Const FCPS2i55 As Double = 0.000408172
'#End Region

'    Dim AH(131) As Double, ACPH(131) As Double, ASi(131) As Double, ACPS(131) As Double
'    Dim FH(131) As Double, FCPH(131) As Double, FS(131) As Double, FCPS(131) As Double

'    Const FLO = 14.948
'    'Const HU = 10250#
'    Const CPTO = 54.52
'    Const RAIR = 29.27
'    Const RF = 29.428
'    Const RW = 47.069

'    Dim AH1(76) As Double, AH2(55) As Double, ACPH1(76) As Double, ACPH2(55) As Double
'    Dim AS1(76) As Double, AS2(55) As Double, ACPS1(76) As Double, ACPS2(55) As Double
'    Dim FH1(76) As Double, FH2(55) As Double, FCPH1(76) As Double, FCPH2(55) As Double
'    Dim FS1(76) As Double, FS2(55) As Double, FCPS1(76) As Double, FCPS2(55) As Double

'    Dim BLTURB(,), SIGMks(), AKCT(), CGMKCT() As Single
'    Dim KER As Integer
'    Dim Hu As Single

'    Dim FatalErr As Boolean
'    'Dim Param() As DataXl
'    Dim ProcessingPossible() As Boolean
'    Dim ErrLog_global As String


'    Dim FAR As Double ' -отношение расхода топлива к расходу воздуха
'    Dim WAR As Double ' -влагосодержание  KgpAPA/Kg сухого воздуха
'    Dim TG As Double '    T   ' T  - температура торможения потока, gPAd K
'    Dim sigmКС As Double '"sкс реж" ' sigmКС
'    Dim HB, HF, HK, HG As Double '  -энтальпия  KKAl/Kg
'    Dim SB, SF As Double 'S    -энтропия KKAl/Kg/gPAd K
'    Dim CPB, CPF As Double '     CP   -теплоёмкость  KKAl/Kg/gPAd K
'    Dim RM As Double '     RM   -газовая постоянная  KgM/Kg/gPAd K
'    Dim AK As Double '     AK   -показатель адиабаты
'    Dim CS As Double '     CS   -скорость звука M/C  (не определяется)
'    Dim DGCA As Double '< - входные величины потерь которое подводится на охолодження турбины до горла СА ТВТ 


'    Public Sub GB1_Tg_v01_1a()

'        'Программа для определения Т*г Gв1 version_01.1a

'        ' Completed 2009.03.20
'        '1)Edited 2009.03.25
'        '  a) Edited 2009.03.31

'        Dim NameSheets(2), Ref(53) As String ' buf,, StrLog, Strn
'        Dim QuantPar, SempQuant As Integer
'        'Dim RepitAdr() As String

'        Stoil()

'        FatalErr = False

'        SempQuant = 7 'количество замеров



'        'ReDim_ProcessingPossible(1 To SempQuant)
'        'For i = 1 To SempQuant
'        '  ProcessingPossible(i) = True
'        'Next

'        Dim QuantParSheet1In As Integer = 12  ' Количество считываемых параметров с листа NameSheets(1)(исходные данные)
'        Dim QuantPerformPar As Integer = 6     'Количество считываемых параметров, заданных характеристик, с листа NameSheets(2)
'        Dim QuantAirBleedAndInput As Integer = 11 'Количество считываемых параметров, заданных отборов-подводов воздуха, с листа NameSheets(2)
'        Dim QuantParSheet2In As Integer = QuantPerformPar + QuantAirBleedAndInput ' Количество считываемых параметров с листа NameSheets(2)
'        Dim QuantParSheet1Out As Integer = 17 ' Количество выводимых параметров на лист NameSheets(1) (результат)
'        Dim ExtraPar As Integer = 7           ' Количество дополнительных параметров, которые участвуют в расчётах и не считываются, не выводятся на печать.


'        QuantPar = QuantParSheet1In + QuantParSheet2In + _
'                   QuantParSheet1Out + ExtraPar ' Количество идентификаторов параметров на листе NameSheets(0) + ExtraPar

'        'NameSheets(0) = "Список_имён" 'Имя листа с идентификаторами переменных и листов с данными
'        'NameSheets(1) = "A2" 'ячейка на листе NameSheets(0), содержащая имя листа ввода-вывода данных
'        'NameSheets(2) = "A50" 'ячейка на листе NameSheets(0), содержащая имя листа с характеристиками узлов, отбором и подводом воздуха.

'        'ReDim_Param(QuantPar), Ref(QuantPar)

'        ' Назначение размерности массива значений параметров (результатов расчёта). Размерность зависит от кол-ва замеров (SempQuant)
'        '    For i = QuantParSheet1In + QuantPerformPar + QuantAirBleedAndInput + 1 To QuantPar
'        '        With Param(i)
'        'ReDim_.Val(1 To SempQuant), .ErrData(1 To SempQuant)
'        '            .QuantSample = SempQuant
'        '            For j = 1 To SempQuant
'        '                .ErrData(j) = True
'        '            Next
'        '        End With
'        '    Next


'        ' На листе NameSheets(1). Исходные данные

'        Ref(1) = "A6"   'Hu
'        Ref(2) = "A7"   'Коэф. полноты сгор. топл. в К.С.
'        Ref(3) = "A8"   'Aг са твд заданное
'        Ref(4) = "A9"   'T*вх реж, K
'        Ref(5) = "A10"  'P*вх реж, кгс/см2
'        Ref(6) = "A11"  'Gт реж
'        Ref(7) = "A12"  'П*кS реж
'        Ref(8) = "A13"  'Gт норм
'        Ref(9) = "A14"  'П*кS норм
'        Ref(10) = "A15" 'Gт пр
'        Ref(11) = "A16" 'П*кS пр
'        Ref(12) = "A17" 'Абсолютный отбор воздуха от КВД на нужды самолёта по режимам (!!!Прежде чем изменить индекс см. ReadRowDefQuant)
'        '1- Hu
'        '2- hг
'        '3- А СА ТВД кр
'        '4- Т*вх реж 
'        '5- Р*вх реж 1
'        '6- Gт реж
'        '7- p*кS реж
'        '8- Gт норм
'        '9- p*кS норм
'        '10- Gт пр 
'        '11- p*кS пр     
'        '12- dGв отб реж


'        ' На листе NameSheets(2). Характеристики узлов и схема отборов и подводов воздуха

'        Ref(13) = "A52"  'Т*КВД пр хар
'        Ref(14) = "A53"  'Р*КВД пр хар
'        Ref(15) = "A55"  'sigmКС хар
'        Ref(16) = "A56"  'Gв пр вх жт КС хар
'        Ref(17) = "A59"  'AСА ТВД хар
'        Ref(18) = "A60"  'П*кS хар

'        Ref(19) = "A63" 'выходные величины потерь которые выбирается на ПС на охолоджение турбины.
'        Ref(20) = "A64" 'выходные величины потерь которые выбирается на КВТ на охолоджение турбины.
'        Ref(21) = "A65" 'абсолютные выходные величины потерь которые выбирается на ПС на лётные потребности
'        Ref(22) = "A66" 'абсолютные выходные величины потерь которые выбирается на КВТ на лётные потребности
'        Ref(23) = "A67" 'выходные величины потерь которые выбирается на охолоджение турбины до горла СА ТВТ.
'        Ref(24) = "A68" 'выходные величины потерь которые выбирается на охолоджение турбины между горлом СА и рабочим колесом ТВТ.
'        Ref(25) = "A69" 'выходные величины потерь которые выбирается на охолоджение турбины за робочим колесом ТВТ.
'        Ref(26) = "A70" 'выходные величины потерь которые выбирается на охолоджение турбины до горла СА ТНТ.
'        Ref(27) = "A71" 'выходные величины потерь которые выбирается на охолоджение турбины между горлом САи рабочим колесом ТНТ.
'        Ref(28) = "A72" 'выходные величины потерь которые выбирается на охолоджение турбины за робочим колесом ТНТ.
'        Ref(29) = "A73" 'выходные величины потерь которые выбирается  на газотурбинный просмотр

'        ' На листе NameSheets(1). Результат
'        '*********************************************
'        Ref(30) = "A22"  'Т*КВД реж
'        Ref(31) = "A23"  'Р*КВД реж
'        Ref(32) = "A24"  'sigmКС реж
'        Ref(33) = "A25"  'Gв1 реж
'        Ref(34) = "A26"  'Т*г "горло" СА реж
'        Ref(35) = "A27"  'Т*г СА реж
'        Ref(36) = "A28"  'Aса твд реж

'        Ref(37) = "A30"  'Т*КВД норм
'        Ref(38) = "A31"  'Р*КВД норм
'        Ref(39) = "A32"  'Gв1 норм
'        Ref(40) = "A33"  'Т*г "горло" СА норм
'        Ref(41) = "A34"  'Т*г СА норм

'        Ref(42) = "A36"  'Т*КВД пр
'        Ref(43) = "A37"  'Р*КВД пр
'        Ref(44) = "A38"  'Gв1 пр
'        Ref(45) = "A39"  'Т*г "горло" СА пр
'        Ref(46) = "A40"  'Т*г СА пр

'        ' Параметры просто участвующие в расчётах (без адреса)
'        Ref(47) = "Gв пр вх жт КС реж"   'Gв пр вх жт КС реж

'        Ref(48) = "Aса твд пр"           'Aса твд пр
'        Ref(49) = "Gв пр вх жт КС пр"    'Gв пр вх жт КС пр
'        Ref(50) = "sigmКС пр"            'sigmКС пр

'        Ref(51) = "Aса твд норм"         'Aса твд норм
'        Ref(52) = "Gв пр вх жт КС норм"  'Gв пр вх жт КС норм
'        Ref(53) = "sigmКС норм"          'sigmКС норм



'        'IsSheetFound(NameSheets(0), Ref(1), Param(1).NameLocation) 'проверка наличия листа с заданным именем в рабочей книге
'        'If FatalErr Then Exit Sub

'        'IsDefinitPar(NameSheets(0), 1) 'проверка наличия индентификатора для параметра
'        'If FatalErr Then Exit Sub

'        'Param(1).NamePar = Param(1).NameLocation.Value

'        ' считывание форматов для 1-го параметра
'        'With Param(1).FormatName
'        '    ReDim_.CharacterI(1 To Len(Param(1).NamePar)), .FormatI(1 To Len(Param(1).NamePar))
'        '        End With
'        '        For fi = 1 To Len(Param(1).NamePar)
'        '            With Param(1).NameLocation
'        '                Param(1).FormatName.CharacterI(fi) = .Characters(Start:=fi, Length:=1).Text
'        '                Param(1).FormatName.FormatI(fi) = .Characters(Start:=fi, Length:=1).Font
'        '            End With
'        '        Next

'        '        'Считывание  имён идентификаторов и их форматов
'        '        For i = 2 To QuantPar - ExtraPar
'        '            Param(i).NameLocation = ThisWorkbook.Sheets(NameSheets(0)).Range(Ref(i))
'        '            IsDefinitPar(NameSheets(0), i) 'проверка наличия индентификатора для параметра
'        '            If FatalErr Then Exit Sub
'        '            Param(i).NamePar = Param(i).NameLocation.Value
'        '            ' считывание форматов
'        '            With Param(i).FormatName
'        '    ReDim_.CharacterI(1 To Len(Param(i).NamePar)), .FormatI(1 To Len(Param(i).NamePar))
'        '            End With
'        '            For fi = 1 To Len(Param(i).NamePar)
'        '                With Param(i).NameLocation
'        '                    Param(i).FormatName.CharacterI(fi) = .Characters(Start:=fi, Length:=1).Text
'        '                    Param(i).FormatName.FormatI(fi) = .Characters(Start:=fi, Length:=1).Font
'        '                End With
'        '            Next
'        '        Next

'        '        For i = QuantPar - ExtraPar + 1 To QuantPar ' Присвоение имён параметрам без адреса
'        '            Param(i).NamePar = Ref(i)
'        '        Next

'        'ParamRepetitionSearch (QuantPar - ExtraPar), NameSheets(0) 'проверка наличия одинаковых идентификаторов на одном листе
'        '        If FatalErr Then Exit Sub

'        '        'Определение индексов для считываемых параметров в зависимости от листа NameSheet(i), на котором они располагаются
'        '        For i = 1 To 2
'        '            If i = 1 Then
'        '                ji = 1
'        '                jn = QuantParSheet1In
'        '            ElseIf i = 2 Then
'        '                ji = QuantParSheet1In + 1
'        '                jn = QuantParSheet1In + QuantParSheet2In
'        '            Else
'        '                MsgBox("Внимание! Обнаружена программная ошибка, которая может привести к непредсказуемому результату.")
'        '            End If

'        '            NameSheets(i) = ThisWorkbook.Sheets(NameSheets(0)).Range(NameSheets(i)).Value
'        '            IsSheetFound(NameSheets(i), "A1", Param(ji).ValLocation) 'проверка наличия листа с заданным именем в рабочей книге
'        '            If FatalErr Then Exit Sub

'        '            'Очистка ячеек вывода результатов расчёта
'        '  If i = 1 Then CleanCells (QuantParSheet1In + QuantParSheet2In + 1), _
'        '                            QuantParSheet1Out - 1, NameSheets(1), RepitAdr()

'        '            'Поиск идентификатора параметра Param(j).NamePar на листе NameSheets(i)
'        '            For j = ji To jn

'        '                SearchWithinColumn(NameSheets(i), j, 1, RepitAdrLog, Repetition, RepitAdr())

'        '                If Repetition > 1 Then        ' Если на листе NameSheets(i) больше одного идентификатора соответствующего заданному Param(j).NamePar
'        '                    MsgBox("Параметр """ & Param(j).NamePar & """ повторяется " & _
'        '                          Repetition & " разa. Смотрите лист """ & NameSheets(i) & """, ячейки " & _
'        '                          RepitAdrLog & ". Выполнение программы прервано.")
'        '                    Exit Sub
'        '                ElseIf Repetition = 0 Then    'Если на листе NameSheets(i) не найдено ни одного идентификатора соответствующего заданному Param(j).NamePar
'        '                    MsgBox("Лист """ & NameSheets(i) & """, параметр """ & Param(j).NamePar & _
'        '                       """ не найден. Выполнение программы прервано.")
'        '                    Exit Sub
'        '                ElseIf Repetition = 1 Then    'Если на листе NameSheets(i) найден только один идентификатор соответствующий заданному Param(j).NamePar
'        '                    '      Set Param(j).ValLocation = ThisWorkbook.Sheets(NameSheets(i)).Range(RepitAdrLog(Repetition))
'        '                    Param(j).ValLocation = ThisWorkbook.Sheets(NameSheets(i)).Range(RepitAdrLog)

'        '                    If i = 1 And (j >= 1 And j <= QuantParSheet1In) Then ' Для считывания данных с листа NameSheets(1)
'        '                        If j >= 1 And j < 4 Then 'Для параметров, значения которых задаются в одной ячейке т.е. по режимам не изменяются.
'        '                            SempQ = 1
'        '                        Else                     'Для изменяющихся параметров по режимам
'        '                            SempQ = SempQuant
'        '                        End If

'        '                        With Param(j)            'Определение размерности массивов Param(j).Val(), Param(j).ErrData() и присвоение всему массиву Param(j).ErrData() значения True
'        '          ReDim_.ErrData(1 To SempQ), .Val(1 To SempQ)
'        '                            For j01 = 1 To SempQ
'        '                                .ErrData(j01) = True
'        '                            Next
'        '                        End With

'        '                        ReadRowDefQuant(SempQ, j) ' Считывание данных с листа NameSheets(1),(если заранее известно количество считываемых значения параметров)

'        '                        If FatalErr Then ' если не заданы значения параметров Hu; Коэф. полноты сгор. топл. в К.С.; Aт г са твд, т. е. которые задаются в одной ячейке не зависимо от количества режимов.
'        '                            ref1 = Param(j).ValLocation.Cells(1, 2).Address
'        '                            ref2 = Param(j).ValLocation.Cells(1, 3).Address
'        '                            MsgBox("Пожалуйста, введите числовое значение параметра """ & _
'        '                                   Param(j).NamePar & """ на листе """ & NameSheets(i) & _
'        '                                   """ в ячейку """ & ref1 & """ или в ячейку """ & ref2 & """.")
'        '                            Exit Sub
'        '                        End If
'        '                    ElseIf i = 2 Then ' Для считывания данных с листа NameSheets(2)
'        '                        If (j > QuantParSheet1In And j <= QuantParSheet1In + QuantPerformPar) Then 'Считывание характеристик узлов
'        '                            ReadRowIndefQuant(Param(j).ValLocation, j, stepj)
'        '                            ArgOrFunk = ArgOrFunk + 1 'Если ArgOrFunk=1 - функция; если ArgOrFunk=2 - аргумент. Переменная необходима для проверки равенства количества считанных значений аргумента функции количеству считанных значений функции, заданной характеристики узла двигателя.
'        '                            If Param(j).QuantSample < 3 Then 'Если количество считанных точек заданной характеристики меньше трех.
'        '                                MsgBox("Ошибка. Лист """ & NameSheets(i) & """ переменная """ & _
'        '                                      Param(j).NamePar & """ - количество считанных значений" & _
'        '                                      " переменой меньше трёх. Выполнение программы прервано")
'        '                                Exit Sub
'        '                            ElseIf ArgOrFunk = 2 Then
'        '                                If Param(j).QuantSample <> Param(j - 1).QuantSample Then ' если количество считанных значений функции не равно количеству считанных значений аргумента функции.
'        '                                    MsgBox("Ошибка. Лист """ & NameSheets(i) & _
'        '                                          """ - количество считанных значений функции """ & Param(j - 1).NamePar & _
'        '                                          " = f(" & Param(j).NamePar & ")""" & " не равно количеству" & _
'        '                                          " считанных значений аргумента """ & Param(j).NamePar & """." & _
'        '                                          " Выполнение программы прервано.")
'        '                                    Exit Sub
'        '                                End If
'        '                                ArgOrFunk = 0
'        '                            End If

'        '                        ElseIf j > QuantParSheet1In + QuantPerformPar And _
'        '                               j <= QuantParSheet1In + QuantParSheet2In Then 'Считывание отборов/подводов воздуха
'        '                            ReadRowIndefQuant(Param(j).ValLocation, j, stepj)

'        '                            If Param(j).QuantSample = 0 Then
'        '                                If j > (QuantParSheet1In + QuantPerformPar + 1) Then
'        '                                    reg = identj_1 + stepj_1 + Param(j - 1).QuantSample        'Максиальный отступ (кол-во колонок) от начала листа для текущих значений отборов/подводов воздуха
'        '                                    Strn = Param(j).ValLocation.Cells(1, stepj_1).Value
'        '                                    i1 = 0
'        '                                    Do Until identj > reg Or IsNumeric(Strn) = True
'        '                                        i1 = i1 + 1
'        '                                        identj = stepj_1 + i1
'        '                                        Strn = Param(j).ValLocation.Cells(1, identj).Value
'        '                                        '                wwww = Param(j).ValLocation.Cells(1, identj).Address
'        '                                    Loop
'        '                                    If IsNumeric(Strn) Then
'        '                                        ReadRowIndefQuant(Param(j).ValLocation.Cells(1, identj - stepj_1), j, stepj)
'        '                                    Else
'        '                                        MsgBox("Значение параметра """ & Param(j).NamePar & _
'        '                                               """ на листе """ & NameSheets(i) & """ не найдено, поэтому " & _
'        '                                               "значение """ & Param(j).NamePar & """ принято равным нулю.")
'        '                                    End If
'        '                                Else
'        '                                    MsgBox("Значение параметра """ & Param(j).NamePar & _
'        '                                            """ на листе """ & NameSheets(i) & """ не найдено, поэтому " & _
'        '                                            "значение """ & Param(j).NamePar & """ принято равным нулю.")
'        '                                End If
'        '                                identj_1 = identj
'        '                                identj = 0
'        '                            Else
'        '                                identj_1 = 0
'        '                            End If
'        '                            stepj_1 = stepj
'        '                        End If
'        '                    End If
'        '                End If
'        '            Next
'        '        Next
'        DataInitialization(SempQuant)
'        'DataOutput (QuantParSheet1In + QuantParSheet2In + 1), QuantParSheet1Out - 1, NameSheets(1)
'        'If FatalErr = True Then Exit Sub
'    End Sub



'    Private Sub DataInitialization(ByVal SempQuant As Integer)

'        Dim IndGT() As Integer  ' массив индексов для Gт


'        'Низшая теплотворная способность топлива
'        Dim j As Integer = 1
'        Hu = 10334.86 '10335 'Param(j).Val(1)


'        'Коэф. полноты сгор. топл. в К.С. hг
'        j = 2
'        Dim EFКоэф_полноты_сгор_топл_КС As Double = 0.995 'Param(j).Val(1)


'        'Аг са твд max харт
'        j = 17
'        'Dim ATmax As Double
'        'For m = 1 To Param(j).QuantSample
'        '    If Param(j).Val(m) > ATmax Then ATmax = Param(j).Val(m) 'Определение max значения Аг твд в заданной характеристике.
'        'Next


'        'Суммарный относительный отбор воздуха от КВД
'        j = 20
'        Dim DGSСум_отн_отбор_воздуха_от_КВД As Double = -0.1752
'        'For m = 1 To Param(j).QuantSample
'        '    DGS = DGS + Param(j).Val(m)
'        'Next

'        'Подводы отобранного воздуха в турбину
'        Dim size_max As Integer
'        'For j = 23 To 29
'        '    If size_max < Param(j).QuantSample Then size_max = Param(j).QuantSample
'        'Next

'        ReDim_BLTURB(size_max, SempQuant)
'        'For j = 1 To 7
'        '    For m = 1 To size_max
'        '        If Param(j + 22).QuantSample < m Then
'        '            BLTURB(m, j) = 0
'        '        Else
'        '            BLTURB(m, j) = Param(j + 22).Val(m)
'        '        End If
'        '    Next
'        'Next

'        ReDim_IndGT(3)
'        'Dim X1Ind, Ind0, Y1Ind As Integer
'        Dim PKBD As Double = 5.256096
'        'For i = 1 To 3 'Варианты расчёта Gв1 и Т*г: i=1 - по замеренным параметрам; i=2 - по нормальным параметрам; i=3 по приведенным параметрам.
'        '    For j = 1 To SempQuant 'количество замеров
'        '        'If ProcessingPossible(j) Then
'        '        If i = 1 Then                        'Для расчёта Gв1 и Т*г по замеренным параметрам
'        '            X1Ind = 31                              'индекс Р*КВД реж

'        '            Ind0 = 11
'        '            'If Param(Ind0).ErrData(j) = False Then
'        '            PKBD = 1.0332 * Param(Ind0).Val(j)    'P*квд пр
'        '            'Param(Ind0).ErrData(j) = False
'        '            'Param(X1Ind).ErrData(j) = False
'        '            'Param(X1Ind).Val(j) = PKBD            'Временно присваивается значение P*квд пр переменной P*квд реж. Это необходимо для определения T*квд реж (T*квд пр = f(P*квд пр)) для расчёта по замеренным параметрам.
'        '            'Else
'        '            '    Param(X1Ind).ErrData(j) = True
'        '            'End If
'        '            Y1Ind = 30                              'индекс Т*КВД реж

'        '            IndGT(i) = 6
'        '            If Param(IndGT(i)).ErrData(j) = False Then GT = Param(IndGT(i)).Val(j) 'Gт реж

'        '            X2Ind = 7                               'индекс П*кS реж
'        '            Y2Ind = 36                              'индекс Aса твд

'        '            X3Ind = 47                              'индекс Gв пр вх жт КС реж
'        '            Y3Ind = 32                              'индекс sigmКС реж

'        '            Y4Ind = 33                              'индекс Gв1 реж
'        '            Y5Ind = 34                              'индекс Т*г "горло" СА реж
'        '            Y6Ind = 35                              'индекс Т*г СА реж

'        '        ElseIf i = 2 Then                    'Для расчёта Gв1 и Т*г по нормальным параметрам
'        '            X1Ind = 38                              'индекс Р*КВД норм
'        '            Y1Ind = 37                              'индекс Т*КВД норм

'        '            IndGT(i) = 8
'        '            If Param(IndGT(i)).ErrData(j) = False Then GT = Param(IndGT(i)).Val(j) 'Gт норм

'        '            X2Ind = 9                               'индекс П*кS норм
'        '            If Param(X2Ind).ErrData(j) = False Then
'        '                PKBD = 1.0332 * Param(X2Ind).Val(j)     'P*квд норм
'        '                Param(X1Ind).Val(j) = PKBD
'        '                Param(X1Ind).ErrData(j) = False
'        '            Else
'        '                Param(X1Ind).ErrData(j) = True
'        '            End If
'        '            Y2Ind = 51                              'индекс Aса твд норм

'        '            X3Ind = 52                              'индекс Gв пр вх жт КС норм
'        '            Y3Ind = 53                              'индекс sigmКС норм

'        '            Y4Ind = 39                              'индекс Gв1 норм
'        '            Y5Ind = 40                              'индекс Т*г "горло" СА норм
'        '            Y6Ind = 41                              'индекс Т*г СА норм

'        '        ElseIf i = 3 Then                    'Для расчёта Gв1 и Т*г по приведенным параметрам
'        '            X1Ind = 43                              'индекс Р*КВД пр
'        '            Y1Ind = 42                              'индекс Т*КВД пр

'        '            IndGT(i) = 10
'        '            If Param(IndGT(i)).ErrData(j) = False Then GT = Param(IndGT(i)).Val(j) 'Gт пр

'        '            X2Ind = 11                              'индекс П*кS пр

'        '            If Param(X2Ind).ErrData(j) = False Then
'        '                PKBD = 1.0332 * Param(X2Ind).Val(j)     'P*квд пр
'        '                Param(X1Ind).Val(j) = PKBD
'        '                Param(X1Ind).ErrData(j) = False
'        '            Else
'        '                Param(X1Ind).ErrData(j) = True
'        '            End If
'        '            Y2Ind = 48                              'индекс Aса твд пр

'        '            X3Ind = 49                              'индекс Gв пр вх жт КС пр
'        '            Y3Ind = 50                              'индекс sigmКС пр

'        '            Y4Ind = 44                              'индекс Gв1 пр
'        '            Y5Ind = 45                              'индекс Т*г "горло" СА пр
'        '            Y6Ind = 46                              'индекс Т*г СА пр

'        '        End If

'        '        ArgInd = 14                                'индекс Р*КВД пр хар
'        '        FunkInd = 13                               'индекс Т*КВД пр хар
'        '        If Param(X1Ind).ErrData(j) = False Then
'        '            TKBD = DPG(X1Ind, Y1Ind, ArgInd, FunkInd, j) 'если i=1 или 3 - Т*КВД пр; i=2 - Т*КВД норм.
'        '            If Len(ErrLog_global) > 0 Then
'        '                MsgBox(ErrLog_global)
'        '                ErrLog_global = ""
'        '            End If
'        '            Param(Y1Ind).Val(j) = TKBD
'        '            Param(Y1Ind).ErrData(j) = False
'        '        Else
'        '            Param(Y1Ind).ErrData(j) = True
'        '        End If
'        '        ArgInd = 18                                 'индекс П*кS хар
'        '        FunkInd = 17                                'индекс AСА ТВД хар

'        '        AT = Param(3).Val(1)                        'Аг са твд заданное

'        '        If Param(X2Ind).ErrData(j) = False Then
'        '            AT = DPG(X2Ind, Y2Ind, ArgInd, FunkInd, j) / ATmax * AT 'если i=1 - Aг са твд реж; i=2 - Aг твд норм; i=3 - Aг твд пр.
'        '            If Len(ErrLog_global) > 0 Then
'        '                MsgBox(ErrLog_global)
'        '                ErrLog_global = ""
'        '            End If
'        '            Param(Y2Ind).ErrData(j) = False
'        '        Else
'        '            Param(Y2Ind).ErrData(j) = True
'        '        End If

'        '        '  Проверка  ###############################################################
'        '        If i = 1 Then
'        '            reng_Pквд = ThisWorkbook.Sheets("Проверка").Range("C2")
'        '            reng_Tквд = ThisWorkbook.Sheets("Проверка").Range("C3")

'        '            If Param(X1Ind).ErrData(j) = False Then
'        '                reng_Pквд.Cells(1, j).Value = Param(X1Ind).Val(j)
'        '                reng_Tквд.Cells(1, j).Value = Param(Y1Ind).Val(j)
'        '            End If
'        '        End If
'        '        '########################################################################

'        '        If i = 1 Then    'Расприведение для расчёта режимных параметров

'        '            Ind0 = 5                                         'Индекс P*вх реж
'        '            If Param(Ind0).ErrData(j) = False And Param(X2Ind).ErrData(j) = False Then
'        '                PKBD = Param(Ind0).Val(j) * Param(X2Ind).Val(j) 'P*квд реж
'        '                Param(X1Ind).Val(j) = PKBD
'        '                Param(X1Ind).ErrData(j) = False
'        '            Else
'        '                Param(X1Ind).ErrData(j) = True
'        '            End If


'        '            Ind0 = 4                                          'Индекс T*вх реж
'        '            If Param(Ind0).ErrData(j) = False And Param(11).ErrData(j) = False Then
'        '                '       TKBD = TKBD * Param(4).Val(j) / 288.15
'        '                TKBD = TKBD * Param(4).Val(j) / 288.15 * _
'        '                       (1 - 0.0001 * (Param(4).Val(j) - 288.15)) 'T*квд реж
'        '                '       TKBD = TKBD * Param(4).Val(j) / 288.15 * _
'        '                '(1 - (1.429 * 13141 - 1716.2) * (Param(4).Val(j) - 288.15) / 100000000)

'        '                '      TKBD = TKBD * Param(4).Val(j) / 288.15 * _
'        '                '(1 - (0.9232 * Param(11).Val(j) + 2.9737) * (Param(4).Val(j) - 288.15) / 100000)
'        '                Param(Y1Ind).Val(j) = TKBD
'        '                Param(Y1Ind).ErrData(j) = False
'        '            Else
'        '                Param(Y1Ind).ErrData(j) = True
'        '            End If
'        '        End If

'        '        DGC = Param(12).Val(j)
'        '        ArgInd = 16                                     'индекс Gв пр вх жт КС хар
'        '        FunkInd = 15                                    'индекс sigmКС хар

'        'If Param(X1Ind).ErrData(j) = False And Param(Y1Ind).ErrData(j) = False And _
'        '   Param(IndGT(i)).ErrData(j) = False And Param(Y2Ind).ErrData(j) = False Then
'        'индекс Т*КВД реж
'        Dim TKBD As Double = 472.723531748952 ' = TKBD * Param(4).Val(j) / 288.15 * _
'        '(1 - 0.0001 * (Param(4).Val(j) - 288.15)) 'T*квд реж
'        Dim GT As Double = 340.9 '= Param(IndGT(i)).Val(j)  'Gт реж
'        Dim DGC As Double = 0 '"dGв отб реж"
'        Dim AT As Double = 54.1431 ' Aса твд пр  '"А СА ТВД кр"
'        Dim GB As Double ' Gв1 пр
'        Dim TG As Double ' Т*г "горло" СА пр



'        '1- Hu
'        '2- hг
'        '3- А СА ТВД кр
'        '4- Т*вх реж 
'        '5- Р*вх реж 1
'        '6- Gт реж
'        '7- p*кS реж
'        '8- Gт норм
'        '9- p*кS норм
'        '10- Gт пр 
'        '11- p*кS пр     
'        '12- dGв отб реж




'        Call GBFAT(0, PKBD, TKBD, GT, DGSСум_отн_отбор_воздуха_от_КВД, DGC, EFКоэф_полноты_сгор_топл_КС, TG, GB, AT) ', X3Ind, Y3Ind, ArgInd, FunkInd, j)

'        'Param(Y2Ind).ErrData(j) = False
'        'Param(X3Ind).ErrData(j) = False
'        'Param(Y3Ind).ErrData(j) = False
'        'Param(Y4Ind).ErrData(j) = False
'        'Param(Y5Ind).ErrData(j) = False
'        'Param(Y2Ind).Val(j) = AT 'индекс Aса твд пр
'        'Param(Y4Ind).Val(j) = GB 'индекс Gв1 пр
'        'Param(Y5Ind).Val(j) = TG 'индекс Т*г "горло" СА пр

'        Dim DGOX As Double ' < - входные величины потерь которое подводится на охолодження турбины между горлом СА и робочим колесом ТВТ 
'        Dim TGPK As Double 'Т*г СА

'        Call TGBL(0, GB, GT, TKBD, TG, DGSСум_отн_отбор_воздуха_от_КВД, DGC, DGOX, TGPK)
'        'Param(Y6Ind).ErrData(j) = False
'        'Param(Y6Ind).Val(j) = TGPK
'        'Else
'        'Param(Y2Ind).ErrData(j) = True
'        'Param(Y4Ind).ErrData(j) = True
'        'Param(Y5Ind).ErrData(j) = True
'        'Param(Y6Ind).ErrData(j) = True
'        'End If

'        '      Param(Y1Ind).Val(j) = TKBD
'        '      Param(Y2Ind).Val(j) = AT
'        'End If
'        'Next
'        'Проверка  ###############################################################################
'        'If i = 1 Then
'        '    For j = 1 To SempQuant
'        '        '    If ProcessingPossible(j) = True Then
'        '        reng_Пи_к_сум_реж = ThisWorkbook.Sheets("Проверка").Range("C8")
'        '        reng_Аг_са_твд = ThisWorkbook.Sheets("Проверка").Range("C9")

'        '        reng_Gв_пр_вх_жт_КС = ThisWorkbook.Sheets("Проверка").Range("C13")
'        '        reng_Sigma_КС = ThisWorkbook.Sheets("Проверка").Range("C14")

'        '        If Param(X2Ind).ErrData(j) = False Then reng_Пи_к_сум_реж.Cells(1, j).Value = Param(X2Ind).Val(j)
'        '        If Param(Y2Ind).ErrData(j) = False Then reng_Аг_са_твд.Cells(1, j).Value = Param(Y2Ind).Val(j)
'        '        If Param(Y2Ind).ErrData(j) = False Then reng_Аг_са_твд.Cells(2, j).Value = Param(Y2Ind).Val(j) / Param(3).Val(1) * ATmax

'        '        If Param(X3Ind).ErrData(j) = False Then reng_Gв_пр_вх_жт_КС.Cells(1, j).Value = Param(X3Ind).Val(j)
'        '        If Param(Y3Ind).ErrData(j) = False Then reng_Sigma_КС.Cells(1, j).Value = Param(Y3Ind).Val(j)
'        '        '    End If
'        '    Next
'        'End If
'        ''########################################################################################
'        'Next
'    End Sub


'    ''' <summary>
'    ''' ПPOГPAMMA OПPEДEЛЯET PACXOД BOЗДУXA ЧEPEЗ ПEPBЫЙ KOHTУP ПO ИЗBECTHOMУ ЗHAЧEHИЮ ПPИBEДEHHEГO PACXOДA ГAЗA ЧEPEЗ 1 CA
'    ''' </summary>
'    ''' <param name="D"></param>
'    ''' <param name="PB"></param>
'    ''' <param name="TB"></param>
'    ''' <param name="GT"></param>
'    ''' <param name="DGSСум_отн_отбор_воздуха_от_КВД"></param>
'    ''' <param name="DGC"></param>
'    ''' <param name="Коэф_полноты_сгор_топл_КС"></param>
'    ''' <param name="TF"></param>
'    ''' <param name="GB"></param>
'    ''' <param name="AT"></param>
'    ''' <remarks></remarks>
'    Private Sub GBFAT(ByVal D As Double, ByVal PB As Double, ByVal TB As Double, ByVal GT As Double, ByVal DGSСум_отн_отбор_воздуха_от_КВД As Double, ByVal DGC As Double, ByRef Коэф_полноты_сгор_топл_КС As Double, ByRef TF As Double, ByRef GB As Double, ByRef AT As Double) ', ByVal XInd, ByVal YInd, ByVal ArgInd, ByVal FunkInd, ByVal j)


'        '  ПPOГPAMMA OПPEДEЛЯET PACXOД BOЗДУXA ЧEPEЗ ПEPBЫЙ KOHTУP ПO ИЗBECTHOMУ ЗHAЧEHИЮ ПPИBEДEHHEГO PACXOДA ГAЗA ЧEPEЗ 1 CA

'        '  BXOДHЫE  ПAPAMETPЫ :
'        '  D - BЛAГOCOДEPЖAHИE
'        '  PB    - ПOЛHOE ДABЛEHИE ВОЗДУХА HA BXOДE B KAMEPУ CГOPAHИЯ (ЗА КВД)
'        '  TB    - ПOЛHAЯ TEMПEPATУPA ВОЗДУХА HA BXOДE B KAMEPУ CГOPAHИЯ (ЗА КВД)
'        '  GT    - PACXOД TOПЛИBA
'        '  DGS   - СУММАРНЫЙ OTHOCИTEЛЬHЫЙ OTБОР BOЗДУХА

'        '  BЫXOДHЫE  ПAPAMETPЫ :
'        '  TF    - TEMПEPATУPA ГAЗA В ГОРЛЕ СА ТВД
'        '  GB    - PACXOД BOЗДУXA ЧEPEЗ ВНУТРЕННИЙ KOHTУP (БЕЗ УЧЕТА ПЕРЕПУСКА ЗА КНД)
'        '  SIGM  - Потери полного давления в жаровой трубе камеры сгорания.
'        '      COMMON/COMBX/AKCT(10),CGMKCT(10),KER

'        '    COMMON/SIGM/SIGMks(10)

'        '     COMMON/GRBLT/
'        '     *BLTURB(6,7),HBL(6),TBL(6)
'        Dim i As Integer

'        DGCA = 0
'        'For i = 1 To 6
'        '    DGCA = DGCA + BLTURB(i, 1)
'        'Next
'        DGCA = 0.0754

'        '            БЛОК 1
'        '     НАЧАЛЬНОЕ ПРИБЛИЖЕНИЕ ДЛЯ ОПРЕДЕЛЕНИЯ РАСХОДА
'        '     ВОЗДУХА В "ГОРЛЕ" СА ТВД (GBTBDG)
'        Dim QF1, GBTBDG, QF, GBTBDGP, GBJT, AKC As Double
'        Dim PF As Double 'по функции

'        QF1 = 0.02
'        GBTBDG = GT / 3600.0# / QF1

'        '            БЛОК 2
'        '     ОПРЕДЕЛЕНИЕ РАСХОДА ВОЗДУХА ЧЕРЕЗ ВНУТРЕННИЙ КОНТУР
'        i = 1
'        Do
'            '1     QF = GT / GBTBDG / 3600#
'            If i > 1 Then GBTBDG = (GBTBDGP + GBTBDG) / 2.0#
'            QF = GT / GBTBDG / 3600.0#

'            Call BURNM(FAR, D, TB, QF, TG, Коэф_полноты_сгор_топл_КС)

'            '     расхода воздуха (GBJT) через жаровую трубу

'            GBJT = (GBTBDG + DGC) * (1.0# + DGSСум_отн_отбор_воздуха_от_КВД) / (1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA) - DGC

'            '     параметр расхода воздуха на входе в жаровую трубу

'            AKC = GBJT * Math.Sqrt(TB) / PB
'            'Param(XInd).Val(j) = AKC 'Gв пр вх жт КС

'            '     полное давление воздуха на входе в ТВД
'            'PF = PB * DPGIteration(XInd, YInd, ArgInd, FunkInd, j)
'            'дано "Gв пр вх жт КС" параметра "sigmКС"  для аргумента со значением YInd найти  "Gв пр вх жт КС хар" в функции "sigmКС хар" для J замера
'            'PF = PB * DPG(XInd, YInd, ArgInd, FunkInd, j)
'            'по оси X - Gв пр вх жт кс
'            'по оси Y - Sig кс
'            PF = 4.904456
'            '     расход воздуха на входе в ТВД

'            GBTBDGP = AT * PF / Math.Sqrt(TG) / (1.0# + QF)
'            i = i + 1
'        Loop Until Math.Abs((GBTBDGP - GBTBDG) / GBTBDGP) <= 0.001 '0.000001
'        GB = (GBTBDGP + DGC) / (1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA)
'        '      IF(ABS((GBTBDGP-GBTBDG)/GBTBDGP) <= 0.0001)  GO  TO  2
'        '      GBTBDG = (GBTBDGP + GBTBDG) / 2#
'        '      GoTo 1
'        '2     GB = GBTBDGP / (1# + DGS + DGCA)
'        'i = i + 1
'        'If (Ki > NDRi) Then i = 1
'        'If (Ki > NDRi) Then NDRi = Ki

'        'Param(YInd).Val(j) = PF / PB '"sкс реж" ' sigmКС
'        sigmКС = PF / PB
'        TF = TG
'        'If Len(ErrLog_global) > 0 Then
'        '    MsgBox(ErrLog_global)
'        '    ErrLog_global = ""
'        'End If
'    End Sub

'    ''' <summary>
'    ''' ПPOГPAMMA OПPEДEЛЯET TEMПEPATУPУ ГAЗA B K.C.
'    ''' </summary>
'    ''' <param name="FARB"></param>
'    ''' <param name="WAR"></param>
'    ''' <param name="TB"></param>
'    ''' <param name="FARF"></param>
'    ''' <param name="TF"></param>
'    ''' <param name="Коэф_полноты_сгор_топл_КС"></param>
'    ''' <remarks></remarks>
'    Private Sub BURNM(ByVal FARB As Double, ByVal WAR As Double, ByVal TB As Double, ByVal FARF As Double, ByRef TF As Double, ByRef Коэф_полноты_сгор_топл_КС As Double)

'        '     ПPOГPAMMA OПPEДEЛЯET TEMПEPATУPУ ГAЗA B K.C.
'        '
'        '     FARB,WAR,TB,EF,FARF - BXOДHЫE ПAPAMETPЫ
'        '     TF - BЫXOДHOЙ ПAPAMETP
'        '          DATA Hu / 10250# /
'        '          DATA HO / 54.52 /
'        Dim HO As Double = CPTO
'        Call CPFT(FARB, WAR, TB, HB, SB, CPB, RM, AK, CS)
'        HF = (HB * (1.0# + FARB + WAR) + (Hu * Коэф_полноты_сгор_топл_КС + HO) * FARF) / (1.0# + FARB + FARF + WAR)
'        FARF = FARF + FARB
'        Call TFH(FARF, WAR, TF, HF, SF, CPF, RM, AK, CS)
'    End Sub

'    ''' <summary>
'    ''' ПPOГPAMMA OПPEДEЛЯET TEMПEPATУPУ ГAЗA ПEPEД PK TBД
'    ''' </summary>
'    ''' <param name="Dвлагосодержание"></param>
'    ''' <param name="GB"></param>
'    ''' <param name="GT"></param>
'    ''' <param name="TK"></param>
'    ''' <param name="TG"></param>
'    ''' <param name="DGSСум_отн_отбор_воздуха_от_КВД"></param>
'    ''' <param name="DGC"></param>
'    ''' <param name="DGPK"></param>
'    ''' <param name="TGPK"></param>
'    ''' <remarks></remarks>
'    Private Sub TGBL(ByVal Dвлагосодержание As Double, ByVal GB As Double, ByVal GT As Double, ByVal TK As Double, ByVal TG As Double, ByVal DGSСум_отн_отбор_воздуха_от_КВД As Double, ByVal DGC As Double, ByVal DGPK As Double, ByVal TGPK As Double)
'        '
'        '     ПPOГPAMMA OПPEДEЛЯET TEMПEPATУPУ ГAЗA ПEPEД PK TBД
'        '
'        '    COMMON/GRBLT/
'        '     *BLTURB(6,7),HBL(6),TBL(6)

'        ' < - выходные величины потерь которые выбирается на охолоджение турбины между горлом СА и рабочим колесом ТВТ
'        DGPK = 0
'        'For i = 1 To 6
'        '    DGPK = DGPK + BLTURB(i, 2)
'        'Next
'        DGPK = 0.0216
'        '< - выходные величины потерь которые выбирается на охолоджение турбины до горла СА ТВТ.
'        DGCA = 0
'        'For i = 1 To 6
'        '    DGCA = DGCA + BLTURB(i, 1)
'        'Next
'        DGCA = 0.0754

'        FAR = 0.0#
'        Dim DGPLANE, QF0, QF1, QF2, DGPKrelative, DGCArelative As Double
'        Dim SK, SG, CP, HGPK, SGPK As Double
'        DGPLANE = DGC / GB
'        'C QF0 = GT / (GB - DGC) / 3600#
'        QF0 = GT / 3600.0# / GB
'        'C QF1 = QF0 / (1# + DGS + DGCA)
'        QF1 = GT / 3600.0# / (GB * (1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA) - DGC)
'        'C QF2 = QF0 / (1# + DGS + DGPK + DGCA)
'        QF2 = GT / 3600.0# / (GB * (1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA + DGPK) - DGC)

'        Call CPFT(FAR, Dвлагосодержание, TK, HK, SK, CP, RM, AK, CS)
'        Call CPFT(QF1, Dвлагосодержание, TG, HG, SG, CP, RM, AK, CS)
'        DGPKrelative = 1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA + DGPK + QF0 - DGPLANE
'        DGCArelative = 1.0# + DGSСум_отн_отбор_воздуха_от_КВД + DGCA + QF0 - DGPLANE
'        HGPK = (HG * DGCArelative + HK * DGPK) / DGPKrelative
'        Call TFH(QF2, Dвлагосодержание, TGPK, HGPK, SGPK, CP, RM, AK, CS)
'    End Sub

'    ''' <summary>
'    ''' ПPOГPAMMA  PACЧETA TEPMOДИHAMИЧECKИX  XAPAKTEPИCTИK  BOЗДУXA
'    ''' </summary>
'    ''' <param name="FAR_Gt_to_Gv"></param>
'    ''' <param name="WARвлагосодержание"></param>
'    ''' <param name="T"></param>
'    ''' <param name="Hэнтальпия"></param>
'    ''' <param name="Sэнтропия"></param>
'    ''' <param name="CPтеплоёмкость"></param>
'    ''' <param name="RM_газовая_постоянная"></param>
'    ''' <param name="AK_показатель_адиабаты"></param>
'    ''' <param name="CS"></param>
'    ''' <remarks></remarks>
'    Private Sub CPFT(ByVal FAR_Gt_to_Gv As Double, ByVal WARвлагосодержание As Double, ByVal T As Double, ByRef Hэнтальпия As Double, ByRef Sэнтропия As Double, ByRef CPтеплоёмкость As Double, ByRef RM_газовая_постоянная As Double, ByRef AK_показатель_адиабаты As Double, ByRef CS As Double)

'        '
'        '   ПPOГPAMMA  PACЧETA TEPMOДИHAMИЧECKИX  XAPAKTEPИCTИK  BOЗДУXA
'        '     И  ПPOДУKTOB  CГOPAHИЯ  B  ЗABИCИMOCTИ OT
'        '     TEMПEPATУPЫ.
'        '
'        '      COMMON/ACP/FLO,HU,CPTO,RAIR,RF,RW,
'        '     1         AH(131),ACPH(131),AS(131),ACPS(131),
'        '     2         FH(131),FCPH(131),FS(131),FCPS(131),
'        '3              WH (131), WCPH(131), WS(131), WCPS(131)
'        '


'        'Const GST As Double = 9.80665
'        Const AHPA As Double = 0.00234192
'        Dim i, r As Integer
'        Dim AFW, TB, DT, TR, DRT As Double

'        If (FAR_Gt_to_Gv < 0.0#) Then
'            FAR_Gt_to_Gv = 0.0#
'        End If
'        If (WARвлагосодержание < 0.0#) Then
'            WARвлагосодержание = 0.0#
'        End If

'        AFW = 1.0# + FAR_Gt_to_Gv + WARвлагосодержание
'        RM_газовая_постоянная = (RAIR + FAR_Gt_to_Gv * (RF + FLO * (RF - RAIR)) + WARвлагосодержание * RW) / AFW


'        If (T >= 220.0#) Then
'            If (T <= 900.0#) Then
'                TB = 200.0#
'                DT = 10.0#
'                TR = (T - TB) / DT
'                i = Fix(TR)
'                r = i
'                DRT = (TR - r) * DT
'                i = i + 1
'            ElseIf (T <= 2400.0#) Then
'                TB = 900.0#
'                DT = 25.0#
'                TR = (T - TB) / DT
'                i = Fix(TR)
'                r = i
'                DRT = (TR - r) * DT
'                i = i + 71
'            Else
'                i = 131
'                DRT = T - 2400.0#
'            End If
'            If (CPтеплоёмкость <> -1.0#) Then
'                CPтеплоёмкость = (ACPH(i) + FAR_Gt_to_Gv * FCPH(i)) / AFW
'                AK_показатель_адиабаты = CPтеплоёмкость / (CPтеплоёмкость - RM_газовая_постоянная * AHPA)
'            End If
'            If (Hэнтальпия <> -1.0#) Then
'                Hэнтальпия = (AH(i) + DRT * ACPH(i) + (FH(i) + DRT * FCPH(i)) * FAR_Gt_to_Gv) / AFW
'            End If
'            If (Sэнтропия <> -1.0#) Then
'                Sэнтропия = (ASi(i) + DRT * ACPS(i) + (FS(i) + DRT * FCPS(i)) * FAR_Gt_to_Gv) / AFW
'            End If
'        Else
'            Call CPFTF(FAR_Gt_to_Gv, WARвлагосодержание, T, Hэнтальпия, Sэнтропия, CPтеплоёмкость, RM_газовая_постоянная, AK_показатель_адиабаты, CS)
'        End If
'    End Sub

'    ''' <summary>
'    ''' Программа расчета термодинамических характеристик газа
'    ''' </summary>
'    ''' <param name="FAR_Gt_to_Gv"></param>
'    ''' <param name="WARвлагосодержание"></param>
'    ''' <param name="T"></param>
'    ''' <param name="Hэнтальпия"></param>
'    ''' <param name="Sэнтропия"></param>
'    ''' <param name="CPтеплоёмкость"></param>
'    ''' <param name="RM_газовая_постоянная"></param>
'    ''' <param name="AK_показатель_адиабаты"></param>
'    ''' <param name="CS"></param>
'    ''' <remarks></remarks>
'    Private Sub CPFTF(ByVal FAR_Gt_to_Gv As Double, ByVal WARвлагосодержание As Double, ByVal T As Double, ByRef Hэнтальпия As Double, ByRef Sэнтропия As Double, ByRef CPтеплоёмкость As Double, ByRef RM_газовая_постоянная As Double, ByRef AK_показатель_адиабаты As Double, ByRef CS As Double)
'        '
'        '  Программа расчета термодинамических характеристик газа
'        '  Программа уровня  CEp
'        '  Входные данные :  FAR,WAR,T
'        '  Выходные данные :  H,S,CP,RM,AK,CS
'        '  Обозначения :
'        '     FAR  -отношение расхода топлива к расходу воздуха
'        '     WAR  -влагосодержание  KgpAPA/Kg сухого воздуха
'        '     T    -температура  gPAd K
'        '     H    -энтальпия  KKAl/Kg
'        '     S    -энтропия KKAl/Kg/gPAd K
'        '     CP   -теплоёмкость  KKAl/Kg/gPAd K
'        '     RM   -газовая постоянная  KgM/Kg/gPAd K
'        '     AK   -показатель адиабаты
'        '     CS   -скорость звука M/C  (не определяется)
'        '
'        '  диапазон температур разбит на 2 поддиапазона :
'        '     1    - T <= 2400.
'        '     2    - T > 2400.
'        '
'        '
'        '      COMMON/ACPF/FLO,HU,CPTO,RAIR,RF,RW,
'        '     *    AR1,AR2,AR3,AR4,AR5,AR6,AR7,AR8,
'        '     *    AF1,AF2,AF3,AF4,AF5,AF6,AF7,AF8,
'        '     *    AW1,AW2,AW3,AW4,AW5,AW6,AW7,AW8
'        '      COMMON/ACPFW/FARF,WARF,RMF
'        '


'        Dim FARF As Double = 1000.0#
'        Dim WARF As Double = 1000.0#

'        ' KOЭФФИЦИEHTЫ AПPOKCИMAЦИИ CP BOЗДУXA OT TEMПEPATУPЫ
'        Const AR1 As Double = 0.238785488
'        Const AR2 As Double = -0.00900857881
'        Const AR3 As Double = 0.0193221258
'        Const AR4 As Double = 0.132711466
'        Const AR5 As Double = -0.190166916
'        Const AR6 As Double = 0.10379398
'        Const AR7 As Double = -0.0258490883
'        Const AR8 As Double = 0.00245365952
'        ' KOЭФФИЦИEHTЫ AПPOKCИMAЦИИ CP ПPOДУKTOB CГOPAHИЯ (ALFA=1)
'        ' OT TEMПEPATУPЫ
'        Const AF1 As Double = 0.367102859
'        Const AF2 As Double = 0.384334584
'        Const AF3 As Double = 0.190633867
'        Const AF4 As Double = -0.361766882
'        Const AF5 As Double = 0.263417998
'        Const AF6 As Double = -0.112618047
'        Const AF7 As Double = 0.0257271586
'        Const AF8 As Double = -0.00237417234
'        ' KOЭФФИЦETЫ AПPOKCИMAЦИИ CP BOДЯHOГO ПAPA OT TEMПEPATУPЫ
'        Const AW1 As Double = 0.4700077
'        Const AW2 As Double = -0.3206047
'        Const AW3 As Double = 1.196138
'        Const AW4 As Double = -1.728544
'        Const AW5 As Double = 1.55052
'        Const AW6 As Double = -0.8291961
'        Const AW7 As Double = 0.237057
'        Const AW8 As Double = -0.02773743

'        '         blOK  b01
'        '   проверка исходных данных
'        '
'        'Const GST As Double = 9.80665
'        Const AHPA As Double = 0.00234192
'        Dim A1, A2, A3, A4, A5, A6, A7, A8 As Double
'        Dim AFW, AF, r, HA, B, SA, c As Double



'        '      IF(FAR < 0.0)  FAR=0.0
'        '     IF(WAR < 0.0)  WAR=0.0
'        '
'        '         blOK  b02
'        '  вычисление газовой постоянной смеси
'        '
'        RM_газовая_постоянная = (RAIR + FAR_Gt_to_Gv * (RF + FLO * (RF - RAIR)) + WARвлагосодержание * RW) / (1.0# + FAR_Gt_to_Gv + WARвлагосодержание)

'        If (T > 2400.0#) Then
'            If (Math.Abs(FAR_Gt_to_Gv - FARF) > 0.00001 Or Math.Abs(WARвлагосодержание - WARF) > 0.00001) Then
'                WARF = WARвлагосодержание
'                FARF = FAR_Gt_to_Gv
'                AFW = (1.0# + FAR_Gt_to_Gv + WARвлагосодержание)
'                AF = (1.0# + FAR_Gt_to_Gv)
'            End If
'            If (FAR_Gt_to_Gv = 0.0#) Then
'                '  dlq  FAR.EQ.0
'                CPтеплоёмкость = 0.30403
'                Hэнтальпия = 658.32 + CPтеплоёмкость * (T - 2400.0#)
'                Sэнтропия = 2.1539 + CPтеплоёмкость * Math.Log(T / 2400.0#)
'                AK_показатель_адиабаты = CPтеплоёмкость / (CPтеплоёмкость - RM_газовая_постоянная * AHPA)
'            Else
'                '  dlq  FAR > 0
'                CPтеплоёмкость = (0.30403 + FAR_Gt_to_Gv * 0.98635) / AF
'                Hэнтальпия = (658.32 + FAR_Gt_to_Gv * 1815.04) / AF + CPтеплоёмкость * (T - 2400.0#)
'                Sэнтропия = (2.1539 + FAR_Gt_to_Gv * 3.1864) / AF + CPтеплоёмкость * Math.Log(T / 2400.0#)
'                AK_показатель_адиабаты = CPтеплоёмкость / (CPтеплоёмкость - RM_газовая_постоянная * AHPA)
'            End If
'        Else
'            If (Math.Abs(FAR_Gt_to_Gv - FARF) > 0.00001 Or Math.Abs(WARвлагосодержание - WARF) > 0.00001) Then
'                WARF = WARвлагосодержание
'                FARF = FAR_Gt_to_Gv
'                AFW = (1.0# + FAR_Gt_to_Gv + WARвлагосодержание)
'                AF = (1.0# + FAR_Gt_to_Gv)

'                If (FAR_Gt_to_Gv = 0.0# And WARвлагосодержание = 0.0#) Then
'                    A1 = AR1
'                    A2 = AR2
'                    A3 = AR3
'                    A4 = AR4
'                    A5 = AR5
'                    A6 = AR6
'                    A7 = AR7
'                    A8 = AR8

'                ElseIf (FAR_Gt_to_Gv = 0.0# And WARвлагосодержание <> 0.0#) Then
'                    A1 = (AR1 + WARвлагосодержание * AW1)
'                    A2 = (AR2 + WARвлагосодержание * AW2)
'                    A3 = (AR3 + WARвлагосодержание * AW3)
'                    A4 = (AR4 + WARвлагосодержание * AW4)
'                    A5 = (AR5 + WARвлагосодержание * AW5)
'                    A6 = (AR6 + WARвлагосодержание * AW6)
'                    A7 = (AR7 + WARвлагосодержание * AW7)
'                    A8 = (AR8 + WARвлагосодержание * AW8)

'                ElseIf (FAR_Gt_to_Gv <> 0.0# And WARвлагосодержание = 0.0#) Then
'                    A1 = (AR1 + FAR_Gt_to_Gv * AF1) / AFW
'                    A2 = (AR2 + FAR_Gt_to_Gv * AF2) / AFW
'                    A3 = (AR3 + FAR_Gt_to_Gv * AF3) / AFW
'                    A4 = (AR4 + FAR_Gt_to_Gv * AF4) / AFW
'                    A5 = (AR5 + FAR_Gt_to_Gv * AF5) / AFW
'                    A6 = (AR6 + FAR_Gt_to_Gv * AF6) / AFW
'                    A7 = (AR7 + FAR_Gt_to_Gv * AF7) / AFW
'                    A8 = (AR8 + FAR_Gt_to_Gv * AF8) / AFW

'                Else
'                    '  dlq  FAR.NE.0.AND.WAR.NE.0
'                    A1 = (AR1 + FAR_Gt_to_Gv * AF1 + WARвлагосодержание * AW1) / AFW
'                    A2 = (AR2 + FAR_Gt_to_Gv * AF2 + WARвлагосодержание * AW2) / AFW
'                    A3 = (AR3 + FAR_Gt_to_Gv * AF3 + WARвлагосодержание * AW3) / AFW
'                    A4 = (AR4 + FAR_Gt_to_Gv * AF4 + WARвлагосодержание * AW4) / AFW
'                    A5 = (AR5 + FAR_Gt_to_Gv * AF5 + WARвлагосодержание * AW5) / AFW
'                    A6 = (AR6 + FAR_Gt_to_Gv * AF6 + WARвлагосодержание * AW6) / AFW
'                    A7 = (AR7 + FAR_Gt_to_Gv * AF7 + WARвлагосодержание * AW7) / AFW
'                    A8 = (AR8 + FAR_Gt_to_Gv * AF8 + WARвлагосодержание * AW8) / AFW
'                End If
'            End If

'            r = T / 1000.0#
'            CPтеплоёмкость = A1 + r * (A2 + r * (A3 + r * (A4 + r * (A5 + r * (A6 + r * (A7 + r * A8))))))
'            AK_показатель_адиабаты = CPтеплоёмкость / (CPтеплоёмкость - RM_газовая_постоянная * AHPA)
'            HA = (-17.742) * FAR_Gt_to_Gv / AF
'            B = r * (A1 + r * (A2 / 2.0# + r * (A3 / 3.0# + r * (A4 / 4.0# + r * (A5 / 5.0# + r * (A6 / 6.0# + r * (A7 / 7.0# + r * A8 / 8.0#)))))))
'            Hэнтальпия = HA + B * 1000.0#
'            SA = A1 * Math.Log(T) + (0.23673 - FAR_Gt_to_Gv * 0.53054) / AF
'            c = r * (A2 + r * (A3 / 2.0# + r * (A4 / 3.0# + r * (A5 / 4.0# + r * (A6 / 5.0# + r * (A7 / 6.0# + r * A8 / 7.0#))))))
'            Sэнтропия = SA + c
'        End If

'        '     BLOCK DATA
'        '     COMMON/ACPF/FLO,HU,CPTO,RAIR,RF,RW,
'        '    *AR1,AR2,AR3,AR4,AR5,AR6,AR7,AR8,
'        '    *AF1,AF2,AF3,AF4,AF5,AF6,AF7,AF8,
'        '    *AW1,AW2,AW3,AW4,AW5,AW6,AW7,AW8
'        '     COMMON/ACPFW/FARF,WARF,RMF
'        '     DATA  FARF/1000./
'        '     DATA  WARF/1000./
'        '     DATA   FLO /14.948 /
'        '     DATA    HU  /10250./
'        '     DATA    CPTO/54.52/
'        '     DATA    RAIR/29.27/
'        '     DATA    RF  /29.428/
'        '     DATA    RW  /47.069/
'        ' KOЭФФИЦИEHTЫ AПPOKCИMAЦИИ CP BOЗДУXA OT TEMПEPATУPЫ
'        '     DATA    AR1/ 0.238785488 E+00/
'        '     DATA    AR2/-0.900857881 E-02/
'        '     DATA    AR3/ 0.193221258 E-01/
'        '     DATA    AR4/ 0.132711466 E+00/
'        '     DATA    AR5/-0.190166916 E+00/
'        '     DATA    AR6/ 0.103793980 E-00/
'        '     DATA    AR7/-0.258490883 E-01/
'        '     DATA    AR8/ 0.245365952 E-02/
'        ' KOЭФФИЦИEHTЫ AПPOKCИMAЦИИ CP ПPOДУKTOB CГOPAHИЯ (ALFA=1)
'        ' OT TEMПEPATУPЫ
'        '     DATA    AF1/  0.367102859 E+00/
'        '     DATA    AF2/  0.384334584 E+00/
'        '     DATA    AF3/  0.190633867 E+00/
'        '     DATA    AF4/ -0.361766882 E+00/
'        '     DATA    AF5/  0.263417998 E+00/
'        '     DATA    AF6/ -0.112618047 E+00/
'        '     DATA    AF7/  0.257271586 E-01/
'        '     DATA    AF8/ -0.237417234 E-02/
'        ' KOЭФФИЦETЫ AПPOKCИMAЦИИ CP BOДЯHOГO ПAPA OT TEMПEPATУPЫ
'        '     DATA    AW1/  0.4700077   E+00/
'        '    DATA    AW2/ -0.3206047   E+00/
'        '   DATA    AW3/  0.1196138   E+01/
'        '   DATA    AW4/ -0.1728544   E+01/
'        '   DATA    AW5/  0.1550520   E+01/
'        '   DATA    AW6/ -0.8291961   E+00/
'        '   DATA    AW7/  0.2370570   E+00/
'        '   DATA    AW8/ -0.2773743   E-01/
'        '   END                                                               
'    End Sub


'    ''' <summary>
'    ''' ПPOГPAMMA  BЫЧИCЛEHИЯ  TEMПEPATУPЫ KAK  ФУHKЦИИ  ЭHTAЛЬПИИ
'    ''' </summary>
'    ''' <param name="far0_Gt_to_Gv"></param>
'    ''' <param name="WARвлагосодержание"></param>
'    ''' <param name="T"></param>
'    ''' <param name="Hэнтальпия"></param>
'    ''' <param name="Sэнтропия"></param>
'    ''' <param name="CPтеплоёмкость"></param>
'    ''' <param name="RM_газовая_постоянная"></param>
'    ''' <param name="AK_показатель_адиабаты"></param>
'    ''' <param name="CS"></param>
'    ''' <remarks></remarks>
'    Private Sub TFH(ByVal far0_Gt_to_Gv As Double, ByVal WARвлагосодержание As Double, ByRef T As Double, ByRef Hэнтальпия As Double, ByRef Sэнтропия As Double, ByRef CPтеплоёмкость As Double, ByRef RM_газовая_постоянная As Double, ByRef AK_показатель_адиабаты As Double, ByRef CS As Double)

'        '       ПPOГPAMMA  BЫЧИCЛEHИЯ  TEMПEPATУPЫ KAK  ФУHKЦИИ  ЭHTAЛЬПИИ

'        '           COMMON/ ACP/FLO,HU,CPTO,RAIR,RF,RW,
'        '     1    AH(131),ACPH(131),AS(131),ACPS(131),
'        '     2    FH(131),FCPH(131),FS(131),FCPS(131),
'        '     3   WH (131), WCPH(131), WS(131), WCPS(131)

'        'Const GST As Double = 9.80665
'        Const AHPA As Double = 0.00234192
'        Dim i, r, n As Integer
'        Dim AFW, HDS, TI, TB, DT, TR, DTD, DRT As Double


'        n = 1
'        AFW = 1.0# + far0_Gt_to_Gv + WARвлагосодержание
'        RM_газовая_постоянная = (RAIR + far0_Gt_to_Gv * (RF + FLO * (RF - RAIR)) + WARвлагосодержание * RW) / AFW

'        If (Hэнтальпия < 52.44) Then
'            Call TFHF(far0_Gt_to_Gv, WARвлагосодержание, T, Hэнтальпия, Sэнтропия, CPтеплоёмкость, RM_газовая_постоянная, AK_показатель_адиабаты, CS)
'        Else
'            If (far0_Gt_to_Gv > 0.0#) Then
'                HDS = (658.32 + far0_Gt_to_Gv * 1815.04) / AFW
'                If (Hэнтальпия > HDS) Then
'                    CPтеплоёмкость = (0.30403 + far0_Gt_to_Gv * 0.98635) / AFW
'                    T = 2400.0# + (Hэнтальпия - HDS) / CPтеплоёмкость
'                Else
'                    HDS = (263.44 + far0_Gt_to_Gv * 590.03) / AFW
'                    If (Hэнтальпия > HDS) Then
'                        T = (Hэнтальпия * AFW + 48.27427 + far0_Gt_to_Gv * 401.6824) / (0.2939106 + far0_Gt_to_Gv * 0.9190792)
'                    Else
'                        T = (Hэнтальпия * AFW + 4.944424 + far0_Gt_to_Gv * 80.67236) / (0.2534112 + far0_Gt_to_Gv * 0.6208404)
'                    End If
'                End If
'            ElseIf (Hэнтальпия > 658.32) Then
'                T = 2400.0# + (Hэнтальпия - 658.32) / 0.30403
'            ElseIf (Hэнтальпия > 263.44) Then
'                T = 164.2481 + 3.402395 * Hэнтальпия
'            Else
'                T = 19.51146 + 3.946155 * Hэнтальпия
'            End If
'            Do
'                If (T > 900.0#) Then
'                    If (T > 2400.0#) Then
'                        i = 131
'                        TI = 2400.0#
'                    Else
'                        TB = 900.0#
'                        DT = 25.0#
'                        TR = (T - TB) / DT
'                        i = Fix(TR)
'                        r = i
'                        TI = (TB + DT * r)
'                        i = i + 71
'                    End If
'                Else
'                    TB = 200.0#
'                    DT = 10.0#
'                    TR = (T - TB) / DT
'                    i = Fix(TR)
'                    r = i
'                    TI = (TB + DT * r)
'                    i = i + 1
'                End If
'                DTD = (Hэнтальпия * AFW - AH(i) - FH(i) * far0_Gt_to_Gv) / (ACPH(i) + FCPH(i) * far0_Gt_to_Gv)
'                T = TI + DTD
'                If ((DTD >= 0.0#) And (DTD <= DT)) Then
'                    Exit Do
'                Else
'                    n = n + 1
'                End If
'            Loop While n <= 2
'            If (CPтеплоёмкость <> -1.0#) Then
'                CPтеплоёмкость = (ACPH(i) + far0_Gt_to_Gv * FCPH(i)) / AFW
'                AK_показатель_адиабаты = CPтеплоёмкость / (CPтеплоёмкость - RM_газовая_постоянная * AHPA)
'            End If
'            If (Sэнтропия <> -1.0#) Then
'                DRT = DTD
'                Sэнтропия = (ASi(i) + DRT * ACPS(i) + (FS(i) + DRT * FCPS(i)) * far0_Gt_to_Gv) / AFW
'            End If
'        End If
'    End Sub

'    ''' <summary>
'    ''' Программа уровня C|p Программа расчета термодинамических характеристик газа
'    ''' </summary>
'    ''' <param name="FAR_Gt_to_Gv"></param>
'    ''' <param name="WARвлагосодержание"></param>
'    ''' <param name="T"></param>
'    ''' <param name="Hэнтальпия"></param>
'    ''' <param name="Sэнтропия"></param>
'    ''' <param name="CPтеплоёмкость"></param>
'    ''' <param name="RM_газовая_постоянная"></param>
'    ''' <param name="AK_показатель_адиабаты"></param>
'    ''' <param name="CS"></param>
'    ''' <remarks></remarks>
'    Private Sub TFHF(ByVal FAR_Gt_to_Gv As Double, ByVal WARвлагосодержание As Double, ByRef T As Double, ByVal Hэнтальпия As Double, ByRef Sэнтропия As Double, ByRef CPтеплоёмкость As Double, ByRef RM_газовая_постоянная As Double, ByRef AK_показатель_адиабаты As Double, ByRef CS As Double)

'        ' Программа уровня C|p Программа расчета термодинамических характеристик газа
'        ' FAR,WAR,H -исходные данные
'        ' T,S,CP    -Результаты вычисления

'        ' Программа уровня C|p
'        ' FAR- отношение расхода топлива к расходу воздуха
'        ' WAR -отношение расхода воды к расходу воздуха
'        ' T  - температура торможения потока, gPAd K
'        ' H  - энтальпия ,KKAl/Kg
'        ' S  - энтропия,KKAl/gPAd K/Kg
'        ' CP - теплоёмкость, KKAl / Kg / gPAd
'        ' RM - газовая постоянная,KgM/Kg/g1
'        ' AK - показатель адиабаты
'        ' CS - скорость звука,M/CEK
'        ' DHDS-интервал неопределенности по энтальпии
'        ' DTDS-интервал неопределенности по температуре

'        '  диапазон температур разбит на 3 поддиапазона :
'        '     1  -T <= 1050
'        '     2  -T > 1050 .AND. T <= 2400
'        '     3  -T > 2400

'        ' используемые подпрограммы CPFT

'        '         blOK b01
'        '  начальные присвоения
'        '
'        CPтеплоёмкость = 1.0#
'        Dim DH As Double = 0.0#
'        Dim X As Double = -1.0#
'        Dim DTDS As Double = 0.01
'        Dim AF, HDS, HCAL, T1 As Double

'        '         blOK  02
'        '  вычисления температуры первого приближения
'        '
'        If (FAR_Gt_to_Gv > 0.0#) Then
'            AF = 1.0# + FAR_Gt_to_Gv
'            HDS = (263.44 + FAR_Gt_to_Gv * 590.03) / AF
'            If (Hэнтальпия <= HDS) Then
'                T = (Hэнтальпия * AF + 4.944424 + FAR_Gt_to_Gv * 80.67236) / (0.2534112 + FAR_Gt_to_Gv * 0.6208404)
'            Else
'                HDS = (658.32 + FAR_Gt_to_Gv * 1815.04) / AF
'                If (Hэнтальпия <= HDS) Then
'                    T = (Hэнтальпия * AF + 48.27427 + FAR_Gt_to_Gv * 401.6824) / (0.2939106 + FAR_Gt_to_Gv * 0.9190792)
'                Else
'                    CPтеплоёмкость = (0.30403 + FAR_Gt_to_Gv * 0.98635) / AF
'                    T = 2400.0# + (Hэнтальпия - HDS) / CPтеплоёмкость
'                End If
'            End If
'        Else
'            FAR_Gt_to_Gv = 0.0#
'            If (Hэнтальпия >= 263.44) Then
'                T = 19.51146 + 3.946155 * Hэнтальпия
'            ElseIf (Hэнтальпия <= 658.32) Then
'                T = 164.2481 + 3.402395 * Hэнтальпия
'            Else
'                T = 2400.0# + (Hэнтальпия - 658.32) / 0.30403
'            End If
'        End If
'        Call CPFTF(FAR_Gt_to_Gv, WARвлагосодержание, T, HCAL, X, CPтеплоёмкость, RM_газовая_постоянная, AK_показатель_адиабаты, CS)
'        Dim n As Integer = 0
'        T1 = T
'        DH = HCAL - Hэнтальпия
'        Do
'            T = T - DH / CPтеплоёмкость
'            n = n + 1
'            If (Math.Abs(T - T1) <= DTDS Or n = 2) Then
'                Exit Do
'            Else
'                Call CPFTF(FAR_Gt_to_Gv, WARвлагосодержание, T, HCAL, X, X, X, X, X)
'                DH = HCAL - Hэнтальпия
'                T1 = T
'            End If
'        Loop
'    End Sub

'    'Private Sub SQINTP(ByVal X, ByVal Y, ByVal n, ByVal A, ByVal F)

'    '    'ПPOГPAMMA  KBAДPATИЧHOЙ  ИHTEPПOЛЯЦИИ

'    '    'Dim A() As Double, F() As Double

'    '    i = 1
'    '    IK = 0
'    '    D2 = 10000000000.0#
'    '    D3 = 10000000000.0#
'    '    Z1 = A(2) - A(1)
'    '    Do
'    '        D1 = A(i) - X
'    '        F1 = F(i)
'    '        If (Z1 >= 0) Then
'    '            D1 = -D1
'    '        End If
'    '        If (IK > 0) Then
'    '            Exit Do
'    '        End If
'    '        If (D1 = 0) Then
'    '            Y = F1
'    '            Exit Sub
'    '        ElseIf (D1 < 0) And ((i - 1) > 0) Then
'    '            IK = 1
'    '            D4 = D3
'    '            F4 = F3
'    '        End If
'    '        i = i + 1
'    '        If ((i - n) > 0) Then
'    '            Exit Do
'    '        End If
'    '        D3 = D2
'    '        D2 = D1
'    '        F3 = F2
'    '        F2 = F1
'    '    Loop
'    '    Y = (D2 * F1 - D1 * F2) / (D2 - D1)
'    '    D4 = (D3 * F2 - D2 * F3) / (D3 - D2)
'    '    Y = (D3 * Y - D1 * D4) / (D3 - D1)
'    'End Sub

'    Private Sub Stoil()

'        '     ПPOГPAMMA  XPAHИT  XAPAKTEPИCTИKИ
'        '   CTAHДAPTHOГO  УГЛEBOДOPOДHOГO  TOПЛИBA


'        AH1(1) = AH1i1
'        AH1(2) = AH1i2
'        AH1(3) = AH1i3
'        AH1(4) = AH1i4
'        AH1(5) = AH1i5
'        AH1(6) = AH1i6
'        AH1(7) = AH1i7
'        AH1(8) = AH1i8
'        AH1(9) = AH1i9
'        AH1(10) = AH1i10
'        AH1(11) = AH1i11
'        AH1(12) = AH1i12
'        AH1(13) = AH1i13
'        AH1(14) = AH1i14
'        AH1(15) = AH1i15
'        AH1(16) = AH1i16
'        AH1(17) = AH1i17
'        AH1(18) = AH1i18
'        AH1(19) = AH1i19
'        AH1(20) = AH1i20
'        AH1(21) = AH1i21
'        AH1(22) = AH1i22
'        AH1(23) = AH1i23
'        AH1(24) = AH1i24
'        AH1(25) = AH1i25
'        AH1(26) = AH1i26
'        AH1(27) = AH1i27
'        AH1(28) = AH1i28
'        AH1(29) = AH1i29
'        AH1(30) = AH1i30
'        AH1(31) = AH1i31
'        AH1(32) = AH1i32
'        AH1(33) = AH1i33
'        AH1(34) = AH1i34
'        AH1(35) = AH1i35
'        AH1(36) = AH1i36
'        AH1(37) = AH1i37
'        AH1(38) = AH1i38
'        AH1(39) = AH1i39
'        AH1(40) = AH1i40
'        AH1(41) = AH1i41
'        AH1(42) = AH1i42
'        AH1(43) = AH1i43
'        AH1(44) = AH1i44
'        AH1(45) = AH1i45
'        AH1(46) = AH1i46
'        AH1(47) = AH1i47
'        AH1(48) = AH1i48
'        AH1(49) = AH1i49
'        AH1(50) = AH1i50
'        AH1(51) = AH1i51
'        AH1(52) = AH1i52
'        AH1(53) = AH1i53
'        AH1(54) = AH1i54
'        AH1(55) = AH1i55
'        AH1(56) = AH1i56
'        AH1(57) = AH1i57
'        AH1(58) = AH1i58
'        AH1(59) = AH1i59
'        AH1(60) = AH1i60
'        AH1(61) = AH1i61
'        AH1(62) = AH1i62
'        AH1(63) = AH1i63
'        AH1(64) = AH1i64
'        AH1(65) = AH1i65
'        AH1(66) = AH1i66
'        AH1(67) = AH1i67
'        AH1(68) = AH1i68
'        AH1(69) = AH1i69
'        AH1(70) = AH1i70
'        AH1(71) = AH1i71
'        AH1(72) = AH1i72
'        AH1(73) = AH1i73
'        AH1(74) = AH1i74
'        AH1(75) = AH1i75
'        AH1(76) = AH1i76


'        AH2(1) = AH2i1
'        AH2(2) = AH2i2
'        AH2(3) = AH2i3
'        AH2(4) = AH2i4
'        AH2(5) = AH2i5
'        AH2(6) = AH2i6
'        AH2(7) = AH2i7
'        AH2(8) = AH2i8
'        AH2(9) = AH2i9
'        AH2(10) = AH2i10
'        AH2(11) = AH2i11
'        AH2(12) = AH2i12
'        AH2(13) = AH2i13
'        AH2(14) = AH2i14
'        AH2(15) = AH2i15
'        AH2(16) = AH2i16
'        AH2(17) = AH2i17
'        AH2(18) = AH2i18
'        AH2(19) = AH2i19
'        AH2(20) = AH2i20
'        AH2(21) = AH2i21
'        AH2(22) = AH2i22
'        AH2(23) = AH2i23
'        AH2(24) = AH2i24
'        AH2(25) = AH2i25
'        AH2(26) = AH2i26
'        AH2(27) = AH2i27
'        AH2(28) = AH2i28
'        AH2(29) = AH2i29
'        AH2(30) = AH2i30
'        AH2(31) = AH2i31
'        AH2(32) = AH2i32
'        AH2(33) = AH2i33
'        AH2(34) = AH2i34
'        AH2(35) = AH2i35
'        AH2(36) = AH2i36
'        AH2(37) = AH2i37
'        AH2(38) = AH2i38
'        AH2(39) = AH2i39
'        AH2(40) = AH2i40
'        AH2(41) = AH2i41
'        AH2(42) = AH2i42
'        AH2(43) = AH2i43
'        AH2(44) = AH2i44
'        AH2(45) = AH2i45
'        AH2(46) = AH2i46
'        AH2(47) = AH2i47
'        AH2(48) = AH2i48
'        AH2(49) = AH2i49
'        AH2(50) = AH2i50
'        AH2(51) = AH2i51
'        AH2(52) = AH2i52
'        AH2(53) = AH2i53
'        AH2(54) = AH2i54
'        AH2(55) = AH2i55


'        ACPH1(1) = ACPH1i1
'        ACPH1(2) = ACPH1i2
'        ACPH1(3) = ACPH1i3
'        ACPH1(4) = ACPH1i4
'        ACPH1(5) = ACPH1i5
'        ACPH1(6) = ACPH1i6
'        ACPH1(7) = ACPH1i7
'        ACPH1(8) = ACPH1i8
'        ACPH1(9) = ACPH1i9
'        ACPH1(10) = ACPH1i10
'        ACPH1(11) = ACPH1i11
'        ACPH1(12) = ACPH1i12
'        ACPH1(13) = ACPH1i13
'        ACPH1(14) = ACPH1i14
'        ACPH1(15) = ACPH1i15
'        ACPH1(16) = ACPH1i16
'        ACPH1(17) = ACPH1i17
'        ACPH1(18) = ACPH1i18
'        ACPH1(19) = ACPH1i19
'        ACPH1(20) = ACPH1i20
'        ACPH1(21) = ACPH1i21
'        ACPH1(22) = ACPH1i22
'        ACPH1(23) = ACPH1i23
'        ACPH1(24) = ACPH1i24
'        ACPH1(25) = ACPH1i25
'        ACPH1(26) = ACPH1i26
'        ACPH1(27) = ACPH1i27
'        ACPH1(28) = ACPH1i28
'        ACPH1(29) = ACPH1i29
'        ACPH1(30) = ACPH1i30
'        ACPH1(31) = ACPH1i31
'        ACPH1(32) = ACPH1i32
'        ACPH1(33) = ACPH1i33
'        ACPH1(34) = ACPH1i34
'        ACPH1(35) = ACPH1i35
'        ACPH1(36) = ACPH1i36
'        ACPH1(37) = ACPH1i37
'        ACPH1(38) = ACPH1i38
'        ACPH1(39) = ACPH1i39
'        ACPH1(40) = ACPH1i40
'        ACPH1(41) = ACPH1i41
'        ACPH1(42) = ACPH1i42
'        ACPH1(43) = ACPH1i43
'        ACPH1(44) = ACPH1i44
'        ACPH1(45) = ACPH1i45
'        ACPH1(46) = ACPH1i46
'        ACPH1(47) = ACPH1i47
'        ACPH1(48) = ACPH1i48
'        ACPH1(49) = ACPH1i49
'        ACPH1(50) = ACPH1i50
'        ACPH1(51) = ACPH1i51
'        ACPH1(52) = ACPH1i52
'        ACPH1(53) = ACPH1i53
'        ACPH1(54) = ACPH1i54
'        ACPH1(55) = ACPH1i55
'        ACPH1(56) = ACPH1i56
'        ACPH1(57) = ACPH1i57
'        ACPH1(58) = ACPH1i58
'        ACPH1(59) = ACPH1i59
'        ACPH1(60) = ACPH1i60
'        ACPH1(61) = ACPH1i61
'        ACPH1(62) = ACPH1i62
'        ACPH1(63) = ACPH1i63
'        ACPH1(64) = ACPH1i64
'        ACPH1(65) = ACPH1i65
'        ACPH1(66) = ACPH1i66
'        ACPH1(67) = ACPH1i67
'        ACPH1(68) = ACPH1i68
'        ACPH1(69) = ACPH1i69
'        ACPH1(70) = ACPH1i70
'        ACPH1(71) = ACPH1i71
'        ACPH1(72) = ACPH1i72
'        ACPH1(73) = ACPH1i73
'        ACPH1(74) = ACPH1i74
'        ACPH1(75) = ACPH1i75
'        ACPH1(76) = ACPH1i76

'        ACPH2(1) = ACPH2i1
'        ACPH2(2) = ACPH2i2
'        ACPH2(3) = ACPH2i3
'        ACPH2(4) = ACPH2i4
'        ACPH2(5) = ACPH2i5
'        ACPH2(6) = ACPH2i6
'        ACPH2(7) = ACPH2i7
'        ACPH2(8) = ACPH2i8
'        ACPH2(9) = ACPH2i9
'        ACPH2(10) = ACPH2i10
'        ACPH2(11) = ACPH2i11
'        ACPH2(12) = ACPH2i12
'        ACPH2(13) = ACPH2i13
'        ACPH2(14) = ACPH2i14
'        ACPH2(15) = ACPH2i15
'        ACPH2(16) = ACPH2i16
'        ACPH2(17) = ACPH2i17
'        ACPH2(18) = ACPH2i18
'        ACPH2(19) = ACPH2i19
'        ACPH2(20) = ACPH2i20
'        ACPH2(21) = ACPH2i21
'        ACPH2(22) = ACPH2i22
'        ACPH2(23) = ACPH2i23
'        ACPH2(24) = ACPH2i24
'        ACPH2(25) = ACPH2i25
'        ACPH2(26) = ACPH2i26
'        ACPH2(27) = ACPH2i27
'        ACPH2(28) = ACPH2i28
'        ACPH2(29) = ACPH2i29
'        ACPH2(30) = ACPH2i30
'        ACPH2(31) = ACPH2i31
'        ACPH2(32) = ACPH2i32
'        ACPH2(33) = ACPH2i33
'        ACPH2(34) = ACPH2i34
'        ACPH2(35) = ACPH2i35
'        ACPH2(36) = ACPH2i36
'        ACPH2(37) = ACPH2i37
'        ACPH2(38) = ACPH2i38
'        ACPH2(39) = ACPH2i39
'        ACPH2(40) = ACPH2i40
'        ACPH2(41) = ACPH2i41
'        ACPH2(42) = ACPH2i42
'        ACPH2(43) = ACPH2i43
'        ACPH2(44) = ACPH2i44
'        ACPH2(45) = ACPH2i45
'        ACPH2(46) = ACPH2i46
'        ACPH2(47) = ACPH2i47
'        ACPH2(48) = ACPH2i48
'        ACPH2(49) = ACPH2i49
'        ACPH2(50) = ACPH2i50
'        ACPH2(51) = ACPH2i51
'        ACPH2(52) = ACPH2i52
'        ACPH2(53) = ACPH2i53
'        ACPH2(54) = ACPH2i54
'        ACPH2(55) = ACPH2i55


'        AS1(1) = AS1i1
'        AS1(2) = AS1i2
'        AS1(3) = AS1i3
'        AS1(4) = AS1i4
'        AS1(5) = AS1i5
'        AS1(6) = AS1i6
'        AS1(7) = AS1i7
'        AS1(8) = AS1i8
'        AS1(9) = AS1i9
'        AS1(10) = AS1i10
'        AS1(11) = AS1i11
'        AS1(12) = AS1i12
'        AS1(13) = AS1i13
'        AS1(14) = AS1i14
'        AS1(15) = AS1i15
'        AS1(16) = AS1i16
'        AS1(17) = AS1i17
'        AS1(18) = AS1i18
'        AS1(19) = AS1i19
'        AS1(20) = AS1i20
'        AS1(21) = AS1i21
'        AS1(22) = AS1i22
'        AS1(23) = AS1i23
'        AS1(24) = AS1i24
'        AS1(25) = AS1i25
'        AS1(26) = AS1i26
'        AS1(27) = AS1i27
'        AS1(28) = AS1i28
'        AS1(29) = AS1i29
'        AS1(30) = AS1i30
'        AS1(31) = AS1i31
'        AS1(32) = AS1i32
'        AS1(33) = AS1i33
'        AS1(34) = AS1i34
'        AS1(35) = AS1i35
'        AS1(36) = AS1i36
'        AS1(37) = AS1i37
'        AS1(38) = AS1i38
'        AS1(39) = AS1i39
'        AS1(40) = AS1i40
'        AS1(41) = AS1i41
'        AS1(42) = AS1i42
'        AS1(43) = AS1i43
'        AS1(44) = AS1i44
'        AS1(45) = AS1i45
'        AS1(46) = AS1i46
'        AS1(47) = AS1i47
'        AS1(48) = AS1i48
'        AS1(49) = AS1i49
'        AS1(50) = AS1i50
'        AS1(51) = AS1i51
'        AS1(52) = AS1i52
'        AS1(53) = AS1i53
'        AS1(54) = AS1i54
'        AS1(55) = AS1i55
'        AS1(56) = AS1i56
'        AS1(57) = AS1i57
'        AS1(58) = AS1i58
'        AS1(59) = AS1i59
'        AS1(60) = AS1i60
'        AS1(61) = AS1i61
'        AS1(62) = AS1i62
'        AS1(63) = AS1i63
'        AS1(64) = AS1i64
'        AS1(65) = AS1i65
'        AS1(66) = AS1i66
'        AS1(67) = AS1i67
'        AS1(68) = AS1i68
'        AS1(69) = AS1i69
'        AS1(70) = AS1i70
'        AS1(71) = AS1i71
'        AS1(72) = AS1i72
'        AS1(73) = AS1i73
'        AS1(74) = AS1i74
'        AS1(75) = AS1i75
'        AS1(76) = AS1i76

'        AS2(1) = AS2i1
'        AS2(2) = AS2i2
'        AS2(3) = AS2i3
'        AS2(4) = AS2i4
'        AS2(5) = AS2i5
'        AS2(6) = AS2i6
'        AS2(7) = AS2i7
'        AS2(8) = AS2i8
'        AS2(9) = AS2i9
'        AS2(10) = AS2i10
'        AS2(11) = AS2i11
'        AS2(12) = AS2i12
'        AS2(13) = AS2i13
'        AS2(14) = AS2i14
'        AS2(15) = AS2i15
'        AS2(16) = AS2i16
'        AS2(17) = AS2i17
'        AS2(18) = AS2i18
'        AS2(19) = AS2i19
'        AS2(20) = AS2i20
'        AS2(21) = AS2i21
'        AS2(22) = AS2i22
'        AS2(23) = AS2i23
'        AS2(24) = AS2i24
'        AS2(25) = AS2i25
'        AS2(26) = AS2i26
'        AS2(27) = AS2i27
'        AS2(28) = AS2i28
'        AS2(29) = AS2i29
'        AS2(30) = AS2i30
'        AS2(31) = AS2i31
'        AS2(32) = AS2i32
'        AS2(33) = AS2i33
'        AS2(34) = AS2i34
'        AS2(35) = AS2i35
'        AS2(36) = AS2i36
'        AS2(37) = AS2i37
'        AS2(38) = AS2i38
'        AS2(39) = AS2i39
'        AS2(40) = AS2i40
'        AS2(41) = AS2i41
'        AS2(42) = AS2i42
'        AS2(43) = AS2i43
'        AS2(44) = AS2i44
'        AS2(45) = AS2i45
'        AS2(46) = AS2i46
'        AS2(47) = AS2i47
'        AS2(48) = AS2i48
'        AS2(49) = AS2i49
'        AS2(50) = AS2i50
'        AS2(51) = AS2i51
'        AS2(52) = AS2i52
'        AS2(53) = AS2i53
'        AS2(54) = AS2i54
'        AS2(55) = AS2i55


'        ACPS1(1) = ACPS1i1
'        ACPS1(2) = ACPS1i2
'        ACPS1(3) = ACPS1i3
'        ACPS1(4) = ACPS1i4
'        ACPS1(5) = ACPS1i5
'        ACPS1(6) = ACPS1i6
'        ACPS1(7) = ACPS1i7
'        ACPS1(8) = ACPS1i8
'        ACPS1(9) = ACPS1i9
'        ACPS1(10) = ACPS1i10
'        ACPS1(11) = ACPS1i11
'        ACPS1(12) = ACPS1i12
'        ACPS1(13) = ACPS1i13
'        ACPS1(14) = ACPS1i14
'        ACPS1(15) = ACPS1i15
'        ACPS1(16) = ACPS1i16
'        ACPS1(17) = ACPS1i17
'        ACPS1(18) = ACPS1i18
'        ACPS1(19) = ACPS1i19
'        ACPS1(20) = ACPS1i20
'        ACPS1(21) = ACPS1i21
'        ACPS1(22) = ACPS1i22
'        ACPS1(23) = ACPS1i23
'        ACPS1(24) = ACPS1i24
'        ACPS1(25) = ACPS1i25
'        ACPS1(26) = ACPS1i26
'        ACPS1(27) = ACPS1i27
'        ACPS1(28) = ACPS1i28
'        ACPS1(29) = ACPS1i29
'        ACPS1(30) = ACPS1i30
'        ACPS1(31) = ACPS1i31
'        ACPS1(32) = ACPS1i32
'        ACPS1(33) = ACPS1i33
'        ACPS1(34) = ACPS1i34
'        ACPS1(35) = ACPS1i35
'        ACPS1(36) = ACPS1i36
'        ACPS1(37) = ACPS1i37
'        ACPS1(38) = ACPS1i38
'        ACPS1(39) = ACPS1i39
'        ACPS1(40) = ACPS1i40
'        ACPS1(41) = ACPS1i41
'        ACPS1(42) = ACPS1i42
'        ACPS1(43) = ACPS1i43
'        ACPS1(44) = ACPS1i44
'        ACPS1(45) = ACPS1i45
'        ACPS1(46) = ACPS1i46
'        ACPS1(47) = ACPS1i47
'        ACPS1(48) = ACPS1i48
'        ACPS1(49) = ACPS1i49
'        ACPS1(50) = ACPS1i50
'        ACPS1(51) = ACPS1i51
'        ACPS1(52) = ACPS1i52
'        ACPS1(53) = ACPS1i53
'        ACPS1(54) = ACPS1i54
'        ACPS1(55) = ACPS1i55
'        ACPS1(56) = ACPS1i56
'        ACPS1(57) = ACPS1i57
'        ACPS1(58) = ACPS1i58
'        ACPS1(59) = ACPS1i59
'        ACPS1(60) = ACPS1i60
'        ACPS1(61) = ACPS1i61
'        ACPS1(62) = ACPS1i62
'        ACPS1(63) = ACPS1i63
'        ACPS1(64) = ACPS1i64
'        ACPS1(65) = ACPS1i65
'        ACPS1(66) = ACPS1i66
'        ACPS1(67) = ACPS1i67
'        ACPS1(68) = ACPS1i68
'        ACPS1(69) = ACPS1i69
'        ACPS1(70) = ACPS1i70
'        ACPS1(71) = ACPS1i71
'        ACPS1(72) = ACPS1i72
'        ACPS1(73) = ACPS1i73
'        ACPS1(74) = ACPS1i74
'        ACPS1(75) = ACPS1i75
'        ACPS1(76) = ACPS1i76

'        ACPS2(1) = ACPS2i1
'        ACPS2(2) = ACPS2i2
'        ACPS2(3) = ACPS2i3
'        ACPS2(4) = ACPS2i4
'        ACPS2(5) = ACPS2i5
'        ACPS2(6) = ACPS2i6
'        ACPS2(7) = ACPS2i7
'        ACPS2(8) = ACPS2i8
'        ACPS2(9) = ACPS2i9
'        ACPS2(10) = ACPS2i10
'        ACPS2(11) = ACPS2i11
'        ACPS2(12) = ACPS2i12
'        ACPS2(13) = ACPS2i13
'        ACPS2(14) = ACPS2i14
'        ACPS2(15) = ACPS2i15
'        ACPS2(16) = ACPS2i16
'        ACPS2(17) = ACPS2i17
'        ACPS2(18) = ACPS2i18
'        ACPS2(19) = ACPS2i19
'        ACPS2(20) = ACPS2i20
'        ACPS2(21) = ACPS2i21
'        ACPS2(22) = ACPS2i22
'        ACPS2(23) = ACPS2i23
'        ACPS2(24) = ACPS2i24
'        ACPS2(25) = ACPS2i25
'        ACPS2(26) = ACPS2i26
'        ACPS2(27) = ACPS2i27
'        ACPS2(28) = ACPS2i28
'        ACPS2(29) = ACPS2i29
'        ACPS2(30) = ACPS2i30
'        ACPS2(31) = ACPS2i31
'        ACPS2(32) = ACPS2i32
'        ACPS2(33) = ACPS2i33
'        ACPS2(34) = ACPS2i34
'        ACPS2(35) = ACPS2i35
'        ACPS2(36) = ACPS2i36
'        ACPS2(37) = ACPS2i37
'        ACPS2(38) = ACPS2i38
'        ACPS2(39) = ACPS2i39
'        ACPS2(40) = ACPS2i40
'        ACPS2(41) = ACPS2i41
'        ACPS2(42) = ACPS2i42
'        ACPS2(43) = ACPS2i43
'        ACPS2(44) = ACPS2i44
'        ACPS2(45) = ACPS2i45
'        ACPS2(46) = ACPS2i46
'        ACPS2(47) = ACPS2i47
'        ACPS2(48) = ACPS2i48
'        ACPS2(49) = ACPS2i49
'        ACPS2(50) = ACPS2i50
'        ACPS2(51) = ACPS2i51
'        ACPS2(52) = ACPS2i52
'        ACPS2(53) = ACPS2i53
'        ACPS2(54) = ACPS2i54
'        ACPS2(55) = ACPS2i55

'        FH1(1) = FH1i1
'        FH1(2) = FH1i2
'        FH1(3) = FH1i3
'        FH1(4) = FH1i4
'        FH1(5) = FH1i5
'        FH1(6) = FH1i6
'        FH1(7) = FH1i7
'        FH1(8) = FH1i8
'        FH1(9) = FH1i9
'        FH1(10) = FH1i10
'        FH1(11) = FH1i11
'        FH1(12) = FH1i12
'        FH1(13) = FH1i13
'        FH1(14) = FH1i14
'        FH1(15) = FH1i15
'        FH1(16) = FH1i16
'        FH1(17) = FH1i17
'        FH1(18) = FH1i18
'        FH1(19) = FH1i19
'        FH1(20) = FH1i20
'        FH1(21) = FH1i21
'        FH1(22) = FH1i22
'        FH1(23) = FH1i23
'        FH1(24) = FH1i24
'        FH1(25) = FH1i25
'        FH1(26) = FH1i26
'        FH1(27) = FH1i27
'        FH1(28) = FH1i28
'        FH1(29) = FH1i29
'        FH1(30) = FH1i30
'        FH1(31) = FH1i31
'        FH1(32) = FH1i32
'        FH1(33) = FH1i33
'        FH1(34) = FH1i34
'        FH1(35) = FH1i35
'        FH1(36) = FH1i36
'        FH1(37) = FH1i37
'        FH1(38) = FH1i38
'        FH1(39) = FH1i39
'        FH1(40) = FH1i40
'        FH1(41) = FH1i41
'        FH1(42) = FH1i42
'        FH1(43) = FH1i43
'        FH1(44) = FH1i44
'        FH1(45) = FH1i45
'        FH1(46) = FH1i46
'        FH1(47) = FH1i47
'        FH1(48) = FH1i48
'        FH1(49) = FH1i49
'        FH1(50) = FH1i50
'        FH1(51) = FH1i51
'        FH1(52) = FH1i52
'        FH1(53) = FH1i53
'        FH1(54) = FH1i54
'        FH1(55) = FH1i55
'        FH1(56) = FH1i56
'        FH1(57) = FH1i57
'        FH1(58) = FH1i58
'        FH1(59) = FH1i59
'        FH1(60) = FH1i60
'        FH1(61) = FH1i61
'        FH1(62) = FH1i62
'        FH1(63) = FH1i63
'        FH1(64) = FH1i64
'        FH1(65) = FH1i65
'        FH1(66) = FH1i66
'        FH1(67) = FH1i67
'        FH1(68) = FH1i68
'        FH1(69) = FH1i69
'        FH1(70) = FH1i70
'        FH1(71) = FH1i71
'        FH1(72) = FH1i72
'        FH1(73) = FH1i73
'        FH1(74) = FH1i74
'        FH1(75) = FH1i75
'        FH1(76) = FH1i76

'        FH2(1) = FH2i1
'        FH2(2) = FH2i2
'        FH2(3) = FH2i3
'        FH2(4) = FH2i4
'        FH2(5) = FH2i5
'        FH2(6) = FH2i6
'        FH2(7) = FH2i7
'        FH2(8) = FH2i8
'        FH2(9) = FH2i9
'        FH2(10) = FH2i10
'        FH2(11) = FH2i11
'        FH2(12) = FH2i12
'        FH2(13) = FH2i13
'        FH2(14) = FH2i14
'        FH2(15) = FH2i15
'        FH2(16) = FH2i16
'        FH2(17) = FH2i17
'        FH2(18) = FH2i18
'        FH2(19) = FH2i19
'        FH2(20) = FH2i20
'        FH2(21) = FH2i21
'        FH2(22) = FH2i22
'        FH2(23) = FH2i23
'        FH2(24) = FH2i24
'        FH2(25) = FH2i25
'        FH2(26) = FH2i26
'        FH2(27) = FH2i27
'        FH2(28) = FH2i28
'        FH2(29) = FH2i29
'        FH2(30) = FH2i30
'        FH2(31) = FH2i31
'        FH2(32) = FH2i32
'        FH2(33) = FH2i33
'        FH2(34) = FH2i34
'        FH2(35) = FH2i35
'        FH2(36) = FH2i36
'        FH2(37) = FH2i37
'        FH2(38) = FH2i38
'        FH2(39) = FH2i39
'        FH2(40) = FH2i40
'        FH2(41) = FH2i41
'        FH2(42) = FH2i42
'        FH2(43) = FH2i43
'        FH2(44) = FH2i44
'        FH2(45) = FH2i45
'        FH2(46) = FH2i46
'        FH2(47) = FH2i47
'        FH2(48) = FH2i48
'        FH2(49) = FH2i49
'        FH2(50) = FH2i50
'        FH2(51) = FH2i51
'        FH2(52) = FH2i52
'        FH2(53) = FH2i53
'        FH2(54) = FH2i54
'        FH2(55) = FH2i55

'        FCPH1(1) = FCPH1i1
'        FCPH1(2) = FCPH1i2
'        FCPH1(3) = FCPH1i3
'        FCPH1(4) = FCPH1i4
'        FCPH1(5) = FCPH1i5
'        FCPH1(6) = FCPH1i6
'        FCPH1(7) = FCPH1i7
'        FCPH1(8) = FCPH1i8
'        FCPH1(9) = FCPH1i9
'        FCPH1(10) = FCPH1i10
'        FCPH1(11) = FCPH1i11
'        FCPH1(12) = FCPH1i12
'        FCPH1(13) = FCPH1i13
'        FCPH1(14) = FCPH1i14
'        FCPH1(15) = FCPH1i15
'        FCPH1(16) = FCPH1i16
'        FCPH1(17) = FCPH1i17
'        FCPH1(18) = FCPH1i18
'        FCPH1(19) = FCPH1i19
'        FCPH1(20) = FCPH1i20
'        FCPH1(21) = FCPH1i21
'        FCPH1(22) = FCPH1i22
'        FCPH1(23) = FCPH1i23
'        FCPH1(24) = FCPH1i24
'        FCPH1(25) = FCPH1i25
'        FCPH1(26) = FCPH1i26
'        FCPH1(27) = FCPH1i27
'        FCPH1(28) = FCPH1i28
'        FCPH1(29) = FCPH1i29
'        FCPH1(30) = FCPH1i30
'        FCPH1(31) = FCPH1i31
'        FCPH1(32) = FCPH1i32
'        FCPH1(33) = FCPH1i33
'        FCPH1(34) = FCPH1i34
'        FCPH1(35) = FCPH1i35
'        FCPH1(36) = FCPH1i36
'        FCPH1(37) = FCPH1i37
'        FCPH1(38) = FCPH1i38
'        FCPH1(39) = FCPH1i39
'        FCPH1(40) = FCPH1i40
'        FCPH1(41) = FCPH1i41
'        FCPH1(42) = FCPH1i42
'        FCPH1(43) = FCPH1i43
'        FCPH1(44) = FCPH1i44
'        FCPH1(45) = FCPH1i45
'        FCPH1(46) = FCPH1i46
'        FCPH1(47) = FCPH1i47
'        FCPH1(48) = FCPH1i48
'        FCPH1(49) = FCPH1i49
'        FCPH1(50) = FCPH1i50
'        FCPH1(51) = FCPH1i51
'        FCPH1(52) = FCPH1i52
'        FCPH1(53) = FCPH1i53
'        FCPH1(54) = FCPH1i54
'        FCPH1(55) = FCPH1i55
'        FCPH1(56) = FCPH1i56
'        FCPH1(57) = FCPH1i57
'        FCPH1(58) = FCPH1i58
'        FCPH1(59) = FCPH1i59
'        FCPH1(60) = FCPH1i60
'        FCPH1(61) = FCPH1i61
'        FCPH1(62) = FCPH1i62
'        FCPH1(63) = FCPH1i63
'        FCPH1(64) = FCPH1i64
'        FCPH1(65) = FCPH1i65
'        FCPH1(66) = FCPH1i66
'        FCPH1(67) = FCPH1i67
'        FCPH1(68) = FCPH1i68
'        FCPH1(69) = FCPH1i69
'        FCPH1(70) = FCPH1i70
'        FCPH1(71) = FCPH1i71
'        FCPH1(72) = FCPH1i72
'        FCPH1(73) = FCPH1i73
'        FCPH1(74) = FCPH1i74
'        FCPH1(75) = FCPH1i75
'        FCPH1(76) = FCPH1i76

'        FCPH2(1) = FCPH2i1
'        FCPH2(2) = FCPH2i2
'        FCPH2(3) = FCPH2i3
'        FCPH2(4) = FCPH2i4
'        FCPH2(5) = FCPH2i5
'        FCPH2(6) = FCPH2i6
'        FCPH2(7) = FCPH2i7
'        FCPH2(8) = FCPH2i8
'        FCPH2(9) = FCPH2i9
'        FCPH2(10) = FCPH2i10
'        FCPH2(11) = FCPH2i11
'        FCPH2(12) = FCPH2i12
'        FCPH2(13) = FCPH2i13
'        FCPH2(14) = FCPH2i14
'        FCPH2(15) = FCPH2i15
'        FCPH2(16) = FCPH2i16
'        FCPH2(17) = FCPH2i17
'        FCPH2(18) = FCPH2i18
'        FCPH2(19) = FCPH2i19
'        FCPH2(20) = FCPH2i20
'        FCPH2(21) = FCPH2i21
'        FCPH2(22) = FCPH2i22
'        FCPH2(23) = FCPH2i23
'        FCPH2(24) = FCPH2i24
'        FCPH2(25) = FCPH2i25
'        FCPH2(26) = FCPH2i26
'        FCPH2(27) = FCPH2i27
'        FCPH2(28) = FCPH2i28
'        FCPH2(29) = FCPH2i29
'        FCPH2(30) = FCPH2i30
'        FCPH2(31) = FCPH2i31
'        FCPH2(32) = FCPH2i32
'        FCPH2(33) = FCPH2i33
'        FCPH2(34) = FCPH2i34
'        FCPH2(35) = FCPH2i35
'        FCPH2(36) = FCPH2i36
'        FCPH2(37) = FCPH2i37
'        FCPH2(38) = FCPH2i38
'        FCPH2(39) = FCPH2i39
'        FCPH2(40) = FCPH2i40
'        FCPH2(41) = FCPH2i41
'        FCPH2(42) = FCPH2i42
'        FCPH2(43) = FCPH2i43
'        FCPH2(44) = FCPH2i44
'        FCPH2(45) = FCPH2i45
'        FCPH2(46) = FCPH2i46
'        FCPH2(47) = FCPH2i47
'        FCPH2(48) = FCPH2i48
'        FCPH2(49) = FCPH2i49
'        FCPH2(50) = FCPH2i50
'        FCPH2(51) = FCPH2i51
'        FCPH2(52) = FCPH2i52
'        FCPH2(53) = FCPH2i53
'        FCPH2(54) = FCPH2i54
'        FCPH2(55) = FCPH2i55

'        FS1(1) = FS1i1
'        FS1(2) = FS1i2
'        FS1(3) = FS1i3
'        FS1(4) = FS1i4
'        FS1(5) = FS1i5
'        FS1(6) = FS1i6
'        FS1(7) = FS1i7
'        FS1(8) = FS1i8
'        FS1(9) = FS1i9
'        FS1(10) = FS1i10
'        FS1(11) = FS1i11
'        FS1(12) = FS1i12
'        FS1(13) = FS1i13
'        FS1(14) = FS1i14
'        FS1(15) = FS1i15
'        FS1(16) = FS1i16
'        FS1(17) = FS1i17
'        FS1(18) = FS1i18
'        FS1(19) = FS1i19
'        FS1(20) = FS1i20
'        FS1(21) = FS1i21
'        FS1(22) = FS1i22
'        FS1(23) = FS1i23
'        FS1(24) = FS1i24
'        FS1(25) = FS1i25
'        FS1(26) = FS1i26
'        FS1(27) = FS1i27
'        FS1(28) = FS1i28
'        FS1(29) = FS1i29
'        FS1(30) = FS1i30
'        FS1(31) = FS1i31
'        FS1(32) = FS1i32
'        FS1(33) = FS1i33
'        FS1(34) = FS1i34
'        FS1(35) = FS1i35
'        FS1(36) = FS1i36
'        FS1(37) = FS1i37
'        FS1(38) = FS1i38
'        FS1(39) = FS1i39
'        FS1(40) = FS1i40
'        FS1(41) = FS1i41
'        FS1(42) = FS1i42
'        FS1(43) = FS1i43
'        FS1(44) = FS1i44
'        FS1(45) = FS1i45
'        FS1(46) = FS1i46
'        FS1(47) = FS1i47
'        FS1(48) = FS1i48
'        FS1(49) = FS1i49
'        FS1(50) = FS1i50
'        FS1(51) = FS1i51
'        FS1(52) = FS1i52
'        FS1(53) = FS1i53
'        FS1(54) = FS1i54
'        FS1(55) = FS1i55
'        FS1(56) = FS1i56
'        FS1(57) = FS1i57
'        FS1(58) = FS1i58
'        FS1(59) = FS1i59
'        FS1(60) = FS1i60
'        FS1(61) = FS1i61
'        FS1(62) = FS1i62
'        FS1(63) = FS1i63
'        FS1(64) = FS1i64
'        FS1(65) = FS1i65
'        FS1(66) = FS1i66
'        FS1(67) = FS1i67
'        FS1(68) = FS1i68
'        FS1(69) = FS1i69
'        FS1(70) = FS1i70
'        FS1(71) = FS1i71
'        FS1(72) = FS1i72
'        FS1(73) = FS1i73
'        FS1(74) = FS1i74
'        FS1(75) = FS1i75
'        FS1(76) = FS1i76

'        FS2(1) = FS2i1
'        FS2(2) = FS2i2
'        FS2(3) = FS2i3
'        FS2(4) = FS2i4
'        FS2(5) = FS2i5
'        FS2(6) = FS2i6
'        FS2(7) = FS2i7
'        FS2(8) = FS2i8
'        FS2(9) = FS2i9
'        FS2(10) = FS2i10
'        FS2(11) = FS2i11
'        FS2(12) = FS2i12
'        FS2(13) = FS2i13
'        FS2(14) = FS2i14
'        FS2(15) = FS2i15
'        FS2(16) = FS2i16
'        FS2(17) = FS2i17
'        FS2(18) = FS2i18
'        FS2(19) = FS2i19
'        FS2(20) = FS2i20
'        FS2(21) = FS2i21
'        FS2(22) = FS2i22
'        FS2(23) = FS2i23
'        FS2(24) = FS2i24
'        FS2(25) = FS2i25
'        FS2(26) = FS2i26
'        FS2(27) = FS2i27
'        FS2(28) = FS2i28
'        FS2(29) = FS2i29
'        FS2(30) = FS2i30
'        FS2(31) = FS2i31
'        FS2(32) = FS2i32
'        FS2(33) = FS2i33
'        FS2(34) = FS2i34
'        FS2(35) = FS2i35
'        FS2(36) = FS2i36
'        FS2(37) = FS2i37
'        FS2(38) = FS2i38
'        FS2(39) = FS2i39
'        FS2(40) = FS2i40
'        FS2(41) = FS2i41
'        FS2(42) = FS2i42
'        FS2(43) = FS2i43
'        FS2(44) = FS2i44
'        FS2(45) = FS2i45
'        FS2(46) = FS2i46
'        FS2(47) = FS2i47
'        FS2(48) = FS2i48
'        FS2(49) = FS2i49
'        FS2(50) = FS2i50
'        FS2(51) = FS2i51
'        FS2(52) = FS2i52
'        FS2(53) = FS2i53
'        FS2(54) = FS2i54
'        FS2(55) = FS2i55

'        FCPS1(1) = FCPS1i1
'        FCPS1(2) = FCPS1i2
'        FCPS1(3) = FCPS1i3
'        FCPS1(4) = FCPS1i4
'        FCPS1(5) = FCPS1i5
'        FCPS1(6) = FCPS1i6
'        FCPS1(7) = FCPS1i7
'        FCPS1(8) = FCPS1i8
'        FCPS1(9) = FCPS1i9
'        FCPS1(10) = FCPS1i10
'        FCPS1(11) = FCPS1i11
'        FCPS1(12) = FCPS1i12
'        FCPS1(13) = FCPS1i13
'        FCPS1(14) = FCPS1i14
'        FCPS1(15) = FCPS1i15
'        FCPS1(16) = FCPS1i16
'        FCPS1(17) = FCPS1i17
'        FCPS1(18) = FCPS1i18
'        FCPS1(19) = FCPS1i19
'        FCPS1(20) = FCPS1i20
'        FCPS1(21) = FCPS1i21
'        FCPS1(22) = FCPS1i22
'        FCPS1(23) = FCPS1i23
'        FCPS1(24) = FCPS1i24
'        FCPS1(25) = FCPS1i25
'        FCPS1(26) = FCPS1i26
'        FCPS1(27) = FCPS1i27
'        FCPS1(28) = FCPS1i28
'        FCPS1(29) = FCPS1i29
'        FCPS1(30) = FCPS1i30
'        FCPS1(31) = FCPS1i31
'        FCPS1(32) = FCPS1i32
'        FCPS1(33) = FCPS1i33
'        FCPS1(34) = FCPS1i34
'        FCPS1(35) = FCPS1i35
'        FCPS1(36) = FCPS1i36
'        FCPS1(37) = FCPS1i37
'        FCPS1(38) = FCPS1i38
'        FCPS1(39) = FCPS1i39
'        FCPS1(40) = FCPS1i40
'        FCPS1(41) = FCPS1i41
'        FCPS1(42) = FCPS1i42
'        FCPS1(43) = FCPS1i43
'        FCPS1(44) = FCPS1i44
'        FCPS1(45) = FCPS1i45
'        FCPS1(46) = FCPS1i46
'        FCPS1(47) = FCPS1i47
'        FCPS1(48) = FCPS1i48
'        FCPS1(49) = FCPS1i49
'        FCPS1(50) = FCPS1i50
'        FCPS1(51) = FCPS1i51
'        FCPS1(52) = FCPS1i52
'        FCPS1(53) = FCPS1i53
'        FCPS1(54) = FCPS1i54
'        FCPS1(55) = FCPS1i55
'        FCPS1(56) = FCPS1i56
'        FCPS1(57) = FCPS1i57
'        FCPS1(58) = FCPS1i58
'        FCPS1(59) = FCPS1i59
'        FCPS1(60) = FCPS1i60
'        FCPS1(61) = FCPS1i61
'        FCPS1(62) = FCPS1i62
'        FCPS1(63) = FCPS1i63
'        FCPS1(64) = FCPS1i64
'        FCPS1(65) = FCPS1i65
'        FCPS1(66) = FCPS1i66
'        FCPS1(67) = FCPS1i67
'        FCPS1(68) = FCPS1i68
'        FCPS1(69) = FCPS1i69
'        FCPS1(70) = FCPS1i70
'        FCPS1(71) = FCPS1i71
'        FCPS1(72) = FCPS1i72
'        FCPS1(73) = FCPS1i73
'        FCPS1(74) = FCPS1i74
'        FCPS1(75) = FCPS1i75
'        FCPS1(76) = FCPS1i76

'        FCPS2(1) = FCPS2i1
'        FCPS2(2) = FCPS2i2
'        FCPS2(3) = FCPS2i3
'        FCPS2(4) = FCPS2i4
'        FCPS2(5) = FCPS2i5
'        FCPS2(6) = FCPS2i6
'        FCPS2(7) = FCPS2i7
'        FCPS2(8) = FCPS2i8
'        FCPS2(9) = FCPS2i9
'        FCPS2(10) = FCPS2i10
'        FCPS2(11) = FCPS2i11
'        FCPS2(12) = FCPS2i12
'        FCPS2(13) = FCPS2i13
'        FCPS2(14) = FCPS2i14
'        FCPS2(15) = FCPS2i15
'        FCPS2(16) = FCPS2i16
'        FCPS2(17) = FCPS2i17
'        FCPS2(18) = FCPS2i18
'        FCPS2(19) = FCPS2i19
'        FCPS2(20) = FCPS2i20
'        FCPS2(21) = FCPS2i21
'        FCPS2(22) = FCPS2i22
'        FCPS2(23) = FCPS2i23
'        FCPS2(24) = FCPS2i24
'        FCPS2(25) = FCPS2i25
'        FCPS2(26) = FCPS2i26
'        FCPS2(27) = FCPS2i27
'        FCPS2(28) = FCPS2i28
'        FCPS2(29) = FCPS2i29
'        FCPS2(30) = FCPS2i30
'        FCPS2(31) = FCPS2i31
'        FCPS2(32) = FCPS2i32
'        FCPS2(33) = FCPS2i33
'        FCPS2(34) = FCPS2i34
'        FCPS2(35) = FCPS2i35
'        FCPS2(36) = FCPS2i36
'        FCPS2(37) = FCPS2i37
'        FCPS2(38) = FCPS2i38
'        FCPS2(39) = FCPS2i39
'        FCPS2(40) = FCPS2i40
'        FCPS2(41) = FCPS2i41
'        FCPS2(42) = FCPS2i42
'        FCPS2(43) = FCPS2i43
'        FCPS2(44) = FCPS2i44
'        FCPS2(45) = FCPS2i45
'        FCPS2(46) = FCPS2i46
'        FCPS2(47) = FCPS2i47
'        FCPS2(48) = FCPS2i48
'        FCPS2(49) = FCPS2i49
'        FCPS2(50) = FCPS2i50
'        FCPS2(51) = FCPS2i51
'        FCPS2(52) = FCPS2i52
'        FCPS2(53) = FCPS2i53
'        FCPS2(54) = FCPS2i54
'        FCPS2(55) = FCPS2i55

'        Dim i As Integer = 1
'        Dim j As Integer = 1
'        Do While i <= 76
'            AH(i) = AH1(i)
'            ACPH(i) = ACPH1(i)
'            ASi(i) = AS1(i)
'            ACPS(i) = ACPS1(i)
'            FH(i) = FH1(i)
'            FCPH(i) = FCPH1(i)
'            FS(i) = FS1(i)
'            FCPS(i) = FCPS1(i)
'            i = i + 1
'        Loop
'        Do While i > 76 And i <= 131 And j <= 55
'            AH(i) = AH2(j)
'            ACPH(i) = ACPH2(j)
'            ASi(i) = AS2(j)
'            ACPS(i) = ACPS2(j)
'            FH(i) = FH2(j)
'            FCPH(i) = FCPH2(j)
'            FS(i) = FS2(j)
'            FCPS(i) = FCPS2(j)
'            i = i + 1
'            j = j + 1
'        Loop
'    End Sub

'    'Sub CleanCells(ByVal Start, ByVal Quant, ByVal NameSheetI, ByVal RepitAdr)

'    '    Dim CurPar As Range

'    '    For j = Start To Start + Quant
'    '        SearchWithinColumn(NameSheetI, j, 1, StrLog, Repetition, RepitAdr)
'    '        If Repetition > 0 Then
'    '            For j1 = 1 To Repetition
'    '                CurPar = ThisWorkbook.Sheets(NameSheetI).Range(RepitAdr(j1))
'    '                For i = 1 To UBound(Param(j).Val, 1)
'    '                    ' Определение отступа от идентификатора параметра
'    '                    If IsNumeric(CurPar.Cells(1, 2).Value) = True And Len(CurPar.Cells(1, 2).Value) > 0 Then
'    '                        indent = 1
'    '                    Else
'    '                        indent = 2
'    '                    End If

'    '                    CurPar.Cells(1, i + indent).Value = Empty

'    '                Next
'    '            Next
'    '        End If
'    '    Next
'    'End Sub

'    'Sub DataOutput(ByVal Start, ByVal Quant, ByVal NameSheetI)

'    '    Dim LimitEmptyRow, EmptyRow, LengthStr As Integer
'    '    Dim RefAdr As String


'    '    Dim NotFoundParam As Boolean
'    '    Dim StartSearch, PrintLocation As Range

'    '    RowLimit = 65536
'    '    'rowlimit = 130
'    '    LimitEmptyRow = 10

'    '    For j = Start To Start + Quant
'    '        SearchWithinColumn(NameSheetI, j, 1, StrLog, Repetition, RepiatAdr)
'    '        If Repetition = 0 Then

'    '            If NotFoundParam = False Then
'    '                NotFoundParam = True
'    '                StartSearch = ThisWorkbook.Sheets(NameSheetI).Range("A1")
'    '                PrintLocation = StartSearch

'    '                'Поиск диапазона, состоящего из количества LimitEmptyRow пустых ячеек, в столбце "A".
'    '                i = 1
'    '                Do Until EmptyRow = LimitEmptyRow Or i > RowLimit
'    '                    '      rrr = PrintLocation.Rows
'    '                    LengthStr = Len(PrintLocation.Value)
'    '                    If LengthStr = 0 Then
'    '                        EmptyRow = EmptyRow + 1
'    '                    Else
'    '                        EmptyRow = 0
'    '                    End If
'    '                    i = i + 1
'    '                    If i <= RowLimit And EmptyRow < 10 Then PrintLocation = StartSearch.Offset(i - 1, 0)
'    '                    '      r = PrintLocation.Address
'    '                Loop

'    '                If EmptyRow < LimitEmptyRow Then
'    '                    Reply = MsgBox("Лист """ & NameSheetI & """, параметр """ & Param(j).NamePar & _
'    '                          """ не найден. Выполнение программы прервано")
'    '                    FatalErr = True
'    '                    Exit Sub
'    '                End If

'    '                'выделить
'    '                PrintLocation.Select()
'    '                With Selection.Interior
'    '                    .ColorIndex = 3
'    '                    .Pattern = xlSolid
'    '                    .PatternColorIndex = xlAutomatic
'    '                End With

'    '                RefAdr = PrintLocation.Address
'    '                Reply = MsgBox("Лист """ & NameSheetI & """, параметр """ & Param(j).NamePar & _
'    '                      """ не найден. Значения параметра """ & Param(j).NamePar & _
'    '                      """, а также последующих ненайденных параметров, будут отображены на листе """ & _
'    '                      NameSheetI & """ начиная с ячейки """ & RefAdr & """. Продолжить выполнение программы?", _
'    '                      vbYesNo)
'    '                'убрать выделение
'    '                With Selection.Interior
'    '                    .ColorIndex = xlNone
'    '                    .Pattern = xlNone
'    '                    .PatternColorIndex = xlNone
'    '                End With
'    '                If Reply = vbYes Then ' продолжить выполнение программы
'    '                    Param(j).ValLocation = PrintLocation
'    '                    '       fi = 1
'    '                    Param(j).ValLocation.Value = Param(j).NamePar
'    '                    For fi = 1 To Len(Param(j).NamePar)

'    '                        'Param(j).ValLocation.Characters(fi, 1).Font.Parent = Param(j).NameLocation.Characters(fi, 1).Font.Parent
'    '                        With Param(j)

'    '                            .ValLocation.Characters(fi, 1).Font.Name = .NameLocation.Characters(fi, 1).Font.Name
'    '                            .ValLocation.Characters(fi, 1).Font.FontStyle = .NameLocation.Characters(fi, 1).Font.FontStyle
'    '                            .ValLocation.Characters(fi, 1).Font.Size = .NameLocation.Characters(fi, 1).Font.Size
'    '                            .ValLocation.Characters(fi, 1).Font.Strikethrough = .NameLocation.Characters(fi, 1).Font.Strikethrough
'    '                            .ValLocation.Characters(fi, 1).Font.Superscript = .NameLocation.Characters(fi, 1).Font.Superscript
'    '                            .ValLocation.Characters(fi, 1).Font.Subscript = .NameLocation.Characters(fi, 1).Font.Subscript
'    '                            .ValLocation.Characters(fi, 1).Font.OutlineFont = .NameLocation.Characters(fi, 1).Font.OutlineFont
'    '                            .ValLocation.Characters(fi, 1).Font.Shadow = .NameLocation.Characters(fi, 1).Font.Shadow
'    '                            .ValLocation.Characters(fi, 1).Font.Underline = .NameLocation.Characters(fi, 1).Font.Underline
'    '                            .ValLocation.Characters(fi, 1).Font.ColorIndex = .NameLocation.Characters(fi, 1).Font.ColorIndex
'    '                        End With
'    '                    Next
'    '                ElseIf Reply = vbNo Then ' прервать выполнение программы
'    '                    FatalErr = True
'    '                    Exit Sub
'    '                End If

'    '            Else
'    '                If PrintLocation.Row < RowLimit Then
'    '                    Param(j).ValLocation = PrintLocation.Offset(1, 0)
'    '                    PrintLocation = Param(j).ValLocation
'    '                    Param(j).ValLocation.Value = Param(j).NamePar
'    '                    For fi = 1 To Len(Param(j).NamePar)
'    '                        With Param(j)
'    '                            .ValLocation.Characters(fi, 1).Font.Name = .NameLocation.Characters(fi, 1).Font.Name
'    '                            .ValLocation.Characters(fi, 1).Font.FontStyle = .NameLocation.Characters(fi, 1).Font.FontStyle
'    '                            .ValLocation.Characters(fi, 1).Font.Size = .NameLocation.Characters(fi, 1).Font.Size
'    '                            .ValLocation.Characters(fi, 1).Font.Strikethrough = .NameLocation.Characters(fi, 1).Font.Strikethrough
'    '                            .ValLocation.Characters(fi, 1).Font.Superscript = .NameLocation.Characters(fi, 1).Font.Superscript
'    '                            .ValLocation.Characters(fi, 1).Font.Subscript = .NameLocation.Characters(fi, 1).Font.Subscript
'    '                            .ValLocation.Characters(fi, 1).Font.OutlineFont = .NameLocation.Characters(fi, 1).Font.OutlineFont
'    '                            .ValLocation.Characters(fi, 1).Font.Shadow = .NameLocation.Characters(fi, 1).Font.Shadow
'    '                            .ValLocation.Characters(fi, 1).Font.Underline = .NameLocation.Characters(fi, 1).Font.Underline
'    '                            .ValLocation.Characters(fi, 1).Font.ColorIndex = .NameLocation.Characters(fi, 1).Font.ColorIndex
'    '                        End With
'    '                    Next
'    '                Else
'    '                    Reply = MsgBox("Лист """ & NameSheetI & """, параметр """ & Param(j).NamePar & _
'    '                          """ не найден. Выполнение программы прервано")
'    '                    FatalErr = True
'    '                    Exit Sub
'    '                End If
'    '            End If
'    '        ElseIf Repetition > 1 Then
'    '            RefAdr = Mid(StrLog, 2, InStr(StrLog, ";") - 3)
'    '            'выделить
'    '            ThisWorkbook.Sheets(NameSheetI).Range(RefAdr).Select()
'    '            With Selection.Interior
'    '                .ColorIndex = 3
'    '                .Pattern = xlSolid
'    '                .PatternColorIndex = xlAutomatic
'    '            End With

'    '            Reply = MsgBox("Лист """ & NameSheetI & """, параметр """ & Param(j).NamePar & _
'    '                    """ найден " & Repetition & " раза в ячейках:" & StrLog & _
'    '                    ". Значения параметра """ & Param(j).NamePar & """ будут отображены на листе """ & _
'    '                    NameSheetI & """ напротив ячейки """ & RefAdr & """. Продолжить выполнение программы?", _
'    '                    vbYesNo)
'    '            'убрать выделение
'    '            With Selection.Interior
'    '                .ColorIndex = xlNone
'    '                .Pattern = xlNone
'    '                .PatternColorIndex = xlNone
'    '            End With

'    '            If Reply = vbYes Then ' продолжить выполнение программы
'    '                Param(j).ValLocation = ThisWorkbook.Sheets(NameSheetI).Range(RefAdr)
'    '            ElseIf Reply = vbNo Then ' прервать выполнение программы
'    '                FatalErr = True
'    '                Exit Sub
'    '            End If

'    '        ElseIf Repetition = 1 Then
'    '            Param(j).ValLocation = ThisWorkbook.Sheets(NameSheetI).Range(StrLog)
'    '        End If
'    '        For i = 1 To UBound(Param(j).Val, 1)
'    '            If Param(j).ErrData(i) = False Then Param(j).ValLocation.Cells(1, i + 2).Value = Param(j).Val(i)
'    '            '     If ProcessingPossible(i) Then Param(j).ValLocation.Cells(1, i + 2).Value = Param(j).Val(i)
'    '        Next
'    '    Next

'    'End Sub


'    'Sub SearchWithinColumn(ByVal NameSheetI, ByVal j, ByVal columni, ByVal StrLog, ByVal Repetition, ByVal RepitAdr)


'    '    'Осуществляет поиск строки с заданным наборов символов в колонке на заданном листе

'    '    ' Входные переменные:
'    '    'NameSheetI - имя листа на котором осуществляется поиск
'    '    'j - индекс параметра  (глобальная переменная Param(j))
'    '    'columni - номер колонки на листе NameSheetI в которой осуществляется поиск.

'    '    'Выходные переменные:

'    '    'Repetition -  количество найденных соответствий.
'    '    'StrLog -

'    '    'Глобальные переменные:
'    '    'Param() - переменная пользовательского типа данных - DataXl (Dim Param() As DataXl), где

'    '    'Type DataXl
'    '    '  NamePar As String
'    '    '  NameLocation As Range
'    '    '  ValLocation As Range
'    '    '  Val() As Single
'    '    '  QuantSample As Integer
'    '    'End Type

'    '    ReDim_RepitAdr(0)
'    '    rowi = 1
'    '    Param(j).ValLocation = ThisWorkbook.Sheets(NameSheetI).Cells(rowi, columni)
'    '    EmptyCell = 0
'    '    Repetition = 0
'    '    StrLog = ""
'    '    Do Until EmptyCell = 50 Or Param(j).ValLocation.Row = 65536
'    '        If IsTheSame(Param(j).NamePar, Param(j).ValLocation.Value, True, Param(j).NameLocation, _
'    '                     Param(j).ValLocation) = True Then
'    '            Repetition = Repetition + 1
'    '    ReDim_Preserve RepitAdr(1 To Repetition)
'    '            RepitAdr(Repetition) = Param(j).ValLocation.Address
'    '            StrLog = StrLog & """" & RepitAdr(Repetition) & """;"
'    '        ElseIf Len(Param(j).ValLocation.Value) = 0 Then
'    '            EmptyCell = EmptyCell + 1
'    '        Else
'    '            EmptyCell = 0
'    '        End If
'    '        rowi = rowi + 1
'    '        Param(j).ValLocation = ThisWorkbook.Sheets(NameSheetI).Cells(rowi, columni)
'    '    Loop

'    '    If Repetition = 1 Then
'    '        StrLog = RepitAdr(Repetition)
'    '    End If
'    'End Sub


'    'Sub ReadRowIndefQuant(RangeJ, j, step)

'    '        Dim CurStr As String
'    '        'wwww = RangeJ.Address
'    '        CurStr = RangeJ.Value
'    '        If IsNumeric(CurStr) Then
'    '  step = 1
'    '        Else
'    '  step = 2
'    '        End If

'    'StartColumn = Param(j).ValLocation.Column + step
'    '        columni = StartColumn
'    '        semp = 0
'    '        CurStr = RangeJ.Cells(1, columni).Value
'    '        'www1 = RangeJ.Cells(1, ColumnI).Address
'    '        Do While columni <= 255 And IsNumeric(CurStr)
'    '            semp = semp + 1
'    '            With Param(j)
'    '    ReDim_Preserve .Val(1 To semp)
'    '                buf = RangeJ.Cells(1, columni).Value
'    '                .QuantSample = semp
'    '                .Val(semp) = buf
'    '                columni = columni + 1
'    '                CurStr = RangeJ.Cells(1, columni).Value
'    '            End With
'    '        Loop
'    '    End Sub

'    '    Sub ReadRowDefQuant(ByVal SempQuant, ByVal j)
'    '        Dim CurStr As String

'    '        'Программа для считывания заданного кол-ва ячеек, расположенных в ряд

'    '        'Входные величины
'    '        'SempQuant-кол-во считываемых ячеек
'    '        'j - индекс параметра в массиве Param()

'    '        Dim buf As String
'    '        'wwww = Param(j).ValLocation.Cells(1, 2).Address
'    '        CurStr = Param(j).ValLocation.Cells(1, 2).Value
'    '        If IsNumeric(CurStr) Then
'    '  step = 1
'    '        Else
'    '  step = 2
'    '        End If

'    'StartColumn = Param(j).ValLocation.Column + step
'    '        semp = 0
'    '        For columni = StartColumn To StartColumn + SempQuant - 1
'    '            semp = semp + 1

'    '            ' If ProcessingPossible(semp) Then
'    '            'RowInd = Param(j).ValLocation.Row
'    '            '    buf = Param(j).ValLocation.Cells(1, columni)

'    '            'buf = Str(Param(j).ValLocation.Cells(1, columni).Value)

'    '            If IsNumeric(Param(j).ValLocation.Cells(1, columni).Value) = True Then
'    '                buf = Param(j).ValLocation.Cells(1, columni)
'    '                If Len(buf) > 0 Then
'    '                    If buf > 0 Or j = 12 Then 'Если значения всех параметров (кроме абсолютных отборов воздуха от КВД на нужды самолёта по режимам) не равно "0" или "Empty"

'    '                        Param(j).QuantSample = semp
'    '                        Param(j).Val(semp) = buf
'    '                        Param(j).ErrData(semp) = False
'    '                    Else
'    '                        If j <= 3 Then
'    '                            FatalErr = True
'    '                            Exit Sub
'    '                        Else
'    '                            '       ProcessingPossible(semp) = False
'    '                            Param(j).ErrData(semp) = True
'    '                        End If
'    '                    End If
'    '                Else
'    '                    If j <= 3 Then
'    '                        FatalErr = True
'    '                        Exit Sub
'    '                    Else
'    '                        ' ProcessingPossible(semp) = False
'    '                        Param(j).ErrData(semp) = True
'    '                    End If
'    '                End If
'    '            Else
'    '                If j <= 3 Then
'    '                    FatalErr = True
'    '                    Exit Sub
'    '                Else
'    '                    ' ProcessingPossible(semp) = False
'    '                    Param(j).ErrData(semp) = True
'    '                End If
'    '            End If
'    '            'End If
'    '        Next

'    '    End Sub

'    'Sub ParamRepetitionSearch(ByVal QuantPar, ByVal NameSheet)

'    '    Dim Repetition(), Quant As Integer

'    '    Quant = 0

'    '    For i = 1 To QuantPar - 1
'    '        For j = i + 1 To QuantPar
'    '            If IsTheSame(Param(i).NamePar, Param(j).NamePar, True, Param(i).NameLocation, _
'    '                         Param(j).NameLocation) = True Then
'    '                Quant = Quant + 1
'    '                ReDim_Preserve Repetition(Quant)
'    '                Repetition(Quant) = j
'    '            End If
'    '        Next
'    '        If Quant > 0 Then
'    '            StrLog = """" & Param(i).NameLocation.Address & """; "
'    '            For ii = 1 To Quant
'    '                StrLog = StrLog & " """ & Param(Repetition(ii)).NameLocation.Address & """;"
'    '            Next
'    '            StrLog = Mid(StrLog, 1, Len(StrLog) - 1) ' убирает последний символ
'    '            MsgBox("Параметр """ & Param(i).NamePar & """ повторяется " & _
'    '                    Quant + 1 & " разa. Смотрите лист """ & NameSheet & """, ячейки " & _
'    '                    StrLog & ". Выполнение программы прервано.")
'    '            FatalErr = True
'    '            Exit Sub
'    '        End If
'    '    Next

'    'End Sub
'    'Sub IsSheetFound(ByVal NameSheet, ByVal Ref, ByVal Location)

'    '    'Входные данные
'    '    ' NameSheet - имя листа
'    '    ' Ref - адрес ячейки

'    '    On Error Resume Next 'Перехват ошибки, если лист NameSheets(0) не существует
'    '    Location = ThisWorkbook.Sheets(NameSheet).Range(Ref)
'    '    If Err <> 0 Then ' Прекращение выполнения программы если лист NameSheets(0) не найден.
'    '        'Set Location = ActiveCell
'    '        'rrr = ActiveSheet.Name
'    '        MsgBox("Лист """ & NameSheet & """ не найден. Выполнение программы прервано.")
'    '        FatalErr = True
'    '    End If
'    'End Sub

'    'Sub IsDefinitPar(ByVal NameSheet, ByVal j)


'    '    buf = Param(j).NameLocation.Value
'    '    If Len(buf) = 0 Then
'    '        Adr = Param(j).NameLocation.Address
'    '        ThisWorkbook.Sheets(NameSheet).Activate()
'    '        ActiveSheet.Range(Adr).Select()
'    '        With Selection.Interior
'    '            .ColorIndex = 3
'    '            .Pattern = xlSolid
'    '            .PatternColorIndex = xlAutomatic
'    '        End With
'    '        MsgBox("Лист """ & NameSheet & """, ячейка """ & Adr & _
'    '               """ - идентификатор параметра не найден. Выполнение программы прервано.")
'    '        FatalErr = True
'    '        With Selection.Interior
'    '            .ColorIndex = xlNone
'    '            .Pattern = xlNone
'    '            .PatternColorIndex = xlNone
'    '        End With
'    '    End If
'    'End Sub
'    'Function IsTheSame(ByVal Str1, ByVal Str2, ByVal CmpFont, ByVal Cell1, ByVal Cell2) As Boolean

'    '    IsTheSame = False

'    '    If Len(Str1) <> Len(Str2) Then
'    '        IsTheSame = False
'    '    Else
'    '        For i = 1 To Len(Str1)
'    '            If Asc(Mid(Str1, i, 1)) <> Asc(Mid(Str2, i, 1)) Then
'    '                IsTheSame = False
'    '                Exit Function
'    '            ElseIf CmpFont = True Then
'    '                If IsTheSameFonts(Cell1.Characters(i, 1).Font, _
'    '                                                       Cell2.Characters(i, 1).Font) = False Then
'    '                    IsTheSame = False
'    '                    Exit Function
'    '                End If
'    '            End If
'    '        Next
'    '        IsTheSame = True
'    '    End If
'    'End Function

'    'Function IsTheSameFonts(ByVal Font1, ByVal Font2) As Boolean
'    '    IsTheSameFonts = False
'    '    'Font2.Name = "Symbol"
'    '    If Font1.Name = Font2.Name Then IsTheSameFonts = True
'    'End Function


'    'Function DPG(ByVal XInd, ByVal YInd, ByVal ArgInd, ByVal FunkInd, ByVal j)



'    '    'Dim A(K), B(K) As Single

'    '    'IF(DP <= A(1)) DPG=B(1)
'    '    'IF(DP.GE.A(K)) DPG=B(K)
'    '    ErrLog_global = ""
'    '    n = Param(ArgInd).QuantSample
'    '    If (Param(XInd).Val(j) <= Param(ArgInd).Val(1)) Then
'    '        DPG = Param(FunkInd).Val(1)
'    '        ErrLog_global = "Замер " & j & ". Параметру """ & Param(YInd).NamePar & _
'    '               """ присвоено минимальное значение функции """ & _
'    '                Param(FunkInd).NamePar & " = f(" & Param(ArgInd).NamePar & ")""."
'    '    ElseIf (Param(XInd).Val(j) >= Param(ArgInd).Val(n)) Then
'    '        DPG = Param(FunkInd).Val(n)
'    '        ErrLog_global = "Замер " & j & ". Параметру """ & Param(YInd).NamePar & _
'    '               """ присвоено максимальное значение функции """ & _
'    '                Param(FunkInd).NamePar & " = f(" & Param(ArgInd).NamePar & ")""."
'    '    Else
'    '        Call SQINTP(Param(XInd).Val(j), DPG, n, _
'    '                    Param(ArgInd).Val(), Param(FunkInd).Val())
'    '    End If
'    'End Function


'End Module
