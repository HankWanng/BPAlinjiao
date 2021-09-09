VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Series"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
Option Compare Text
DefLng A-Z

Private Sub Worksheet_Change(ByVal Target As Range)

'Sub main()
'Set sheet_lj = Worksheets("lj")
Dim rownbst, rownbend, datenb As Integer
Chart.Range(Chart.[a7], Chart.[BR9]).ClearContents


'=====================
'呼叫OA API接口取得手工輸入相關數據
Call APImain
'=====================




'=====================
'判斷日期是否是月底
datenb = Day(Cnfg.[B9] + 1)

Select Case datenb
'Case 1
'rownbst = 15: rownbend = 13 '0：00~16：00  (15-13)
Case 3
rownbst = 16: rownbend = 12 '16：00~00：00   (16-12)
Case Else
rownbst = 16: rownbend = 13 '0：00~00：00   (16-13)
End Select
'=====================
'define TagName at first value of columns of  series

Dim TimeIndex_minu2 As Range, TimeIndex_Today As Range, TimeIndex_minu1 As Range, TimeIndex_series As Range

Dim FI420_Sum_previous As Range, FI420_Sum_today As Range, FI420_Sum_series As Range, FI420_Result As Range
Dim FI420_Sum_previous_chart As Range, FI420_Sum_today_chart As Range
Dim FI120_Sum_previous As Range, FI120_Sum_today As Range, FI120_Sum_series As Range, FI120_Result As Range
Dim FI120_Sum_previous_chart As Range, FI120_Sum_today_chart As Range
Dim FI112_Sum_previous As Range, FI112_Sum_today As Range, FI112_Sum_series As Range, FI112_Result As Range
Dim FI112_Sum_previous_chart As Range, FI112_Sum_today_chart As Range

Dim Silo_T741_previous As Range, Silo_T741_today As Range
Dim Silo_T742_previous As Range, Silo_T742_today As Range
Dim Silo_T743_previous As Range, Silo_T743_today As Range
Dim Silo_T744_previous As Range, Silo_T744_today As Range
Dim Silo_T745_previous As Range, Silo_T745_today As Range
Dim Silo_T746_previous As Range, Silo_T746_today As Range
Dim Silo_T741_previous_chart As Range, Silo_T741_today_chart As Range
Dim Silo_T742_previous_chart As Range, Silo_T742_today_chart As Range
Dim Silo_T743_previous_chart As Range, Silo_T743_today_chart As Range
Dim Silo_T744_previous_chart As Range, Silo_T744_today_chart As Range
Dim Silo_T745_previous_chart As Range, Silo_T745_today_chart As Range
Dim Silo_T746_previous_chart As Range, Silo_T746_today_chart As Range
Dim Silo_Sum_today As Range, Silo_Sum_previous As Range, Silo_Delta As Range, Silo_chart As Range
Dim zf_Sum_previous As Range, zf_Sum_today As Range, zf_Sum_series As Range, zf_Chart As Range


Set TimeIndex_series = [a1]
'Set TimeIndex_minu2 = Chart.[a7]
Set TimeIndex_minu1 = Chart.[a7]
Set TimeIndex_Today = Chart.[a8]


TimeIndex_Today = TimeIndex_series.Offset(rownbst - 1, 0)
TimeIndex_minu1 = TimeIndex_series.Offset(rownbend - 1, 0)

'TimeIndex_minu2 = Date - 2

Set FI420_Sum_series = [f1]
Set FI420_Sum_previous = FI420_Sum_series.Offset(rownbend - 1, 0)
Set FI420_Sum_today = FI420_Sum_series.Offset(rownbst - 1, 0)
Set FI420_Sum_previous_chart = Chart.[p7]
Set FI420_Sum_today_chart = Chart.[p8]
Set FI420_Result = Chart.[q8]

Set FI120_Sum_series = [ac1]
Set FI120_Sum_previous = FI120_Sum_series.Offset(rownbend - 1, 0)
Set FI120_Sum_today = FI120_Sum_series.Offset(rownbst - 1, 0)
Set FI120_Sum_previous_chart = Chart.[r7]
Set FI120_Sum_today_chart = Chart.[r8]
Set FI120_Result = Chart.[s8]

Set FI112_Sum_series = [c1]
Set FI112_Sum_previous = FI112_Sum_series.Offset(rownbend - 1, 0)
Set FI112_Sum_today = FI112_Sum_series.Offset(rownbst - 1, 0)
Set FI112_Sum_previous_chart = Chart.[aa7]
Set FI112_Sum_today_chart = Chart.[aa8]
Set FI112_Result = Chart.[ab8]

Set zf_Sum_previous = Chart.[q8].Offset(rownbend - 1, 0)
Set zf_Sum_today = Chart.[q9].Offset(rownbst - 1, 0)
Set zf_Sum_series = [c1]
Set zf_Chart = Chart.[ba8]

Set FI112_Sum_series = [c1]
Set Silo_Delta = Chart.[AC8]

Set Silo_T741_previous = [ad1].Offset(rownbst - 1, 0)
Set Silo_T742_previous = [ae1].Offset(rownbst - 1, 0)
Set Silo_T743_previous = [af1].Offset(rownbst - 1, 0)
Set Silo_T744_previous = [ag1].Offset(rownbst - 1, 0)
Set Silo_T745_previous = [ah1].Offset(rownbst - 1, 0)
Set Silo_T746_previous = [ai1].Offset(rownbst - 1, 0)
Set Silo_T741_previous_chart = Chart.[ag7]
Set Silo_T742_previous_chart = Chart.[ah7]
Set Silo_T743_previous_chart = Chart.[ai7]
Set Silo_T744_previous_chart = Chart.[aj7]
Set Silo_T745_previous_chart = Chart.[ak7]
Set Silo_T746_previous_chart = Chart.[al7]

Set Silo_T741_today = [ad1].Offset(rownbend - 1, 0)
Set Silo_T742_today = [ae1].Offset(rownbend - 1, 0)
Set Silo_T743_today = [af1].Offset(rownbend - 1, 0)
Set Silo_T744_today = [ag1].Offset(rownbend - 1, 0)
Set Silo_T745_today = [ah1].Offset(rownbend - 1, 0)
Set Silo_T746_today = [ai1].Offset(rownbend - 1, 0)
Set Silo_T741_today_chart = Chart.[ag8]
Set Silo_T742_today_chart = Chart.[ah8]
Set Silo_T743_today_chart = Chart.[ai8]
Set Silo_T744_today_chart = Chart.[aj8]
Set Silo_T745_today_chart = Chart.[ak8]
Set Silo_T746_today_chart = Chart.[al8]


Silo_T741_previous_chart = Silo_T741_previous
Silo_T742_previous_chart = Silo_T742_previous
Silo_T743_previous_chart = Silo_T743_previous
Silo_T744_previous_chart = Silo_T744_previous
Silo_T745_previous_chart = Silo_T745_previous
Silo_T746_previous_chart = Silo_T746_previous

Silo_T741_today_chart = Silo_T741_today
Silo_T742_today_chart = Silo_T742_today
Silo_T743_today_chart = Silo_T743_today
Silo_T744_today_chart = Silo_T744_today
Silo_T745_today_chart = Silo_T745_today
Silo_T746_today_chart = Silo_T746_today



Set Silo_Sum_previous = Chart.[am7]
Set Silo_Sum_today = Chart.[am8]

'========================================================================================================================================================================
'計算過程
 FI420_Result = (Int(FI420_Sum_today.value) - Int(FI420_Sum_previous.value))
 Dim result As Single
 
 result = FI420_Result.Value2 * 1000
 Debug.Print result
 FI420_Result = result
 FI420_Sum_today_chart.Value2 = FI420_Result
 FI420_Sum_previous_chart = FI420_Sum_previous
 
 FI120_Result = (FI120_Sum_today - FI120_Sum_previous)
 FI120_Sum_today_chart = FI120_Sum_today
 FI120_Sum_previous_chart = FI120_Sum_previous
 
 FI112_Result = (FI112_Sum_today - FI112_Sum_previous)
 FI112_Sum_today_chart = FI112_Sum_today
 FI112_Sum_previous_chart = FI112_Sum_previous
 

 
 
 Silo_Sum_today = Int(Silo_T741_today + Silo_T742_today) + Int(Silo_T743_today + Silo_T744_today + Silo_T745_today + Silo_T746_today)
 
 Silo_Sum_previous = Int(Silo_T741_previous + Silo_T742_previous) + Int(Silo_T743_previous + Silo_T744_previous + Silo_T745_previous + Silo_T746_previous)
 
 Dim Warehouse_series As Range, Warehouse_chart As Range
 Dim tank_series As Range, tank_chart As Range
 Dim seabulk_series As Range, seabulk_chart As Range
 
 Set Warehouse_series = Calc.[g7] '自動倉
  'Set Warehouse_series = [n1].Offset(rownbend - 1, 0) '自動倉
 Set Warehouse_chart = Chart.[an8]
 Set tank_series = Calc.[h7]   '槽車出貨
 'Set tank_series = [s1].Offset(rownbend - 1, 0)   '槽車出貨
 Set tank_chart = Chart.[ao8]
 'Set seabulk_series = [t1].Offset(rownbend - 1, 0)'seabulk
 Set seabulk_series = Calc.[i7]
 Set seabulk_chart = Chart.[ap8]
 
 
     
Warehouse_chart = Warehouse_series / 1000
tank_chart = tank_series / 1000
seabulk_chart = seabulk_series / 1000



Dim BPA_product As Range
Dim BPAF_series As Range, BPAF_chart As Range
Dim BPAOFF_series As Range, BPAOFF_chart As Range

Set BPA_product = Chart.[aq8]

 '=====================bpa產量
 BPA_product = (Warehouse_chart + tank_chart + seabulk_chart - Silo_Sum_today + Silo_Sum_previous) * 1000
 
 '=====================bpa產量
 
 'zf
 
 Set zf_Chart = Chart.[k7]
 Set zf_Sum_previous = [aq1].Offset(rownbst - 1, 0)
 Set zf_Sum_today = [aq1].Offset(rownbend - 1, 0)
 zf_Chart = (zf_Sum_previous - zf_Sum_today) '8.26gaiwei zhengshu
 
 Dim Unit_P101A As Range, Unit_A230A As Range
 Set Unit_P101A = Chart.[U8]
 Set Unit_A230A = Chart.[AC8]
 Dim P101A_fromBPA As Single
 P101A_fromBPA = FI420_Result + FI120_Result
 Set BPAF_series = Calc.[r7] '細粉
 Set BPAOFF_series = Calc.[v7] '次級品
 Unit_P101A = P101A_fromBPA / (BPA_product + BPAF_series + BPAOFF_series)
 Unit_A230A = FI112_Result / (BPA_product + BPAF_series + BPAOFF_series)
  
  ' do not know what is this yet.
 'Chart.[l7] = sheet_lj.[e7]
 'Chart.[m7] = (sheet_lj.[b7] - sheet_lj.[c7]) * 1000
 
 'Chart.[p7] = Chart.[m7] / (sheet_lj.[j7] + sheet_lj.[r7] + sheet_lj.[v7])
 'Chart.[q7] = Chart.[l7] / (sheet_lj.[j7] + sheet_lj.[r7] + sheet_lj.[v7])
 
 'BPAF

 Dim A203A_FromF As Range
 Dim P101A_FromF As Range
'Set BPAF_series = [r1].Offset(rownbend - 1, 0) '細粉

Set BPAF_chart = Chart.[At8]
Set P101A_FromF = Chart.[bc8]
Set A203A_FromF = Chart.[bd8]
 BPAF_chart = BPAF_series
 
 P101A_FromF = BPAF_series * Unit_P101A
 A203A_FromF = BPAF_series * Unit_A230A
'
 
 'BPAOFF
 
 Dim A203A_FromOff As Range
 Dim P101A_FromOff As Range
 
'Set BPAOFF_series = [v1].Offset(rownbend - 1, 0) '次級品

Set BPAOFF_chart = Chart.[aw8]
Set P101A_FromOff = Chart.[be8]
Set A203A_FromOff = Chart.[bf8]
 BPAOFF_chart = BPAOFF_series
 
 P101A_FromOff = BPAOFF_series * Unit_P101A
 A203A_FromOff = BPAOFF_series * Unit_A230A
 '=====================
 
 
 
 '=====================
 '第一張工單 No.1 (bpa1,zf,pn,ac)
 Dim bpa1, bpa1zf, a203a, p101a, p101are As Long
 bpa1 = BPA_product
 bpa1zf = zf_Chart
 'a203a = FI112_Result * Unit_A230A
 'p101a = P101A_fromBPA * Unit_P101A
  a203a = bpa1 * Unit_A230A
 p101a = bpa1 * Unit_P101A
 p101are = 0  'not yet
 
 
 '=====================
  '第二張工單 No.2 (bpa1f,a203af,p101af)
  Dim bpa1f, a203af, p101af As Long
 bpa1f = BPAF_chart
 a203af = A203A_FromF
 p101af = P101A_FromF
 
 
 '=====================
  '第三張工單 No.3 (bpa1off,a203aoff,p101aoff)
    Dim bpa1off, a203aoff, p101aoff As Long
 bpa1off = BPAOFF_chart
 a203aoff = A203A_FromOff
 p101aoff = P101A_FromOff
 
 
 '=====================
 Dim BBmmm As Integer
 BBmmm = Format(Date, "d")
 If BBmmm <> 1 Then
    Call MakeJsonbpa1(bpa1, bpa1zf, a203a, p101a, p101are, 21235)
    Call MakeJsonbpa1f(bpa1f, a203af, p101af, 56789)
    Call MakeJsonbpa1off(bpa1off, a203aoff, p101aoff, 21231)
 End If

End Sub


