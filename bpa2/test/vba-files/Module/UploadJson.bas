Attribute VB_Name = "UploadJson"
Sub MakeJsonbpa1(i_bpa1, i_bpa1zf, i_a203a, i_p101a, i_p101are, M)
'=====================
'bpa1:����A;
'bpa1zf:���o
'a203a�G���
'p101a�G��ƭܭf��
'p101are�G�^����
'=====================
Dim result, duedates, lists As String
Dim currentdate As String
Dim Target As Range
'Set sheet_lj = Worksheets("lj")



'=====================unit
cplotno = Format(Cnfg.[B8], "yyyymmdd")
cplotno_str = "L" + cplotno   '�帹 'Str
'GLTRP_str = Chr(34) + Format(Date, "yyyymmdd") + Chr(34) '�ɶ�
GLTRP_str = Format(Cnfg.[B9], "yyyymmdd")  '�ɶ�
'=====================



'=====================���~
'bpa1 = Chr(34) + Str(Round(bpa1)) + Chr(34) '����AStr
bpa1name = Chr(34) + Chart.[AQ10] + Chr(34)
BATCHNO = Chr(34) + "B" + Chr(34)
BATCHNOX = Chr(34) + "X" + Chr(34)
'=====================
    

'=====================���
   
bpa1zf = Str(Round(i_bpa1zf)) '���o
a203a = Str(Round(i_a203a, 0)) '���
p101a = Str(Round(i_p101a, 0))   '�f��
p101are = Str(Round(i_p101are, 0))  '�^���f��
phoh = "0"   '�~����
bpa1zfname = Chart.[BA10]
a203aname = Chart.[AB10]
p101aname = Chart.[Q10]
p101arename = Chart.[S10]
phohname = Chart.[Q10]
'=====================
    
    
MATNR_str = "BPA1"
MAKTX_str = Chart.[AQ10]
GWEMG_str = Str(Round(Abs(i_bpa1)))
BATCHNO = "B"
LGORT = "BP42"
POST_DATE = Format(Cnfg.[B9], "yyyymmdd")

ProductData = maininfo("BPA2", "BPA1") + productJson("1", cplotno, MATNR_str, MAKTX_str, GWEMG_str, BATCHNO, LGORT, POST_DATE)
'==========================================
'NUMB:�������Ǹ��X           exa:"1"
'GLTRP_str:�u����
'MATNR_str:�N�X              exa:"BPA1"
'MAKTX_str:�W��                 bpa1name
'GWEMG_str:���~���q
'BATCHNO:                   �����|     "X"
'LGORT:�w��                 "BP42"
'POST_DATE:�u��W�Ǥ��
'==========================================



BPA1ZF_1 = materialJson("1", "BPA1ZF", "B", "BP42", bpa1zfname, bpa1zf, cplotno_str, Calc.[AE2].Offset(0, 0))
A203A_2 = materialJson("2", "A203A", "X", "BP02", a203aname, a203a, cplotno_str, Calc.[AE2].Offset(1, 0))
A203A_3 = materialJson("3", "A203A", "X", "BP22", a203aname, "0", cplotno_str, Calc.[AE2].Offset(2, 0))
P101A_4 = materialJson("4", "P101A", " ", "BP02", p101aname, p101a, cplotno_str, Calc.[AE2].Offset(3, 0))
P101A_5 = materialJson("5", "P101A", " ", "BP22", p101arename, p101are, cplotno_str, Calc.[AE2].Offset(4, 0))
PHOH_6 = materialJson("6", "PHOH", " ", "BP02", phohname, phoh, cplotno_str, Calc.[AE2].Offset(5, 0))
PHOH_7 = materialJson("7", "PHOH", " ", "BP22", phohname, phoh, cplotno_str, Calc.[AE2].Offset(6, 0))

MaterialData = materialInfo() + BPA1ZF_1 + A203A_2 + A203A_3 + P101A_4 + P101A_5 + PHOH_6 + PHOH_7
'==========================================
'NUMB:�������Ǹ��X exa:"1"
'IDNRK_I_str:²�X    "BPA1ZF"
'LGORT_I_str�G�w��    "BP42"
'MAKTX_I_str:����y�z
'MENGE_I_str:���~���q
'cplotno_str:�帹
'==========================================


'==========================================
'����JSON
JsonData = ProductData + "]}," + MaterialData

result = Left(JsonData, Len(JsonData) - 2) + "]}"

Debug.Print (result)
'==========================================

cloo = M

Call UploadJson(result, cloo)

End Sub

Sub MakeJsonbpa1f(i_bpa1f, i_a203af, i_p101af, M)

Dim result, duedates, lists As String
Dim currentdate As String
Dim Target As Range


  cplotno = Format(Cnfg.[B8], "yyyymmdd")
 
cplotno_str = "L" + cplotno   '�帹 'Str
'GLTRP_str = Chr(34) + Format(Date, "yyyymmdd") + Chr(34) '�ɶ�
GLTRP_str = Format(Cnfg.[B9], "yyyymmdd")  '�ɶ�
  
    '���
    bpa1f = Str(Round(i_bpa1f))
    a203af = Str(Round(i_a203af))
    p101af = Str(Round(i_p101af))
    phoh = "0"
    
    bpa1fname = Chart.[AS2]
    a203afname = Chart.[AB10]
    p101afname = Chart.[Q10]
    
    
MAKTX_str = Chart.[AQ10]

BATCHNO = "B"
LGORT = "BP42"
POST_DATE = Format(Cnfg.[B9], "yyyymmdd")

ProductData = maininfo("BPA2", "BPA1F") + productJson("1", cplotno, "BPA1F", bpa1fname, bpa1f, BATCHNO, LGORT, POST_DATE)
'==========================================
'NUMB:�������Ǹ��X           exa:"1"
'GLTRP_str:�u����
'MATNR_str:�N�X              exa:"BPA1"
'MAKTX_str:�W��                 bpa1name
'GWEMG_str:���~���q
'BATCHNO:                   �����|     "X"
'LGORT:�w��                 "BP42"
'POST_DATE:�u��W�Ǥ��
'==========================================




A203A_1 = materialJson("1", "A203A", "X", "BP02", a203afname, a203af, cplotno_str, Calc.[AH2].Offset(0, 0))
A203A_2 = materialJson("2", "A203A", "X", "BP22", a203afname, "0", cplotno_str, Calc.[AH2].Offset(1, 0))
P101A_3 = materialJson("3", "P101A", " ", "BP02", p101afname, p101af, cplotno_str, Calc.[AH2].Offset(2, 0))
P101A_4 = materialJson("4", "P101A", " ", "BP22", p101afname, "0", cplotno_str, Calc.[AH2].Offset(3, 0))
PHOH_5 = materialJson("5", "PHOH", " ", "BP02", p101afname, "0", cplotno_str, Calc.[AH2].Offset(4, 0))
PHOH_6 = materialJson("6", "PHOH", " ", "BP22", p101afname, "0", cplotno_str, Calc.[AH2].Offset(5, 0))
MaterialData = materialInfo() + A203A_1 + A203A_2 + P101A_3 + P101A_4 + PHOH_5 + PHOH_6
'==========================================
'NUMB:�������Ǹ��X exa:"1"
'IDNRK_I_str:²�X    "BPA1ZF"
'LGORT_I_str�G�w��    "BP42"
'MAKTX_I_str:����y�z
'MENGE_I_str:���~���q
'cplotno_str:�帹
'==========================================


'==========================================
'����JSON
JsonData = ProductData + "]}," + MaterialData

result = Left(JsonData, Len(JsonData) - 2) + "]}"
    

'End Select

    Debug.Print result
cloo = M

Call UploadJson(result, cloo)


    

End Sub
Sub MakeJsonbpa1off(i_bpa1off, i_a203aoff, i_p101aoff, M)

Dim result, duedates, lists As String
Dim currentdate As String
Dim Target As Range

  cplotno = Format(Cnfg.[B8], "yyyymmdd")

cplotno_str = "L" + cplotno   '�帹 'Str
'GLTRP_str = Chr(34) + Format(Date, "yyyymmdd") + Chr(34) '�ɶ�
GLTRP_str = Format(Cnfg.[B8], "yyyymmdd") '�ɶ�
  
    '���
    bpa1off = Str(Round(i_bpa1off))
    a203aoff = Str(Round(i_a203aoff))
    p101aoff = Str(Round(i_p101aoff))
    phoh = "0"
    
    bpa1offname = Chart.[AV2]
    a203aoffname = Chart.[AB10]
    p101aoffname = Chart.[Q10]
    
    
MAKTX_str = Chart.[AQ10]

BATCHNO = "B"
LGORT = "BP42"
POST_DATE = Format(Cnfg.[B9], "yyyymmdd")

ProductData = maininfo("BPA2", "BPA1OFF") + productJson("1", GLTRP_str, "BPA1OFF", bpa1offname, bpa1off, BATCHNO, LGORT, POST_DATE)
'==========================================
'NUMB:�������Ǹ��X           exa:"1"
'GLTRP_str:�u����
'MATNR_str:�N�X              exa:"BPA1"
'MAKTX_str:�W��                 bpa1name
'GWEMG_str:���~���q
'BATCHNO:                   �����|     "X"
'LGORT:�w��                 "BP42"
'POST_DATE:�u��W�Ǥ��
'==========================================




A203A_1 = materialJson("1", "A203A", "X", "BP02", a203aoffname, a203aoff, cplotno_str, Calc.[AK2].Offset(0, 0))
A203A_2 = materialJson("2", "A203A", "X", "BP22", a203aoffname, "0", cplotno_str, Calc.[AK2].Offset(1, 0))
P101A_3 = materialJson("3", "P101A", " ", "BP02", p101aoffname, p101aoff, cplotno_str, Calc.[AK2].Offset(2, 0))
P101A_4 = materialJson("4", "P101A", " ", "BP22", p101aoffname, "0", cplotno_str, Calc.[AK2].Offset(3, 0))
PHOH_5 = materialJson("5", "PHOH", " ", "BP02", p101aoffname, "0", cplotno_str, Calc.[AK2].Offset(4, 0))
PHOH_6 = materialJson("6", "PHOH", " ", "BP22", p101aoffname, "0", cplotno_str, Calc.[AK2].Offset(5, 0))

MaterialData = materialInfo() + A203A_1 + A203A_2 + P101A_3 + P101A_4 + PHOH_5 + PHOH_6
'==========================================
'NUMB:�������Ǹ��X exa:"1"
'IDNRK_I_str:²�X    "BPA1ZF"
'LGORT_I_str�G�w��    "BP42"
'MAKTX_I_str:����y�z
'MENGE_I_str:���~���q
'cplotno_str:�帹
'==========================================


'==========================================
'����JSON
JsonData = ProductData + "]}," + MaterialData

result = Left(JsonData, Len(JsonData) - 2) + "]}"
    

'End Select

    Debug.Print result
cloo = M

Call UploadJson(result, cloo)

    






End Sub

Sub UploadJson(result, clo)
    'itemp = itemp + 1
    'Sheets("setup").Range("E20").Offset(itemp, 0).Value = result
    'Sheets("setup").Range("E20").Offset(itemp, 1).Value = Len(result)
    'Dim MyData As New DataObject
    'MyData.SetText result
    'MyData.PutInClipboard
    '����Ĵ���Ϊdebugʱ����ʾjson����

  Dim http
  
  
  myfile = "E:\CCJS_PIMS_RM\SAMPLE\BPA\linjiao\bpa2logs\"
'TargetPath = "D:\Desktop\bpatest\12.xlsm"

  myname = "upload_" & Format(Cnfg.[B9], "yyyymmdd") & clo & ".log"

sText = "success post!"
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(myfile & myname)) Then
       GoTo 100
    Else
    
     
  Set http = CreateObject("Microsoft.XMLHTTP")
  'URL = "http://192.168.218.142/ccpAPI/PIMS/PIMSOA_API4.jsp?action=create"   ' for test
  'URL = "https://oa.ccpgp.com.cn/ccpAPI/PIMS/PIMSOA_API4.jsp?action=create"
  URL = "http://192.168.218.142/ccpAPI/PIMS/PIMSOA_API5.jsp?action=create"   ' for test
  
  http.Open "POST", URL, False
  http.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
  http.send "&info=" & result
  Do Until http.readyState = 4
    DoEvents
  Loop
  If http.Status = 200 Then
    sLog = Now() & ": " & result
    Set fs = fso.OpenTextFile(myfile & myname, 8, True)
    fs.WriteLine sLog
    fs.Close
    Set fs = Nothing
   ' MsgBox "�ϴ�" & Create & "�ɹ�"
   Call test1
    'Debug.Print http.ResponseText
  Else
   ' MsgBox "�ϴ�" & Create & "ʧ�ܣ�������룺" & http.Status
  End If
    
    
    
    
    
    End If
  
    
    Set fso = Nothing
  
  
  
  
  
  
 
  
100:
End Sub

Sub test1()

myfile = "E:\CCJS_PIMS_RM\SAMPLE\BPA\linjiao\bpa2logs\"
'TargetPath = "D:\Desktop\bpatest\12.xlsm"

 
  myname = "upload_" & Format(Cnfg.[B9], "yyyymmdd") & ".log"

sText = "success post!"
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(myfile & myname)) Then
       GoTo 100
    Else
    sLog = Now() & ": " & sText
    Set fs = fso.OpenTextFile(myfile & myname, 8, True)
    fs.WriteLine sLog
    fs.Close
    Set fs = Nothing
    Set fso = Nothing
    End If

100:

End Sub

Sub dataselect()
'���|/�O�|
Set er = Worksheets(7)

cpsum = 0: cp = 0: bdo = 0: pta = 0
cpxsum = 0: cpx = 0: bdox = 0: ptax = 0


'THF���|
thfcj = er.Cells(28, "B")
THFCL = er.Cells(28, "C")
THFFJ = er.Cells(28, "D")
THFWJ = er.Cells(28, "E")


'THF�O�|

thfcjX = er.Cells(27, "B")
THFCLX = er.Cells(27, "C")
THFFJX = er.Cells(27, "D")
THFWJX = er.Cells(27, "E")

grad = thfcj + THFCL + THFFJ + THFWJ + thfcjX + THFCLX + THFFJX + THFWJX
If grad <> 0 Then
 M = 12354
    Call MakeJson(thfcj, THFCL, THFFJ, THFWJ, thfcjX, THFCLX, THFFJX, THFWJX, M)


End If




End Sub


