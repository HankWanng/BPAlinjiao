Attribute VB_Name = "Utility"
'==========================================

'0726 �N����JSON�i��ʸˡA��K�ե�
'productJson�����~����
'materialJson��ƫ~����
'�i�����եΦh���W�[����
'==========================================

Function productJson(ByVal NUMB As String, ByVal GLTRP_str As String, ByVal MATNR_str As String, ByVal MAKTX_str As String, ByVal GWEMG_str As String, ByVal BATCHNO As String, ByVal LGORT As String, ByVal POST_DATE As String)
'==========================================
'NUMB:�������Ǹ��X exa:"1"
'GLTRP_str:�u����
'MATNR_str:�N�X
'MAKTX_str:�W��
'GWEMG_str:���~���q
'BATCHNO:�帹
'LGORT:�w��
'POST_DATE:�u��W�Ǥ��
'==========================================

NUMB = Chr(34) + NUMB + Chr(34)
GLTRP_str = Chr(34) + GLTRP_str + Chr(34)
MATNR_str = Chr(34) + MATNR_str + Chr(34)
MAKTX_str = Chr(34) + MAKTX_str + Chr(34)
GWEMG_str = Chr(34) + GWEMG_str + Chr(34)
BATCHNO = Chr(34) + BATCHNO + Chr(34)
LGORT = Chr(34) + LGORT + Chr(34)
POST_DATE = Chr(34) + POST_DATE + Chr(34)
If NUMB = Chr(34) + "1" + Chr(34) Then
A = """SNO"": " + NUMB + "," + """GLTRP"":" + GLTRP_str
Else
A = ",{" + """SNO"": " + NUMB + "," + """GLTRP"":" + GLTRP_str
End If
b = ",""MATNR"":" + MATNR_str + ",""MAKTX"":" + MAKTX_str + ",""GWEMG"":" + GWEMG_str
C = ",""ERFME"":""KG"",""CHARG"":" + BATCHNO '����
d = ",""LGORT"":" + LGORT + ",""POST_DATE"":" + POST_DATE  '�w����
cplotno = Format(Cnfg.[B8], "yyyymmdd")
cplot = ",""cplotno"":" + Chr(34) + "L" + cplotno + Chr(34)
e = ",""sapgdh"":""""" + cplot
productJson = A + b + C + d + e + "}"
End Function



Function maininfo(ByVal plant As String, ByVal matnr As String)
'==========================================
'plant �u�t�W
'==========================================
plant = Chr(34) + plant + Chr(34)
matnr = Chr(34) + matnr + Chr(34)
maininfo = "{""main"":{ ""BUKRS"":""1010"",""AUART"":" + plant + ",""MATNR"":" + matnr + ",""maininfo"":[{"
End Function



Function materialInfo()
materialInfo = """detail"":[{"
End Function



Function materialJson(ByVal NUMB As String, ByVal IDNRK_I_str As String, ByVal BATCH_I_str As String, ByVal LGORT_I_str As String, ByVal MAKTX_I_str As String, ByVal MENGE_I_str As String, ByVal cplotno_str As String, ByVal REMARK_str As String)
'==========================================
'NUMB:�������Ǹ��X exa:"1"
'IDNRK_I_str:²�X    "BPA1ZF"
'LGORT_I_str�G�w��    "BP42"
'MAKTX_I_str:����y�z
'MENGE_I_str:���~���q
'BATCH_I_str:�����|     "X"
'cplotno_str:�����|     "X"
'==========================================
NUMB = Chr(34) + NUMB + Chr(34)
IDNRK_I_str = Chr(34) + IDNRK_I_str + Chr(34)
BATCH_I_str = Chr(34) + BATCH_I_str + Chr(34)
LGORT_I_str = Chr(34) + LGORT_I_str + Chr(34)
MAKTX_I_str = Chr(34) + MAKTX_I_str + Chr(34)
MENGE_I_str = Chr(34) + MENGE_I_str + Chr(34)
cplotno_str = Chr(34) + cplotno_str + Chr(34)
REMARK_str = Chr(34) + REMARK_str + Chr(34)
C = ",""sapgdh"":"""""
A = """SNO"": " + NUMB + "," + """IDNRK_I"":" + IDNRK_I_str + ",""MAKTX_I"":" + MAKTX_I_str + ",""MENGE_I"":" + MENGE_I_str
b = ",""MEINS_I"":""KG"",""BATCH_I"":" + BATCH_I_str + ",""LGORT_I"":" + LGORT_I_str + ",""REMARK"":" + REMARK_str + ",""cplotno"":" + cplotno_str + C
materialJson = A + b + "},{"
End Function

Sub txt()
plant = "BPA2"
NUMB = "1"
GLTRP_str = Format(Date, "yyyymmdd")
product_name = "BPA2"
GWEMG_str = "BPA2"
MATNR_str = "PN"
MAKTX_str = "P25"
BATCHNO = "BPA2"
LGORT = "BPA2"
POST_DATE = Format(Date, "yyyymmdd")
x = Test(plant, GLTRP_str, product_name, GWEMG_str, BATCHNO, LGORT, POST_DATE)
y = Test2(NUMB, GLTRP_str, MATNR_str, MAKTX_str, GWEMG_str, BATCHNO, LGORT, POST_DATE)
y1 = Test2("2", GLTRP_str, MATNR_str, MAKTX_str, GWEMG_str, BATCHNO, LGORT, POST_DATE)
f = "]},"

IDNRK_I_str = "bpa"
LGORT_I_str = "bp02"
MAKTX_I_str = "zhongwen"
MENGE_I_str = "121332"


M = materialInfo()
n = materialJson("1", IDNRK_I_str, LGORT_I_str, MAKTX_I_str, MENGE_I_str, BATCHNO)
n1 = materialJson("2", IDNRK_I_str, LGORT_I_str, MAKTX_I_str, MENGE_I_str, BATCHNO)
n2 = materialJson("3", IDNRK_I_str, LGORT_I_str, MAKTX_I_str, MENGE_I_str, BATCHNO)
Z = maininfo(plant) + y + y1 + f + M + n + n1 + n2

Debug.Print (Left(Z, Len(Z) - 2) + "]}")

End Sub

Sub APImain()
'Call Readjson("BPA1")
Call Readjson("BPA2")
'Call Readjson("BPA3")
End Sub
Sub Readjson(PlantName)
Dim http
Dim aa As String
  Set http = CreateObject("Microsoft.XMLHTTP")
  da = Format(Cnfg.[B8], "yyyy-mm-dd")
  URL = "https://oa.ccpgp.com.cn/ccpAPI/PIMS/pimsapi_bpa.jsp?reqdate=" + da + "&deptscode=" + PlantName
  'Debug.Print URL http://192.168.218.142/ccpAPI/PIMS/pimsapi_bpa.jsp
  http.Open "POST", URL, False
  http.send ""
  If http.Status = 200 Then
    aa = http.responseText
    Debug.Print aa
   Else
    aa = 0
End If
Call TestJson(aa, PlantName)
End Sub
Sub TestJson(value, PlantName)
Dim jsstr As String
Set wk6 = Worksheets("Calc Sheet")
  'vb����ַ���Ҫ��n�����ţ����Ƿ���
jsstr = Trim(value)
 'ǰ�ڰ󶨷���ʹ��������ʾ
 'Dim scobj As New MSScriptControl.ScriptControl
Dim age As Integer
Set scobj = CreateObject("ScriptControl")
 'ScriptControlʹ�õĽű����ԡ�����js��Ҳ֧��Vbscript
scobj.Language = "JavaScript"
 '���ű�����Ӵ��룬�������ַ���
scobj.AddCode ("var query = " & jsstr)
 'JSON�����ȡ���Եı�ʾ����������.����
 '���Ե�ֵ����Ǹ����������������飬����ʹ��������ʾȡ�ö��󣺶���.����[0]
 'Eval�Ǳ��ʽ��ֵ
Select Case PlantName
Case "BPA1"
BPAF_X = scobj.Eval("query.ex[0].BPA1F_X")
BPAF = scobj.Eval("query.ex[0].BPA1F")
BPAOFF = scobj.Eval("query.ex[0].BPA1OFF")
NUM = 0
Case "BPA2"
BPAF_X = scobj.Eval("query.ex[0].BPA1F_X")
BPAF = scobj.Eval("query.ex[0].BPA1F")
BPAOFF = scobj.Eval("query.ex[0].BPA1OFF")
BPATANK = scobj.Eval("query.ex[0].BPA1TANK")
BPA_WAREHOUSE = scobj.Eval("query.ex[0].BPA1_WAREHOUSE")
seabulk = scobj.Eval("query.ex[0].seabulk")
NUM = 1
Case "BPA3"
BPAF_X = scobj.Eval("query.ex[0].BPA3F_X")
BPAF = scobj.Eval("query.ex[0].BPA3F")
BPAOFF = scobj.Eval("query.ex[0].BPA3OFF")
NUM = 2
End Select
'Debug.Print BPA1F_X & BPA1F & BPA1OFF
wk6.Range("R7").Offset(0, 0).value = BPAF_X + BPAF

wk6.Range("V7").Offset(0, 0).value = BPAOFF
wk6.Range("H7").Offset(0, 0).value = BPATANK
wk6.Range("G7").Offset(0, 0).value = BPA_WAREHOUSE
wk6.Range("I7").Offset(0, 0).value = seabulk

End Sub

