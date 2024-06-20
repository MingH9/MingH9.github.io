<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conn.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_conn_STRING
    MM_editCmd.CommandText = "INSERT INTO zxdd (u_name, u_password, u_iphone, chanpin, shuliang, xc_wz, xc_bz, xc_ds, xc_sd, lynr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 50, Request.Form("name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("password")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("phone")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("radio6")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 50, Request.Form("shuliang")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 50, MM_IIF(Request.Form("where1"), "Y", "N")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 50, MM_IIF(Request.Form("where2"), "Y", "N")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 202, 1, 50, MM_IIF(Request.Form("where3"), "Y", "N")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 202, 1, 50, MM_IIF(Request.Form("where4"), "Y", "N")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 203, 1, 536870910, Request.Form("textarea")) ' adLongVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "dingdan_index.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_conn_STRING
Recordset1_cmd.CommandText = "SELECT * FROM zxdd" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>在线订单</title>


<link href="../other/css.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>


<style type="text/css">
<!--
.STYLE4 {font-size: 24px;
	font-weight: bold;
	letter-spacing: 10px;
	color: #028DC6;}
-->
</style>
</head>

<body onload="MM_preloadImages('../files/images/12.jpg','../files/images/22.jpg','../files/images/32.jpg','../files/images/42.jpg','../files/images/52.jpg','../files/images/62.jpg','../images/72.jpg')">
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="120"><img src="../images/logo.png" width="1000" height="200" /></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="166" height="43" align="left" valign="middle"><a href="../index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image10','','../images/12.jpg',1)"><img src="../images/11.jpg" name="Image10" width="142" height="43" border="0" id="Image10" /></a></td>
    <td width="166" align="left" valign="middle"><a href="../shouji/shouji_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image5','','../images/22.jpg',1)"><img src="../images/21.jpg" name="Image5" width="143" height="43" border="0" id="Image5" /></a></td>
    <td width="166" align="left" valign="middle"><a href="../pingban/pingban_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image6','','../images/32.jpg',1)"><img src="../images/31.jpg" name="Image6" width="143" height="43" border="0" id="Image6" /></a></td>
    <td width="166" align="left" valign="middle"><a href="../bijiben/bijiben_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image7','','../images/42.jpg',1)"><img src="../images/41.jpg" name="Image7" width="143" height="43" border="0" id="Image7" /></a></td>
    <td width="166" align="left" valign="middle"><a href="../xiangji/xiangji_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image8','','../images/52.jpg',1)"><img src="../images/51.jpg" name="Image8" width="143" height="43" border="0" id="Image8" /></a></td>
    <td width="166" align="left" valign="middle"><a href="../dayinji/dayinji_index.html" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image9','','../images/62.jpg',1)"><img src="../images/61.jpg" name="Image9" width="143" height="43" border="0" id="Image9" /></a></td>
    <td width="166" align="left" valign="middle"><a href="dingdan_index.asp" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image11','','../images/72.jpg',1)"><img src="../images/71.jpg" name="Image11" width="143" height="43" border="0" id="Image11" /></a></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="1000" height="558" border="0" align="center" cellpadding="0" cellspacing="0" class="wai">
      <tr>
        <td height="556" align="center" valign="top"><span class="STYLE4"><br />  
          在 线 订 单</span><br />
            <br />
            <form METHOD="POST" action="<%=MM_editAction%>" name="form1" id="form1">
          <table width="75%" border="0" cellpadding="1" cellspacing="1" bgcolor="#17A1DC">
            <tr>
              <td width="23%" height="40" align="left" valign="middle" bgcolor="#FFFFFF">姓名(或用户名)：</td>
              <td width="77%" height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                <input type="text" name="name" id="name" />
              </label></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">密码：</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><input type="password" name="password" id="password" /></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">联系电话：</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                <input type="text" name="phone" id="phone" />
              </label></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">产品类型</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                  <input name="radio6" type="radio" id="radio6" value="智能手机" checked="checked" />
              智能手机
              <input type="radio" name="radio6" id="radio7" value="平板电脑" />
              平板电脑
              <input type="radio" name="radio6" id="radio8" value="笔记本" />
              笔记本
              <input type="radio" name="radio6" id="radio9" value="数码相机" />
              数码相机
              <input type="radio" name="radio6" id="radio10" value="打印机" />
              打印机
              </label></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">预订数量：</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                <select name="shuliang" size="1" id="shuliang">
                  <option>请选择数量...</option>
                  <option value="1">1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="10">更多</option>
                </select>
              </label></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">你从哪里看到产品宣传：</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                <input name="where1" type="checkbox" id="where1" value="网站" checked="checked" />
                网站
                <input name="where2" type="checkbox" id="where2" value="报纸" />
                报纸
                <input name="where3" type="checkbox" id="where3" value="电视" />
                电视
                <input name="where4" type="checkbox" id="where4" value="商店" />
                商店
                <br />
              </label></td>
            </tr>
            <tr>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF">留言内容：</td>
              <td height="40" align="left" valign="middle" bgcolor="#FFFFFF"><label>
                <textarea name="textarea" id="textarea" cols="60" rows="8"></textarea>
              </label></td>
            </tr>
            <tr>
              <td height="40" colspan="2" align="center" valign="middle" bgcolor="#FFFFFF"><label>
                <input type="submit" name="button" id="button" value="我要预订" />
                &nbsp;&nbsp;
                <input type="reset" name="button2" id="button2" value="取消" />
              </label></td>
            </tr>
                                                  </table>
        
              
            
          <input type="hidden" name="MM_insert" value="form1" />
            </form>
          <br />              </td>
      </tr>
</table>      </td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" class="end">
  <tr>
    <td height="130" align="center">数码资讯网 版权所有</td>
  </tr>
</table>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
