<% @ LANGUAGE=VBScript CODEPAGE=65001 %> 
<% Response.CharSet = "utf-8"
   Response.Buffer = true
   Session.codepage = 65001
   nowprogram_ls = "appcheck"
   Program_name = "提供app程式呼叫"
   
   '2013-09-27 因為上線的app無法取得限制自訓查訓的鴿會故不呈現自訓查詢的鴿會
   show_trainOpenFg = "N"
   'opkind = A001  開啟程式檢查，
   '               回傳
   '               "mobile"   :"手機號碼",
   '               "smsfg"    :"Y or N",    是否可接收簡訊回傳
   '               "sitedata" : {[]}可接收回傳的鴿會和鴿舍(array) sysid , pigsiteno
   '               "nosysid"  : {[]} 鴿會需要驗證(array)  sysid   
   'opkind = A070  取得即時資料
   '               "mobile"   :"手機號碼",
   '               "opkind"   :"A070",
   '               "type"     :"3",  1:比賽、2:測試環、3:自訓 
   '               "sysid"    :"鴿會代碼",
   '               "pigsiteno":"",
   '               "page"     :"1",
   '               "pagenum"  :"10",
   '               "orderby"  :"1"  1:後到達、2:先到達、3:鴿舍前、4:鴿舍後 
   '               回傳 
   '               "mobile"     :"手機號碼",
   '               "totpage"    :"10",    資料全部頁數
   '               "totnum"     :"10",    資料全部筆數
   '               "returndata" : {[]} 回傳(array)  seq,pigsiteno,ringno,backtime(時間),lo,la,uid
   'opkind = A011  判斷輸入鴿舍、密碼驗證是否有誤 
   '               "mobile"   :"手機號碼",
   '               "opkind"   :"A011",
   '               "sysid"    :"鴿會代碼",
   '               "pigsiteno":"",
   '               "sitepw":""
   '               回傳 
   '               "mobile"   :"手機號碼",
   '               "opkind"   :"A011",
   '               "sysid"    :"鴿會代碼",
   '               "pigsiteno":"",
   '               "sitepw"   :"",
   '               "rtnfg"    :"Y or N",
   'opkind = A020  簡訊回報 
   '               "mobile"   :"手機號碼",
   '               "opkind"   :"A020",
   '               "smsrp"  : {[]} 鴿會、鴿舍(array)  sysid,pigsiteno
   '               回傳 
   '               "mobile"     :"手機號碼",
   '               "smsdata"    : {[]} 回傳(array)  sysid,pigsiteno,ringno,bkdate,bktime(時間)
   'opkind = A030  取得鴿會資料 
   '               "mobile"     :"手機號碼",
   '               "opkind"    :"A011",
   '               "sysid"     :"鴿會代碼",
   '               "pigsiteno" :"輸入鴿舍編號",
   '               "sitepw"    :"輸入的密碼
   '               回傳 
   '               "mobile"     :"手機號碼"
   '               "sysiddata"  : {[]} 鴿會資料(array) sysid , sysname 

%> 
<%'產生json的函數%>
<!--#include file="JSON_latest.asp"-->  
<%'資料庫和常用函數%>
<!--#include file="connections/conn.asp" -->
<!--#include file="SMS_lib.asp" -->
<%
  nowdatetime = year(date()) &  right("0" & month(date()),2) & right("0" & day(date()),2) _
              & right("0" & hour(now()),2) & right("0" & minute(now()),2) & right("0" & second(now()),2)
  nowdate = year(date()) &  right("0" & month(date()),2) & right("0" & day(date()),2)
  nowbkdate = year(date()) & "/" & right("0" & month(date()),2) & "/" & right("0" & day(date()),2)
  set rs=Server.Createobject("ADODB.Recordset")
  rs.CursorLocation=3   

  '增加多筆資料的方法 
  Set mutil_rec = jsArray()    
  Sub AddMembernew(field_list , value_list)          
    Set mutil_rec(Null) = jsObject()   
    field_arr = Split(field_list,",")
    value_arr = Split(value_list,",")
    for i = LBound(field_arr ) to UBound(field_arr )
      field_ls = trim(field_arr(i))
      mutil_rec(Null)(field_ls) = trim(value_arr(i))
    next
  End Sub 

  '單筆資料
  set return_json = jsObject() 

  '解析傳入的JSON
  Dim sc4Json 

  Sub InitScriptControl
    Set sc4Json = Server.CreateObject("MSScriptControl.ScriptControl")
    sc4Json.Language = "JavaScript"
    sc4Json.AddCode "var itemTemp=null;function getJSArray(arr, index){itemTemp=arr[index];}"
  End Sub 

  Function getJSONObject(strJSON)
    sc4Json.AddCode "var jsonObject = " & strJSON
    Set getJSONObject = sc4Json.CodeObject.jsonObject
  End Function 

  Sub getJSArrayItem(objDest,objJSArray,index)
    On Error Resume Next
    sc4Json.Run "getJSArray",objJSArray, index
    Set objDest = sc4Json.CodeObject.itemTemp
    If Err.number=0 Then Exit Sub
    objDest = sc4Json.CodeObject.itemTemp
  End Sub
  
  '*******************************************/
  '呼叫 
  '****************************/
  '產生子json陣列
  Sub AddMember(item_id , field_list , value_list)          
    Set return_json(item_id)(Null) = jsObject()   
    field_arr = Split(field_list,",")
    value_arr = Split(value_list,",")
    for i = LBound(field_arr ) to UBound(field_arr )
      field_ls = trim(field_arr(i))
      return_json(item_id)(Null)(field_ls) = trim(value_arr(i))
    next
  End Sub  

  '檢查是否可以接收簡訊回傳A001
  json_data = request("jsondata")
  savefilename = "../AppCheckLog/" & left(nowdate,6)&"_appcheck.log" '以每個月的檔名方式記錄LOG
  writefiledata savefilename,nowdatetime &"    "&json_data
  'showmessage(json_data)
 
  Dim objINjson
  Call InitScriptControl '呼叫 getJSArray 宣告
 
  Set objINjson = getJSONObject(json_data)
  
  opkind_ls = objINjson.opkind

'======================================================================================================================
  if ( opkind_ls = "A000" ) then
    '檢查帳密是否存在，輸入帳號、密碼傳回對應的電話號碼
    'appcheck.asp?jsondata={"opkind":"A000","app_id":"webtest","app_pw":"testweb"}

    app_id_ls = objINjson.app_id
    app_pw_ls = objINjson.app_pw
    sql = " select phone from free " _
        & "  where id = '" & app_id_ls & "' " _
        & "    and pw = '" & app_pw_ls & "' " 
    rs.open sql,conn,1,3
    if rs.eof then
      mobile_ls = ""
      return_json("mobile") = "" 
      return_json.Flush
    else
      mobile_ls = rs("phone")
      opkind_ls = "A001"
    end if
    rs.close
  else
    mobile_ls = objINjson.mobile
  end if
  
   'test
  if mobile_ls = "0972917220" then
    mobile_ls = "0939209078"
    writefiledata savefilename,nowdatetime &"    nobile_ls = 0972917220--->0939209078"
  end if
  if mobile_ls = "0927999630" then
    mobile_ls = "0910817355"
    writefiledata savefilename,nowdatetime &"    nobile_ls = 0927999630--->0910817355"
  end if
  if mobile_ls = "0920417450" then
    mobile_ls = "0937319188"
    writefiledata savefilename,nowdatetime &"    nobile_ls = 0920417450--->0937319188"
  end if
  if mobile_ls = "0931660810" then
    mobile_ls = "0915566009"
    writefiledata savefilename,nowdatetime &"    nobile_ls = 0931660810--->0915566009"
  end if
  if mobile_ls = "1234321234" then
    mobile_ls = "0921543687"
    writefiledata savefilename,nowdatetime &"    nobile_ls = 1234321234--->0921543687"
  end if
  
'======================================================================================================================
  if ( opkind_ls = "A001" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A001"}
    'appcheck.asp?jsondata={"mobile":"0930768736","opkind":"A001"}
    'appcheck.asp?jsondata={"mobile":"","opkind":"A001"}
    return_json("mobile") = mobile_ls    
    
    '檢查手機號碼是否有在當季賽事 
    if IsEmptyEJ(mobile_ls) = "Y" then  
      return_json("smsfg") = "N"
    else
      sql = " select sysid , pigsiteno from sms_sitedata "_
          & "  where ( mobil = '" & mobile_ls & "' " _
          & "          or  recievermobil like '%" & mobile_ls & "%' )" _
          & "    and exists ( select raceno  from sms_raceperiod " _
          & "                  where raceno = sms_sitedata.raceno " _
          & "                    and sysid = sms_sitedata.sysid " _
          & "                    and startdate <= CONVERT(varchar(8),GETDATE(),112) " _
          & "                    and enddate >= CONVERT(varchar(8),GETDATE(),112) " _
          & "               ) "
      'showmessage(sql)
      rs.open sql,conn,1,3
      if rs.eof then
        return_json("smsfg") = "N"
      else
        return_json("smsfg") = "Y"
        Set return_json("sitedata") = jsArray()
        do while ( not rs.eof )
          A001_ls = rs("sysid") & "," & rs("pigsiteno")
          AddMember "sitedata","sysid,pigsiteno" , A001_ls
          rs.MoveNext
        loop
      end if
      rs.close
    end if 
    'Set return_json("sitedata") = jsArray()
    'AddMember "sitedata","sysid,pigsiteno" , "1012,144"
    'AddMember "sitedata","sysid,pigsiteno" , "1105,244"
    '判斷是否有限制開放查詢，table 加一個欄位 
    sql = " select sysid from sms_sysid " _
        & "  where showfg='Y' and TrainOpenfg = 'N' order by seqno "
    rs.open sql,conn,1,3
    if not rs.eof then 
      Set return_json("nosysid") = jsArray()
      do while ( not rs.eof )
        AddMember "nosysid","sysid" , rs("sysid")
        rs.MoveNext
      loop
    end if
    rs.close
    return_json.Flush
  end if
  
'======================================================================================================================
  '判斷鴿舍是否有權限進入
  '輸入:鴿會代號、鴿舍編號、鴿舍密碼
  '輸出:有資料Y、沒資料N
  if ( opkind_ls = "A002" ) then 
    'appcheck.asp?jsondata={"opkind":"A002","sysid":"1003","pigsiteno":"0012",","app_pw":"testweb"}
    
    sysid = objINjson.sysid
    pigsiteno = objINjson.pigsiteno
    pwdata = objINjson.app_pw
    
    
    if len(pigsiteno) > 2 then pigsiteno = right("00" & pigsiteno , 4)
    nowdate = year(date()) &  right("0" & month(date()),2) & right("0" & day(date()),2)
    
    sql = "select raceno from sms_raceperiod where sysid='" & sysid & "' and startdate<= '" & nowdate & "' and enddate>= '" & nowdate & "' "
    RaceNo = getSQLFieldValue(sql,"raceno")
    
    '檢查有沒有資料
    if IsEmptyEJ(pwdata)="N" then '有輸入密碼
      sql = " select * from sms_sitedata where sysid = '" & sysid & "' and pigsiteno = '" & pigsiteno & "' and DefPassword ='" & pwdata & "' and RaceNo ='"  & RaceNo & "' "
      ok_ls = CheckDataExist(sql)
    end if
 
 		if ok_ls = "Y" then
 			
	 	else
		end if
		return_json.Flush

		do while ( not rs.eof )
			'2018/08/23 增加比賽、測試、自訓
			tmodeType = "NNN"
			if Trim(rs("TrainOpenfg")) = "N" Then tmodeType = "NNY" 
			
		  A030_ls = rs("sysid") & "," & rs("sysname") & "," & tmodeType
		  'showmessage(A030_ls)
		  AddMember "sysiddata","sysid,sysname,modeType" , A030_ls
		  rs.MoveNext
		loop	

    rs.close
    return_json.Flush
    		
		
    if rs.eof then
      mobile_ls = ""
      return_json("mobile") = "" 
      return_json.Flush
    else
      mobile_ls = rs("phone")
      opkind_ls = "A001"
    end if
    rs.close
    return_json.Flush
 
  end if
 
'======================================================================================================================
  if ( opkind_ls = "A010" ) then 
    '賽事：appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A010","type":"1","sysid":"1003","pigsiteno":""}
    '自訓：appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A010","type":"3","sysid":"1003","sitepw":"Y","pigsiteno":"0015"}
    'appcheck.asp?jsondata={"mobile":"0930768736","opkind":"A010","type":"3","sysid":"1012","sitepw":"Y","pigsiteno":"0709"}
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A010","type":"3","sysid":"1012","sitepw":"D52C49BAC4","pigsiteno":"0709","country":"TWN"}
    '要先檢查參數是否存在 
    if InStr(json_data,"type") then
      type_ls = objINjson.type
    else
      type_ls = ""
    end if 
    if InStr(json_data,"sysid") then
      sysid_ls = objINjson.sysid
    else
      sysid_ls = ""
    end if 
    if InStr(json_data,"pigsiteno") then
      pigsiteno_ls = objINjson.pigsiteno
    else
      pigsiteno_ls = ""
    end if 
    if InStr(json_data,"page") then
      page_ls = objINjson.page
    else
      page_ls = ""
    end if 
    if InStr(json_data,"pagenum") then
      pagenum_ls = objINjson.pagenum
    else
      pagenum_ls = ""
    end if 
    if InStr(json_data,"orderby") then
      orderby_ls = objINjson.orderby 
    else
      orderby_ls = ""
    end if 
    if InStr(json_data,"sitepw") then
      sitepw_ls = objINjson.sitepw 
    else
      sitepw_ls = ""
    end if 
    '2013-06-28
    if InStr(json_data,"qdate") then
      qdate_ls = objINjson.qdate
    else
      qdate_ls = nowbkdate
    end if 

    if InStr(json_data,"country") then
      country_ls = objINjson.country
    else
      country_ls = "TWN"
    end if 

    run_fg = "Y"

	
	'台灣
	if country_ls = "TWN" then 
			if IsEmptyEJ(sysid_ls) = "Y" then 
			  run_fg = "N"
			else
			  '判斷自訓時才要檢查密碼，1:比賽、2:測試環、3:自訓 
			  if type_ls = "3" then
					'判斷鴿會是否開放       
					sql = " select sysid from sms_sysid where sysid = '" & sysid_ls & "' and TrainOpenfg = 'N' "
					if CheckDataExist(sql) = "Y" then 
					  '1.未開放鴿會
					  '  1-1 檢查要有輸入鴿舍  
					  '  1-2 判斷是否有傳送sitepw      
					  if ( IsEmptyEJ(pigsiteno_ls) = "Y" ) or ( IsEmptyEJ(sitepw_ls) = "Y" ) then 
							run_fg = "N"
					  else            
							if sitepw_ls = "Y" then 
							  '  1-3 無sitepw 則表示鴿會、鴿舍、手機存在sms_sitedata的sysid、pigsiteno、mobil(or recievermobil) 內
							  sql = " select sysid , pigsiteno from sms_sitedata "_
								  & "  where ( mobil = '" & mobile_ls & "' " _
								  & "          or  recievermobil like '%" & mobile_ls & "%' )" _
								  & "    and exists ( select raceno  from sms_raceperiod " _
								  & "                  where raceno = sms_sitedata.raceno " _
								  & "                    and sysid = sms_sitedata.sysid " _
								  & "                    and startdate <= CONVERT(varchar(8),GETDATE(),112) " _
								  & "                    and enddate >= CONVERT(varchar(8),GETDATE(),112) " _
								  & "               ) " _
								  & "    and sysid = '" & sysid_ls & "' " _
								  & "    and pigsiteno = '" & pigsiteno_ls & "' " 
							  run_fg = CheckDataExist(sql)
							else						 
								 '  1-4 有sitepw 則表示鴿會、鴿舍、sitepw存在sms_sitedata的sysid、pigsiteno、defPassword 內 
							  sql = " select sysid , pigsiteno from sms_sitedata "_
								  & "  where exists ( select raceno  from sms_raceperiod " _
								  & "                  where raceno = sms_sitedata.raceno " _
								  & "                    and sysid = sms_sitedata.sysid " _
								  & "                    and startdate <= CONVERT(varchar(8),GETDATE(),112) " _
								  & "                    and enddate >= CONVERT(varchar(8),GETDATE(),112) " _
								  & "               ) " _
								  & "    and sysid = '" & sysid_ls & "' " _
								  & "    and pigsiteno = '" & pigsiteno_ls & "' " _
								  & "    and defPassword = '" & sitepw_ls & "' "   
							  run_fg = CheckDataExist(sql) 
							end if 
					  end if
					end if 
			  end if '判斷自訓時才要檢查密碼，1:比賽、2:測試環、3:自訓 結束
			end if ' IsEmptyEJ(sysid_ls) = "Y" 結束
			
			if run_fg = "N" then 
			  return_json("mobile") = mobile_ls 
			  return_json("totpage") = 0 
			  Set return_json("returndata") = jsArray()
			  AddMember "returndata","seq,pigsiteno,ringno,backtime,lo,la,uid" , ",,,,,,"
			  return_json.Flush
			else      
			  '比賽測試用
			  'if type_ls = "1" or type_ls = "2" then
			  '  nowbkdate = "2013/03/24"
			  '  nowdate = "20130324"
			  'end if 
			  'CONVERT(varchar(8),GETDATE(),112)
			  race_date_ls = left(qdate_ls,4) & mid(qdate_ls,6,2) &right(qdate_ls,2)
			  select_ls = " select a.pigeon_site_no , a.pring_no , a.backtime , a.lo , a.la , a.uid "
			  where_ls = " where 1=1 " _
					   & "   and a.sysid= '" & sysid_ls & "' " _
					   & "   and left(a.backtime,10) = '" & qdate_ls & "' "
					   
				if(sysid_ls <> "0801") then where_ls = where_ls & "   and a.Pigeon_site_no <> '9998' "
			  where_ls = where_ls & "   and a.Pigeon_site_no <> '9999' "
			  race_check_ls = ""
			  
			  '給預設值 
			  if IsEmptyEJ(page_ls) = "Y" then page_ls = 1
			  if IsEmptyEJ(pagenum_ls) = "Y" then pagenum_ls = 10
			  if IsEmptyEJ(orderby_ls) = "Y" then orderby_ls = 1
			  if IsEmptyEJ(type_ls) = "Y" then type_ls = 3
			  	
			  '排序方式 
			  if orderby_ls = 1 then orderby_str = " order by backtime desc "
			  if orderby_ls = 2 then orderby_str = " order by backtime asc "
			  if orderby_ls = 3 then orderby_str = " order by Pigeon_site_no asc "
			  if orderby_ls = 4 then orderby_str = " order by Pigeon_site_no desc "
			  	
			  if IsEmptyEJ(pigsiteno_ls) = "N" then 
				'因為自強在比賽時使用備用鴿鐘時也能取得正確的鴿舍編號 
					if ( (sysid_ls = "1012") and (type_ls=1) ) then
					  where_ls = where_ls & "    and ( a.Pigeon_site_no like '%" & pigsiteno_ls & "' or b.pigsiteno like '%" & pigsiteno_ls & "' ) "
					else
					  where_ls = where_ls & "    and a.Pigeon_site_no like '%" & pigsiteno_ls & "' "
					end if 
			  end if
				  
			  '判斷測試環 
			  if type_ls=2 then 
					where_ls = where_ls & " and left(a.ringno,2) = '00' "
			  else
					where_ls = where_ls & " and left(a.ringno,2) <> '00' "
			  end if
			  
			  '自強、中原增加判斷
			  leftjoin_ls = ""
			  if ( sysid_ls = "1107" ) or ( (sysid_ls = "1012") and (type_ls=1) ) then
					leftjoin_ls = "   left join sms_ringdata b "  _
								& "     on left(b.pring_No,8) = left(a.ringno,8) "  _
								& "    and b.sysid = a.sysid "  _
								& "    and exists ( select raceno  from sms_raceperiod " _
								& "                  where raceno = b.raceno " _
								& "                    and sysid = b.sysid " _
								& "                    and startdate <=  '" & race_date_ls &"'"_
								& "                    and enddate >=  '" & race_date_ls &"'" _
								& "               ) "
					select_ls = select_ls & " , b.pigsiteno "
			  end if  
			  
			  '取得要查詢的資料表 
			  if type_ls = 3 then 
					table_name = "pigsms_train a"
			  else
					table_name = "pigsms a "
					'判斷賽事 
					if ( type_ls = 1 ) then 
					  if not ( (s1="0702") or (sysid_ls="0804") or (sysid_ls="0907") or (sysid_ls="0908") or (sysid_ls="0909") or (sysid_ls="0910")  or (sysid_ls="9999") ) then
						race_check_ls = "  and exists ( select sysid from sms_racestation " _
									  & "                where sysid = a.sysid " _
									  & "                  and racedate = '" & race_date_ls & "' " _
									  & "                  and ( contestno = a.contestno or isnull(contestno,'')='' or ltrim(contestno) = '') " _
									  & "              ) "
					  end if
					end if
					
					if ( type_ls = 1 ) and (sysid_ls="0907" or sysid_ls="0908" or sysid_ls="0909" or sysid_ls="0910") then
					  table_name = "pigsms_train a "
					  race_check_ls = ""
					end if 
			  end if 'if type_ls = 3 then 結束
			  
			  '組合要查詢的sql語法 
			  sql = select_ls _
				  & " from " & table_name _
				  & leftjoin_ls _
				  & where_ls _
				  & race_check_ls _
				  & orderby_str
			  rs.open sql,conn,1,3   

			  IF RS.BOF=false THEN
					totnum_ls = rs.RecordCount  '全部筆數 
					rs.PageSize = pagenum_ls
					If CLng(page_ls) >= rs.PageCount Then page_ls = rs.PageCount
						totpage_ls = rs.PageCount  '全部頁數 
						return_json("mobile") = mobile_ls
						return_json("totpage") = totpage_ls
						return_json("totnum") = totnum_ls
						rs.AbsolutePage = page_ls
						Set return_json("returndata") = jsArray()
						For iPage = 1 To rs.PageSize
						  'showmessage(iPage)
						  if orderby_ls = 1 then 
								RecNo = rs.RecordCount - (page_ls - 1) * rs.PageSize - iPage +1
						  else 
								RecNo =(page_ls - 1) * rs.PageSize + iPage
						  end if
						  
						  if ( sysid_ls = "1107" ) or ( (sysid_ls = "1012") and (type_ls=1) ) then
								pigeon_site_no_ls = rs("pigsiteno")
						  else
								pigeon_site_no_ls = rs("pigeon_site_no")
						  end if
						  
						  ringno_ls = right(trim(rs("pring_no")),2)
						  backtime_ls = mid(rs("backtime"),12,8)
						  
						  lo1=rs("lo").value
						  lo = "000" & " ° "  & "00" & "’"  & "00" & "”"
						  if lo1<>"" then  lo = mid(lo1,1,4) & " ° "  & mid(lo1,6,2) & "’"  & mid(lo1,9,2) & "”"
					 
						  la1=rs("la").value
						  la = "000" & " ° "  & "00" & "’"  & "00" & "”"
						  if la1<>"" then la = mid(la1,1,4) & " ° "  & mid(la1,6,2) & "’"  & mid(la1,9,2) & "”"
		
						  UID = rs("uid")
						  A010_ls = RecNo _
								  & "," & pigeon_site_no_ls _
								  & "," & ringno_ls _
								  & "," & backtime_ls _
								  & "," & lo _
								  & "," & la _
								  & "," & UID
						  'showmessage(rs("pring_no"))
						  AddMember "returndata","seq,pigsiteno,ringno,backtime,lo,la,uid" , A010_ls
						  rs.MoveNext
						  If rs.EOF Then Exit For 
						next 
				  else
						return_json("mobile") = mobile_ls
						return_json("totpage") = "0"
						return_json("totnum") = "0"        
				  end if 'If CLng(page_ls) >= rs.PageCount Then page_ls = rs.PageCount 結束
			  rs.close
			  return_json.Flush
			end If 'if run_fg = "N" then 結束
	'台灣end--------------------------------------------------------------------------------
	else
		  '給預設值 
		  if IsEmptyEJ(page_ls) = "Y" then page_ls = 1
		  if IsEmptyEJ(pagenum_ls) = "Y" then pagenum_ls = 10
		  if IsEmptyEJ(orderby_ls) = "Y" then orderby_ls = 1
		  if IsEmptyEJ(type_ls) = "Y" then type_ls = 3
		  '排序方式 
		  if orderby_ls = 1 then orderby_str = " order by backtime desc "
		  if orderby_ls = 2 then orderby_str = " order by backtime asc "
		  if orderby_ls = 3 then orderby_str = " order by Pigeon_site_no asc "
		  if orderby_ls = 4 then orderby_str = " order by Pigeon_site_no desc "

			sql_select = " SELECT a.sysid , ISNULL(c.ENname, a.sysid) sysname , a.pigeon_site_no , a.pring_no ,case WHEN (V8_flag ='2' or V8_flag ='3') and isnull(timeZone,'') <> ''THEN case WHEN len(timeZone) < 3 then CONVERT(VARCHAR(24),DATEADD(n, convert(int, timeZone), CONVERT(VARCHAR(24), backtime, 121)),120) else CONVERT(VARCHAR(24),DATEADD(n, convert(int, timeZone), CONVERT(VARCHAR(24), backtime, 121)),120) END else backtime END backtime , a.lo , a.la , a.uid "

			'判斷查詢為比賽、自訓、測試環。 
			if (type_ls="1") or ( type_ls="2" ) then 
				sql_from =  "   FROM pigsms a left join sms_sysid c on c.sysid=a.sysid"
			else
				sql_from =  "   FROM pigsms_train a left join sms_sysid c on c.sysid=a.sysid"
			end if

			sql_where = "  where 1=1 and "
			sql_where = sql_where & " (case WHEN (V8_flag ='2' or V8_flag ='3') and isnull(timeZone,'') <> '' THEN case WHEN len(timeZone) < 3 and len(timeZone) > 1 then CONVERT(VARCHAR(10),DATEADD(n,CAST(isnull(timeZone,'0') AS float),CONVERT(VARCHAR(24),backtime,121)),111) else CONVERT(VARCHAR(10),DATEADD(n,CAST(isnull(timeZone,'0') AS float),CONVERT(VARCHAR(24),backtime, 121)),111) END else CONVERT(VARCHAR(10), backtime, 111) END) = '" & qdate_ls & "' "

			'判斷是否查詢測試環 
			if (type_ls="1") or ( type_ls="3" ) then
				sql_where = sql_where & "    and left(a.ringno,2)<>'00' "
			else
				sql_where = sql_where & "    and left(a.ringno,2)='00' "
			end if
			 
			'判斷是否有輸入鴿舍編號 
			if isnull(pigsiteno_ls) or (trim(pigsiteno_ls) ="") then 
			else
				sql_where = sql_where & "    and a.Pigeon_site_no like '%" & pigsiteno_ls & "' "
			end If

			'判斷是否有輸入設備編號 
			if isnull(sitepw_ls) or (trim(sitepw_ls) ="") then 
			else
				sql_where = sql_where & "    and a.DeviceNo = '" & sitepw_ls & "' "
			end If

			If (IsEmptyEJ(sysid_ls) = "N") Then	sql_where = sql_where & "    and a.sysid='" & sysid_ls & "' "
				
			If (country_ls = "Other") Then
				RaceFg_ls = "T"  'T自訓 R比賽 
				If ((type_ls="1") or ( type_ls="2" )) Then RaceFg_ls = "R" 

				sql_exists = " and exists(select '' from pigsms_item b where sysidFg='N' and associFg = 'N' and RaceFg = '"& RaceFg_ls &"' "
				sql_exists = sql_exists & " and BackTime = '"& qdate_ls &"' and a.AssociationCode = b.AssociationCode and a.SysId = b.SysId and a.ContestNo = b.ContestNo)"

			ElseIf (IsEmptyEJ(country_ls) = "N") Then
				'鴿會
				sql_exists =  " and exists(select '' from sms_sysid b where country = '"& country_ls &"' and a.sysid = b.sysid) "
				'協會
				sql_exists =  sql_exists &" and exists(select '' from sms_racestation c where a.sysid = c.sysid and a.ContestNo = c.ContestNo "
				sql_exists =  sql_exists &" and racedate <= '"& qdate_ls &"' and isnull(ContestNo,'') <> '' and isnull(sheetDateLimit,'') <> '' "
				sql_exists =  sql_exists &" and a.backtime <= convert(varchar, convert(datetime, sheetDateLimit), 111) +' '+  "
				sql_exists =  sql_exists &" substring(RaceInTime,1,2)+':'+substring(RaceInTime,3,2)+':'+substring(RaceInTime,5,2)  "
				sql_exists =  sql_exists &" ) "

				sql_exists_sms_sysid = " and exists(select '' from sms_sysid b where country = '"& country_ls &"' and a.sysid = b.sysid)  " 

				sql_exists_association = " and exists(select '' from SMS_association b where isvalfg = 'Y' and showfg='Y' and countrycode = '"& country_ls &"' and a.AssociationCode = b.fancierCode )  "
			End If 'If (country_ls = "Other") Then 結束
			
			sql = ""
			If (country_ls = "Other") Then
				sql = sql_select & sql_from & sql_where & sql_exists & orderby_str 
			ElseIf (IsEmptyEJ(country_ls) = "N") Then
				sql = sql_select & sql_from & sql_where & sql_exists
				sql =  sql & " union "
				sql =  sql & sql_select & sql_from & sql_where & sql_exists_sms_sysid
				sql =  sql & " union "
				sql =  sql & sql_select & sql_from & sql_where & sql_exists_association
				sql =  sql & orderby_str
			End if

		  rs.open sql,conn_en,1,3
		  IF RS.BOF=false THEN
				totnum_ls = rs.RecordCount  '全部筆數 
				rs.PageSize = pagenum_ls
				
				If CLng(page_ls) >= rs.PageCount Then page_ls = rs.PageCount
					totpage_ls = rs.PageCount  '全部頁數 
					return_json("mobile") = mobile_ls
					return_json("totpage") = totpage_ls
					return_json("totnum") = totnum_ls
					rs.AbsolutePage = page_ls
					Set return_json("returndata") = jsArray()
					For iPage = 1 To rs.PageSize
					  if orderby_ls = 1 then 
							RecNo = rs.RecordCount - (page_ls - 1) * rs.PageSize - iPage +1
					  else 
							RecNo =(page_ls - 1) * rs.PageSize + iPage
					  end if
					  
					  if ( sysid_ls = "1107" ) or ( (sysid_ls = "1012") and (type_ls=1) ) then
							pigeon_site_no_ls = rs("pigsiteno")
					  else
							pigeon_site_no_ls = rs("pigeon_site_no")
					  end if
	
					  ringno_ls = trim(rs("pring_no"))
					  backtime_ls = mid(rs("backtime"),12,8)
	 
					  lo1=rs("lo").value
					  lo = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if lo1<>"" then lo = mid(lo1,1,4) & " ° "  & mid(lo1,6,2) & "’"  & mid(lo1,9,2) & "”"
				 
					  la1=rs("la").value
					  la = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if la1<>"" then la = mid(la1,1,4) & " ° "  & mid(la1,6,2) & "’"  & mid(la1,9,2) & "”"
	
					  UID = rs("uid")
					  sysid_ls = rs("sysid")
					  sysname_ls = rs("sysname")
					  A010_ls = RecNo _
							  & "," & sysid_ls _
							  & "," & sysname_ls _
							  & "," & pigeon_site_no_ls _
							  & "," & ringno_ls _
							  & "," & backtime_ls _
							  & "," & lo _
							  & "," & la _
							  & "," & UID
					  'showmessage(rs("pring_no"))
					  AddMember "returndata","seq,sysid,sysname,pigsiteno,ringno,backtime,lo,la,uid" , A010_ls
					  rs.MoveNext
					  If rs.EOF Then Exit For 
					next 
			  else
					return_json("mobile") = mobile_ls
					return_json("totpage") = "0"
					return_json("totnum") = "0"        
			  end if 'If CLng(page_ls) >= rs.PageCount Then page_ls = rs.PageCount
		  rs.close
		  return_json.Flush
			end If ' 判斷 country_ls ，IF RS.BOF=false THEN 結束
  end If ' A010 結束

'======================================================================================================================  
  if  ( opkind_ls = "A011" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A011","sysid":"1012","pigsiteno":"0111","sitepw":"E61D47B68F"}
    sysid_ls = objINjson.sysid
    pigsiteno_ls = trim(objINjson.pigsiteno)
    sitepw_ls = objINjson.sitepw
    sql = " select sysid , pigsiteno from sms_sitedata "_
        & "  where sysid = '" & sysid_ls & "' " _
        & "    and pigsiteno = '" & pigsiteno_ls & "' " _
        & "    and defPassword = '" & sitepw_ls & "' " _
        & "    and exists ( select raceno  from sms_raceperiod " _
        & "                  where raceno = sms_sitedata.raceno " _
        & "                    and sysid = sms_sitedata.sysid " _
        & "                    and startdate <= CONVERT(varchar(8),GETDATE(),112) " _
        & "                    and enddate >= CONVERT(varchar(8),GETDATE(),112) " _
        & "               ) "
    rs.open sql,conn,1,3
    return_json("mobile") = mobile_ls 
    return_json("sysid") = sysid_ls 
    return_json("pigsiteno") = pigsiteno_ls 
    return_json("sitepw") = sitepw_ls 
    if rs.eof then       
      return_json("rtnfg") = "N" 
    else
      return_json("rtnfg") = "Y" 
    end if
    rs.close
    return_json.Flush
  end if 
  
'======================================================================================================================  
  if  ( opkind_ls = "A020" ) then
    'appcheck.asp?jsondata={"mobile":"0930768736","opkind":"A020","smsrp":[{"sysid":"1012","pigsiteno":"0709"}]}
    'appcheck.asp?jsondata={"mobile":"0929369536","opkind":"A020","smsrp":[{"sysid":"1003","pigsiteno":"0003"}]}
    startpos = instr(json_data,"[")
    endpos = instr(json_data,"]")
    smsrp_ls = mid(json_data,startpos+1,(endpos-startpos)-1)
    sysidsite_num = objINjson.smsrp.length

    return_json("mobile") = mobile_ls
    Set return_json("smsdata") = jsArray()
    A020_field = "type,sysid,pigsiteno,ringno,bkdate,bktime"
    
    smsdata_num = 0
    for i = 0 to (sysidsite_num - 1)
      Set detail_obj = getJSONObject(mid(smsrp_ls,(i*35)+1+i,35))
      temp_sysid = trim(detail_obj.sysid)
      temp_pigsiteno = right("000"+trim(detail_obj.pigsiteno),4)

      '檢查手機、鴿會、鴿舍是否存在      
      sql = " select sysid , pigsiteno from sms_sitedata "_
          & "  where ( mobil = '" & mobile_ls & "' " _
          & "          or  recievermobil like '%" & mobile_ls & "%' )" _
          & "    and sysid = '" & temp_sysid &"' "_
          & "    and pigsiteno = '" & temp_pigsiteno &"' "_
          & "    and exists ( select raceno  from sms_raceperiod " _
          & "                  where raceno = sms_sitedata.raceno " _
          & "                    and sysid = sms_sitedata.sysid " _
          & "                    and startdate <= CONVERT(varchar(8),GETDATE(),112) " _
          & "                    and enddate >= CONVERT(varchar(8),GETDATE(),112) " _
          & "               ) "
          
      if CheckDataExist(sql) = "Y" then
        '檢查查核資料是否存在
        sql = " select receivefg from app_smssite " _
            & "  where sysid = '" & temp_sysid &"' " _
            & "    and pigsiteno = '" & temp_pigsiteno &"' "_
            & "    and mobile = '" & mobile_ls &"' "
        
        rs.open sql,conn,1,3
        if rs.eof then 
          '新增資料 
          ins_ls = " insert into app_smssite " _
                 & "      ( sysid , pigsiteno , mobile " _
                 & "      , receivefg , Syscheck , ins_dt ) " _
                 & " values "_
                 & "      ( '" & temp_sysid & "' " _
                 & "      , '" & temp_pigsiteno & "' " _
                 & "      , '" & mobile_ls & "' " _
                 & "      , 'Y' " _
                 & "      , 'Y' " _
                 & "      , '" & nowdatetime & "' " _
                 & "      ) "
          RunExecSql(ins_ls)
        else
          if rs("receivefg") = "N" then
            '更新資料
            upd_ls = " update app_smssite set receivefg = 'Y' " _
                   & "      , Syscheck = 'Y' "_
                   & "  where sysid = '" & temp_sysid &"' "_
                   & "    and pigsiteno = '" & temp_pigsiteno &"' "_
                   & "    and mobile = '" & mobile_ls &"' "
            RunExecSql(upd_ls)
          end if
        end if
        rs.close
        
        '取得資料
        sql = " select * from APP_SMS " _
            & "  where sysid = '" & temp_sysid &"' "_
            & "    and pigsiteno = '" & temp_pigsiteno &"' "_
            & "    and mobile = '" & mobile_ls &"' "_
            & "    and ( Sendfg = 'N' or Sendfg = '' or Sendfg is null ) "_
            & "    and left(backtime,10) = '" & nowbkdate &"' "_
            & "  order by backtime desc "
        
        rs.open sql,conn,1,3
        do while ( not rs.eof )
        	'A020_field = "type,sysid,pigsiteno,ringno,bkdate,bktime"
        	A020_ls = rs("Stype") _
        	        & "," & rs("sysid") _
        	        & "," & rs("pigsiteno") _
        	        & "," & rs("ringno") _
        	        & "," & rs("bkdate") _
        	        & "," & rs("bktime")
        	AddMember "smsdata", A020_field , A020_ls
        	'更新資料
        	upd_ls = " update APP_SMS set Sendfg = 'Y' " _
                   & "    , mod_dt = '" & nowdatetime & "' " _
                 & "  where sysid = '" & temp_sysid &"' "_
                 & "    and pigsiteno = '" & temp_pigsiteno &"' "_
                 & "    and mobile = '" & mobile_ls &"' " _
                 & "    and backtime = '" & rs("backtime") &"' " _
                 & "    and ringno = '" & rs("ringno") &"' "
          RunExecSql(upd_ls)
          rs.MoveNext
        loop
        rs.close
      end if 'if CheckDataExist(sql) = "Y" then 結束
    next 'for i = 0 to (sysidsite_num - 1) 結束
    return_json.Flush
  end if
  
'======================================================================================================================    
  if  ( opkind_ls = "A030" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A030","country":"TWN"}
    return_json("mobile") = mobile_ls    
    Set return_json("sysiddata") = jsArray()
    
		if objINjson.country = "TWN" then
			sql = " select sysid , sysname, TrainOpenfg from sms_sysid " _
				& "  where showfg='Y' " _
				& "  order by seqno "
				
'				'2018/08/23 修改成APP和PC版的鴿會選項是相同的	
'				if show_trainOpenFg = "N" then 
'				  sql = " select sysid , sysname from sms_sysid " _
'					  & "  where showfg='Y' and ( trainOpenFg <> 'N' or trainOpenFg is null ) " _
'					  & "  order by seqno "
'				end If

			rs.open sql,conn,1,3
		End If

		if objINjson.country = "Other" then
			sql = " select distinct sysid, sysid as sysname, '' as TrainOpenfg from pigsms_item " _
				& "  where associfg = 'N' and sysidfg='N' " _
				& "  order by sysid "
			rs.open sql,conn_en,1,3
		End If

		'不屬於其他或是台灣的，就連線到國外資料庫去做協會的判斷，管理協會代碼、SMS_association 協會對應鴿會
		if Not (objINjson.country = "Other" or objINjson.country = "TWN") then
			sql = " select sysid , enname sysname, TrainOpenfg from sms_sysid " _
				& "  where showfg='Y' and enflag = 'Y' " _
				& "  and Country = '" & objINjson.country & "' " _
				& "  union " _
				& "  select distinct sysid, sysid as sysname, '' as TrainOpenfg " _
				& "  from pigsms_item a  " _
				& "  left join SMS_association b on a.associationcode = b.fanciercode  " _
				& "  where b.IsValFg='Y' and not (a.associfg = 'N' and a.sysidfg='N')  " _
				& "  and b.Countrycode = '" & objINjson.country & "'  " _
				& "  order by sysid "
			rs.open sql,conn_en,1,3
		End If

		do while ( not rs.eof )
			'2018/08/23 增加比賽、測試、自訓
			tmodeType = "NNN"
			if Trim(rs("TrainOpenfg")) = "N" Then tmodeType = "NNY" 
			
		  A030_ls = rs("sysid") & "," & rs("sysname") & "," & tmodeType
		  'showmessage(A030_ls)
		  AddMember "sysiddata","sysid,sysname,modeType" , A030_ls
		  rs.MoveNext
		loop	

    rs.close
    return_json.Flush
  end If 'A030 結束

'======================================================================================================================  
  if  ( opkind_ls = "A040" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A040"}
    return_json("mobile") = mobile_ls    
    Set return_json("countrydata") = jsArray()
    sql = " select ShowValue, showdesc from com_Gcode " _
        & "  where M_code='0017' and useflag='Y' " _
        & "  order by showdesc asc "

    rs.open sql,conn_en,1,3
    do while ( not rs.eof )
      A040_ls = rs("ShowValue") & "," & rs("showdesc")
      AddMember "countrydata","showvalue,showdesc" , A040_ls
      rs.MoveNext
    loop
    rs.close
    return_json.Flush
  end If
  
'======================================================================================================================  
  if  ( opkind_ls = "A050" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A050","country":"ROU"}
    country_ls = objINjson.country

		If ( country_ls <> "")  then
				Set return_json("associationdata") = jsArray()
				sql = " select associationNo, associationName from SMS_association " _
					& "  where countrycode='" & country_ls & "'  " _
					& "  order by associationName asc "
	
				rs.open sql,conn_en,1,3
				do while ( not rs.eof )
				  A050_ls = rs("associationNo") & "," & rs("associationName")
				  AddMember "associationdata","associationNo,associationName" , A050_ls
				  rs.MoveNext
				loop
				rs.close
				return_json.Flush
			
		else if ( country_ls = "" )  then
				Set return_json("associationdata") = jsArray()
				sql = " select distinct  associationcode associationNo, associationcode associationName from pigsms_item " _
					& "  order by associationcode asc "
				'showmessage(sql)
				rs.open sql,conn_en,1,3
				do while ( not rs.eof )
				  A050_ls = rs("associationNo") & "," & rs("associationName")
				  'showmessage(A040_ls)
				  AddMember "associationdata","associationNo,associationName" , A050_ls
				  rs.MoveNext
				loop
				rs.close
				return_json.Flush
		end if
	end if

  end If '?

'======================================================================================================================  
  if  ( opkind_ls = "A060" ) then 
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A060","country":"TWN","association":""}

    country_ls = objINjson.country
    association_ls = objINjson.association

		If ( country_ls = "TWN" and association_ls = "" )  then
				Set return_json("sysdata") = jsArray()
				sql = " select sysid , sysname from sms_sysid " _
					& "  where showfg='Y' and sysid <> '9999' order by seqno "
				'showmessage(sql)
				rs.open sql,conn,1,3
				do while ( not rs.eof )
				  A060_ls = rs("sysid") & "," & rs("sysname")
				  'showmessage(A040_ls)
				  AddMember "sysdata","sysid,sysname" , A060_ls
				  rs.MoveNext
				loop
				rs.close
				return_json.Flush
			
		Else if ( country_ls = "" and association_ls <> "" )  then
				Set return_json("sysdata") = jsArray()
				sql = " select  distinct sysid sysid, sysid sysname from pigsms_item " _
					& "  where associationcode='" & association_ls & "'  " _
					& "  order by sysid asc "
				'showmessage(sql)
				rs.open sql,conn_en,1,3
				do while ( not rs.eof )
				  A060_ls = rs("sysid") & "," & rs("sysname")
				  'showmessage(A040_ls)
				  AddMember "sysdata","sysid,sysname" , A060_ls
				  rs.MoveNext
				loop
				rs.close
				return_json.Flush
		Else
				Set return_json("sysdata") = jsArray()
				sql = " select sysid , enname sysname from sms_sysid "
				sql = sql & "  where 1=1  "
	
				if country_ls <> "" then sql = sql & "  and country='" & country_ls & "'  "
				if association_ls <> "" then	sql = sql & "  and associationNo='" & association_ls & "'  "
	
				'showmessage(country_ls)			
				if ( country_ls = "ROU" ) or ( country_ls = "rou" ) Then
				else
					sql = sql & "  and showfg='Y' "
				end if
				sql = sql & "  and sysid <> '9999' and sysid <> 'AVANCE' order by seqno "
	
				rs.open sql,conn_en,1,3
				do while ( not rs.eof )
				  A060_ls = rs("sysid") & "," & rs("sysname")
				  'showmessage(A040_ls)
				  AddMember "sysdata","sysid,sysname" , A060_ls
				  rs.MoveNext
				loop
				rs.close
				return_json.Flush			
			end if
		end if '?
  end If '?

'====================================================================================================================== 
  '20170810 增加AssociationCode查詢，顯示筆數，改為收到第幾筆，回傳幾筆
  if ( opkind_ls = "A070" ) then 
    '賽事：appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A070","type":"1","sysid":"1003","pigsiteno":""}
    '自訓：appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A070","type":"3","sysid":"1003","sitepw":"Y","pigsiteno":"0015"}
    'appcheck.asp?jsondata={"mobile":"0930768736","opkind":"A070","type":"3","sysid":"1012","sitepw":"Y","pigsiteno":"0709"}
    'appcheck.asp?jsondata={"mobile":"0920417450","opkind":"A070","type":"3","sysid":"1012","sitepw":"D52C49BAC4","pigsiteno":"0709","country":"TWN"}
    '要先檢查參數是否存在 
    if InStr(json_data,"type") then
      type_ls = objINjson.type
    else
      type_ls = ""
    end if 
    
    if InStr(json_data,"sysid") then
      sysid_ls = objINjson.sysid
    else
      sysid_ls = ""
    end if 
    
    if InStr(json_data,"pigsiteno") then
      pigsiteno_ls = objINjson.pigsiteno
    else
      pigsiteno_ls = ""
    end if 
    
    if InStr(json_data,"RowNumls") then
      RowNumls_ls = objINjson.RowNumls
    else
      RowNumls_ls = ""
    end if 
    
    if InStr(json_data,"pagenum") then
      pagenum_ls = objINjson.pagenum
    else
      pagenum_ls = ""
    end if 
    
    if InStr(json_data,"orderby") then
      orderby_ls = objINjson.orderby 
    else
      orderby_ls = ""
    end if 
    
    if InStr(json_data,"sitepw") then
      sitepw_ls = objINjson.sitepw 
    else
      sitepw_ls = ""
    end if 
    
    '2013-06-28
    if InStr(json_data,"qdate") then
      qdate_ls = objINjson.qdate
    else
      qdate_ls = nowbkdate
    end if 

    if InStr(json_data,"associationcode") then
      associationcode_ls = objINjson.associationcode
    else
      associationcode_ls = ""
    end if 

    if InStr(json_data,"country") then
      country_ls = objINjson.country
    else
      country_ls = "TWN"
    end if 

    run_fg = "Y"
	
	'台灣
	if country_ls = "TWN" then 
			if IsEmptyEJ(sysid_ls) = "Y" then 
			  run_fg = "N"
			Else
			  '107.08.14 光老闆要全部自訓都可以查詢
			end if
			if run_fg = "N" then 
			  return_json("mobile") = mobile_ls 
			  return_json("totpage") = 0 
			  Set return_json("returndata") = jsArray()
			  AddMember "returndata","seq,pigsiteno,ringno,backtime,lo,la,uid" , ",,,,,,"
			  return_json.Flush
			else      

			  race_date_ls = left(qdate_ls,4) & mid(qdate_ls,6,2) &right(qdate_ls,2)
			  '給預設值 
			  if IsEmptyEJ(RowNumls_ls) = "Y" then RowNumls_ls = 1
			  if IsEmptyEJ(pagenum_ls) = "Y" then pagenum_ls = 10
			  if IsEmptyEJ(orderby_ls) = "Y" then orderby_ls = 1
			  if IsEmptyEJ(type_ls) = "Y" then type_ls = 3
			   pagenum_ls = pagenum_ls + RowNumls_ls
			  '排序方式 
			  if orderby_ls = 1 then orderby_str = " order by backtime desc "
			  if orderby_ls = 2 then orderby_str = " order by backtime asc "
			  if orderby_ls = 3 then orderby_str = " order by Pigeon_site_no asc "
			  if orderby_ls = 4 then orderby_str = " order by Pigeon_site_no desc "
			  select_ls = " select ROW_NUMBER() OVER (" & orderby_str & ") sort ,a.pigeon_site_no , a.pring_no , a.backtime , a.lo , a.la , a.uid "
			  where_ls = " where 1=1 " _
					   & "   and a.sysid= '" & sysid_ls & "' " _
					   & "   and left(a.backtime,10) = '" & qdate_ls & "' "
			  where_ls = where_ls & "   and a.Pigeon_site_no <> '9998' "
			  where_ls = where_ls & "   and a.Pigeon_site_no <> '9999' "
			  race_check_ls = ""

			  if IsEmptyEJ(pigsiteno_ls) = "N" then 
					'因為自強在比賽時使用備用鴿鐘時也能取得正確的鴿舍編號 
					if ( (sysid_ls = "1012") and (type_ls=1) ) then
					  where_ls = where_ls & "    and ( a.Pigeon_site_no like '%" & pigsiteno_ls & "' or b.pigsiteno like '%" & pigsiteno_ls & "' ) "
					else
					  where_ls = where_ls & "    and a.Pigeon_site_no like '%" & pigsiteno_ls & "' "
					end if 
			  end if
			  
			  '判斷測試環 
			  if type_ls=2 then 
					where_ls = where_ls & " and left(a.ringno,2) = '00' "
			  else
					where_ls = where_ls & " and left(a.ringno,2) <> '00' "
			  end if
			  
			  '自強、中原增加判斷
			  leftjoin_ls = ""
			  if (type_ls=1)  then
					leftjoin_ls = "   left join sms_ringdata b "  _
								& "     on left(b.pring_No,8) = left(a.ringno,8) "  _
								& "    and b.sysid = a.sysid "  _
								& "    and exists ( select raceno  from sms_raceperiod " _
								& "                  where raceno = b.raceno " _
								& "                    and sysid = b.sysid " _
								& "                    and startdate <=  '" & race_date_ls &"'"_
								& "                    and enddate >=  '" & race_date_ls &"'" _
								& "               ) "
					select_ls = select_ls & " , b.pigsiteno "
			  end if  
			  
			  '取得要查詢的資料表 
			  if type_ls = 3 then 
					table_name = "pigsms_train a"
			  else
					table_name = "pigsms a "
					
					'判斷賽事 
					if ( type_ls = 1 ) then 
					  if not ( (s1="0702") or (sysid_ls="0804") or (sysid_ls="0907") or (sysid_ls="0908") or (sysid_ls="0909") or (sysid_ls="0910")  or (sysid_ls="9999") ) then
						race_check_ls = "  and exists ( select sysid from sms_racestation " _
									  & "                where sysid = a.sysid " _
									  & "                  and racedate = '" & race_date_ls & "' " _
									  & "                  and ( contestno = a.contestno or isnull(contestno,'')='' or ltrim(contestno) = '') " _
									  & "              ) "
					  end if
					end if
				
					if ( type_ls = 1 ) and (sysid_ls="0907" or sysid_ls="0908" or sysid_ls="0909" or sysid_ls="0910") then
					  table_name = "pigsms_train a "
					  race_check_ls = ""
					end if 
			  end if
			  
			  '組合要查詢的sql語法 
			  sql = "select * from (" _
				  & select_ls _
				  & " from " & table_name _
				  & leftjoin_ls _
				  & where_ls _
				  & race_check_ls _
				  & " ) a where a.sort >= '" & RowNumls_ls & "' and a.sort < '" & pagenum_ls & "'       "
			  rs.open sql,conn,1,3   

			  IF RS.BOF=false THEN
					return_json("mobile") = mobile_ls
					return_json("totnum") = pagenum_ls - RowNumls_ls
					Set return_json("returndata") = jsArray()
					For  X= 1  To  pagenum_ls  
					  'showmessage(iPage)
					  if ( sysid_ls = "1107" ) or ( (sysid_ls = "1012") and (type_ls=1) ) then
							pigeon_site_no_ls = rs("pigsiteno")
					  else
							pigeon_site_no_ls = rs("pigeon_site_no")
					  end if
					  
					  ringno_ls = right(trim(rs("pring_no")),2)
					  backtime_ls = mid(rs("backtime"),12,8)

					  lo1=rs("lo").value
					  lo = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if lo1<>"" then lo = mid(lo1,1,4) & " ° "  & mid(lo1,6,2) & "’"  & mid(lo1,9,2) & "”"
				 
					  la1=rs("la").value
					  la = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if la1<>"" then la = mid(la1,1,4) & " ° "  & mid(la1,6,2) & "’"  & mid(la1,9,2) & "”"

					  UID = rs("uid")
					  A070_ls = RecNo _
							  & "," & pigeon_site_no_ls _
							  & "," & ringno_ls _
							  & "," & backtime_ls _
							  & "," & lo _
							  & "," & la _
							  & "," & UID
					  AddMember "returndata","seq,pigsiteno,ringno,backtime,lo,la,uid" , A070_ls
					  rs.MoveNext
					  If rs.EOF Then Exit For 
					next 
			  else
					return_json("mobile") = mobile_ls
					return_json("totnum") = "0"        
			  end if
			  rs.close
			  return_json.Flush
			end If
	'台灣end
	else

			  '給預設值 
			  if IsEmptyEJ(RowNumls_ls) = "Y" then RowNumls_ls = 1
			  if IsEmptyEJ(pagenum_ls) = "Y" then pagenum_ls = 10
			  if IsEmptyEJ(orderby_ls) = "Y" then orderby_ls = 1
			  if IsEmptyEJ(type_ls) = "Y" then type_ls = 3
			   pagenum_ls = pagenum_ls + RowNumls_ls - 1
			  '排序方式 
			  if orderby_ls = 1 then orderby_str = " order by backtime desc "
			  if orderby_ls = 2 then orderby_str = " order by backtime asc "
			  if orderby_ls = 3 then orderby_str = " order by Pigeon_site_no asc "
			  if orderby_ls = 4 then orderby_str = " order by Pigeon_site_no desc "

			  if orderby_ls_top = 1 then orderby_str_top = " order by backtime asc "
			  if orderby_ls_top = 2 then orderby_str_top = " order by backtime desc "
			  if orderby_ls_top = 3 then orderby_str_top = " order by Pigeon_site_no desc "
			  if orderby_ls_top = 4 then orderby_str_top = " order by Pigeon_site_no asc "

			sql = " SELECT  a.sysid , ISNULL(c.ENname, a.sysid) sysname , a.pigeon_site_no , a.pring_no , case WHEN (V8_flag ='2' or V8_flag ='3') and isnull(timeZone,'') <> '' THEN case WHEN len(timeZone) < 3 then CONVERT(VARCHAR(24),DATEADD(hh, convert(int, timeZone), CONVERT(VARCHAR(24), backtime, 121)),120) else CONVERT(VARCHAR(24),DATEADD(n, convert(int, timeZone), CONVERT(VARCHAR(24), backtime, 121)),120) END else backtime END as backtime , a.lo , a.la , a.uid "

			'判斷查詢為比賽、自訓、測試環。 
			if (type_ls="1") or ( type_ls="2" ) then 
				sql = sql & "   FROM pigsms a left join sms_sysid c on c.sysid=a.sysid"
			else
				sql = sql & "   FROM pigsms_train a left join sms_sysid c on c.sysid=a.sysid"
			end if

			sql = sql & "  where 1=1 "
			sql = sql & " and left(a.backtime,10) = '" & qdate_ls & "' "

			'判斷是否查詢測試環 
			if (type_ls="1") or ( type_ls="3" ) then
				sql = sql & "    and left(a.ringno,2)<>'00' "
			else
				sql = sql & "    and left(a.ringno,2)='00' "
			end if
			 
			'判斷是否有輸入鴿舍編號 
			if isnull(pigsiteno_ls) or (trim(pigsiteno_ls) ="") then 
			else
				sql = sql & "    and a.Pigeon_site_no like '%" & pigsiteno_ls & "' "
			end If
			
			'判斷是否有associationcode 
			if isnull(associationcode_ls) or (trim(associationcode_ls) ="") then 
			else
				sql = sql & "    and a.associationcode like '%" & associationcode_ls & "' "
			end If

			'判斷是否有輸入設備編號 
			if isnull(sitepw_ls) or (trim(sitepw_ls) ="") then 
			else
				sql = sql & "    and a.DeviceNo = '" & sitepw_ls & "' "
			end If

			If (IsEmptyEJ(sysid_ls) = "N") Then	sql = sql & "    and a.sysid='" & sysid_ls & "' "
				
			If (country_ls = "Other") Then
				RaceFg_ls = "T"  'T自訓 R比賽 
				If ((type_ls="1") or ( type_ls="2" )) Then	RaceFg_ls = "R" 
				sql = sql & " and exists(select '' from pigsms_item b where sysidFg='N' and associFg = 'N' and RaceFg = '"& RaceFg_ls &"' "
				sql = sql & " and BackTime = '"& qdate_ls &"' and a.AssociationCode = b.AssociationCode and a.SysId = b.SysId and a.ContestNo = b.ContestNo)"

			ElseIf (IsEmptyEJ(country_ls) = "N") Then
				'鴿會
				sqlC =  sql &" and exists(select '' from sms_sysid b where country = '"& country_ls &"' and a.sysid = b.sysid) "
				'協會
				sqlH =  sql &" and exists(select '' from SMS_association b where isvalfg = 'Y' and countrycode = '"& country_ls &"' "
				sqlH =  sqlH &" and a.AssociationCode = b.fancierCode "
				sqlH =  sqlH &" ) "

				sql = ""
				sql = sqlC &" union "& sqlH
			End if

			sql = sql & orderby_str 

			  '組合要查詢的sql語法 
			  sql_top = " select c.*  "
			  sql_top = sql_top & " from ( "
			  sql_top = sql_top & " select c.* "
			  sql_top = sql_top & " from ( "
			  sql_top = sql_top & " select top " & RowNumls_ls & " b.* "
			  sql_top = sql_top & " from ( "
			  sql_top = sql_top & " select top " & pagenum_ls & " a.* "
			  sql_top = sql_top & " from ( "
			  sql_top = sql_top &  sql 
			  sql_top = sql_top & " ) a "
			  sql_top = sql_top & " ) b "
			  sql_top = sql_top & orderby_str_top
			  sql_top = sql_top & " ) c "
			  sql_top = sql_top & orderby_str

			  showmessage(sql) '看起來沒用到A070
			  Response.end

			  IF RS.BOF=false THEN
					return_json("mobile") = mobile_ls
					return_json("totnum") = pagenum_ls - RowNumls_ls
					Set return_json("returndata") = jsArray()
					For  X= 1  To  pagenum_ls  
					  if orderby_ls = 1 then 
							RecNo = rs.RecordCount - (page_ls - 1) * rs.PageSize - iPage +1
					  else 
							RecNo =(page_ls - 1) * rs.PageSize + iPage
					  end if
	
					  if ( sysid_ls = "1107" ) or ( (sysid_ls = "1012") and (type_ls=1) ) then
							pigeon_site_no_ls = rs("pigsiteno")
					  else
							pigeon_site_no_ls = rs("pigeon_site_no")
					  end if
					  ringno_ls = right(trim(rs("pring_no")),2)
					  backtime_ls = mid(rs("backtime"),12,8)
	
					  lo1=rs("lo").value
					  lo = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if lo1<>"" then lo = mid(lo1,1,4) & " ° "  & mid(lo1,6,2) & "’"  & mid(lo1,9,2) & "”"
				 
					  la1=rs("la").value
					  la = "000" & " ° "  & "00" & "’"  & "00" & "”"
					  if la1<>"" then la = mid(la1,1,4) & " ° "  & mid(la1,6,2) & "’"  & mid(la1,9,2) & "”"
	
					  UID = rs("uid")
					  sysid_ls = rs("sysid")
					  sysname_ls = rs("sysname")
					  A070_ls = RecNo _
							  & "," & sysid_ls _
							  & "," & sysname_ls _
							  & "," & pigeon_site_no_ls _
							  & "," & ringno_ls _
							  & "," & backtime_ls _
							  & "," & lo _
							  & "," & la _
							  & "," & UID
					  AddMember "returndata","seq,sysid,sysname,pigsiteno,ringno,backtime,lo,la,uid" , A070_ls
					  rs.MoveNext
					  If rs.EOF Then Exit For 
					next 
			  else
					return_json("mobile") = mobile_ls
					return_json("totpage") = "0"
					return_json("totnum") = "0"        
			  end if
			  rs.close
			  return_json.Flush
	end If

  end If

conn_en.close  
conn.close
%>
