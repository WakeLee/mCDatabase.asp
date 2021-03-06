<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<!--#include virtual="/mCDatabase/__public.asp"-->
<!--#include virtual="/mCDatabase/mCDatabase.js.asp"-->

<%var title = "ASP / MySQL"; %>
<title><%=title%></title>
<%=title%>
<table border="1" cellspacing="5" cellpadding="5">
<%			
var db = new mCDatabase();
db.OpenDb( filetojson("_mysql.json") );

// 增
// Insert
{
	var rs = db.OpenRs();
		
	rs.SetInt("m_tinyint", 127);
	rs.SetInt("m_smallint", 32767);
	rs.SetInt("m_int", 2147483647);
	rs.SetInt("m_bigint", 2147483647);
	
	rs.SetDouble("m_double", 1234.5678);
	
	rs.SetString("m_char5", "12345");
	rs.SetString("m_varchar5", "12345");
	rs.SetString("m_text", title);
	
	rs.SetDateTime("m_datetime", "");

	rs.Insert("mctable");
	db.CloseRs(rs);
}

// 删
// Delete
{
	var rs = db.OpenRs();
	rs.SetWhere("ID >= 71 and ID <= 76");
	rs.Delete("mctable");
	db.CloseRs(rs);
}

// 改
// Update
{
	var rs = db.OpenRs();
	rs.SetInt("m_bigint", 88);
	rs.SetWhere("ID >= 66 and ID <= 80");
	rs.Update("mctable");
	db.CloseRs(rs);
}

// 查
// Query
{
	var table = "";

	var rs = db.OpenRs();
	rs.Query("select *,FROM_UNIXTIME(UNIX_TIMESTAMP(m_datetime), '%Y-%m-%d %H:%i:%S') AS m_datetimeF from mctable order by ID desc limit 10");
	
	table += "<tr>";
	var iColumnCount = rs.GetColumnCount();
	for(var i = 0; i < iColumnCount; i++)
	{
		table += "<td>";
		table += rs.GetColumnName(i);
		table += "</td>";
	}
	table += "</tr>";
	
	while(!rs.eof)
	{
		table += "<tr>";
		table += "<td>" + rs.GetInt("ID") + "</td>";
		
		table += "<td>" + rs.GetInt("m_tinyint") + "</td>";
		table += "<td>" + rs.GetInt("m_smallint") + "</td>";
		table += "<td>" + rs.GetInt("m_int") + "</td>";
		table += "<td>" + rs.GetInt("m_bigint") + "</td>";
		
		table += "<td>" + rs.GetDouble("m_double") + "</td>";
		
		table += "<td>" + rs.GetString("m_char5") + "</td>";
		table += "<td>" + rs.GetString("m_varchar5") + "</td>";
		table += "<td>" + rs.GetString("m_text") + "</td>";
		
		table += "<td>" + rs.GetDateTime("m_datetime") + "</td>";
		table += "<td>" + rs.GetString("m_datetimeF") + "</td>";
		table += "</tr>";
		
		rs.MoveNext();
	}
	db.CloseRs(rs);
	
	Response.Write(table);
}

db.CloseDb();
%>
</table>
