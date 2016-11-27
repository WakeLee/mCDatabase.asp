<script language="javascript" runat="server">
// 作者: 李辙
// 找我: wakelee.coderwriter.com
// Author: Wake Lee
// FindMe: wakelee.coderwriter.com
function mCDatabase()
{
	//---- ObjectStateEnum Values ----
	var adStateClosed = 0x00000000;
	var adStateOpen = 0x00000001;
	var adStateConnecting = 0x00000002;
	var adStateExecuting = 0x00000004;
	var adStateFetching = 0x00000008;

	//---- CursorTypeEnum Values ----
	var adOpenForwardOnly = 0;
	var adOpenKeyset = 1;
	var adOpenDynamic = 2;
	var adOpenStatic = 3;

	//---- LockTypeEnum Values ----
	var adLockReadOnly = 1;
	var adLockPessimistic = 2;
	var adLockOptimistic = 3;
	var adLockBatchOptimistic = 4;

	//---- CursorLocationEnum Values ----
	var adUseServer = 2;
	var adUseClient = 3;
	
	//---- CommandTypeEnum Values ----
	var adCmdUnknown = 0x0008;
	var adCmdText = 0x0001;
	var adCmdTable = 0x0002;
	var adCmdStoredProc = 0x0004;
	var adCmdFile = 0x0100;
	var adCmdTableDirect = 0x0200;
	
	var bLog = true;
	var sLogPath = "";
	this.SetLog = function(_bLog, _sLogPath)
	{
		bLog = _bLog;
		sLogPath = _sLogPath;
	}

	Show = function(tip)
	{
		try
		{
			var html = "";
			html += "<div style='background-color:#b22222;color:#ffffe0;font-size:24px;padding:10px;margin:10px 0px;'>";
			html += tip;
			html += "</div>";
			
			Response.Write(html);
			
			if(bLog)
			{
				var sAbsolute = Server.MapPath(sLogPath + "/mCDatabase-logs");
				var fs = new ActiveXObject("Scripting.FileSystemObject");
				if( !fs.FolderExists(sAbsolute) ) fs.CreateFolder(sAbsolute);
				var f = fs.OpenTextFile(sAbsolute + "/mCDatabase-log-" + DateToStr() + ".txt", 8, true, -1);
				f.WriteLine("[ " + DateTimeToStr() + " ] [ " + GetUrl() + " ] [ " + tip + " ]");
				f.Close();
			}
		}
		catch(e)
		{
		}
	}
	
	var ErrorCode = 0;
	IsError = function(tip)
	{		
		switch(ErrorCode)
		{
			case 0: break;
			
			case 1: Show("Error : open connection error " + tip); break;
			case 2: Show("Error : close connection error " + tip); break;
			
			case 3: Show("Error : open recordset error " + tip); break;
			case 4: Show("Error : close recordset error " + tip); break;
			
			case 5: Show("Error : insert error " + tip); break;
			case 6: Show("Error : delete error " + tip); break;
			case 7: Show("Error : update error " + tip); break;
			case 8: Show("Error : query error " + tip); break;
			
			case 9: Show("Error : get record total count error " + tip); break;
			case 10: Show("Error : get field total count error " + tip); break;
			case 11: Show("Error : get field name error " + tip); break;
			
			case 12: Show("Error : get int error " + tip); break;
			case 13: Show("Error : get double error " + tip); break;
			case 14: Show("Error : get string error " + tip); break;
			case 15: Show("Error : get datetime error " + tip); break;
			
			case 16: Show("Error : move next error " + tip); break;
			
			case 17: Show("Error : driver error " + tip); break;
		}
		
		return ErrorCode == 0 ? false : true;
	}

	DateTimeToStr = function(DateTime)
	{
		DateTime = DateTime ? new Date(DateTime) : new Date();
		
		var str = "";
		str += DateTime.getFullYear() + "-";
		str += FillZero( DateTime.getMonth() + 1 ) + "-";
		str += FillZero( DateTime.getDate() ) + " ";
		str += FillZero( DateTime.getHours() ) + ":";
		str += FillZero( DateTime.getMinutes() ) + ":";
		str += FillZero( DateTime.getSeconds() );
		
		return str;
	}
	DateToStr = function(DateTime)
	{
		DateTime = DateTime ? new Date(DateTime) : new Date();
		
		var str = "";
		str += DateTime.getFullYear() + "-";
		str += FillZero( DateTime.getMonth() + 1 ) + "-";
		str += FillZero( DateTime.getDate() );
		
		return str;
	}
	FillZero = function(iNum, iBit)
	{
		if(!iBit) iBit = 2;
		
		var sZero = "";
		
		var sNum = iNum.toString();
		for(var i = 0; i < iBit - sNum.length; i++)
		{
			sZero += "0";
		}
		
		return sZero + sNum;
	}
	GetUrl = function()
	{
		var host = "" + Request.ServerVariables("HTTP_HOST");
		var url = "" + Request.ServerVariables("URL");
		var str = "" + Request.ServerVariables("QUERY_STRING");
		
		var full_url = "http://" + host + url;
		if(str != "") full_url += "?" + str;
	
		return full_url;
	}
	
	this.conn = Server.CreateObject("ADODB.Connection");
	this.dbtype = "";

	// 打开数据库
	// Open database
	this.OpenDb = function(option)
	{
		this.dbtype = option.dbtype;
		var ConnectionString;
		
		switch(this.dbtype)
		{
			case "mysql":
			{
				ConnectionString = 
				"Driver={" + option.dbdriver + "};" +
				"Server=" + option.dblocation + ";Port=" + option.dbport + ";" +
				"UID=" + option.uid + ";" +
				"PWD=" + option.pwd + ";" +
				"Database=" + option.dbname + ";";
			}
			break;

			case "sqlserver":
			{
				ConnectionString = 
				"Provider=" + option.dbdriver + ";" +
				"Data Source=" + option.dblocation + "," + option.dbport + ";" +
				"UID=" + option.uid + ";" +
				"PWD=" + option.pwd + ";" +
				"Initial Catalog=" + option.dbname + ";";
			}
			break;
			
			case "access":
			{
				ConnectionString = 
				"Provider=" + option.dbdriver + ";" +
				"Data Source=" + option.dblocation + ";";
			}
			break;
		}
		
		try
		{
			this.conn.Open(ConnectionString);
		}
		catch(e)
		{
			ErrorCode = 1;
			if( IsError("OpenDb() " + e.description) ) return;
		}
	}
	
	// 关闭数据库
	// Close database
	this.CloseDb = function()
	{				
		if( IsError("CloseDb()") ) return;
		
		try
		{
			if(this.conn && this.conn.State != adStateClosed)
			{
				this.conn.Close();
				this.conn = null;
			}
		}
		catch(e)
		{
			ErrorCode = 2;
			if( IsError("CloseDb() " + e.description) ) return;
		}
	}
	
	// 打开记录集
	// Open recordset
	this.OpenRs = function()
	{
		var rs = new CRs(this.dbtype, this.conn, bLog, sLogPath)
		
		if( IsError("OpenRs()") ) return rs;
		
		try
		{
			rs.rs = Server.CreateObject("ADODB.Recordset");
		}
		catch(e)
		{
			ErrorCode = 3;
			if( IsError("OpenRs() " + e.description) ) return;
		}
		
		return rs;
	}
	
	// 关闭记录集
	// Close recordset
	this.CloseRs = function(rs)
	{
		if( IsError("CloseRs()") ) return;
		
		try
		{
			if(rs.rs.State != adStateClosed)
			{
				rs.rs.Close();
				rs.rs = null;
			}
		}
		catch(e)
		{
			ErrorCode = 4;
			if( IsError("CloseRs() " + e.description) ) return;
		}
	}
	
	// 记录集类
	// Recordset class
	function CRs(dbtype, conn, bLog, sLogPath)
	{
		this.dbtype = dbtype;
		this.conn = conn;
		bLog = bLog;
		sLogPath = sLogPath;
		this.rs = null;
		this.eof = true;
		
		this.kvs = []; // key/value array
		this.where = "";
		
		// 增
		// Insert
		this.Insert = function(TableName)
		{
			if( IsError("Insert()") ) return;

			var sql = "insert into " + TableName;
			sql += "(";
			for(var i = 0; i < this.kvs.length; i++)
			{
				sql += this.kvs[i].key;
				if(i != this.kvs.length - 1) sql += ",";
			}
			sql += ")";
			sql += "values";
			sql += "(";
			for(var i = 0; i < this.kvs.length; i++)
			{
				switch( typeof(this.kvs[i].value) )
				{
					case "number":
					{
						sql += this.kvs[i].value;
					}
					break;
					
					case "string":
					{
						sql += "'" + this.kvs[i].value + "'";
					}
					break;
				}
				if(i != this.kvs.length - 1) sql += ",";
			}
			sql += ")";
			
			try
			{
				this.rs.Open(sql, this.conn, adOpenKeyset, adLockOptimistic);
			}
			catch(e)
			{
				ErrorCode = 5;
				if( IsError("Insert() " + e.description) ) return;
			}
		}
		
		// 删
		// Delete
		this.Delete = function(TableName)
		{
			if( IsError("Delete()") ) return;

			var sql = "delete from " + TableName + " where " + this.where;
			
			try
			{
				this.rs.Open(sql, this.conn, adOpenKeyset, adLockOptimistic);
			}
			catch(e)
			{
				ErrorCode = 6;
				if( IsError("Delete() " + e.description) ) return;
			}
		}
		
		// 改
		// Update
		this.Update = function(TableName)
		{
			if( IsError("Update()") ) return;

			var sql = "update " + TableName + " set ";
			for(var i = 0; i < this.kvs.length; i++)
			{
				switch( typeof(this.kvs[i].value) )
				{
					case "number":
					{
						sql += " " + this.kvs[i].key + "=" + this.kvs[i].value + " ";
					}
					break;
					
					case "string":
					{
						sql += " " + this.kvs[i].key + "='" + this.kvs[i].value + "' ";
					}
					break;
				}
				if(i != this.kvs.length - 1) sql += ",";
			}
			sql += " where " + this.where;
						
			try
			{
				this.rs.Open(sql, this.conn, adOpenKeyset, adLockOptimistic);
			}
			catch(e)
			{
				ErrorCode = 7;
				if( IsError("Update() " + e.description) ) return;
			}
		}
		
		// 查
		// Query
		this.Query = function(sql)
		{
			if( IsError("Query()") ) return;

			try
			{
				this.rs.CursorLocation = adUseClient; // 使RecordCount可用 // Make RecordCount available
				this.rs.Open(sql, this.conn, adOpenKeyset, adLockPessimistic);
				this.eof = this.rs.EOF;
			}
			catch(e)
			{
				ErrorCode = 8;
				if( IsError("Query() " + e.description) ) return;
			}
		}
		
		this.GetRecordCount = function()
		{
			var value = 0;
			
			if( IsError("GetRecordCount()") ) return value;
			
			try
			{
				value = parseInt(this.rs.RecordCount);
			}
			catch(e)
			{
				ErrorCode = 9;
				if( IsError("GetRecordCount() " + e.description) ) return value;
			}
			
			return value;
		}
		this.GetColumnCount = function()
		{
			var value = 0;
			
			if( IsError("GetColumnCount()") ) return value;
			
			try
			{
				value = parseInt(this.rs.Fields.Count);
			}
			catch(e)
			{
				ErrorCode = 10;
				if( IsError("GetColumnCount() " + e.description) ) return value;
			}
			
			return value;
		}
		this.GetColumnName = function(index)
		{
			var value = "";
			
			if( IsError("GetRecordCount()") ) return value;
			
			try
			{
				value = this.rs(index).Name;
			}
			catch(e)
			{
				ErrorCode = 11;
				if( IsError("GetColumnName() " + e.description) ) return value;
			}
			
			return value;
		}
		
		this.SetInt = function(key, value)
		{
			this.kvs.push( {"key":key, "value":parseInt(value)} );
		}
		this.SetDouble = function(key, value)
		{
			this.kvs.push( {"key":key, "value":"" + value} );
		}
		this.SetString = function(key, value)
		{
			this.kvs.push( {"key":key, "value":value} );
		}
		this.SetDateTime = function(key, value)
		{
			if(value == "") value = DateTimeToStr();
			
			this.kvs.push( {"key":key, "value":value} );
		}
		
		this.GetInt = function(key)
		{
			var value = 0;
			
			if( IsError("GetInt()") ) return value;
			
			try
			{
				if( !isNaN( "" + this.rs(key) ) ) value = parseInt( "" + this.rs(key) );
			}
			catch(e)
			{
				ErrorCode = 12;
				if( IsError("GetInt() " + e.description) ) return value;
			}
			
			return value;
		}
		this.GetDouble = function(key)
		{
			var value = 0;
			
			if( IsError("GetDouble()") ) return value;
			
			try
			{
				if( !isNaN( "" + this.rs(key) ) ) value = parseFloat( "" + this.rs(key) );
			}
			catch(e)
			{
				ErrorCode = 13;
				if( IsError("GetDouble() " + e.description) ) return value;
			}
			
			return value;
		}
		this.GetString = function(key)
		{
			var value = "";
			
			if( IsError("GetString()") ) return value;
			
			try
			{
				value = "" + this.rs(key);
			}
			catch(e)
			{
				ErrorCode = 14;
				if( IsError("GetString() " + e.description) ) return value;
			}
			
			return value;
		}
		this.GetDateTime = function(key)
		{
			var value = "";
			
			if( IsError("GetDateTime()") ) return value;
			
			switch(this.dbtype)
			{
				case "mysql":
				case "access":
				{
					try
					{
						value = DateTimeToStr( "" + this.rs(key) );
					}
					catch(e)
					{
						ErrorCode = 15;
						if( IsError("GetDateTime() " + e.description) ) return value;
					}
				}
				break;
				
				case "sqlserver":
				{
					try
					{
						var _value = "" + this.rs(key);
						value = _value.substring( 0, _value.lastIndexOf(".") );
					}
					catch(e)
					{
						ErrorCode = 15;
						if( IsError("GetDateTime() " + e.description) ) return value;
					}
				}
				break;
			}
			
			return value;
		}
		
		this.SetWhere = function(where)
		{
			this.where = where;
		}
		
		this.MoveNext = function()
		{
			this.eof = true;
			
			if( IsError("MoveNext()") ) return;

			try
			{
				this.rs.MoveNext();
				this.eof = this.rs.EOF;
			}
			catch(e)
			{
				ErrorCode = 16;
				if( IsError("MoveNext() " + e.description) ) return;
			}
		}
	}
}
</script>