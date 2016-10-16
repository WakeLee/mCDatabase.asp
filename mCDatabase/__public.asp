<script language="javascript" runat="server" src="__json_m.js"></script>

<script language="javascript" runat="server">
function filetojson(file)
{
	var fs = Server.CreateObject("Scripting.FileSystemObject");
	var f = fs.OpenTextFile( Server.MapPath(file), 1 );
	
	var option = strtojson( f.ReadAll() );
	
	f.Close();
	f = null;
	fs = null;
	
	return option;
}
function strtojson(str)
{
	var JSON = new JSON_M;			

	return JSON.parse(str);
}
function jsontostr(json)
{
	var JSON = new JSON_M;			
	
	return JSON.stringify(json);
}
</script>