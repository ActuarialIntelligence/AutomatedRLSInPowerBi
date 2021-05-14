using System;
using System.Collections.Generic;
using Microsoft.AnalysisServices;
using Microsoft.AnalysisServices.Tabular;

namespace CoreRLS.Implementation
{
    class Program
    {
        static void Main(string[] args)
        {
            DynamicRoleCreation("<roleName>", "<Server>", "<databaseId>",
"<filterExpression>");
        }
        private static void DynamicRoleCreation(string roleName,
string serverAddress, string databaseId, string filterExpression)
        {
            string ConnectionString = @"Provider=MSOLAP;Data Source=" + serverAddress +
           @";Initial Catalog=" + databaseId + ";Integrated Security=SSPI;ImpersonationLevel = Impersonate; persist security info = True; ";
            var RoleIDNamePair = new Dictionary<int, string>();
            // Get Roles Where Rolename == 'NewlycreatedRoleName=@Param'
            var GetRolesIDXMLA = @"<Batch xmlns
=""http://schemas.microsoft.com/analysisservices/2003/engine""
Transaction=""true""><Discover xmlns = ""urn:schemas-microsoft-com:xmlanalysis""><RequestType>TMSCHEMA_ROLES</RequestType><Restrictions/><Properties/></Discover
></Batch>";
            // Take Param: @RoleName for use in XMLA to create role
            var xmlaCreateRole = @"<Batch Transaction=""false""
xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><Create
xmlns=""http://schemas.microsoft.com/analysisservices/2014/engine""><DatabaseID>" + databaseId + @"</DatabaseID><Roles><xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""
xmlns:sql=""urn:schemas-microsoft-com:xmlsql""><xs:element><xs:complexType><xs:sequence><xs:element
type=""row""/></xs:sequence></xs:complexType></xs:element><xs:complexType
name=""row""><xs:sequence><xs:element name=""Name"" type=""xs:string"" sql:field=""Name""
minOccurs=""0""/><xs:element name=""Description"" type=""xs:string""
sql:field=""Description"" minOccurs=""0""/><xs:element name=""ModelPermission""
type=""xs:long"" sql:field=""ModelPermission""
minOccurs=""0""/></xs:sequence></xs:complexType></xs:schema><row xmlns=""urn:schemasmicrosoft-com:xmlanalysis:rowset""><Name>" + roleName + @"</Name><ModelPermission>2</ModelPermission></row></Ro
les></Create></Batch>";
            // Get Role ID Where Rolename == @PRoleName: XMLA??
            var objServer = new Microsoft.AnalysisServices.Tabular.Server();
            objServer.Connect(ConnectionString);
            objServer.Execute(xmlaCreateRole);
            var reader = objServer.ExecuteReader(GetRolesIDXMLA, out XmlaResultCollection
           resultsOut, null, true);
            // Add all roles to Cached List
            while (reader.Read())
            {
                RoleIDNamePair.Add(int.Parse(reader[0].ToString()), reader[2].ToString());
            }
            reader.Close();
            reader.Dispose();
            var RoleIDByRoleName = "";//= RoleIDNamePair.Single(s => s.Value == roleName).Key.ToString();
            foreach (var pair in RoleIDNamePair)
            {
                if (pair.Value == roleName)
                {
                    RoleIDByRoleName = pair.Key.ToString();
                    break;
                }
            }
            var AddFilterTorole = @"<Batch Transaction=""false""
xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine""><Create
xmlns=""http://schemas.microsoft.com/analysisservices/2014/engine""><DatabaseID>" + databaseId + @"</DatabaseID><TablePermissions><xs:schema
xmlns:xs=""http://www.w3.org/2001/XMLSchema"" xmlns:sql=""urn:schemas-microsoft-com:xmlsql""><xs:element><xs:complexType><xs:sequence><xs:element
type=""row""/></xs:sequence></xs:complexType></xs:element><xs:complexType
name=""row""><xs:sequence><xs:element name=""RoleID"" type=""xs:unsignedLong""
sql:field=""RoleID"" minOccurs=""0""/><xs:element name=""RoleID.Role"" type=""xs:string""
sql:field=""RoleID.Role"" minOccurs=""0""/><xs:element name=""TableID""
type=""xs:unsignedLong"" sql:field=""TableID"" minOccurs=""0""/><xs:element
name=""TableID.Table"" type=""xs:string"" sql:field=""TableID.Table""
minOccurs=""0""/><xs:element name=""FilterExpression"" type=""xs:string""
sql:field=""FilterExpression"" minOccurs=""0""/><xs:element name=""MetadataPermission""
type=""xs:long"" sql:field=""MetadataPermission""
minOccurs=""0""/></xs:sequence></xs:complexType></xs:schema><row xmlns=""urn:schemasmicrosoft-com:xml-analysis:rowset""><RoleID>" + RoleIDByRoleName +
       @"</RoleID><TableID>364</TableID><FilterExpression>" + filterExpression + @"</FilterExpression
></row></TablePermissions></Create><SequencePoint
xmlns=""http://schemas.microsoft.com/analysisservices/2014/engine""><DatabaseID>" +
       databaseId + @"</DatabaseID></SequencePoint></Batch>";
            var applyRoleQuery = AddFilterTorole;
            objServer.Execute(applyRoleQuery);
        }

        private static XmlaResultCollection DynamicXMLAParser(
string serverAddress, string database, string XMLAQuery, bool outputToCsv)
        {
            string ConnectionString = @"Provider=MSOLAP;Data Source=" + serverAddress +
           @";Initial Catalog=" + database + ";Integrated Security=SSPI;ImpersonationLevel = Impersonate; persist security info = True; ";

            var XMLA = XMLAQuery;
            var objServer = new Microsoft.AnalysisServices.Tabular.Server();
            objServer.Connect(ConnectionString);
            var reader = objServer.ExecuteReader(XMLA, out XmlaResultCollection
           resultsOut, null, true);
            // Add all roles to Cached List
            //while (reader.Read())
            //{

            //}
            reader.Close();
            reader.Dispose();
            objServer.Dispose();
            return resultsOut;
        }
    }
}
