using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;

namespace Neo4jDriverExcelAddin
{
    using System;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using Neo4j.Driver;

    public partial class ThisAddIn
    {
        private const string defaultConnection = "bolt://localhost:7687/";
        private CustomTaskPane _customTaskPane;
        private IDriver _driver;
        private Boolean _connected;
        private Neo4jDriverExcelAddinRibbon _ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Neo4jDriverExcelAddinRibbon();
            _ribbon.ShowHide += RibbonShowHide;
            return _ribbon;
        }

        private void RibbonShowHide(object sender, EventArgs e)
        {
            bool forceVisible = false;
            if (_customTaskPane == null)
            {
                InitializePane();
                forceVisible = true;
            }

            if (_customTaskPane != null)
                _customTaskPane.Visible = forceVisible || !_customTaskPane.Visible;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ConnectDatabase(defaultConnection);
        }

        private void ConnectDatabase(object sender, ConnectDatabaseArgs e)
        {
            ConnectDatabase(e.ConnectionString);
        }

        private void ConnectDatabase(string connectionString)
        {
            _driver = GraphDatabase.Driver(new Uri(connectionString));
            _connected = true;
            //todo : Add serialization code here.
        }

        private async void CreateNodes(object sender, SelectionArgs e)
        {
            var control = _customTaskPane.Control as ExecuteQuery;
            if (_connected == false)
            {

                ConnectDatabase(this, new ConnectDatabaseArgs { ConnectionString = control.ConnectionString() });
            }
            var session = _driver.AsyncSession();
            try
            {
                var worksheet = ((Worksheet)Application.ActiveSheet);
                var inputrange = e.SelectionRange;




                control.progress.Report(0);

                if (inputrange.Columns.Count <= 1)
                {
                    CurrentControl.SetMessage("Select more than 1 column");
                }

                string[] properties = new string[inputrange.Columns.Count];
                for (int i = 2; i <= inputrange.Columns.Count; i++)
                {
                    try
                    {
                        properties[i - 2] = Convert.ToString(inputrange.Cells[1, i].Value2);
                    }
                    catch
                    {
                        properties[i - 2] = "property" + (i - 1).ToString();
                    }
                }



                for (int r = 2; r <= inputrange.Rows.Count; r++)
                {
                    control.progress.Report(r / inputrange.Rows.Count * 100);
                    var row = inputrange.Rows[r];
                    var label = "";
                    try
                    {
                        label = row.Cells[1, 1].Value2.ToString();
                    }
                    catch
                    {
                        label = "NewExcelNode";
                    }

                    string cypher = "MERGE (a: " + label + " { ";
                    int i = 2;
                    {
                        string propval = Convert.ToString(row.Cells[1, i].Value2);

                        if (properties[i - 2].Length > 0 && propval.Length > 0)
                        {
                            cypher += "`" + properties[i - 2] + "`" + ": \"" + propval + "\",";
                        }
                    }
                    cypher = cypher.TrimEnd(',');
                    cypher += "})";



                    if (row.columns.count > 2)
                    {
                        cypher += " SET a += { ";
                        for (i = 3; i <= row.Columns.Count; i++)
                        {
                            string propval = Convert.ToString(row.Cells[1, i].Value2);

                            if (properties[i - 2] != null && propval != null)
                            {
                                if (properties[i - 2].Length > 0 && propval.Length > 0)
                                {
                                    cypher += "`" + properties[i - 2] + "`" + ": \"" + propval + "\",";
                                }
                            }
                        }
                        cypher = cypher.TrimEnd(',');
                        cypher += "}";
                    }



                    try
                    {
                        IResultCursor cursor = await session.RunAsync(cypher);

                        var records = await cursor.ToListAsync();

                        var summary = await cursor.ConsumeAsync();
                        string message = summary.ToString();
                        if (r == inputrange.Rows.Count)
                        {
                            CurrentControl.SetMessage(message);
                        }
                    }
                    catch (Neo4jException ee)
                    {
                        CurrentControl.SetMessage(ee.Message);
                    }

                }
                await session.CloseAsync();




            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
            }


        }

        private async void ExecuteSelection(object sender, SelectionArgs e)
        {
            var session = _driver.AsyncSession();
            try
            {
                var worksheet = ((Worksheet)Application.ActiveSheet);
                var inputrange = e.SelectionRange;

                if (_connected == false)
                {
                    var control = _customTaskPane.Control as ExecuteQuery;
                    ConnectDatabase(this, new ConnectDatabaseArgs { ConnectionString = control.ConnectionString() });
                }
                /*
                using (session)
                {
                    for (int r = 1; r <= inputrange.Rows.Count; r++)
                    {

                        string cypher = "";

                        foreach (Range col in inputrange.Rows[r].Columns)
                        {
                            try
                            {
                                cypher += col.Cells[1, 1].Value2.ToString();
                            }
                            catch
                            {

                            }
                        }
                        var result = session.Run(cypher);

                        if (r == inputrange.Rows.Count)
                        {
                            CurrentControl.SetMessage(result.Summary.Statement.Text);
                        }

                    }
                }*/
            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
            }
            finally
            {
                await session.CloseAsync();
            }
        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            RemoveOrphanedTaskPanes();
            _driver?.Dispose();
        }

        private void RemoveOrphanedTaskPanes()
        {
            try
            {
                for (var i = CustomTaskPanes.Count; i > 0; i--)
                {
                    var ctp = CustomTaskPanes[i - 1];
                    if (ctp.Window == null)
                    {
                        CustomTaskPanes.Remove(ctp);
                        var control = ctp.Control as ExecuteQuery;
                        control?.Dispose();
                    }
                }
            }
            catch (ObjectDisposedException)
            {
            }
        }

        internal ExecuteQuery InitializePane()
        {
            try
            {
                var gotPane = GetPane();
                if (gotPane != null)
                {
                    _customTaskPane = gotPane;

                    return _customTaskPane.Control as ExecuteQuery;
                }

                var executeQueryControl = new ExecuteQuery();
                executeQueryControl.ExecuteCypher += ExecuteCypher;
                executeQueryControl.ConnectDatabase += ConnectDatabase;
                executeQueryControl.CreateNodes += CreateNodes;
                executeQueryControl.ExecuteSelection += ExecuteSelection;
                executeQueryControl.CreateRelationships += CreateRelationships;
                executeQueryControl.LoadButtonEventHandler += LoadAllNodes;
                executeQueryControl.UpdateButtonEventHandler += UpdateAllNodes;


                _customTaskPane = CustomTaskPanes.Add(executeQueryControl, "Execute Query");

                _customTaskPane.Visible = true;
                return executeQueryControl;
            }
            catch
            {
                return null;
            }
        }

        private void UpdateAllNodes(object sender, SelectionArgs e)
        {
            CurrentControl.SetMessage("Update DB Nodes");

        }

        private void LoadAllNodes(object sender, SelectionArgs e)
        {
            ExecuteLoadAllNodes(sender, e);

            CurrentControl.SetMessage("Load All Nodes");
        }

        private async void ExecuteLoadAllNodes(object sender, SelectionArgs e)
        {
            string cypherGetAllNodes = "CALL db.labels()";
            List<IRecord> records = await ExecuteCypherQuery(cypherGetAllNodes);
            await CreateWorkSheet(records);

        }

        private async Task<List<IRecord>> ExecuteCypherQuery(string queryString)
        {
            var session = _driver.AsyncSession();
            List<IRecord> records = new List<IRecord>();
            try
            {
                if (_connected == false)
                {
                    var control = _customTaskPane.Control as ExecuteQuery;
                    ConnectDatabase(this, new ConnectDatabaseArgs {ConnectionString = control.ConnectionString()});
                }

                var worksheet = ((Worksheet) Application.ActiveSheet);

                try
                {
                    IResultCursor cursor = await session.RunAsync(queryString);
                    records = await cursor.ToListAsync();

                    var summary = await cursor.ConsumeAsync();
                    string summaryText = summary.ToString();

                    CurrentControl.SetMessage("Execution Summary :" + "\n\n" + summaryText);
                    return records;

                }
                finally
                {
                    await session.CloseAsync();
                    
                }
                
            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
            }
            finally
            {
                await session.CloseAsync();
            }

            return records;
        }

        private async Task CreateWorkSheet(List<IRecord> records)
        {
            var labels = GetAllLables(records);
            
            foreach (var wsName in labels)
            {
               var ws = GetOrCreateWorksheet(wsName);
               string getAllNodesOfLabel = $"Match (n:{wsName}) Return n ";
               List<IRecord> nodeRecords = await ExecuteCypherQuery(getAllNodesOfLabel);
               LoadRowsFromRootNode(nodeRecords,ws); 
            }
            
        }

        private  List<string> GetAllLables(List<IRecord> records)
        {
            List<string> labels = new List<string>();
            if (records != null)
            {
                foreach (var r in records)
                {
                    labels.Add(r[0].ToString());
                }
            }
            return labels;
        }

        private Worksheet GetOrCreateWorksheet(string wsName)
        {
            Sheets allSheets = Application.Sheets;
            Application.ActiveWorkbook.Save();
            int count = allSheets.Count;
            
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    var s = (Worksheet)Application.Worksheets[i];

                    Debug.Print(s.Name);
                    if (wsName == s.Name)
                    {
                        s.Cells.Clear();
                        return s;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
                
            }

            return CreateNewWorksheet(wsName);
        }

        private Worksheet CreateNewWorksheet(string wsName)
        {
            var newSheet = (Worksheet) Application.Sheets.Add();
            newSheet.Name = wsName;
            Debug.Print(wsName);
            return newSheet;
        }

        internal ExecuteQuery CurrentControl => _customTaskPane.Control as ExecuteQuery;

        /// <summary>
        /// Gets the appropriate Excel column name given a number index.
        /// </summary>
        /// <remarks>Initial source: http://stackoverflow.com/questions/4583191/incrementation-of-char </remarks>
        private static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        private async void ExecuteCypher(object sender, ExecuteQueryArgs e)
        {
            await ExecuteCypher(e.Cypher);
        }

        private async Task ExecuteCypher(string cypherQuery)
        {
            ExecuteQueryArgs e;
            var session = _driver.AsyncSession();
            try
            {
                if (_connected == false)
                {
                    var control = _customTaskPane.Control as ExecuteQuery;
                    ConnectDatabase(this, new ConnectDatabaseArgs {ConnectionString = control.ConnectionString()});
                }

                var worksheet = ((Worksheet) Application.ActiveSheet);

                try
                {
                    IResultCursor cursor = await session.RunAsync(cypherQuery);
                    var records = await cursor.ToListAsync();
                    LoadRowsFromRecords(records, worksheet);


                    var summary = await cursor.ConsumeAsync();
                    string summaryText = summary.ToString();

                    CurrentControl.SetMessage("Execution Summary :" + "\n\n" + summaryText);
                }
                finally
                {
                    await session.CloseAsync();
                }
            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
            }
            finally
            {
                await session.CloseAsync();
            }
        }

        private static void LoadRowsFromRootNode(List<IRecord> records, Worksheet worksheet)
        {
            bool isFirstRow = true;
            int row = 2;
            Dictionary<string, int> propertiesIndex = new Dictionary<string, int>();
            foreach (var r in records)
            {
                var node = r[0] as INode;
                var properties = node.Properties;
                int i = 0;
                foreach (var k in properties.Keys)
                {
                    isFirstRow = false;
                    if (!propertiesIndex.TryGetValue(k, out i))
                    {
                        i = propertiesIndex.Count;
                        propertiesIndex.Add(k,i);
                        isFirstRow = true;
                    }

                    var colName = GetColNameFromIndex(i + 1);
                    if (isFirstRow)
                        worksheet.Range[$"{colName}1"].Value2 = k;
                    worksheet.Range[$"{colName}{row}"].Value2 = properties[k].ToString();
                    
                }

                row++;
                isFirstRow = false;
            }
        }
        private static void LoadRowsFromRecords(List<IRecord> records, Worksheet worksheet)
        {
            bool isFirstRow = true;
            int row = 2;
            foreach (var r in records)
            {
                for (int i = 0; i < r.Keys.Count; i++)
                {
                    var colName = GetColNameFromIndex(i + 1);
                    var key = r.Keys[i];
                    if (isFirstRow)
                        worksheet.Range[$"{colName}1"].Value2 = key;
                    worksheet.Range[$"{colName}{row}"].Value2 = r.Values[key].As<string>();
                }

                row++;
                isFirstRow = false;
            }
        }

        private async void CreateRelationships(object sender, SelectionArgs e)
        {
            if (_connected == false)
            {
                var control = _customTaskPane.Control as ExecuteQuery;
                ConnectDatabase(this, new ConnectDatabaseArgs { ConnectionString = control.ConnectionString() });
            }
            var session = _driver.AsyncSession();
            try
            {
                var worksheet = ((Worksheet)Application.ActiveSheet);
                var inputrange = e.SelectionRange;




                if (inputrange.Columns.Count <= 1)
                {
                    CurrentControl.SetMessage("Select 3 columns with nodes and relationship to create");
                    await session.CloseAsync();
                    return;
                }
                if (inputrange.Rows.Count <= 2)
                {
                    CurrentControl.SetMessage("Select 2 header rows and 1 data row");
                    await session.CloseAsync();
                    return;
                }

                string label1 = Convert.ToString(inputrange.Cells[1, 1].Value2);
                string label2 = Convert.ToString(inputrange.Cells[1, 2].Value2);
                string relationshiptype = Convert.ToString(inputrange.Cells[1, 3].Value2);

                if (label1.Length == 0 || label2.Length == 0 || relationshiptype.Length == 0)
                {
                    CurrentControl.SetMessage("Labels and relationship type must not be empty");
                    await session.CloseAsync();
                    return;
                }

                string[] properties = new string[inputrange.Columns.Count];
                for (int i = 1; i <= inputrange.Columns.Count; i++)
                {
                    try
                    {
                        properties[i - 1] = Convert.ToString(inputrange.Cells[2, i].Value2);
                    }
                    catch
                    {
                    }
                }

                for (int r = 3; r <= inputrange.Rows.Count; r++)
                {

                    var row = inputrange.Rows[r];
                    var input = inputrange;

                    string cypher = "MATCH (a: {0}),(b: {1} ) WHERE a.`{2}` = '{3}' and b.`{4}` = '{5}' MERGE (a)-[r:`{6}`]->(b) {7}";


                    string relprop = "";
                    if (properties.Length > 2)
                    {
                        relprop = "SET r += { ";
                        bool addcoma = false;
                        for (int p = 2; p < properties.Length; p++ )
                        {
                            
                            string propvalue = Convert.ToString(inputrange.Cells[r, p].Value2);
                            if (properties[p].Length >0 && propvalue.Length>0)
                            {
                                string prop = "`{0}`:\"{1}\"";
                                prop = String.Format(prop, properties[p], propvalue);
                                if (addcoma)
                                {
                                    relprop += " , ";
                                }
                                addcoma = true;
                                relprop += prop;
                            }

                           
                        }
                        relprop += " }";
                        if (relprop.Length == "SET r += { ".Length)
                        {
                            relprop = "";
                        }
                    }
                            

                    string formatedcypher = String.Format(
                    cypher,
                    label1,
                    label2,
                    properties[0],
                    inputrange.Cells[r, 1].Value2.ToString(),
                    properties[1],
                    inputrange.Cells[r, 2].Value2.ToString(),
                    relationshiptype, relprop);

                    try
                    {
                        IResultCursor cursor = await session.RunAsync(formatedcypher);
                        var records = await cursor.ToListAsync();

                        var summary = await cursor.ConsumeAsync();
                        if (r == inputrange.Rows.Count)
                        {
                            string message = summary.ToString();
                            CurrentControl.SetMessage(message);
                        }
                    }
                    catch (Exception ex)
                    {
                        CurrentControl.SetMessage(ex.Message);
                    }
                }

                await session.CloseAsync();



            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
            }



        }

        /// <summary></summary>
        /// <remarks>
        ///     Based on:
        ///     http://svn.alfresco.com/repos/alfresco-open-mirror/alfresco/COMMUNITYTAGS/V4.0d/root/projects/extensions/AlfrescoOffice2007/AlfrescoWord2007/ThisAddIn.cs
        /// </remarks>
        /// <returns></returns>
        private CustomTaskPane GetPane()
        {
            try
            {
                if (CustomTaskPanes.Count > 0)
                {
                    foreach (var ctp in CustomTaskPanes)
                    {
                        try
                        {
                            if (ctp.Window == Application.ActiveWindow)
                            {
                                return ctp;
                            }
                        }
                        catch
                        {
                            // Likely due to no active window
                            if (ctp.Window == null)
                            {
                                // This is the one
                                return ctp;
                            }
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}