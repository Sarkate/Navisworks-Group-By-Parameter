using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using sysForm =  System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisPlugin
{
    [PluginAttribute("NavisPlugin", //Plugin name
                     "ADSK", //Developer ID or GUID
                     ToolTip = "Rad Urban Addin tool",
                     //The tooltip for the item in the ribbon
                     DisplayName = "RAD Urban")]


    public class ClAddin: AddInPlugin
    {
          public override int Execute(params string[] parameters)
            {
            Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            //ModelItem rootItem = oDoc.Models[0].RootItem;
 
            //select all visiable items
            IEnumerable<ModelItem> allItems = oDoc.Models.First.RootItem.Descendants;//.Where(x => !x.IsHidden);
            //oDoc.CurrentSelection.CopyFrom(allItems);
            int ic = allItems.Count();

            //get all sequence names
            List<string> seqVars = new List<string>();
            foreach (ModelItem oItem in allItems)
            {
               foreach (PropertyCategory oPc in oItem.PropertyCategories)
                {
                    //oPcNames.Add(oPc.DisplayName);
                    if(oPc.DisplayName == "Element")
                    {
                        foreach (DataProperty oDP in oPc.Properties)
                        {
                            if (oDP.Value.IsDisplayString)
                            {   
                               if(oDP.DisplayName == "Erection Sequence")
                                {
                                    seqVars.Add(oDP.Value.ToDisplayString());
                                }
                            }
                        }
                    }

                }
            }

            var unique_seq = new HashSet<string>(seqVars);

            List<Search> sLst = new List<Search>();
            string error = "No Errors";
            var ses = oDoc.SelectionSets;
            foreach (string seqVar in unique_seq)
            {
                Search search= new Search();
                search.Selection.SelectAll();
                search.SearchConditions.Add(SearchCondition.HasPropertyByDisplayName("Element", "Erection Sequence").EqualValue(VariantData.FromDisplayString(seqVar)));
                ModelItemCollection setItems = search.FindAll(oDoc, false);
                //highlight
                oDoc.CurrentSelection.CopyFrom(setItems);
                sLst.Add(search);

                try
                {
                    SelectionSet oMySet = new SelectionSet(search);
                    oMySet.DisplayName = seqVar;
                    ses.AddCopy(oMySet);
                }
                catch
                {
                    error = "Some errors";
                }
            }


            #region create folder for selection set (not used)
            //var cs = oDoc.CurrentSelection;
            //var se = oDoc.SelectionSets;

            //var fn = "S_Set";
            //var sn = Guid.NewGuid().ToString();
            //Search s = sLst[0];
            //try
            //{
            //    var fi = se.Value.IndexOfDisplayName(fn);
            //    if (fi == -1)
            //    {
            //        se.AddCopy(new FolderItem() { DisplayName = fn });
            //    }
            //    var set = new SelectionSet(s) { DisplayName = sn };
            //    se.AddCopy(set);

            //    var fo = se.Value[se.Value.IndexOfDisplayName(fn)] as FolderItem;
            //    var ns = se.Value[se.Value.IndexOfDisplayName(sn)] as SavedItem;

            //    se.Move(ns.Parent, se.Value.IndexOfDisplayName(sn), fo, 0);

            //    //SelectionSet oMySet = new SelectionSet(s);
            //    //oMySet.DisplayName = seqVar;
            //    //se.AddCopy(oMySet);

            //}
            //catch (Exception)
            //{
            //    error += "Some Errors";
            //}
            #endregion

            string info = "";
            foreach (string sName in unique_seq)
            {
                info += sName + ",";
            }
                sysForm.MessageBox.Show(error, "Info");
            return 0;
            }
     }
}
