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
            string error = "No Errors";
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

            List<string> fns = new List<string>();
            foreach (string seqVar in unique_seq)
            {
                try
                {
                    fns.Add(seqVar.Split('-')[0]);
                }
                catch (Exception)
                {
                    error += "Some String Errors";
                }
            }
            var unique_fns = new HashSet<string>(fns);

            //create folders
            List<FolderItem> newFolders = new List<FolderItem>();
            foreach (string fn in unique_fns)
            {
                //var sn = Guid.NewGuid().ToString();
                // Search s = sLst[0];
                try
                {
                    FolderItem oFolder = new FolderItem();
                    oFolder.DisplayName = fn;
                    newFolders.Add(oFolder);
                    //   se.AddCopy(new FolderItem() { DisplayName = fn });
                }
                catch (Exception)
                {
                    error += "Some Errors";
                }
            }

            List<Search> sLst = new List<Search>();
            var oSets = oDoc.SelectionSets;
            List<string> uniPrefixs = new List<string>();
            List<SelectionSet> oSSets = new List<SelectionSet>();
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
                    SelectionSet oSubSet = new SelectionSet(search);
                    oSubSet.DisplayName = seqVar;
                    oSets.AddCopy(oSubSet);
                    oSSets.Add(oSubSet);
                }
                catch (Exception ex)
                {
                    error = ex.ToString();
                }
            }
            
            //Add sets to folder not used
            string sn = "";
            foreach (FolderItem fi in newFolders)
            {
                if (fi.DisplayName.Contains("PANEL"))
                {
                    oSets.AddCopy(fi);
                    
                    foreach (SelectionSet subSet in oSSets)
                    {
                        if (subSet.DisplayName.Contains(fi.DisplayName))
                        {
                            int groupNdx = oSets.Value.IndexOfDisplayName(subSet.DisplayName);
                            //if (oSets.RootItem.Children.Count - groupNdx > 2)
                            try
                            {
                              //  oSets.Move(groupNdx,groupNdx + 1, fi, 0);
                            }
                            catch (Exception ex)
                            {
                                error = ex.ToString();
                            }
                        }
                       
                    }
                    
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
