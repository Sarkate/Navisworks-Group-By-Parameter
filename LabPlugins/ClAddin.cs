using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using sysForm = System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisPlugin
{
    [PluginAttribute("NavisPlugin", //Plugin name
                     "ADSK", //Developer ID or GUID
                     ToolTip = "Addin tool",
                     //The tooltip for the item in the ribbon
                     DisplayName = "Company Name")]

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

            //create set
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

            //create folders
            List<FolderItem> newFolders = new List<FolderItem>();
            foreach (string fn in unique_fns)
            {
                try
                {
                    FolderItem oFolder = new FolderItem();
                    oFolder.DisplayName = fn;
                    newFolders.Add(oFolder);
                    if (fn.Contains("PANEL"))
                    {
                        oSets.AddCopy(oFolder);
                    }
                }
                catch (Exception)
                {
                    error += "Some Errors";
                }
            }

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
