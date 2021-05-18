#region Copyright
//
// Copyright (C) 2010-2015 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
#endregion // Copyright

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;


using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;



namespace Lab_Plugins
{
    #region "Plug-in AddIn Tab"
    [PluginAttribute("NWAPI_HelloWorldPlugin_AddInTab", //Plugin name
                      "ADSK", //Developer ID or GUID
                      ToolTip = "Hello World Plugin in AddIn Tab",
                      //The tooltip for the item in the ribbon
                      DisplayName = "HelloWorld AddInTab")]
    //Display name for the Plugin in the Ribbon

    //  Specify the location of the plug-in
    //[AddInPlugin(AddInLocation.Help)]
    //[AddInPlugin(AddInLocation.CurrentSelectionContextMenu)]
    public class HelloWorld : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            ICollection ic = 
            MessageBox.Show("Hello World in AddIns Tab");
            return 0;
        }

    }

    #endregion

    #region "Plug-in in Context Menu"
    //Plugin name
    [PluginAttribute("NWAPI_HelloWorld_ContextMenu",//Plugin name                      
                     "ADSK",//Developer ID or GUID                     
                            //The tooltip for the item in the ribbon    
                     ToolTip = "Hello World Plugin in AddIn Tab in Context Menu",
                      //Display name for the Plugin in the Ribbon
                      DisplayName = "HelloWorld ContextMenu")]

    [AddInPlugin(AddInLocation.CurrentSelectionContextMenu)]
    public class HelloWorldContextMenu : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            MessageBox.Show("Hello World");
            return 0;
        }

    }

    #endregion

    #region "Dock Pane Plug-in" 
    [Plugin("NWAPI_HelloWorld_DockPane", //Plugin name
            "ADSK", //Developer ID or GUID         
            DisplayName = "HelloWorld DockPane Plugin",
            ToolTip = "HelloWorld DockPane Plugin")]
    [DockPanePlugin(100, 300)]
    public class BasicDockPanePlugin : DockPanePlugin
    {

        public override Control CreateControlPane()
        {
            //create the control that will be used to display in the pane
            HelloWorldControl control = new HelloWorldControl();
            control.Dock = DockStyle.Fill;
            //localisation
            control.Text = this.TryGetString("HelloWorldText");
            //create the control
            control.CreateControl();
            return control;
        }

        public override void DestroyControlPane(Control pane)
        {
            pane.Dispose();
        }
    }
    #endregion 
}
