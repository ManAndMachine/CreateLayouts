// (C) Copyright 2021 by  
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

using AcCoreAp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

// This line is not mandatory, but improves loading performances
[assembly: ExtensionApplication(typeof(CreateLayouts.MyPlugin))]

namespace CreateLayouts
{
    // This class is instantiated by AutoCAD once and kept alive for the 
    // duration of the session. If you don't do any one time initialization 
    // then you should remove this class.
    public class MyPlugin : IExtensionApplication
    {

        public void Initialize()
        {
            AcCoreAp.Idle += OnIdle;
        }

        private void OnIdle(object sender, EventArgs e)
        {
            var doc = AcCoreAp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                AcCoreAp.Idle -= OnIdle;
                doc.Editor.WriteMessage("\nCreate Layouts addin loaded.\n");
            }
        }

        public void Terminate()
        { }

    }

}
