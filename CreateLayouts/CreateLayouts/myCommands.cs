// (C) Copyright 2021 by Alex Fielder 
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Autodesk.AutoCAD.Colors;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.PlottingServices;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(CreateLayouts.MyCommands))]

namespace CreateLayouts
{
    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl,
 EntryPoint = "?acedSetCurrentVPort@@YA?AW4ErrorStatus@Acad@@PBVAcDbViewport@@@Z")]
        extern static private int acedSetCurrentVPort(IntPtr AcDbVport);

        /// <summary>
        /// Boolean that states whether we inserted new blocks or not.
        /// </summary>
        public bool Blockswereinserted { get; set; }
        /// <summary>
        /// A string denoting Visibility of something which I forget.
        /// </summary>
        public string VisProp;
        //if we fail to new our dictionaries or lists we'll get a fatal error!
        /// <summary>
        /// a list of Jobdetails
        /// </summary>
        public static List<Jobdetails> detailslist = new List<Jobdetails>();
        /// <summary>
        /// a list of Layouts
        /// </summary>
        public static List<Layouts> LayoutList = new List<Layouts>();
        /// <summary>
        /// a list of Vports
        /// </summary>
        public static List<Vports> vport = new List<Vports>();
        /// <summary>
        /// A list for holding existing viewport details.
        /// </summary>
        public static List<Vports> ExistingvportList = new List<Vports>();
        /// <summary>
        /// A Dictionary that holds details of the new blocks we've inserted.
        /// </summary>
        public static Dictionary<int, ObjectId> blockdict = new Dictionary<int, ObjectId>();
        /// <summary>
        /// Contains an Array of available Floor values.
        /// </summary>
        public static ArrayList Floors = new ArrayList();
        /// <summary>
        /// Contains a Dictionary of Strings, ObjectIds relating to the keyplans required for our Tool.
        /// </summary>
        public static Dictionary<string, ObjectId> kpdict = new Dictionary<string, ObjectId>();
        Database db = HostApplicationServices.WorkingDatabase;
        Document dwg = AcadApp.DocumentManager.MdiActiveDocument;
        Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;
        int ISPOSP = 0; // default to zero = ISP.
        String BlockName = "";

        [CommandMethod("CreateLayouts")]
        public void CreateLayouts()
        {
            // Before we change anything, we should save some details about the drawing - don't want to break anything!
            //need to check whether we've already run the tool otherwise we'll get an error when we try to save the layer state.
            int i = 0;
            int floorcount = 0;

            switch (MessageBox.Show("About to delete all layouts, (and Custom UCS entries) are you sure you wish to do this?", "Delete Layouts?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    LayoutCommands.DeleteAllLayouts(false);
                    EverythingIsNew();
                    Document tmpdoc = AcadApp.DocumentManager.MdiActiveDocument;
                    LayerStateManager lyrStMan = tmpdoc.Database.LayerStateManager;
                    Blockswereinserted = false;
                    string tmplyrstName = "test";
                    if (lyrStMan.HasLayerState(tmplyrstName))
                    {
                        lyrStMan.DeleteLayerState(tmplyrstName);
                    }
                    // save the current layer states before we change/add anything.
                    LayerCommands layerclass = new LayerCommands();
                    layerclass.SaveLayerstate();
                    /*instead of deleting our custom ucs entries, I think I need to instead compare their details to those of the blocks 
                     that may or may not have already been inserted.*/
                    //UCSTools.RemoveCustomUCS();
                    break;
                case DialogResult.No:
                    goto theend;
                case DialogResult.Cancel:
                    goto theend;
            }
            SelectJobType:
            switch (MessageBox.Show("Is this an ISP Job?", "ISP?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    ISPOSP = 0;
                    BlockName = "ISPVIEWPORT";
                    break;
                case DialogResult.No:
                    switch (MessageBox.Show("Is this an OSP Job?", "OSP?", MessageBoxButtons.YesNo))
                    {
                        case DialogResult.Yes:
                            ISPOSP = 1;
                            BlockName = "OSPVIEWPORT";
                            break;
                        case DialogResult.No:
                            //Application.Exit;
                            break;
                    }
                    break;
            }
            if (BlockName == "")
            {
                goto SelectJobType;
            }
            Boolean InsertVPBlock = true;
            Boolean FirstTime = true;
            //this will count the number of blocks already in the drawing prior to inserting more.
            Dictionary<int, ObjectId> origblockdict = GetListofInsertedBlocks(BlockName, true);
            do
            {
                switch (MessageBox.Show("Would you like to insert a viewport block?", "Insert block", MessageBoxButtons.YesNo))
                {
                    case DialogResult.Yes:
                        InsertBlockJig BlkClass = new InsertBlockJig();
                        BlkClass.RotationAngleUse = false;
                        /*need to add this here in order to account for already inserted blocks! 
                         but need we also should only need to do this once as after that blockdict
                         will have a complete list of available blocks.*/
                        if (FirstTime)
                        {
                            blockdict = GetListofInsertedBlocks(BlockName, false);
                            FirstTime = false;
                        }
                        if (blockdict.Count > 0)
                        {
                            if (blockdict.Count == 1)
                            {
                                i = 0;
                            }
                            else
                            {
                                i = blockdict.Count;
                            }
                            //zoom to the last block we inserted.
                            ZoomCommands.ZoomToEntity(blockdict[i]);
                            blockdict.Add(blockdict.Count + 1, BlkClass.StartInsert(BlockName, ISPOSP));
                        }
                        else
                        {
                            blockdict.Add(i, BlkClass.StartInsert(BlockName, ISPOSP));
                        }
                        if (blockdict.Count > origblockdict.Count)
                        {
                            Blockswereinserted = true;
                        }
                        else
                        {
                            Blockswereinserted = false;
                        }
                        BlkClass = null;
                        db.UpdateExt(true);
                        //zoom extents
                        ZoomCommands.ZoomWindow(ed, db.Extmin, db.Extmax);
                        i++;
                        break;
                    case DialogResult.No:
                        InsertVPBlock = false;
                        if (BlockName == "")
                        { goto SelectJobType; }
                        else
                        {
                            blockdict = GetListofInsertedBlocks(BlockName, false);
                        }
                        if (blockdict.Count > origblockdict.Count)
                        {
                            Blockswereinserted = true;
                        }
                        else
                        {
                            Blockswereinserted = false;
                        }
                        //ExistingvportList = RetrieveExistingViewportDetails(blockdict);
                        break;
                }
            } while (InsertVPBlock != false);
            switch (MessageBox.Show("Should we continue to the layout creation?", "Create New Layouts?", MessageBoxButtons.YesNo))
            {
                case DialogResult.Yes:
                    //db.UpdateExt(true);
                    ZoomCommands.ZoomWindow(ed, db.Extmin, db.Extmax);
                    /*as we have by this point collected the ObjectIds of the blocks we inserted it should be easy to sort them into groups
                     and assign a keyplan block accordingly. */
                    KeyplanCommands kpclass = new KeyplanCommands();
                    //if we don't new this here, we get multiple keyplans we don't need?!
                    kpdict = new Dictionary<string, ObjectId>();
                    List<Keyplans> keyplans = new List<Keyplans>();
                    foreach (string floor in Floors)
                    {
                        foreach (KeyValuePair<int, ObjectId> kp in blockdict)
                        {
                            Vports keyplan = (from Vports vp in ExistingvportList
                                              where vp.blockid == kp.Value
                                              select vp).First();
                            if (keyplan.FloorName == floor)
                            {
                                Keyplans newkp = new Keyplans();
                                newkp.groupint = floorcount;
                                newkp.oid = keyplan.blockid;
                                keyplans.Add(newkp);
                            }
                        }
                        floorcount++;
                    }

                    kpdict = kpclass.DrawKeyplans(keyplans, floorcount);
                    //kpdict = kpclass.DrawKeyplans();
                    if (blockdict.Count == 0)
                    {
                        MessageBox.Show("Cannot continue, there aren't any blocks inserted!");
                        break;
                    }
                    SetupNewLayoutDetails();
                    BeginCreateLayouts(kpdict);
                    LayoutCommands.DeleteAllLayouts(true);
                    //Sort the layouts
                    LayoutCommands.SortLayouts(db);
                    break;
                case DialogResult.No:
                    break;
            }
            EverythingIsNull();
            db.Dispose();
            ed.Regen();
            theend:
            return;
        }

        private void EverythingIsNew()
        {
            detailslist = new List<Jobdetails>();
            LayoutList = new List<Layouts>();
            vport = new List<Vports>();
            ExistingvportList = new List<Vports>();
            blockdict = new Dictionary<int, ObjectId>();
            Floors = new ArrayList();
            kpdict = new Dictionary<string, ObjectId>();
        }

        private void EverythingIsNull()
        {
            detailslist = null;
            LayoutList = null;
            vport = null;
            ExistingvportList = null;
            blockdict = null;
            Floors = null;
            kpdict = null;
        }

        /// <summary>
        /// Sets up some of the details we need for each sheet of the new layouts.
        /// </summary>
        /// <returns>Returns a list of details about this job/drawing.</returns>
        public List<Jobdetails> SetupNewLayoutDetails()
        {
            PromptResult res = null;
            Jobdetails details = new Jobdetails();
            Boolean SITENumNotDefault = false;
            //SiteNo.
            do
            {
                PromptStringOptions Sitepso = new PromptStringOptions("\nWhat is the Site Number?");
                Sitepso.DefaultValue = "1234";
                res = ed.GetString(Sitepso);
                if (res.Status == PromptStatus.OK)
                {
                    if (res.StringResult != "1234")
                    {
                        details.SiteNo = res.StringResult;
                        SITENumNotDefault = true;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Please select a value other than the default!");
                    }
                }
            } while (SITENumNotDefault != true);
            //BuildingNo.
            Boolean BLDGNumNotDefault = false;
            do
            {
                if (ISPOSP != 1) // this drawing is ISP!
                {
                    PromptStringOptions BLDGpso = new PromptStringOptions("\nWhat is the Building Number?");

                    BLDGpso.DefaultValue = "2345";
                    BLDGpso.AllowSpaces = false;
                    res = ed.GetString(BLDGpso);
                    if (res.Status == PromptStatus.OK)
                    {
                        if (res.StringResult != "2345")
                        {
                            BLDGNumNotDefault = true;
                            details.BuildingNo = res.StringResult;
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Please select a value other than the default!");
                        }
                    }
                }
            } while (BLDGNumNotDefault != true);


            //Date.
            PromptStringOptions datepso = new PromptStringOptions("\nPlease state the date the drawing is to be issued to the client!");
            datepso.DefaultValue = DateTime.Now.ToString("yyyy-MM-dd");

            res = ed.GetString(datepso);
            if (res.Status == PromptStatus.OK)
            {
                // assign a new date value!
                details.Datestr = res.StringResult;
            }

            //Username (Initials)
            PromptStringOptions userpso = new PromptStringOptions("\nPlease input your initials:");
            string userinitials = Environment.UserName;
            string username = userinitials.Substring(0, 2);
            //Need to insert the same filtering we used in the vba adaptation of this code! 2009-01-09 AF
            userpso.DefaultValue = username.ToUpper();

            res = ed.GetString(userpso);
            if (res.Status == PromptStatus.OK)
            {
                // assign a new date value!
                details.Initials = res.StringResult;
            }
            //stores the values gathered above.
            detailslist.Add(details);
            return detailslist;
        }
        /// <summary>
        /// Begins to create new Layouts based on information gathered so far.
        /// </summary>
        /// <param name="kpdict">The dictionary we've compiled containing the keyplans we need for each layout.</param>
        public void BeginCreateLayouts(Dictionary<string, ObjectId> kpdict)
        {
            foreach (KeyValuePair<int, ObjectId> kvpblock in blockdict)
            {
                Vports vportitem = (from Vports n in ExistingvportList
                                    where n.blockid == kvpblock.Value
                                    select n).Single();
                LayoutCommands.CreateNewLayouts(detailslist, kvpblock.Value, LayoutCommands.SetupNewViewports(vportitem, ISPOSP, vportitem.blockid, vportitem.vpNumber, vportitem.UCSName));
                vportitem = null;
            }
        }

        /// <summary>
        /// A function to return the current Layout, you can get properties as name and id from the returned object.
        /// Ex: hzGetCurrentLayout.Name
        /// To check if the current layout is Modelspace, use .modeltype as property (true is Model, false is Paper).
        /// </summary>
        /// <returns>Returns a Layout object</returns>
        public static Layout hzGetCurrentLayout()
        {
            //' Get the current document and database, and start a transaction
            Document acDoc = AcadApp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            LayoutManager acLayoutMgr = default(LayoutManager);
            Layout acLayout = default(Layout);

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                acLayoutMgr = LayoutManager.Current;
                // Get the current layout
                acLayout = (Layout)acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead);
                // Close transaction
                acTrans.Dispose();
            }

            return acLayout;

        }

        /// <summary>
        /// Gets a list of inserted blocks
        /// </summary>
        /// <param name="BlockName">A String relating to the block we're looking for.</param>
        /// <param name="FirstTime">A Boolean denoting whether we've done this before in this instance.</param>
        /// <returns>Returns a Dictionary of integers, ObjectIds relating to BlockName.</returns>
        private Dictionary<int, ObjectId> GetListofInsertedBlocks(string BlockName, bool FirstTime)
        {
            //need to "new" the collections otherwise we get a fatal error!
            Dictionary<int, ObjectId> blockids = new Dictionary<int, ObjectId>();
            ObjectIdCollection count = new ObjectIdCollection();
            Document tmpdoc = AcadApp.DocumentManager.MdiActiveDocument;
            Database db = tmpdoc.Database;
            //if we don't make a new list every time this is run, we end up with duplicate entries.
            ExistingvportList = new List<Vports>();
            int i = 1;
            string UCSName = "";
            Point3d BlockPos = new Point3d(0, 0, 0);
            Scale3d BlockScale = new Scale3d(1);
            Double BlockAngle = 0;
            BlockReference blkRef = null;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
            getblocks:
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead, false) as BlockTable;
                BlockTableRecord btr = null;
                try
                {
                    btr = (BlockTableRecord)tr.GetObject(bt[BlockName], OpenMode.ForRead);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    if (ex.ErrorStatus == ErrorStatus.KeyNotFound)
                    {
                        BlockHelperClass.ImportBlocks();
                        goto getblocks;
                    }
                }
                count = DBUtils.GetBlockReferenceIds(btr.ObjectId);
                foreach (ObjectId blockrefid in count)
                {
                    blockids.Add(i, blockrefid);
                    i++;
                    #region "Work with BlockReferences"
                    // and then extract/store the information we need!
                    Vports Existingvport = new Vports();
                    blkRef = tr.GetObject(blockrefid, OpenMode.ForRead, false) as BlockReference;
                    BlockPos = blkRef.Position;
                    BlockScale = blkRef.ScaleFactors;
                    BlockAngle = blkRef.Rotation;
                    Existingvport.blockid = blockrefid;
                    //Dynamic Block commands:
                    DynamicBlockReferencePropertyCollection DynpropsCol = blkRef.DynamicBlockReferencePropertyCollection;
                    foreach (DynamicBlockReferenceProperty DynpropsRef in DynpropsCol)
                    {
                        string DynPropsName = DynpropsRef.PropertyName.ToUpper();
                        switch (DynPropsName)
                        {
                            case "LOOKUP":
                                Existingvport.PageSize = Convert.ToString(DynpropsRef.Value);
                                Existingvport.vpVisProp = Convert.ToString(DynpropsRef.Value);
                                break;
                        }
                    }
                    #region CollectViewportBlockAtts
                    // fill out the attributes for this block.
                    //int i = 1;
                    AttributeCollection attrefids = blkRef.AttributeCollection;

                    foreach (ObjectId attrefid in attrefids)
                    {
                        AttributeReference attref = tr.GetObject(attrefid, OpenMode.ForWrite, false) as AttributeReference;
                        //vport = new List<Vports>();
                        //vport.Capacity = i + 1;
                        switch (attref.Tag)
                        {
                            case "VIEWNO_NONVIS":
                                Existingvport.vpName = attref.TextString;
                                Existingvport.vpNumber = Convert.ToInt32(attref.TextString);
                                break;
                            case "FULLORPARTIAL":
                                Existingvport.FullorPartial = attref.TextString;
                                break;
                            case "FLOOR":
                                if (!Floors.Contains(attref.TextString))
                                {
                                    Floors.Add(attref.TextString);
                                }
                                Existingvport.FloorName = attref.TextString;
                                break;
                            case "XLOCATION":
                                Existingvport.Xaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                            case "YLOCATION":
                                Existingvport.Yaxis = BlockPos.GetVectorTo(attref.Position);
                                break;
                            case "VPSCALE":
                                Existingvport.vpScale = attref.TextString;
                                break;
                            case "VIEWCENTRE":
                                //Position3d attpos = new Position3d(attref.X,attref.,attref.Z);
                                Existingvport.vpMSCentrePoint = new Point2d(attref.Position.X, attref.Position.Y);
                                break;
                            case "":
                                break;
                        }
                        //get ready to add the next entries.

                        // i++;
                    }
                    //blockdict.Add(i, Existingvport.blockid);
                    //i++;
                    //if (FirstTime)
                    //{
                    UCSName = "UCS_" + Existingvport.vpNumber;
                    Existingvport.UCSName = UCSTools.GetorCreateUCS(UCSName, BlockPos, Existingvport.Xaxis, Existingvport.Yaxis);
                    //}
                    ExistingvportList.Add(Existingvport);
                    Existingvport = null;
                    #endregion //EditViewportBlockAtts
                    #endregion
                }
            }
            return blockids;
        }
    }

}
