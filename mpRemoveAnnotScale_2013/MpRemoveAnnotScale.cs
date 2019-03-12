﻿using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using ModPlusAPI;
using ModPlusAPI.Windows;


namespace mpRemoveAnnotScale
{
    public class MpRemoveAnnotScale
    {
        private const string LangItem = "mpRemoveAnnotScale";
        [CommandMethod("ModPlus", "mpRemoveAnnotScale", CommandFlags.UsePickSet)]
        public static void Main()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());

            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var pso = new PromptSelectionOptions
                    {
                        MessageForAdding = "\n" + Language.GetItem(LangItem, "msg1")
                    };
                    var psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK) return;

                    var cm = db.ObjectContextManager;
                    var occ = cm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                    var currentscale = occ.CurrentContext as AnnotationScale;

                    foreach (var objId in psr.Value.GetObjectIds())
                    {
                        using (var ent = tr.GetObject(objId, OpenMode.ForWrite, false, false))
                        {
                            if (ent is MLeader)
                            {
                                var ml = ent as MLeader;
                                if (ml.Annotative == AnnotativeStates.True)
                                {
                                    AddScale(ml, currentscale, occ);
                                    if (ml.ContentType == ContentType.MTextContent)
                                    {
                                        MText mt = ml.MText;
                                        if (mt.Annotative == AnnotativeStates.True)
                                        {
                                            AddScale(mt, currentscale, occ);
                                            ml.MText = mt;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (ent.Annotative == AnnotativeStates.True)
                                    AddScale(ent, currentscale, occ);
                            }
                        }
                    }
                    tr.Commit();
                }
            }
            catch (Exception ex)
            {
                ExceptionBox.Show(ex);
            }
        }
        private static void AddScale(DBObject ent, ObjectContext curScale, ObjectContextCollection occ)
        {
            ent.AddContext(curScale);
            foreach (var scale in occ)
            {
                if (scale.Name != curScale.Name && ent.HasContext(scale))
                    ent.RemoveContext(scale);
            }
        }
    }
}