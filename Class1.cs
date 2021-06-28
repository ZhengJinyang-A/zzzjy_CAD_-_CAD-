using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using cadSer = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Colors;
using cadWin = Autodesk.AutoCAD.Windows;
using System.IO;
using Autodesk.AutoCAD.Geometry;
using MgdApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Timers;
using System.Threading;
using Autodesk.AutoCAD.ApplicationServices;

[assembly: CommandClass(typeof(zjy.zjyCAD))]
namespace zjy
{
    class zjyCAD
    {
        [CommandMethod("BPZFromText")]
        public void BPZFromText()
        {
           
            Editor ed = cadSer.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            var middoc = cadSer.Application.DocumentManager.MdiActiveDocument;
            List<BlockReference> blockReferencesList = new List<BlockReference>();
            List<DBText> dbTextList = new List<DBText>();
            ed.WriteMessage("=========================注意事项：==============================\n");
            ed.WriteMessage("1、本命令功能为：将高程数据文字Text的内容复制到BlockReferce中的Z值;\n");
            ed.WriteMessage("2、本命令只可以识别出范围为:500单位，在使用过程应自己核对范围。\n");
            ed.WriteMessage("3、当多个高程数据Text与BlockRerferce的位置相接近时,容易造成误判;\n");
            ed.WriteMessage("   在使用过程应仔细核对，尤其要注意判断线交叉位置;\n");
            ed.WriteMessage("4、本命令对未使用的Text以及未找到高程数据的BlockReferce采用了不同颜色的圆标记,\n");
            ed.WriteMessage("   黄色的圆标记未使用的Text,绿色的圆标记BlockReferce;\n");

            try
            {
                #region 获取范围
                //List<ObjectId> objectIds = ed.GetSelection().Value.GetObjectIds().ToList<ObjectId>();
                ObjectId[] objectIdsArr = ed.GetSelection().Value.GetObjectIds();
                ed.WriteMessage("==================================================\n");
                // ed.WriteMessage(objectIds.Count.ToString() + "\n");
                var doc_lock = middoc.LockDocument();
                DBObject dbo;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId tempObjectId in objectIdsArr)
                    {
                        dbo = trans.GetObject(tempObjectId, OpenMode.ForRead);
                        //trans.Commit();

                        if (dbo.GetType().Equals(typeof(BlockReference)))
                        {
                            BlockReference dbo_bR = (BlockReference)dbo;
                            blockReferencesList.Add(dbo_bR);
                        }
                        else if (dbo.GetType().Equals(typeof(DBText)))
                        {
                            DBText dbo_dT = (DBText)dbo;
                            dbTextList.Add(dbo_dT);
                        }
                    }
                }
                ed.WriteMessage("blockReferce的个数：" + blockReferencesList.Count().ToString() + "\n");
                ed.WriteMessage("DBtext的个数：" + dbTextList.Count().ToString() + "\n");
                #endregion
                ed.WriteMessage("=============开始转换============\n");
                ed.WriteMessage("=============开始转换============\n");

                double scopeIni = 20;
                int FindNum = 1;
                //int BR_Not_Text = 1;
                //这样Scope的最大范围20+100*5
                int LimitFindNum = 500;
                while (blockReferencesList.Count() > 0)
                {

                    DBText DBTUsed = null;

                    ed.WriteMessage("====================查找次数====={0}====================\n", FindNum++);
                    ed.WriteMessage("dbTextList个数{0}\n", dbTextList.Count);
                    ed.WriteMessage("blockReferencesList个数{0}\n", blockReferencesList.Count);
                    ed.WriteMessage("scopeIni:{0}\n", scopeIni);
                    //ed.WriteMessage("blockReferce无对应的text个数:{0}\n", BR_Not_Text);

                    List<BlockReference> bRNoUsedList = new List<BlockReference>();

                    foreach (BlockReference tmp_BR in blockReferencesList)
                    {
                        bool BRNotUsed_bool = true;
                        foreach (DBText tmp_dBText in dbTextList)
                        {
                            if (PointsCompare(tmp_BR.Position, tmp_dBText.Position, scopeIni))
                            {
                                using (Transaction trs = db.TransactionManager.StartTransaction())
                                {
                                    //是否存在层，无则新建
                                    //为了放辅助的线
                                    LayerTable lt = (LayerTable)trs.GetObject(db.LayerTableId, OpenMode.ForWrite);
                                    string layerName = "zjy_Temp_Layer";
                                    if (!lt.Has(layerName))
                                    {
                                        LayerTableRecord ltr = new LayerTableRecord();
                                        ltr.Name = layerName;
                                        lt.Add(ltr);
                                        ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 7);
                                        trs.AddNewlyCreatedDBObject(ltr, true);
                                    }

                                    //往数据库写数据
                                    BlockTable bt = (BlockTable)trs.GetObject(db.BlockTableId, OpenMode.ForRead);
                                    BlockTableRecord modelSpace = (BlockTableRecord)trs.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                                    // double positonZ;
                                    //对DBText文本格式进行限定，一个字符，以及转换不成功，以及转化结果为0.0都不可以，
                                    //都进行标注
                                    BlockReference tmp_BR_Change = (BlockReference)trs.GetObject(tmp_BR.ObjectId, OpenMode.ForWrite);
                                    DBText tmp_dBText_Change = (DBText)trs.GetObject(tmp_dBText.ObjectId, OpenMode.ForWrite);
                                    string DbTextString = tmp_dBText.TextString;
                                    try
                                    {
                                        double positonZ = Convert.ToDouble(DbTextString);
                                        //对文本格式进行限定，一个字符，以及转换不成功，以及转化结果为0.0都不可以，
                                        //都进行标注
                                        if (DbTextString.Length > 1 || positonZ != 0.0)
                                        {

                                            // ed.WriteMessage(positonZ.ToString());
                                            tmp_BR_Change.Position = new Point3d(tmp_BR_Change.Position.X, tmp_BR_Change.Position.Y, positonZ);
                                            tmp_BR_Change.Color = Color.FromColorIndex(ColorMethod.ByAci, 5);
                                            tmp_dBText.Color = Color.FromColorIndex(ColorMethod.ByAci, 6);
                                        }
                                        else
                                        {

                                            Entity cir_tmp = (Entity)new Circle(tmp_dBText.Position, new Vector3d(0, 0, 1), 30);
                                            cir_tmp.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                                            modelSpace.AppendEntity(cir_tmp);
                                            trs.AddNewlyCreatedDBObject(cir_tmp, true);
                                            trs.Commit();
                                            continue;

                                        }

                                    }
                                    catch
                                    {
                                        Entity cir_tmp = (Entity)new Circle(tmp_dBText.Position, new Vector3d(0, 0, 1), 30);
                                        cir_tmp.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                                        modelSpace.AppendEntity(cir_tmp);
                                        trs.AddNewlyCreatedDBObject(cir_tmp, true);

                                        ed.WriteMessage("查找到的文字内容非数字为:{0}\n", tmp_dBText.TextString);
                                        trs.Commit();
                                        continue;
                                    }


                                    //在两点间绘制线

                                    Entity line3D = (Entity)new Line(new Point3d(tmp_BR.Position.X, tmp_BR.Position.Y, 0.0), tmp_dBText.Position);
                                    line3D.Layer = layerName;
                                    line3D.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                                    modelSpace.AppendEntity(line3D);
                                    trs.AddNewlyCreatedDBObject(line3D, true);

                                    //tmp_BR_Change.Layer

                                    trs.Commit();
                                }

                                //BRNotUsed_bool标记 BlockReferce已找到相应的DBText
                                //DBTUsed标记 BlockReferce对应的DBText
                                BRNotUsed_bool = false;
                                DBTUsed = tmp_dBText;
                                break;

                            }

                        }


                        if (DBTUsed != null)
                        {
                            dbTextList.Remove(DBTUsed);
                        }

                        if (BRNotUsed_bool)
                        {
                            bRNoUsedList.Add(tmp_BR);
                        }

                    }
                    blockReferencesList = bRNoUsedList;
                    //ed.WriteMessage("bRNoUsedList个数{0}\n", bRNoUsedList.Count);

                    scopeIni += 5;
                    //对超出一定范围的BlockReferce则跳出
                    if (scopeIni > LimitFindNum && blockReferencesList.Count() > 0)
                    {
                        break;
                    }

                }
            }
            catch
            {
                ed.WriteMessage("出错\n");

            }

            //========================================================================================
            //对超出一定范围的BlockReferce进行标记   对未使用的Text进行标记
            ed.WriteMessage("===================BlockReferce与Text剩余个数===================\n");
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //是否存在层，无则新建
                //为了放辅助的线
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForWrite);
                string layerName = "zjy_Temp_Layer";
                if (!lt.Has(layerName))
                {
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    lt.Add(ltr);
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }

                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                //对超出一定范围的BlockReferce进行标记
              //  ed.WriteMessage("对未使用的TBlockReference进行标记\n");
                ed.WriteMessage("剩余blockReferencesList的个数：" + blockReferencesList.Count().ToString() + "\n");
                foreach (BlockReference bfTemp in blockReferencesList)
                {
                    //对未找到BlockReferce相对应的text的进行标注
                    Entity cir_tmp = (Entity)new Circle(bfTemp.Position, new Vector3d(0, 0, 1), 50);
                    cir_tmp.Color = Color.FromColorIndex(ColorMethod.ByAci, 3);
                    cir_tmp.Layer = layerName;
                    modelSpace.AppendEntity(cir_tmp);
                    trans.AddNewlyCreatedDBObject(cir_tmp, true);

                }

                // 对未使用的Text进行标记
             //   ed.WriteMessage("对未使用的Text进行标记\n");
                ed.WriteMessage("剩余dbTextList的个数：" + dbTextList.Count().ToString() + "\n");
                foreach (DBText dBTemp in dbTextList)
                {
                    DBText tmp_dBText_Change = (DBText)trans.GetObject(dBTemp.ObjectId, OpenMode.ForWrite);
                    Entity cir_tmp = (Entity)new Circle(dBTemp.Position, new Vector3d(0, 0, 1), 30);
                    cir_tmp.Color = Color.FromColorIndex(ColorMethod.ByAci, 2);
                    cir_tmp.Layer = layerName;
                    modelSpace.AppendEntity(cir_tmp);
                    trans.AddNewlyCreatedDBObject(cir_tmp, true);
                }

                trans.Commit();
            }

            ed.WriteMessage("=========================完成!!!!!==============================\n");
        }



        bool PointsCompare(Point3d p1, Point3d p2, double scope)
        {
            double distance = Math.Sqrt((p1.X - p2.X) * (p1.X - p2.X) + (p1.Y - p2.Y) * (p1.Y - p2.Y));
            if (distance > scope) return false;

            return true;

        }

        ////创建层
        //public ObjectId LayerAdd(string LayerName,Database db)
        //{
        //    ObjectId layerId = ObjectId.Null;
        //    using (tra)
        //}
    }
}
