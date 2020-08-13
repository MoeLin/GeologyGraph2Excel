using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;

namespace GeologyGraph2Excel
{
    public class GeologyGraph2Excel : IExtensionApplication
    {
        public static string BH_Keyword = "钻孔编号";
        public static string KKGC_Keyword = "孔口高程";
        public static string FCHD_Keyword = "分层厚度";
        public static string DCBH_Keyword = "地层编号";
        public static double PaperWidth = 210;
        public static double PaperHeight = 297;
        public static double Distance_alw = 0.5;
        public static double Distance_Column = 3.5;
        public static bool isA3 = false;
        public static bool isRead_DCBH = true;
        public static bool isMatch_FCHD = true;
        public void Initialize()
        {
            System.Windows.Forms.MessageBox.Show(
                "1.插件启动命令为\"RG\"；\n" +
                "2.适用于地层编号带椭圆、散列文字的情况；\n" +
                "3.当同一钻孔地质柱状图由多页组成时，程序默认把靠右侧的一页的信息认为后续页；\n" +
                "4.使用时，识别的表头和内容均不能在块里面，柱状图的外框必须为多段线（不能是二维多段线，请自行使用convert命令转换）。\n" +
                "5.柱状图需要读取地层编号列的底部，不能有其余无关信息，比如编制人等单行文字，请手工删除。");
        }
        public void Terminate()
        {

        }

        public static List<double> Measure()
        {

            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;//获取当前的活动文档 
            Database acDb = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
            {
                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("");
                // 提示起点
                pPtOpts.Message = "\n输入左上框";

                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                Point3d ptStart = pPtRes.Value;


                // 如果用户按 ESC 键或取消命令，就退出
                if (pPtRes.Status == PromptStatus.Cancel) return null;
                // 提示终点
                pPtOpts.Message = "\n输入右下框";
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = ptStart;
                pPtRes = acDoc.Editor.GetCorner("\n输入右下框", ptStart);
                Point3d ptEnd = pPtRes.Value;
                if (pPtRes.Status == PromptStatus.Cancel) return null;
                List<double> distance = new List<double>();
                distance.Add(Math.Abs(Math.Round(ptStart.X - ptEnd.X, 3)));
                distance.Add(Math.Abs(Math.Round(ptStart.Y - ptEnd.Y, 3)));
                return distance;
            }

        }

        [CommandMethod("RG")]
        public void Read_GeologyGraph()
        {

            Editor acDocEd = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;//获取当前的活动文档 
            Database acDb = Application.DocumentManager.MdiActiveDocument.Database;
            FileInfo outputdwg = new FileInfo(acDoc.Name);
            FileInfo outputcsv = new FileInfo(outputdwg.FullName.Substring(0, outputdwg.FullName.Length - outputdwg.Extension.Length) + ".csv");
            List<string> Exist_ZK = new List<string>();
            List<List<Graph>> Graphs_Lists = new List<List<Graph>>();
            string output = "";
            List<string> err_list = new List<string>();
            using (DocumentLock acLckDoc = acDoc.LockDocument())
            {
                using (Transaction acTrans = acDb.TransactionManager.StartTransaction())
                {
                    PromptPointResult pPtRes;
                    PromptPointOptions pPtOpts = new PromptPointOptions("");
                    pPtOpts.Keywords.Add("Setting");
                    // 提示起点
                    pPtOpts.Message = "\n输入左上框";

                    pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                    Point3d ptStart = pPtRes.Value;

                    while (pPtRes.Status == PromptStatus.Keyword)
                    {
                        new Form_Setting().ShowDialog();
                        pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                        ptStart = pPtRes.Value;
                    }

                    // 如果用户按 ESC 键或取消命令，就退出
                    if (pPtRes.Status == PromptStatus.Cancel) return;
                    // 提示终点
                    pPtOpts.Message = "\n输入右下框";
                    pPtOpts.UseBasePoint = true;
                    pPtOpts.BasePoint = ptStart;
                    pPtRes = acDoc.Editor.GetCorner("\n输入右下框", ptStart);
                    Point3d ptEnd = pPtRes.Value;
                    if (pPtRes.Status == PromptStatus.Cancel) return;


                    TypedValue[] acTypValAr = new TypedValue[1];
                    acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                    // 将过滤器条件赋值给 SelectionFilter 对象
                    SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
                    SelectionSet acSSet;
                    PromptSelectionResult acSSPrompt;
                    acSSPrompt = acDocEd.SelectCrossingWindow(ptStart, ptEnd, acSelFtr);
                    // 如果提示状态 OK，表示已选择对象
                    if (acSSPrompt.Status == PromptStatus.OK)
                    {
                        acSSet = acSSPrompt.Value;
                    }
                    else
                    {
                        return;
                    }


                    List<OutBound_Point> Graph_Form = new List<OutBound_Point>();
                    for (int i = 0; i < acSSet.Count; i++) //每个地质图的矩形框遍历
                    {
                        Entity obj = acTrans.GetObject(acSSet[i].ObjectId, OpenMode.ForRead) as Entity;
                        Point3d pmin = obj.GeometricExtents.MinPoint;
                        Point3d pmax = obj.GeometricExtents.MaxPoint;
                        if (Math.Truncate(pmax.X - pmin.X - PaperWidth) == 0 && Math.Truncate(pmax.Y - pmin.Y - PaperHeight) == 0)
                        {
                            if (isA3 == false)
                            {
                                Graph_Form.Add(new OutBound_Point(pmin, pmax));
                            }
                            else
                            {
                                Graph_Form.Add(new OutBound_Point(pmin, new Point3d((pmin.X + pmax.X) / 2, pmax.Y, pmax.Z)));
                                Graph_Form.Add(new OutBound_Point(new Point3d((pmin.X + pmax.X) / 2, pmin.Y, pmin.Z), pmax));
                            }

                        }
                    }

                    foreach (OutBound_Point obj in OutBound_Point_Distinct(Graph_Form).OrderBy(n => n.Pmin.X))
                    {
                        Point3d pmin = obj.Pmin;
                        Point3d pmax = obj.Pmax;

                        Zoom(pmin, pmax, new Point3d(), 1);


                        TypedValue[] acTypValAr_1 = new TypedValue[4];
                        acTypValAr_1.SetValue(new TypedValue((int)DxfCode.Operator, "<or"), 0);
                        acTypValAr_1.SetValue(new TypedValue((int)DxfCode.Start, "TEXT"), 1);
                        acTypValAr_1.SetValue(new TypedValue((int)DxfCode.Start, "MTEXT"), 2);
                        acTypValAr_1.SetValue(new TypedValue((int)DxfCode.Operator, "or>"), 3);
                        // 将过滤器条件赋值给 SelectionFilter 对象
                        SelectionFilter acSelFtr1 = new SelectionFilter(acTypValAr_1);
                        SelectionSet acSSet1;
                        PromptSelectionResult acSSPrompt1;
                        acSSPrompt1 = acDocEd.SelectCrossingWindow(pmin, pmax, acSelFtr1);
                        // 如果提示状态 OK，表示已选择对象
                        if (acSSPrompt1.Status == PromptStatus.OK)
                        {
                            acSSet1 = acSSPrompt1.Value;
                        }
                        else
                        {
                            continue;
                        }
                        List<DBText> dBTexts = new List<DBText>();
                        List<MText> mTexts = new List<MText>();
                        foreach (ObjectId objectId in acSSet1.GetObjectIds())
                        {

                            DBText sset1_text = acTrans.GetObject(objectId, OpenMode.ForRead) as DBText;
                            if (sset1_text != null)
                            {
                                dBTexts.Add(sset1_text);
                            }
                            MText sset1_mtext = acTrans.GetObject(objectId, OpenMode.ForRead) as MText;
                            if (sset1_mtext != null)
                            {
                                mTexts.Add(sset1_mtext);
                            }
                        }

                        List<Graph> ZK_result = null;
                        try
                        {
                            ZK_result = Text2List(dBTexts, mTexts);
                        }
                        catch (System.Exception e)
                        {
                            if (Convert.ToInt32(e.Source) < 0)
                            {
                                System.Windows.Forms.MessageBox.Show("请检查设置！\n" +
                                    e.Message);
                            }
                            else if (Convert.ToInt32(e.Source) == 1)
                            {
                                err_list.Add(e.Message);
                                continue;
                            }
                        }

                        if (err_list.All(n => n != ZK_result[0].Content))//存在第一页出错，则第二页就不读取,后面在处理如果第一页成功第二页出错情况
                        {
                            if (Exist_ZK.Any(n => n == ZK_result[0].Content)) //存在第二页的情况
                            {
                                int i = Exist_ZK.IndexOf(ZK_result[0].Content);
                                List<Graph> tmp = Graphs_Lists[i];
                                for (int j = 2; j < ZK_result.Count; j++)
                                {
                                    tmp.Add(ZK_result[j]);
                                }
                                Graphs_Lists[i] = tmp;
                            }
                            else
                            {
                                Exist_ZK.Add(ZK_result[0].Content);
                                Graphs_Lists.Add(ZK_result);
                            }
                        }
                    }

                    //====================如果第一页成功第二页出错情况===============
                    foreach (string exist_err in err_list)
                    {
                        if (Exist_ZK.Any(n => n == exist_err))
                        {
                            Exist_ZK.Remove(exist_err);
                            Graphs_Lists = Graphs_Lists.Where(n => n[0].Content != exist_err).ToList();
                        }
                    }
                    //====================如果第一页成功第二页出错情况===============

                    foreach (List<Graph> ZK_result in Graphs_Lists)
                    {
                        foreach (Graph graph in ZK_result)
                        {
                            output = output + graph.Title_name + ",";
                        }
                        output += "\n";
                        foreach (Graph graph in ZK_result)
                        {
                            output = output + graph.Content + ",";
                        }
                        output += "\n";
                    }
                    try
                    {
                        StreamWriter sw = new StreamWriter(outputcsv.FullName, false, Encoding.Default);
                        sw.Write(output);
                        sw.Close();
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message);
                        return;
                    }
                    try
                    {
                        Combine(outputcsv.FullName);
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show("合并结果列表出错，请检查设置是否勾选分层厚度匹配！\n" + e.Message);

                    }

                }
            }
            string tmp_err = "";
            err_list.ForEach(n => tmp_err += n + "、");
            if (tmp_err != "")
            {
                tmp_err = tmp_err.Substring(0, tmp_err.Length - 1);
            }
            System.Windows.Forms.MessageBox.Show("转换完成（共2个文件）：\n"
                + outputcsv.DirectoryName + "\n"
                + " " + outputcsv.Name + "\n"
                + " " + outputcsv.Name.Substring(0, outputcsv.Name.Length - 4) + "_conbine.csv" + "\n\n"
                + "识别图表共" + Exist_ZK.Count + "个。\n"
                + "无法识别图表共" + err_list.Count + "个：(按下ctrl+c复制信息)\n"
                + tmp_err
                );

            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + outputcsv.FullName;
            System.Diagnostics.Process.Start(psi);

        }

        private List<OutBound_Point> OutBound_Point_Distinct(List<OutBound_Point> outBound_Points)
        {
            for (int i = 0; i < outBound_Points.Count; i++)
            {
                for (int j = outBound_Points.Count - 1; j > i; j--)
                {
                    if (outBound_Points[i].Pmin.DistanceTo(outBound_Points[j].Pmin) < Distance_alw && outBound_Points[i].Pmax.DistanceTo(outBound_Points[j].Pmax) < Distance_alw)
                    {
                        outBound_Points.RemoveAt(j);
                    }
                }
            }
            return outBound_Points;
        }
        private List<Graph> Text2List(List<DBText> dBTexts, List<MText> mTexts)
        {
            //===========================获取钻孔编号==========================================================
            string ZKBH_Name;
            try
            {
                ZKBH_Name = Get_Content_By_Title(BH_Keyword, dBTexts, mTexts);
            }
            catch (System.Exception e)
            {
                System.Exception exception = new System.Exception("钻孔编号位置Y误差过小，无法读取钻孔编号！");
                exception.Source = "-1";
                throw exception;
            }
            //===========================获取孔口高程==========================================================
            string KKGC;
            try
            {
                KKGC = Get_Content_By_Title(KKGC_Keyword, dBTexts, mTexts);
            }
            catch (System.Exception e)
            {
                System.Exception exception = new System.Exception("孔口高程位置Y误差过小，无法读取钻孔编号！");
                exception.Source = "-2";
                throw exception;
            }
            //===========================获取分层厚度==========================================================
            List<double> FCHD = new List<double>();
            foreach (IGrouping<double, DBText> group in dBTexts.OrderBy(n => n.Position.X).ThenByDescending(n => n.Position.Y).GroupBy(n => Math.Round(n.Position.X, 0)))
            {
                Debug.WriteLine(Jointext(group.OrderByDescending(n => n.Position.Y).ToList()));
                if (group.Count() > 1)
                {
                    if (Jointext(group.OrderByDescending(n => n.Position.Y).ToList()).Contains(FCHD_Keyword))
                    {
                        FCHD = dBTexts.Where(n => Math.Abs(n.Position.X - group.OrderByDescending(m => m.Position.Y).ToList().First().Position.X) < Distance_Column
                          && n.Position.Y < group.OrderByDescending(m => m.Position.Y).ToList().First().Position.Y)
                            .Where(n => IsNumeric(n.TextString) == true)
                          .OrderByDescending(n => Math.Round(n.Position.Y, 3))
                          .Select(n => Convert.ToDouble(n.TextString)).ToList();
                        break;
                    }

                }

            }
            //===========================获取分层厚度==========================================================

            List<string> DCBH = new List<string>();
            if (isRead_DCBH == true)
            {
                //===========================获取地层编号==========================================================
                foreach (IGrouping<double, DBText> group in dBTexts.OrderBy(n => n.Position.X).ThenByDescending(n => n.Position.Y).GroupBy(n => Math.Round(n.Position.X, 0)))
                {
                    if (group.Count() > 1)
                    {
                        if (Jointext(group.OrderByDescending(n => n.Position.Y).ToList()).Contains(DCBH_Keyword))
                        {
                            var tmp_DCBH_Group = dBTexts.Where(n => Math.Abs(n.Position.X - group.OrderByDescending(m => m.Position.Y).ToList().First().Position.X) < Distance_Column
                              && n.Position.Y < group.OrderByDescending(m => m.Position.Y).ToList().First().Position.Y
                              && DCBH_Keyword.Contains(n.TextString) == false)
                                .OrderByDescending(n => n.Position.Y)
                              .GroupBy(n => Math.Round(n.Height, 3)).OrderByDescending(n => n.Key);
                            if (tmp_DCBH_Group.Count() == 1)//不带椭圆的地层编号
                            {
                                List<DBText> a = tmp_DCBH_Group.First().ToList();
                                if (a.GroupBy(n => Math.Round(n.Position.Y, 3)).Count() == a.Count)
                                {
                                    foreach (DBText dB in a)
                                    {
                                        DCBH.Add(dB.TextString);
                                    }
                                }
                                else//地铁院的，地层编号竟然是一堆等高散的文字放一起
                                {
                                    foreach (IGrouping<double, DBText> g in a.GroupBy(n => Math.Round(n.Position.Y, 3)).OrderByDescending(n => n.Key).ToList())
                                    {
                                        string tmp = "";
                                        foreach (DBText dB in g.OrderBy(n => n.Position.X))
                                        {
                                            tmp = tmp + dB.TextString + "-";
                                        }
                                        DCBH.Add(tmp.Substring(0, tmp.Length - 1));
                                    }
                                }
                            }
                            else if (tmp_DCBH_Group.Count() >= 2)//带椭圆的地层编号
                            {
                                List<List<DBText>> list_num = new List<List<DBText>>();
                                for (int i = 0; i < tmp_DCBH_Group.Count(); i++)
                                {
                                    list_num.Add(tmp_DCBH_Group.ElementAt(i).OrderByDescending(n => n.Position.Y).ToList());
                                }
                                
                                for (int i = 0; i < list_num[0].Count(); i++)
                                {
                                    string tmp_num = "";
                                    tmp_num += list_num[0][i].TextString;
                                    DBText mark = list_num[0][i];
                                    bool find_subscript = false;
                                    for (int j = 1; j < list_num.Count(); j++)
                                    {
                                        for (int k = 0; k < list_num[j].Count(); k++)
                                        {
                                            if(list_num[j][k].Position.Y <= mark.Position.Y + list_num[j][k].Height
                                                && list_num[j][k].Position.Y >= mark.Position.Y - list_num[j][k].Height)
                                            {
                                                tmp_num += "-" + list_num[j][k].TextString;
                                                mark = list_num[j][k];
                                                find_subscript = true;
                                                break;
                                            }
                                        }
                                        if (find_subscript == false)
                                        {
                                            break;
                                        }
                                        find_subscript = false;
                                    }
                                    DCBH.Add(tmp_num);
                                }
                                /*只支持两个元素的字高递减模式
                                List<DBText> a = tmp_DCBH_Group.First().OrderByDescending(n => n.Position.Y).ToList();//可能没有次编号 a的数量永远大于等于b
                                List<DBText> b = tmp_DCBH_Group.Last().OrderByDescending(n => n.Position.Y).ToList();
                                int mark_b = 0;
                                for (int i = 0; i < a.Count; i++)
                                {
                                    if (i != a.Count - 1)
                                    {
                                        if (mark_b < b.Count)
                                        {
                                            if (b[mark_b].Position.DistanceTo(a[i].Position) < a[i].Position.DistanceTo(a[i + 1].Position)) //有后缀
                                            {
                                                DCBH.Add(a[i].TextString + "-" + b[mark_b].TextString);
                                                mark_b += 1;
                                            }
                                            else //无后缀
                                            {
                                                DCBH.Add(a[i].TextString);
                                            }
                                        }
                                        else//后缀用完了
                                        {
                                            DCBH.Add(a[i].TextString);
                                        }
                                    }
                                    else
                                    {
                                        if (mark_b < b.Count)
                                        {
                                            DCBH.Add(a[i].TextString + "-" + b[mark_b].TextString);
                                        }
                                        else//后缀用完了
                                        {
                                            DCBH.Add(a[i].TextString);
                                        }
                                    }
                                }*/
                            }
                            break;
                        }

                    }

                }
                //===========================获取地层编号==========================================================
            }


            //===========================获取岩土名称==========================================================
            List<string> YTMC = new List<string>();
            if (mTexts.Count != 0)//岩土名称是多行文字的情况
            {
                foreach (MText mText in mTexts.OrderByDescending(n => n.Location.Y))
                {
                    Debug.WriteLine(mText.Contents);
                    try
                    {
                        String pattern = @"([^\d;]+?)[:：]";
                        Match m = Regex.Match(mText.Contents, pattern);
                        if (m.Groups[1].Value != "")
                            YTMC.Add(m.Groups[1].Value);
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message + "\n" + mText.Contents);
                        continue;
                    }
                }
            }
            else //岩土名称是单行文字的情况
            {
                foreach (DBText dBText in dBTexts.Where(n => IsYTMC(n.TextString) == true).OrderByDescending(n => n.Position.Y))
                {
                    Debug.WriteLine(dBText.TextString);
                    try
                    {
                        String pattern = @"([^\d]+?)[:：]";
                        Match m = Regex.Match(dBText.TextString, pattern);
                        if (m.Groups[1].Value != "")
                            YTMC.Add(m.Groups[1].Value);
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message + "\n" + dBText.TextString);
                        continue;
                    }
                }
            }
            //===========================获取岩土名称==========================================================
            List<Graph> result = new List<Graph>();
            result.Add(new Graph(BH_Keyword, ZKBH_Name));
            result.Add(new Graph(KKGC_Keyword, KKGC.ToString()));


            if (isMatch_FCHD == true)
            {
                if (DCBH.Count == YTMC.Count && DCBH.Count == FCHD.Count
                && YTMC.Count == FCHD.Count && isRead_DCBH == true)
                {
                    for (int i = 0; i < FCHD.Count; i++)
                    {
                        result.Add(new Graph(DCBH[i] + " " + YTMC[i], FCHD[i].ToString()));
                    }
                }
                else if (YTMC.Count == FCHD.Count && isRead_DCBH == false)
                {
                    for (int i = 0; i < FCHD.Count; i++)
                    {
                        result.Add(new Graph(YTMC[i], FCHD[i].ToString()));
                    }
                }
                else
                {
                    System.Exception exception = new System.Exception(ZKBH_Name);
                    exception.Source = "1";
                    throw exception;
                }
            }
            else
            {
                try
                {
                    if (isRead_DCBH == true)
                    {
                        for (int i = 0; i < FCHD.Count; i++)
                        {
                            result.Add(new Graph(DCBH[i] + " " + YTMC[i], FCHD[i].ToString()));
                        }
                    }
                    else
                    {
                        for (int i = 0; i < FCHD.Count; i++)
                        {
                            result.Add(new Graph(YTMC[i], FCHD[i].ToString()));
                        }
                    }
                }
                catch (System.Exception e)
                {
                    System.Exception exception = new System.Exception(ZKBH_Name);
                    exception.Source = "1";
                    throw exception;
                }
            }
            return result;
        }
        private string Jointext(List<DBText> dBTexts)
        {
            string tmp = "";
            foreach (DBText dBText in dBTexts)
            {
                tmp += dBText.TextString;
            }
            return tmp;
        }
        private string Get_Content_By_Title(string Ref_String, List<DBText> dBTexts, List<MText> mTexts)
        {
            string ZKBH_Name = "";
            foreach (DBText dBText in dBTexts)
            {
                if (dBText.TextString.Contains(Ref_String))
                {
                    Point3d ZK_BH_Ref = dBText.Position;

                    double ZKBH_Distance = PaperWidth;
                    foreach (DBText ZY_BH in dBTexts)
                    {
                        if (ZY_BH.Position.X > ZK_BH_Ref.X && Math.Abs(ZY_BH.Position.Y - ZK_BH_Ref.Y) <= Distance_alw
                            && ZY_BH.Position.X - ZK_BH_Ref.X < ZKBH_Distance)
                        {
                            ZKBH_Name = ZY_BH.TextString;
                            ZKBH_Distance = ZY_BH.Position.X - ZK_BH_Ref.X;
                        }
                    }

                    foreach (MText ZY_BH in mTexts)
                    {
                        if (ZY_BH.Location.X > ZK_BH_Ref.X && Math.Abs(ZY_BH.Location.Y - ZK_BH_Ref.Y) <= Distance_alw
                            && ZY_BH.Location.X - ZK_BH_Ref.X < ZKBH_Distance)
                        {
                            ZKBH_Name = ZY_BH.Contents;
                            ZKBH_Distance = ZY_BH.Location.X - ZK_BH_Ref.X;
                        }
                    }

                    Debug.WriteLine(ZKBH_Name);
                    break;
                }
            }
            return ZKBH_Name;
        }
        private void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            //获取当前文档及数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            int nCurVport =
            System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            // 没提供点或只提供了一个中心点时，获取当前空间的范围
            // 检查当前空间是否为模型空间
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                // 检查当前空间是否为图纸空间
                if (nCurVport == 1)
                {
                    // 获取图纸空间范围
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    // 获取模型空间范围
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }
            // 启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 获取当前视图
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;
                    // 将 WCS 坐标变换为 DCS 坐标
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) *
                    matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                    acView.ViewDirection,
                    acView.Target) * matWCS2DCS;
                    //如果指定了中心点，就为中心模式和比例模式
                    //设置显示范围的最小点和最大点；
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                        pCenter.Y - (acView.Height / 2), 0);
                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                        (acView.Height / 2) + pCenter.Y, 0);
                    }
                    // 用直线创建范围对象；
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                        acLine.Bounds.Value.MaxPoint);
                    }
                    // 计算当前视图的宽高比
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);
                    // 变换视图范围
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);
                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;
                    //检查是否提供了中心点（中心模式和比例模式）
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;
                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }
                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // 窗口、范围和界限模式下
                    {
                        // 计算当前视图的宽高新值；
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;
                        // 获取视图中心点
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X +
                        eExtents.MinPoint.X) * 0.5),
                        ((eExtents.MaxPoint.Y +
                        eExtents.MinPoint.Y) * 0.5));
                    }
                    // 检查宽度新值是否适于当前窗口
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;
                    // 调整视图大小；
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }
                    // 设置视图中心；
                    acView.CenterPoint = pNewCentPt;
                    // 更新当前视图；
                    acDoc.Editor.SetCurrentView(acView);
                }
                // 提交更改；
                acTrans.Commit();
            }
        }

        private static bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }
        private static bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }
        private static bool IsUnsign(string value)
        {
            return Regex.IsMatch(value, @"^\d*[.]?\d*$");
        }

        private static bool isTel(string strInput)
        {
            return Regex.IsMatch(strInput, @"\d{3}-\d{8}|\d{4}-\d{7}");
        }
        private static bool IsYTMC(string value)
        {
            return !(Regex.IsMatch(value, @"\d+?[:：]\d+")) && Regex.IsMatch(value, @"[^\d]+?[:：]");
        }

        private void Combine(string path)
        {
            StreamReader reader = new StreamReader(path, Encoding.Default);
            List<string> ZK_Name = new List<string>();
            List<List<Soil>> ZK_YT = new List<List<Soil>>();
            List<Soil> ZK_Soil = new List<Soil>();
            List<string> Base_Soil_Name = new List<string>();
            int cnt = 1;
            string[] s1 = null;
            string[] s2 = null;
            while (!reader.EndOfStream)
            {
                int result;
                Math.DivRem(cnt, 2, out result);
                string txt = reader.ReadLine();

                if (result != 0)
                {
                    string[] s = new string[1];
                    s[0] = ",";
                    s1 = txt.Split(s, StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {

                    string[] s = new string[1];
                    s[0] = ",";
                    s2 = txt.Split(s, StringSplitOptions.RemoveEmptyEntries);
                    ZK_Name.Add(s2[0]);
                    ZK_Soil = new List<Soil>();
                    for (int i = 1; i < s1.Count(); i++)
                    {
                        ZK_Soil.Add(new Soil(s1[i], s2[i]));
                        if (cnt == 2)
                        {
                            Base_Soil_Name.Add(s1[i]);
                        }
                    }
                    ZK_YT.Add(ZK_Soil);
                    if (cnt > 2)
                    {
                        int mark = 0;
                        for (int i = 0; i < ZK_Soil.Count; i++)
                        {
                            bool found = false;
                            for (int j = mark; j < Base_Soil_Name.Count; j++)
                            {
                                if (ZK_Soil[i].Soild_Name == Base_Soil_Name[j])
                                {
                                    mark = j + 1;
                                    found = true;
                                    break;
                                }
                            }
                            if (found == false)
                            {
                                Base_Soil_Name = Insert_element(mark, ZK_Soil[i].Soild_Name, Base_Soil_Name);
                                mark += 1;
                            }
                        }
                    }


                }
                cnt += 1;
                //Console.WriteLine(reader.ReadLine());  //读取一行数据
            }
            //================output
            string output;
            output = BH_Keyword + ",";
            Base_Soil_Name.ForEach(n => output += n + ",");
            output += "\n";
            for (int i = 0; i < ZK_Name.Count; i++)
            {
                output += ZK_Name[i] + ",";
                int mark = 0;
                for (int j = 0; j < Base_Soil_Name.Count; j++)
                {
                    bool found = false;
                    if (mark < ZK_YT[i].Count)
                    {
                        if (Base_Soil_Name[j] == ZK_YT[i][mark].Soild_Name)
                        {
                            found = true;
                            output += ZK_YT[i][mark].Soild_Thick + ",";
                            mark += 1;
                        }
                    }
                    if (found == false)
                    {
                        output += "0,";
                    }

                }
                output += "\n";
            }
            StreamWriter sw = new StreamWriter(path.Substring(0, path.Length - 4) + "_conbine.csv", false, Encoding.Default);
            sw.Write(output);
            sw.Close();
            //Console.WriteLine(output);
            //Console.ReadKey();
            reader.Close();
        }
        private List<string> Insert_element(int Insert_Place, string Insert_Object, List<string> operate)
        {
            List<string> tmp = new List<string>();
            for (int j = 0; j < operate.Count + 1; j++)
            {
                if (j < Insert_Place)
                {
                    tmp.Add(operate[j]);
                }
                else if (j == Insert_Place)
                {
                    tmp.Add(Insert_Object);
                }
                else if (j > Insert_Place)
                {
                    tmp.Add(operate[j - 1]);
                }
            }
            return tmp;
        }

    }

    class Graph
    {
        private string title_name;
        private string content;
        public Graph(string name, string content)
        {
            Content = content;
            Title_name = name;
        }

        public string Title_name { get => title_name; set => title_name = value; }
        public string Content { get => content; set => content = value; }

        public override string ToString()
        {
            return Title_name + "   " + Content;
        }

    }

    class Soil
    {
        private string soild_name;
        private string soild_thick;
        public Soil(string sn, string st)
        {
            Soild_Thick = st;
            Soild_Name = sn;
        }
        public string Soild_Thick { get => soild_thick; set => soild_thick = value; }
        public string Soild_Name { get => soild_name; set => soild_name = value; }

        public override string ToString()
        {
            return Soild_Name + "   " + Soild_Thick;
        }
    }

    class OutBound_Point
    {
        private Point3d pmin;
        private Point3d pmax;
        public OutBound_Point(Point3d pmin, Point3d pmax)
        {
            Pmin = pmin;
            Pmax = pmax;
        }

        public Point3d Pmin { get => pmin; set => pmin = value; }
        public Point3d Pmax { get => pmax; set => pmax = value; }
        public override string ToString()
        {
            return Pmin.ToString() + "   " + Pmax.ToString();
        }

    }
}
