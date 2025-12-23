using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms; // 需要在 .csproj 中开启 <UseWindowsForms>true</UseWindowsForms>
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using View = Autodesk.Revit.DB.View;

namespace RevitMezzaninePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class DrawFloorWithWallsAndInput : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<XYZ> points = new List<XYZ>();
            List<ElementId> tempVisualIds = new List<ElementId>();

            try
            {
                // ================= 第一步：绘制轮廓 =================
                TaskDialog.Show("步骤 1/2", "请绘制夹层轮廓：\n1. 点击添加顶点。\n2. 按【ESC】完成绘制并进入参数设置。");

                while (true)
                {
                    XYZ pickedPoint = null;
                    try
                    {
                        string prompt = points.Count == 0 ? "点击起点" : "点击下一个点 (按 ESC 完成)";
                        // 开启最近点捕捉，方便点在墙上
                        pickedPoint = uiDoc.Selection.PickPoint(ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections | ObjectSnapTypes.Nearest, prompt);
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        break;
                    }

                    if (pickedPoint != null)
                    {
                        points.Add(pickedPoint);
                        if (points.Count > 1)
                        {
                            // 实时预览
                            XYZ start = points[points.Count - 2];
                            XYZ end = points[points.Count - 1];
                            end = new XYZ(end.X, end.Y, start.Z); // 强制高度对齐
                            points[points.Count - 1] = end;

                            DrawTempLine(doc, start, end, tempVisualIds);
                            uiDoc.RefreshActiveView();
                        }
                    }
                }

                if (points.Count < 3)
                {
                    DeleteTempGraphics(doc, tempVisualIds);
                    return Result.Cancelled; // 点数不足，直接取消
                }

                // ================= 第二步：参数输入 =================
                // 弹出输入框获取用户想要的夹层高度
                double inputHeightMM = ShowInputDialog("请输入夹层离地高度 (mm):", "2800");
                
                // 如果用户点击取消或输入无效，终止程序
                if (inputHeightMM < 0) 
                {
                    DeleteTempGraphics(doc, tempVisualIds);
                    return Result.Cancelled; 
                }

                // ================= 第三步：环境分析 (标高识别) =================
                // 1. 确定底标高 (Base Level)
                Level baseLevel = GetBaseLevel(doc, points, tempVisualIds);
                if (baseLevel == null)
                {
                    DeleteTempGraphics(doc, tempVisualIds);
                    message = "无法确定当前标高，请在平面视图或有标高的区域操作。";
                    return Result.Failed;
                }

                // 2. 确定顶标高 (Top Level / Next Level) - 用于墙体约束
                Level topLevel = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .Where(l => l.Elevation > baseLevel.Elevation)
                    .OrderBy(l => l.Elevation)
                    .FirstOrDefault();

                // 3. 获取构建类型
                FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().FirstOrDefault(x => !x.IsFoundationSlab);
                WallType wallType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(x => x.Kind == WallKind.Basic);

                if (floorType == null || wallType == null)
                {
                    DeleteTempGraphics(doc, tempVisualIds);
                    message = "项目中缺少基本的楼板或墙类型。";
                    return Result.Failed;
                }

                // ================= 第四步：生成模型 =================
                using (Transaction t = new Transaction(doc, "生成夹层与围墙"))
                {
                    t.Start();
                    
                    // 清理预览线
                    doc.Delete(tempVisualIds);

                    // 1. 构建 CurveLoop
                    CurveLoop profile = new CurveLoop();
                    double zBase = points[0].Z;
                    for (int i = 0; i < points.Count; i++)
                    {
                        XYZ p1 = points[i];
                        XYZ p2 = (i == points.Count - 1) ? points[0] : points[i + 1];
                        XYZ p1Flat = new XYZ(p1.X, p1.Y, zBase);
                        XYZ p2Flat = new XYZ(p2.X, p2.Y, zBase);
                        if (!p1Flat.IsAlmostEqualTo(p2Flat))
                            profile.Append(Line.CreateBound(p1Flat, p2Flat));
                    }

                    if (profile.IsOpen())
                    {
                        message = "轮廓未闭合。";
                        return Result.Failed;
                    }

                    // 2. 创建楼板
                    Floor newFloor = Floor.Create(doc, new List<CurveLoop> { profile }, floorType.Id, baseLevel.Id);
                    
                    // 设置楼板高度偏移
                    double heightOffsetFeet = UnitUtils.ConvertToInternalUnits(inputHeightMM, UnitTypeId.Millimeters);
                    newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).Set(heightOffsetFeet);

                    // 3. 创建周围的墙
                    // 如果找到了顶标高，墙连接到顶标高；否则给一个默认未连接高度 (如 3000mm)
                    double unconnectedHeight = UnitUtils.ConvertToInternalUnits(3000, UnitTypeId.Millimeters);
                    
                    foreach (Curve curve in profile)
                    {
                        // Wall.Create: 路径, 墙类型, 底标高, 高度, 底部偏移, ...
                        Wall newWall = Wall.Create(doc, curve, wallType.Id, baseLevel.Id, unconnectedHeight, heightOffsetFeet, false, false);

                        // 3.1 设置墙底部偏移 (踩在夹层上)
                        newWall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(heightOffsetFeet);

                        // 3.2 设置墙顶部约束
                        if (topLevel != null)
                        {
                            newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(topLevel.Id); // 顶部约束设为下一层
                            newWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(0); // 顶部偏移归零
                        }
                        else
                        {
                            // 如果没有二层标高，这就变成了未连接高度。
                            // 你可以根据需要修改这里的逻辑，比如让墙高 = 总层高 - 夹层高
                        }
                        
                        // 3.3 (可选) 调整墙的定位线为 "面层外部" 
                        // newWall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set((int)WallLocationLine.FinishFaceExterior); 
                    }

                    t.Commit();
                }

                TaskDialog.Show("完成", $"夹层已创建！\n高度: {inputHeightMM}mm\n标高基准: {baseLevel.Name}\n围墙已生成。");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DeleteTempGraphics(doc, tempVisualIds);
                message = ex.Message;
                return Result.Failed;
            }
        }

        // --- 辅助方法 1: 获取输入的简易弹窗 ---
        private double ShowInputDialog(string title, string defaultValue)
        {
            // 使用 Windows Forms 创建一个简单的输入框
            System.Windows.Forms.Form form = new System.Windows.Forms.Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button();
            System.Windows.Forms.Button buttonCancel = new System.Windows.Forms.Button();

            form.Text = title;
            label.Text = title;
            textBox.Text = defaultValue;

            buttonOk.Text = "确定";
            buttonCancel.Text = "取消";
            buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | System.Windows.Forms.AnchorStyles.Right;
            buttonOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
            buttonCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;

            form.ClientSize = new System.Drawing.Size(396, 107);
            form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk, buttonCancel });
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            System.Windows.Forms.DialogResult dialogResult = form.ShowDialog();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                if (double.TryParse(textBox.Text, out double result))
                {
                    return result;
                }
            }
            return -1; // 返回负数表示取消或错误
        }

        // --- 辅助方法 2: 获取基础标高 ---
        private Level GetBaseLevel(Document doc, List<XYZ> points, List<ElementId> tempLines)
        {
            // 策略 A: 尝试通过轮廓寻找相交的墙 (之前的高级逻辑)
            // 这里为了简化代码且保证稳定性，先优先使用视图标高
            // 如果你想用上一条回复的"相交检测"，可以把那个函数 copy 过来

            // 策略 B: 使用当前视图标高
            Level lvl = doc.ActiveView.GenLevel;
            if (lvl != null) return lvl;

            // 策略 C: 3D视图下，取最低标高
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .FirstOrDefault();
        }

        // --- 辅助方法 3: 绘制/清理临时线 ---
        private void DrawTempLine(Document doc, XYZ start, XYZ end, List<ElementId> idList)
        {
            using (Transaction t = new Transaction(doc, "TempLine"))
            {
                t.Start();
                Line line = Line.CreateBound(start, end);
                DirectShape ds = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
                ds.SetShape(new List<GeometryObject> { line });
                idList.Add(ds.Id);
                t.Commit();
            }
        }
        private void DeleteTempGraphics(Document doc, List<ElementId> idList)
        {
            if (idList != null && idList.Count > 0)
            {
                using (Transaction t = new Transaction(doc, "Cleanup"))
                {
                    t.Start();
                    doc.Delete(idList);
                    t.Commit();
                }
            }
        }
    }
}
