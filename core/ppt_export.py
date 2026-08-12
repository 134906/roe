# ppt_export.py
import os
import tempfile
import win32com.client
from win32com.client import constants

def export_charts_to_ppt(excel_path, ppt_path, template_path=None, sheet_names=None):
    """
    将 Excel 工作簿中指定工作表的图表复制到 PowerPoint。
    优先尝试复制为可编辑的 OLE 对象（与 Ctrl+C/Ctrl+V 行为一致），
    失败则导出为图片插入。
    """
    excel_path = os.path.abspath(excel_path)
    ppt_path = os.path.abspath(ppt_path)
    if template_path:
        template_path = os.path.abspath(template_path)

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    # 启动 Excel
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    try:
        wb = excel_app.Workbooks.Open(excel_path)
    except Exception as e:
        excel_app.Quit()
        raise RuntimeError(f"无法打开 Excel 文件 {excel_path}: {e}")

    # 启动 PowerPoint
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        ppt_app.Visible = False
    except Exception:
        ppt_app.Visible = True

    if template_path and os.path.exists(template_path):
        pres = ppt_app.Presentations.Open(template_path)
    else:
        pres = ppt_app.Presentations.Add()

    # 收集图表
    charts = []
    if sheet_names is None:
        for sheet in wb.Charts:
            charts.append(('ChartSheet', sheet.Name, sheet))
    else:
        for sheet in wb.Charts:
            if sheet.Name in sheet_names:
                charts.append(('ChartSheet', sheet.Name, sheet))

    for sheet in wb.Worksheets:
        if sheet_names is None or sheet.Name in sheet_names:
            for shape in sheet.Shapes:
                if shape.Type == 3:  # xlChart
                    charts.append(('Worksheet', sheet.Name, shape.Chart))

    print(f"共收集到 {len(charts)} 个图表")

    if not charts:
        print("未找到任何图表，PPT 将不包含图表内容。")
    else:
        for source, sheet_name, chart_obj in charts:
            success = False
            # ---- 尝试复制为 OLE（可编辑） ----
            for attempt in range(2):  # 最多尝试两次
                try:
                    excel_app.CutCopyMode = False
                    if source == 'Worksheet':
                        chart_obj.Parent.Activate()
                        try:
                            chart_obj.Activate()
                        except:
                            pass
                        try:
                            chart_obj.ChartArea.Select()
                        except:
                            pass
                        try:
                            chart_obj.Select()
                        except:
                            pass
                    # 关键修改：使用 ChartArea.Copy() 复制完整对象
                    try:
                        chart_obj.ChartArea.Copy()
                    except AttributeError:
                        # 若为 Chart Sheet，直接 Copy
                        chart_obj.Copy()
                    success = True
                    break
                except Exception as e:
                    print(f"复制尝试 {attempt+1} 失败（{sheet_name}）: {e}")

            if success:
                slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # ppLayoutTitleOnly
                # ---- 使用 ppPasteDefault (0) 模拟默认粘贴行为 ----
                try:
                    # 显式指定 DataType=0, Link=False, DisplayAsIcon=False
                    slide.Shapes.PasteSpecial(DataType=0, Link=False, DisplayAsIcon=False)
                    print(f"成功粘贴图表（{sheet_name}）为可编辑对象 (Default)")
                    continue
                except Exception as e:
                    print(f"PasteSpecial Default 失败，尝试 OLE 对象: {e}")
                    # 后备：显式尝试 OLE 对象 (10)
                    try:
                        slide.Shapes.PasteSpecial(DataType=10, Link=False, DisplayAsIcon=False)
                        print(f"成功粘贴图表（{sheet_name}）为 OLE 对象")
                        continue
                    except Exception as e2:
                        print(f"OLE 粘贴失败，回退到常规 Paste: {e2}")
                        # 最终保底：普通 Paste（可能为图片）
                        try:
                            slide.Shapes.Paste()
                            print(f"使用 Paste 粘贴成功（{sheet_name}）")
                            continue
                        except Exception as e3:
                            print(f"Paste 也失败，将尝试导出为图片: {e3}")

            # ---- 后备方案：导出为图片 ----
            try:
                print(f"尝试导出 {sheet_name} 为图片...")
                tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                tmp_path = tmp.name
                tmp.close()
                chart_obj.Export(Filename=tmp_path, FilterName="PNG")
                slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
                slide.Shapes.AddPicture(FileName=tmp_path, LinkToFile=False, SaveWithDocument=True,
                                        Left=0, Top=0, Width=-1, Height=-1)
                os.unlink(tmp_path)
                print(f"成功插入图片（{sheet_name}）")
            except Exception as e:
                print(f"导出为图片失败（{sheet_name}）: {e}")
                # 若全部失败，跳过该图表

    # 保存并清理
    pres.SaveAs(ppt_path)
    pres.Close()
    ppt_app.Quit()
    wb.Close(SaveChanges=False)
    excel_app.Quit()
    del pres, ppt_app, wb, excel_app
    print(f"PPT 已生成: {ppt_path}")