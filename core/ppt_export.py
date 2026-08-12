# ppt_export.py
import os
import time
import win32com.client
from win32com.client import constants

def export_charts_to_ppt(excel_path, ppt_path, template_path=None, sheet_names=None):
    """
    将 Excel 工作簿中指定工作表的图表复制为可编辑的 OLE 对象到 PowerPoint。
    - excel_path: 已填充数据的 Workfile (.xlsx)
    - ppt_path: 输出 PPT 文件路径 (.pptx)
    - template_path: 可选，PPT 模板文件路径
    - sheet_names: 可选，要提取图表的工作表名称列表，若为None则提取所有工作表中的图表
    """
    # ---- 转为绝对路径，避免相对路径问题 ----
    excel_path = os.path.abspath(excel_path)
    ppt_path = os.path.abspath(ppt_path)
    if template_path:
        template_path = os.path.abspath(template_path)

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    # 启动 Excel（后台）
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    try:
        wb = excel_app.Workbooks.Open(excel_path)
    except Exception as e:
        excel_app.Quit()
        raise RuntimeError(f"无法打开 Excel 文件 {excel_path}: {e}")

    # 启动 PowerPoint（后台，若不允许隐藏则设为可见）
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        ppt_app.Visible = False
    except Exception:
        ppt_app.Visible = True   # 某些环境不允许隐藏

    if template_path and os.path.exists(template_path):
        pres = ppt_app.Presentations.Open(template_path)
    else:
        pres = ppt_app.Presentations.Add()

    # ---- 收集所有图表 ----
    charts = []

    # 1) 独立图表工作表（Chart Sheets）
    if sheet_names is None:
        for sheet in wb.Charts:
            charts.append(('ChartSheet', sheet.Name, sheet))
    else:
        for sheet in wb.Charts:
            if sheet.Name in sheet_names:
                charts.append(('ChartSheet', sheet.Name, sheet))

    # 2) 嵌入在工作表中的图表对象
    for sheet in wb.Worksheets:
        if sheet_names is None or sheet.Name in sheet_names:
            for shape in sheet.Shapes:
                if shape.Type == 3:  # xlChart
                    charts.append(('Worksheet', sheet.Name, shape.Chart))

    print(f"共收集到 {len(charts)} 个图表")

    if not charts:
        print("在指定工作表中未找到任何图表，PPT 将不包含图表内容。")
    else:
        for source, sheet_name, chart_obj in charts:
            try:
                # ---- 激活图表所在工作表（针对嵌入图表） ----
                if source == 'Worksheet':
                    # 激活工作表并选中图表，确保图表可复制
                    chart_obj.Parent.Activate()
                    chart_obj.Select()
                # ---- 复制图表（可能失败，跳过） ----
                chart_obj.Copy()
            except Exception as e:
                print(f"复制图表失败（源: {source}, 工作表: {sheet_name}）: {e}")
                continue   # 跳过无法复制的图表

            # 添加空白幻灯片
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # 12 = ppLayoutBlank
            # 粘贴为可编辑 OLE 对象
            try:
                slide.Shapes.PasteSpecial(DataType=constants.ppPasteOLEObject)
            except Exception as e:
                print(f"PasteSpecial 失败，使用 Paste: {e}")
                try:
                    slide.Shapes.Paste()
                except Exception as e2:
                    print(f"Paste 也失败，跳过该图表: {e2}")
                    continue

    # 保存 PPT
    pres.SaveAs(ppt_path)
    pres.Close()
    ppt_app.Quit()

    wb.Close(SaveChanges=False)
    excel_app.Quit()

    # 释放 COM 对象
    del pres, ppt_app, wb, excel_app
    print(f"PPT 已生成（图表可编辑）: {ppt_path}")