# ppt_export.py
import os
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
            success = False
            # ---- 多次尝试复制（包括后备方案） ----
            for attempt in range(3):  # 最多尝试3次
                try:
                    # 每次尝试前清空剪贴板，避免干扰
                    excel_app.CutCopyMode = False

                    if source == 'Worksheet':
                        # 激活工作表
                        ws = chart_obj.Parent
                        ws.Activate()
                        # 尝试多种方式激活图表本身（部分方法可能失败，忽略异常）
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

                    # 尝试复制图表
                    chart_obj.Copy()
                    success = True
                    break  # 成功则跳出重试循环
                except Exception as e:
                    print(f"复制尝试 {attempt+1} 失败（源: {source}, 工作表: {sheet_name}）: {e}")
                    # 若第三次仍失败，尝试复制为图片
                    if attempt == 2:
                        try:
                            print("尝试复制为图片（CopyPicture）...")
                            chart_obj.CopyPicture()
                            success = True
                            break
                        except Exception as e2:
                            print(f"复制为图片也失败: {e2}")
                            success = False

            if not success:
                print(f"跳过图表（源: {source}, 工作表: {sheet_name}），无法复制")
                continue

            # 添加空白幻灯片
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # 12 = ppLayoutBlank
            # 粘贴为可编辑 OLE 对象（优先）
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
    print(f"PPT 已生成（图表可编辑或图片）: {ppt_path}")