# ppt_export.py
import os
import tempfile
import win32com.client
from win32com.client import constants

def _adjust_chart_data_range(chart):
    """
    根据表头（第一行和第一列）的有效性调整图表数据源范围。
    剔除表头为空白或0的列/行（从右侧和底部开始截断）。
    """
    try:
        src = chart.GetSourceData()
        if not src:
            return
        src = src.lstrip('=')
        if '!' in src:
            sheet_part, ref_part = src.split('!', 1)
            if '[' in sheet_part:
                sheet_part = sheet_part.split(']')[-1]
            sheet_name = sheet_part.strip("'")
        else:
            sheet_name = chart.Parent.Parent.Name
            ref_part = src
        ws = chart.Application.Worksheets(sheet_name)
        rng = ws.Range(ref_part)
        total_rows = rng.Rows.Count
        total_cols = rng.Columns.Count
        if total_rows < 2 or total_cols < 2:
            return  

        # ---- 检查列标题（第一行）----
        # 第一列（行标签列）总是保留，从第2列开始检查
        last_valid_col = 1
        for j in range(total_cols, 1, -1):
            cell = rng.Cells(1, j)
            val = cell.Value
            # 判定无效：None、空字符串、数字0、字符串"0"
            if val is None or (isinstance(val, str) and val.strip() == '') or val == 0 or val == '0':
                continue
            else:
                last_valid_col = j
                break

        # ---- 检查行标题（第一列）----
        # 第一行（列标题行）总是保留，从第2行开始检查
        last_valid_row = 1
        for i in range(total_rows, 1, -1):
            cell = rng.Cells(i, 1)
            val = cell.Value
            if val is None or (isinstance(val, str) and val.strip() == '') or val == 0 or val == '0':
                continue
            else:
                last_valid_row = i
                break

        # 如果只有标题行/列，则不调整
        if last_valid_col == 1 and last_valid_row == 1:
            return

        # 构建新的连续区域（从(1,1)到(last_valid_row, last_valid_col)）
        new_start = ws.Cells(1, 1).Address(ReferenceStyle=1)
        new_end = ws.Cells(last_valid_row, last_valid_col).Address(ReferenceStyle=1)
        new_ref = f"={sheet_name}!{new_start}:{new_end}"
        chart.SetSourceData(Source=new_ref)
        print(f"图表数据范围已调整: {new_ref} (行数: {last_valid_row}, 列数: {last_valid_col})")
    except Exception as e:
        print(f"调整图表数据范围时出错: {e}")

def export_charts_to_ppt(excel_path, ppt_path, template_path=None, sheet_names=None, copy_tables=True):
    """
    将 Excel 工作簿中指定工作表的图表复制到 PowerPoint。
    优先尝试复制为可编辑的 OLE 对象（与 Ctrl+C/Ctrl+V 行为一致），
    失败则导出为图片插入。
    若 copy_tables=True，还会复制工作表中的 UsedRange 表格区域。
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

    # ---- 收集需要处理的图表 ----
    charts = []
    target_sheets = []
    if sheet_names is None:
        for sheet in wb.Charts:
            charts.append(('ChartSheet', sheet.Name, sheet))
        for sheet in wb.Worksheets:
            target_sheets.append(sheet)
            for shape in sheet.Shapes:
                if shape.Type == 3:  # xlChart
                    charts.append(('Worksheet', sheet.Name, shape.Chart))
    else:
        for name in sheet_names:
            found = False
            for sheet in wb.Charts:
                if sheet.Name == name:
                    charts.append(('ChartSheet', name, sheet))
                    found = True
                    break
            if not found:
                try:
                    sheet = wb.Worksheets(name)
                    target_sheets.append(sheet)
                    for shape in sheet.Shapes:
                        if shape.Type == 3:
                            charts.append(('Worksheet', name, shape.Chart))
                except Exception:
                    print(f"警告：未找到工作表或图表 {name}")

    print(f"共收集到 {len(charts)} 个图表，{len(target_sheets)} 个工作表用于复制表格")

    # ---- 处理图表 ----
    for source, sheet_name, chart_obj in charts:
        # 若图表嵌入在工作表中，先调整数据范围（根据表头裁剪）
        if source == 'Worksheet':
            _adjust_chart_data_range(chart_obj)

        # 复制图表
        success = False
        for attempt in range(2):
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
                try:
                    chart_obj.ChartArea.Copy()
                except AttributeError:
                    chart_obj.Copy()
                success = True
                break
            except Exception as e:
                print(f"复制尝试 {attempt+1} 失败（{sheet_name}）: {e}")

        if success:
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
            try:
                slide.Shapes.PasteSpecial(DataType=0, Link=False, DisplayAsIcon=False)
                print(f"成功粘贴图表（{sheet_name}）为可编辑对象 (Default)")
                continue
            except Exception as e:
                print(f"PasteSpecial Default 失败，尝试 OLE 对象: {e}")
                try:
                    slide.Shapes.PasteSpecial(DataType=10, Link=False, DisplayAsIcon=False)
                    print(f"成功粘贴图表（{sheet_name}）为 OLE 对象")
                    continue
                except Exception as e2:
                    print(f"OLE 粘贴失败，回退到常规 Paste: {e2}")
                    try:
                        slide.Shapes.Paste()
                        print(f"使用 Paste 粘贴成功（{sheet_name}）")
                        continue
                    except Exception as e3:
                        print(f"Paste 也失败，将尝试导出为图片: {e3}")

        # 后备：导出为图片
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

    # ---- 复制表格区域（UsedRange） ----
    if copy_tables and target_sheets:
        for sheet in target_sheets:
            try:
                sheet.Activate()
                used_range = sheet.UsedRange
                if used_range is None or used_range.Cells.Count == 0:
                    continue
                used_range.Copy()
                slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
                try:
                    slide.Shapes.PasteSpecial(DataType=0, Link=False, DisplayAsIcon=False)
                    print(f"成功粘贴表格区域（{sheet.Name}）")
                except Exception as e:
                    print(f"粘贴表格区域失败（{sheet.Name}），尝试普通粘贴: {e}")
                    try:
                        slide.Shapes.Paste()
                        print(f"普通粘贴表格区域（{sheet.Name}）成功")
                    except Exception as e2:
                        print(f"表格区域粘贴完全失败（{sheet.Name}）: {e2}")
            except Exception as e:
                print(f"处理表格区域（{sheet.Name}）时出错: {e}")

    # 保存并清理
    pres.SaveAs(ppt_path)
    pres.Close()
    ppt_app.Quit()
    wb.Close(SaveChanges=False)
    excel_app.Quit()
    del pres, ppt_app, wb, excel_app
    print(f"PPT 已生成: {ppt_path}")