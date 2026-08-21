# ppt_export.py
import os
import tempfile
import win32com.client
from win32com.client import constants

# ---------------------------------------------------------------------------
# 辅助函数（用于清理空白列）
# ---------------------------------------------------------------------------
def _is_valid_cell(val):
    """判断单元格是否包含有意义的数据（非空、非0、非'-'）"""
    if val is None:
        return False
    if isinstance(val, str):
        stripped = val.strip()
        if stripped == '' or stripped == '-':
            return False
    if isinstance(val, (int, float)) and val == 0:
        return False
    return True

def _is_empty_cell(val):
    """只判断是否为 None 或空字符串（用于数据块分割）"""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == '':
        return True
    return False

def _read_range_values(ws, row_start, col_start, row_end, col_end):
    """高效读取矩形区域的值"""
    if row_start > row_end or col_start > col_end:
        return ()
    try:
        rng = ws.Range(ws.Cells(row_start, col_start), ws.Cells(row_end, col_end))
        vals = rng.Value
    except Exception:
        return ()
    if vals is None:
        return ()
    if not isinstance(vals, tuple):
        return ((vals,),)
    if len(vals) > 0 and not isinstance(vals[0], tuple):
        return tuple((v,) for v in vals)
    return vals

def _find_data_blocks(ws, log_callback=print):
    """按空白行切分数据块，返回 [(row_start, col_start, row_end, col_end), ...]"""
    used = ws.UsedRange
    if used is None:
        return []
    total_rows = used.Rows.Count
    total_cols = used.Columns.Count
    if total_rows < 2 or total_cols < 2:
        return []
    all_vals = _read_range_values(ws, 1, 1, total_rows, total_cols)
    if not all_vals:
        return []
    non_empty_rows = []
    for row_idx in range(total_rows):
        row_has_data = False
        row_vals = all_vals[row_idx] if row_idx < len(all_vals) else ()
        for col_idx in range(len(row_vals)):
            if not _is_empty_cell(row_vals[col_idx]):
                row_has_data = True
                break
        non_empty_rows.append(row_has_data)

    row_groups = []
    in_group = False
    group_start = 0
    for i, has_data in enumerate(non_empty_rows):
        row_num = i + 1
        if has_data and not in_group:
            in_group = True
            group_start = row_num
        elif not has_data and in_group:
            in_group = False
            row_groups.append((group_start, row_num - 1))
    if in_group:
        row_groups.append((group_start, total_rows))

    blocks = []
    for row_start, row_end in row_groups:
        non_empty_cols = set()
        for row_idx in range(row_start - 1, row_end):
            if row_idx >= len(all_vals):
                continue
            row_vals = all_vals[row_idx]
            for col_idx in range(len(row_vals)):
                if not _is_empty_cell(row_vals[col_idx]):
                    non_empty_cols.add(col_idx + 1)
        if non_empty_cols:
            col_start = min(non_empty_cols)
            col_end = max(non_empty_cols)
            blocks.append((row_start, col_start, row_end, col_end))
            log_callback(f"    数据块: 行 {row_start}‑{row_end}, 列 {col_start}‑{col_end}")
    return blocks

# ---------------------------------------------------------------------------
# 清理函数：仅删除空白列（保留行）
# ---------------------------------------------------------------------------
def _trim_block_only_columns(ws, top, left, bottom, right, log_callback=print):
    """只删除空白列，保留所有行"""
    try:
        valid_cols = set()
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                val = ws.Cells(r, c).Value
                if _is_valid_cell(val):
                    valid_cols.add(c)
        if not valid_cols:
            log_callback("    无有效列，跳过")
            return
        min_col = min(valid_cols)
        max_col = max(valid_cols)
        # 从右向左删除右侧多余列
        for c in range(right, max_col, -1):
            ws.Columns(c).Delete()
            log_callback(f"    删除空列：列 {c}")
        # 从左向右删除左侧多余列（注意列号变化）
        for c in range(min_col - 1, left - 1, -1):
            ws.Columns(c).Delete()
            log_callback(f"    删除空列：列 {c}")
        log_callback(f"    清理后保留列 {min_col}‑{max_col}")
    except Exception as e:
        log_callback(f"    清理列时出错: {e}")

def clean_workfile_columns(workbook_path, output_path=None, log_callback=print):
    """
    仅删除 Excel 工作簿中的空白列，保留所有行。
    删除操作基于每个数据块（按空白行分隔）的有效列。
    若 output_path 为 None，则覆盖原文件。
    """
    from win32com.client import Dispatch
    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    try:
        wb = excel_app.Workbooks.Open(os.path.abspath(workbook_path))
    except Exception as e:
        excel_app.Quit()
        raise RuntimeError(f"无法打开工作簿 {workbook_path}: {e}")

    for ws in wb.Worksheets:
        blocks = _find_data_blocks(ws, log_callback=log_callback)
        if not blocks:
            continue
        # 从右下到左上处理，避免列号偏移
        sorted_blocks = sorted(blocks, key=lambda b: (-b[2], -b[3]))
        for block in sorted_blocks:
            log_callback(f"清理工作表 {ws.Name} 数据块: 行 {block[0]}-{block[2]}, 列 {block[1]}-{block[3]}")
            _trim_block_only_columns(ws, block[0], block[1], block[2], block[3], log_callback=log_callback)

    if output_path is None:
        wb.Save()
    else:
        wb.SaveAs(os.path.abspath(output_path))
    wb.Close()
    excel_app.Quit()
    log_callback(f"空白列清理完成，文件保存至: {output_path or workbook_path}")

# ---------------------------------------------------------------------------
# 复制数据表到PPT（辅助）
# ---------------------------------------------------------------------------
def _copy_table_to_slide(pres, ws, block, log_callback=print):
    top, left, bottom, right = block
    try:
        rng = ws.Range(ws.Cells(top, left), ws.Cells(bottom, right))
        rng.Copy()
        slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
        try:
            slide.Shapes.PasteSpecial(DataType=0, Link=False, DisplayAsIcon=False)
            log_callback(f"  数据表粘贴为可编辑对象")
            return True
        except Exception:
            pass
        try:
            slide.Shapes.PasteSpecial(DataType=10, Link=False, DisplayAsIcon=False)
            log_callback(f"  数据表粘贴为 OLE 对象")
            return True
        except Exception:
            pass
        slide.Shapes.Paste()
        log_callback(f"  数据表粘贴为常规对象")
        return True
    except Exception as e:
        log_callback(f"  数据表粘贴失败: {e}")
        return False

# ---------------------------------------------------------------------------
# 主导出函数（简化：不修改 Excel 文件）
# ---------------------------------------------------------------------------
def export_charts_to_ppt(excel_path, ppt_path, template_path=None, sheet_names=None, copy_tables=False):
    """
    将 Excel 文件中的图表和数据表导出到 PowerPoint。
    不修改 Excel 文件，直接复制图表。
    """
    excel_path = os.path.abspath(excel_path)
    ppt_path = os.path.abspath(ppt_path)
    if template_path:
        template_path = os.path.abspath(template_path)

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    try:
        wb = excel_app.Workbooks.Open(excel_path)
    except Exception as e:
        excel_app.Quit()
        raise RuntimeError(f"无法打开 Excel 文件 {excel_path}: {e}")

    # 确定目标工作表
    if sheet_names is None:
        target_sheets = list(wb.Worksheets)
    else:
        target_sheets = []
        for name in sheet_names:
            try:
                sheet = wb.Worksheets(name)
                target_sheets.append(sheet)
            except Exception:
                print(f"警告：未找到工作表 {name}")

    # 收集图表
    charts = []
    for chart_sheet in wb.Charts:
        charts.append(('ChartSheet', chart_sheet.Name, chart_sheet))
    for sheet in target_sheets:
        for shape in sheet.Shapes:
            if shape.Type == 3:  # xlChart
                charts.append(('Worksheet', sheet.Name, shape.Chart))

    print(f"\n共收集到 {len(charts)} 个图表")

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

    # 复制图表到幻灯片
    for source, sheet_name, chart_obj in charts:
        success = False
        for attempt in range(2):
            try:
                excel_app.CutCopyMode = False
                if source == 'Worksheet':
                    try:
                        chart_obj.Parent.Activate()
                    except:
                        pass
                    try:
                        chart_obj.Activate()
                    except:
                        pass
                    try:
                        chart_obj.ChartArea.Copy()
                    except AttributeError:
                        chart_obj.Copy()
                else:
                    chart_obj.Activate()
                    chart_obj.ChartArea.Copy()
                success = True
                break
            except Exception as e:
                print(f"复制尝试 {attempt+1} 失败（{sheet_name}）: {e}")

        if success:
            slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
            pasted = False
            # 尝试多种粘贴方式
            for data_type in [0, 10, 2]:
                try:
                    slide.Shapes.PasteSpecial(DataType=data_type, Link=False, DisplayAsIcon=False)
                    print(f"成功粘贴图表（{sheet_name}）(DataType={data_type})")
                    pasted = True
                    break
                except Exception:
                    continue
            if not pasted:
                try:
                    slide.Shapes.Paste()
                    print(f"使用 Paste 粘贴成功（{sheet_name}）")
                    pasted = True
                except Exception as e:
                    print(f"Paste 失败（{sheet_name}）: {e}")

            if not pasted:
                # 最终后备：导出为图片
                try:
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

    # 可选：复制数据表到PPT（使用 UsedRange 作为区块）
    if copy_tables:
        print("\n复制数据表到PPT...")
        for sheet in target_sheets:
            used = sheet.UsedRange
            if used:
                block = (1, 1, used.Rows.Count, used.Columns.Count)
                _copy_table_to_slide(pres, sheet, block, log_callback=print)

    # 保存并清理
    pres.SaveAs(ppt_path)
    pres.Close()
    ppt_app.Quit()

    wb.Close(SaveChanges=False)
    excel_app.Quit()

    del pres, ppt_app, wb, excel_app
    print(f"\nPPT 已生成: {ppt_path}")