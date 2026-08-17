# ppt_export.py
import os
import tempfile
import win32com.client
from win32com.client import constants


# ---------------------------------------------------------------------------
# 辅助函数
# ---------------------------------------------------------------------------

def _is_empty_or_zero(val):
    """判断单元格值是否为空、空白字符串或零。"""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == '':
        return True
    if val == 0:
        return True
    return False


def _read_range_values(ws, row_start, col_start, row_end, col_end):
    """高效读取工作表一个矩形区域的值，返回二维元组。"""
    if row_start > row_end or col_start > col_end:
        return ()
    try:
        rng = ws.Range(ws.Cells(row_start, col_start), ws.Cells(row_end, col_end))
        vals = rng.Value
    except Exception:
        return ()
    if vals is None:
        return ()
    # 单单元格 -> ((val,),)
    if not isinstance(vals, tuple):
        return ((vals,),)
    # 单行 -> ((v1, v2, ...),)  需要包装成 ((v1,), (v2,), ...) 吗？
    # 实际上单行时 vals 是 (v1, v2, ...)，每个元素不是 tuple
    if len(vals) > 0 and not isinstance(vals[0], tuple):
        return tuple((v,) for v in vals)
    return vals


# ---------------------------------------------------------------------------
# 数据表检测：按空白行分隔，找出每个独立的数据块
# ---------------------------------------------------------------------------

def _find_data_blocks(ws, log_callback=print):
    """
    扫描工作表，找出所有独立的数据块（按空白行分隔）。
    返回 [(top_row, left_col, bottom_row, right_col), ...] 列表。
    """
    used = ws.UsedRange
    if used is None:
        return []

    total_rows = used.Rows.Count
    total_cols = used.Columns.Count

    if total_rows < 2 or total_cols < 2:
        return []

    # 一次性读取整个 UsedRange 的值
    all_vals = _read_range_values(ws, 1, 1, total_rows, total_cols)
    if not all_vals:
        return []

    # 找出哪些行有非空单元格
    non_empty_rows = []
    for row_idx in range(total_rows):
        row_has_data = False
        row_vals = all_vals[row_idx] if row_idx < len(all_vals) else ()
        for col_idx in range(len(row_vals)):
            if not _is_empty_or_zero(row_vals[col_idx]):
                row_has_data = True
                break
        non_empty_rows.append(row_has_data)

    # 将连续的非空行归为一组 -> 一个数据块
    row_groups = []
    in_group = False
    group_start = 0
    for i, has_data in enumerate(non_empty_rows):
        row_num = i + 1  # 1‑based
        if has_data and not in_group:
            in_group = True
            group_start = row_num
        elif not has_data and in_group:
            in_group = False
            row_groups.append((group_start, row_num - 1))
    if in_group:
        row_groups.append((group_start, total_rows))

    log_callback(f"  检测到 {len(row_groups)} 个数据区域")

    # 对每个行组，确定列范围
    blocks = []
    for row_start, row_end in row_groups:
        non_empty_cols = set()
        for row_idx in range(row_start - 1, row_end):  # 转 0‑based
            if row_idx >= len(all_vals):
                continue
            row_vals = all_vals[row_idx]
            for col_idx in range(len(row_vals)):
                if not _is_empty_or_zero(row_vals[col_idx]):
                    non_empty_cols.add(col_idx + 1)  # 1‑based

        if non_empty_cols:
            col_start = min(non_empty_cols)
            col_end = max(non_empty_cols)
            blocks.append((row_start, col_start, row_end, col_end))
            log_callback(f"    数据块: 行 {row_start}‑{row_end}, 列 {col_start}‑{col_end}")

    return blocks


# ---------------------------------------------------------------------------
# 裁剪单个数据块：从右向左删除全空列，从下向上删除全空行
# ---------------------------------------------------------------------------

def _trim_block(ws, top, left, bottom, right, log_callback=print):
    """
    裁剪单个数据块中尾部全空/全零的行和列。
    循环执行直到没有可删除项。
    """
    try:
        deleted = True
        while deleted:
            deleted = False

            # ---- 从右向左扫描列 ----
            last_nonempty_col = 0
            for col_idx in range(right, left - 1, -1):
                col_has_data = False
                for row_idx in range(top, bottom + 1):
                    cell = ws.Cells(row_idx, col_idx)
                    val = cell.Value
                    if not _is_empty_or_zero(val):
                        col_has_data = True
                        break
                if col_has_data:
                    last_nonempty_col = col_idx
                    break

            if last_nonempty_col > 0 and last_nonempty_col < right:
                for col_idx in range(right, last_nonempty_col, -1):
                    try:
                        ws.Columns(col_idx).Delete()
                        log_callback(f"    删除空列：列 {col_idx}")
                        deleted = True
                        right -= 1
                    except Exception as e:
                        if "array" in str(e).lower():
                            log_callback(f"    列 {col_idx} 包含数组公式，尝试清除内容后删除")
                            try:
                                ws.Columns(col_idx).ClearContents()
                                ws.Columns(col_idx).Delete()
                                log_callback(f"    删除空列（清除后）：列 {col_idx}")
                                deleted = True
                                right -= 1
                            except Exception as e2:
                                log_callback(f"    仍无法删除列 {col_idx}: {e2}")
                        else:
                            log_callback(f"    删除列 {col_idx} 失败: {e}")
                if deleted:
                    continue

            # ---- 从下向上扫描行 ----
            last_nonempty_row = 0
            for row_idx in range(bottom, top - 1, -1):
                row_has_data = False
                for col_idx in range(left, right + 1):
                    cell = ws.Cells(row_idx, col_idx)
                    val = cell.Value
                    if not _is_empty_or_zero(val):
                        row_has_data = True
                        break
                if row_has_data:
                    last_nonempty_row = row_idx
                    break

            if last_nonempty_row > 0 and last_nonempty_row < bottom:
                for row_idx in range(bottom, last_nonempty_row, -1):
                    try:
                        ws.Rows(row_idx).Delete()
                        log_callback(f"    删除空行：行 {row_idx}")
                        deleted = True
                        bottom -= 1
                    except Exception as e:
                        if "array" in str(e).lower():
                            log_callback(f"    行 {row_idx} 包含数组公式，尝试清除内容后删除")
                            try:
                                ws.Rows(row_idx).ClearContents()
                                ws.Rows(row_idx).Delete()
                                log_callback(f"    删除空行（清除后）：行 {row_idx}")
                                deleted = True
                                bottom -= 1
                            except Exception as e2:
                                log_callback(f"    仍无法删除行 {row_idx}: {e2}")
                        else:
                            log_callback(f"    删除行 {row_idx} 失败: {e}")
                if deleted:
                    continue

            if not deleted:
                break

        log_callback(f"    数据块 [{top},{left}]-[{bottom},{right}] 裁剪完成")
    except Exception as e:
        log_callback(f"    裁剪数据块时出错: {e}")


# ---------------------------------------------------------------------------
# 图表坐标轴：保存/恢复设置，防止删除行/列后坐标轴变形
# ---------------------------------------------------------------------------

def _save_chart_axes(wb, log_callback=print):
    """
    保存工作簿中所有图表的坐标轴设置，返回 {key: [axis_info, ...]} 字典。
    修复了原代码中"纵轴最大值变成 1000%"的问题。
    """
    saved = {}
    # 嵌入图表
    for ws in wb.Worksheets:
        for shape in ws.Shapes:
            if shape.Type == 3:  # xlChart
                try:
                    chart = shape.Chart
                    axes_info = _get_chart_axes_info(chart)
                    key = f"ws_{ws.Name}_chart_{shape.Name}"
                    saved[key] = axes_info
                except Exception as e:
                    log_callback(f"  保存图表轴设置失败 ({ws.Name}/{shape.Name}): {e}")
    # 独立图表页
    for chart_sheet in wb.Charts:
        try:
            axes_info = _get_chart_axes_info(chart_sheet)
            key = f"chart_sheet_{chart_sheet.Name}"
            saved[key] = axes_info
        except Exception as e:
            log_callback(f"  保存图表轴设置失败 ({chart_sheet.Name}): {e}")

    log_callback(f"  已保存 {len(saved)} 个图表的轴设置")
    return saved


def _get_chart_axes_info(chart):
    """提取图表所有坐标轴的设置。"""
    axes_info = []
    try:
        for ax in chart.Axes():
            try:
                info = {
                    'minimum_scale': ax.MinimumScale,
                    'minimum_scale_is_auto': ax.MinimumScaleIsAuto,
                    'maximum_scale': ax.MaximumScale,
                    'maximum_scale_is_auto': ax.MaximumScaleIsAuto,
                    'major_unit': ax.MajorUnit,
                    'major_unit_is_auto': ax.MajorUnitIsAuto,
                    'minor_unit': ax.MinorUnit,
                    'minor_unit_is_auto': ax.MinorUnitIsAuto,
                }
                axes_info.append(info)
            except Exception:
                pass
    except Exception:
        pass
    return axes_info


def _restore_chart_axes(wb, saved_axes, log_callback=print):
    """恢复之前保存的图表坐标轴设置。"""
    restored = 0
    for ws in wb.Worksheets:
        for shape in ws.Shapes:
            if shape.Type == 3:
                key = f"ws_{ws.Name}_chart_{shape.Name}"
                if key in saved_axes:
                    try:
                        _set_chart_axes_info(shape.Chart, saved_axes[key])
                        restored += 1
                    except Exception as e:
                        log_callback(f"  恢复图表轴设置失败 ({ws.Name}/{shape.Name}): {e}")
    for chart_sheet in wb.Charts:
        key = f"chart_sheet_{chart_sheet.Name}"
        if key in saved_axes:
            try:
                _set_chart_axes_info(chart_sheet, saved_axes[key])
                restored += 1
            except Exception as e:
                log_callback(f"  恢复图表轴设置失败 ({chart_sheet.Name}): {e}")

    log_callback(f"  已恢复 {restored} 个图表的轴设置")


def _set_chart_axes_info(chart, axes_info):
    """将保存的轴设置应用回图表。"""
    try:
        axes = list(chart.Axes())
    except Exception:
        return

    for i, info in enumerate(axes_info):
        if i >= len(axes):
            break
        try:
            ax = axes[i]
            if not info.get('minimum_scale_is_auto', True):
                ax.MinimumScale = info['minimum_scale']
            if not info.get('maximum_scale_is_auto', True):
                ax.MaximumScale = info['maximum_scale']
            if not info.get('major_unit_is_auto', True):
                ax.MajorUnit = info['major_unit']
            if not info.get('minor_unit_is_auto', True):
                ax.MinorUnit = info['minor_unit']
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 将单个数据表复制到 PPT 的一个独立页面
# ---------------------------------------------------------------------------

def _copy_table_to_slide(pres, ws, block, log_callback=print):
    """
    将 Excel 中一个数据表区域复制到 PowerPoint 的一个新幻灯片。
    block: (top, left, bottom, right)
    """
    top, left, bottom, right = block
    try:
        rng = ws.Range(ws.Cells(top, left), ws.Cells(bottom, right))
        rng.Copy()

        slide = pres.Slides.Add(pres.Slides.Count + 1, 12)  # ppLayoutBlank

        # 尝试粘贴为可编辑对象
        try:
            slide.Shapes.PasteSpecial(DataType=0, Link=False, DisplayAsIcon=False)
            log_callback(f"  数据表 [{top},{left}]-[{bottom},{right}] 粘贴为可编辑对象")
            return True
        except Exception:
            pass

        # 尝试 OLE 对象粘贴
        try:
            slide.Shapes.PasteSpecial(DataType=10, Link=False, DisplayAsIcon=False)
            log_callback(f"  数据表粘贴为 OLE 对象")
            return True
        except Exception:
            pass

        # 回退到普通 Paste
        slide.Shapes.Paste()
        log_callback(f"  数据表粘贴为常规对象")
        return True

    except Exception as e:
        log_callback(f"  数据表粘贴失败: {e}")

    # 最终后备：导出为图片
    try:
        tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        tmp_path = tmp.name
        tmp.close()
        rng = ws.Range(ws.Cells(top, left), ws.Cells(bottom, right))
        rng.CopyPicture(Format=2)  # xlBitmap
        slide = pres.Slides.Add(pres.Slides.Count + 1, 12)
        slide.Shapes.Paste()
        log_callback(f"  数据表粘贴为图片")
        return True
    except Exception as e2:
        log_callback(f"  数据表导出失败: {e2}")
        return False


# ---------------------------------------------------------------------------
# 主函数
# ---------------------------------------------------------------------------

def export_charts_to_ppt(excel_path, ppt_path, template_path=None, sheet_names=None, copy_tables=False):
    """
    将 Excel 工作簿中指定工作表的图表和数据表复制到 PowerPoint。

    核心改进：
    - 检测每个工作表中的多个独立数据表，分别裁剪空白行/列
    - 裁剪前保存图表坐标轴，裁剪后恢复，防止纵轴比例变形
    - 每个数据表单独占一个 PPT 页面
    """
    excel_path = os.path.abspath(excel_path)
    ppt_path = os.path.abspath(ppt_path)
    if template_path:
        template_path = os.path.abspath(template_path)

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 文件不存在: {excel_path}")

    # ---- 启动 Excel ----
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    try:
        wb = excel_app.Workbooks.Open(excel_path)
    except Exception as e:
        excel_app.Quit()
        raise RuntimeError(f"无法打开 Excel 文件 {excel_path}: {e}")

    # ---- 确定目标工作表 ----
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

    # ---- 1. 检测所有工作表中的数据块 ----
    all_blocks = {}  # sheet_name -> [(top, left, bottom, right), ...]
    for sheet in target_sheets:
        print(f"\n处理工作表：{sheet.Name}")
        blocks = _find_data_blocks(sheet, log_callback=print)
        all_blocks[sheet.Name] = blocks

    # ---- 2. 保存图表坐标轴设置（防止裁剪后变形） ----
    print("\n保存图表坐标轴设置...")
    saved_axes = _save_chart_axes(wb, log_callback=print)

    # ---- 3. 裁剪每个数据块的空白行/列 ----
    # 从右下到左上处理，避免行列删除后索引偏移
    for sheet in target_sheets:
        blocks = all_blocks.get(sheet.Name, [])
        # 按右下优先排序
        sorted_blocks = sorted(blocks, key=lambda b: (-b[2], -b[3]))
        for block in sorted_blocks:
            print(f"\n  裁剪数据块: 行 {block[0]}-{block[2]}, 列 {block[1]}-{block[3]}")
            _trim_block(sheet, block[0], block[1], block[2], block[3], log_callback=print)

    # ---- 4. 恢复图表坐标轴设置 ----
    print("\n恢复图表坐标轴设置...")
    _restore_chart_axes(wb, saved_axes, log_callback=print)

    # ---- 5. 保存工作簿，使图表数据源自动更新 ----
    wb.Save()

    # ---- 6. 收集所有图表 ----
    charts = []
    # 独立图表（ChartSheet）
    for chart_sheet in wb.Charts:
        charts.append(('ChartSheet', chart_sheet.Name, chart_sheet))
    # 嵌入图表
    for sheet in target_sheets:
        for shape in sheet.Shapes:
            if shape.Type == 3:  # xlChart
                charts.append(('Worksheet', sheet.Name, shape.Chart))

    print(f"\n共收集到 {len(charts)} 个图表")

    # ---- 7. 启动 PowerPoint ----
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        ppt_app.Visible = False
    except Exception:
        ppt_app.Visible = True

    if template_path and os.path.exists(template_path):
        pres = ppt_app.Presentations.Open(template_path)
    else:
        pres = ppt_app.Presentations.Add()

    # ---- 8. 复制图表到 PPT ----
    for source, sheet_name, chart_obj in charts:
        success = False
        for attempt in range(2):
            try:
                excel_app.CutCopyMode = False
                if source == 'Worksheet':
                    chart_obj.Parent.Activate()
                    try:
                        chart_obj.Activate()
                    except Exception:
                        pass
                    try:
                        chart_obj.ChartArea.Select()
                    except Exception:
                        pass
                    try:
                        chart_obj.Select()
                    except Exception:
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

    # ---- 9. 复制数据表到 PPT（每个表一个独立页面） ----
    if copy_tables:
        print("\n开始复制数据表到 PPT...")
        for sheet in target_sheets:
            blocks = all_blocks.get(sheet.Name, [])
            for block in blocks:
                print(f"\n  复制数据表 [{sheet.Name}]: 行 {block[0]}-{block[2]}, 列 {block[1]}-{block[3]}")
                _copy_table_to_slide(pres, sheet, block, log_callback=print)

    # ---- 10. 保存并清理 ----
    pres.SaveAs(ppt_path)
    pres.Close()
    ppt_app.Quit()
    wb.Close(SaveChanges=False)  # 已保存过，关闭时不再保存
    excel_app.Quit()
    del pres, ppt_app, wb, excel_app
    print(f"\nPPT 已生成: {ppt_path}")
