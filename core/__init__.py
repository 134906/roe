# -*- coding: utf-8 -*-
from .config import load_definition
from .data_io import read_spss, recode_variables, merge_variables, prepare_weight, _lower_to_col, parse_filter_expr
from .descriptives import get_descriptives, perform_sample_check
from .factor_analysis import perform_factor_analysis
from .regression import run_regression, parse_factor_indicator
from .crosstabs import seen_vs_noseen
from .output import save_results
from .sample_extract import extract_samples_from_full_table
from .workfile_io import fill_workfile
import pandas as pd
import numpy as np


def run_analysis(spss_path, def_path, output_path="ROE_Results", workfile_path=None, log_callback=print, generate_ppt=True, ppt_template=None):
    # 第1步：加载配置
    log_callback("开始加载定义文件...")
    config = load_definition(def_path)

    filter_expr = config.get('filter_expr', '')
    log_callback(f"筛选表达式：{filter_expr}")

    # 第2步：读取SPSS数据（仅传入 filter_expr）
    log_callback("读取SPSS数据...")
    df, meta, name_label_map, var_to_value_labels = read_spss(
        spss_path,
        filter_expr=filter_expr
    )
    log_callback(f"数据读取成功，样本数: {len(df)}")

    # ---- 定义解析函数（加强版，兜底遍历列） ----
    def resolve_var(name):
        if name is None:
            return None
        if not isinstance(name, str):
            return name
        name_clean = name.strip()
        if name_clean in df.columns:
            return name_clean
        lower_name = name_clean.lower()
        if lower_name in _lower_to_col:
            return _lower_to_col[lower_name]
        for col in df.columns:
            if col.strip().lower() == lower_name:
                return col
        return name_clean

    # ---- 标签到变量名的映射（不区分大小写） ----
    label_to_name_lower = {}
    for name, label in zip(meta.column_names, meta.column_labels):
        if label is not None and str(label).strip():
            label_to_name_lower[str(label).strip().lower()] = name

    # ---- 解析 X_Var 和 Y_Var ----
    x_vars_resolved = [resolve_var(v) for v in config['x_vars'] if resolve_var(v) in df.columns]
    y_vars_resolved = [resolve_var(v) for v in config['y_vars'] if resolve_var(v) in df.columns]

    merged_x_vars = config.get('merged_x_vars', [])

    if config.get('weight_var'):
        config['weight_var'] = resolve_var(config['weight_var']) or config['weight_var']

    # 第3步：X 变量合并
    y_var_recode = config.get('y_var_recode', True)
    x_var_merge = config.get('x_var_merge', False)
    merge_configs = config.get('merge_configs', [])

    if x_var_merge and merge_configs:
        log_callback("执行 X 变量合并...")
        df = merge_variables(df, merge_configs, log_callback)
        log_callback(f"合并完成，新增变量: {merged_x_vars}")

    # 第4步：重编码变量
    log_callback(f"重编码变量... (Y变量recode: {y_var_recode}, X变量merge: {x_var_merge})")
    all_x_for_recode = list(config['x_vars'])
    df = recode_variables(df, all_x_for_recode)
    if y_var_recode:
        df = recode_variables(df, config['y_vars'])

    # 权重预处理
    weight_var = config.get('weight_var', '')
    if weight_var:
        weight_series, weight_valid_mask = prepare_weight(df, weight_var, log_callback)
        if weight_series is not None:
            log_callback(f"权重变量 '{weight_var}' 已使用，有效权重样本数：{weight_valid_mask.sum()}")
        else:
            log_callback(f"警告：权重变量 '{weight_var}' 无效或不在数据中，将不使用权重")
    else:
        weight_series = None
        weight_valid_mask = None
        log_callback("未配置权重变量")

    quasi_orig_x = merged_x_vars if x_var_merge else x_vars_resolved

    # 确定 recoded X 变量名（用于因子分析）
    used_x_vars = []
    for x in quasi_orig_x:
        if x in merged_x_vars:
            actual = resolve_var(x)
            if actual in df.columns:
                used_x_vars.append(actual)
            else:
                log_callback(f"警告：合并变量 {x} 不存在，已忽略")
        else:
            recoded = f'd{x}'
            actual = resolve_var(recoded)
            if actual in df.columns:
                used_x_vars.append(actual)
            else:
                if resolve_var(x) in df.columns:
                    used_x_vars.append(resolve_var(x))
                else:
                    log_callback(f"警告：变量 {x} 重编码后不存在，已忽略")

    # 第5步：Factor_Target 映射
    target = config.get('factor_target')
    if target is not None:
        mapped_index = []
        for item in target.index:
            item_str = str(item).strip()
            item_lower = item_str.lower()
            mapped_name = None

            mapped_name = resolve_var(item_str)
            if mapped_name not in df.columns:
                if item_str in label_to_name_lower:
                    mapped_name = label_to_name_lower[item_str]
                elif item_lower in label_to_name_lower:
                    mapped_name = label_to_name_lower[item_lower]
                if mapped_name is not None:
                    mapped_name = resolve_var(mapped_name)
            if mapped_name not in df.columns:
                raise ValueError(f"目标矩阵行标签 '{item}' 无法对应到任何数据列，请检查")
            mapped_index.append(mapped_name)

        target.index = mapped_index
        config['factor_target'] = target
        used_x_vars = mapped_index

    config['used_x_vars'] = used_x_vars

    # 第6步：Y 变量清单
    used_y_vars_orig = [resolve_var(v) for v in config['y_vars'] if resolve_var(v) in df.columns]
    used_y_vars_recoded = []
    if y_var_recode:
        for y in config['y_vars']:
            var_d = f'd{y}'
            actual_d = resolve_var(var_d)
            if actual_d in df.columns:
                used_y_vars_recoded.append(actual_d)
                log_callback(f"Y 变量重编码成功：{y} -> {actual_d}")
                continue
            actual_orig = resolve_var(y)
            if actual_orig in df.columns:
                used_y_vars_recoded.append(actual_orig)
                log_callback(f"Y 变量使用原始列：{y} -> {actual_orig}（未重编码）")
                continue
            log_callback(f"⚠️ Y 变量 '{y}' 不存在（包括 d 前缀），将忽略该变量")
    else:
        used_y_vars_recoded = used_y_vars_orig

    config['used_y_vars_orig'] = used_y_vars_orig
    config['used_y_vars_recoded'] = used_y_vars_recoded

    # 获取 X 变量标准差
    all_x_std = None
    if quasi_orig_x:
        existing_vars = [v for v in quasi_orig_x if v in df.columns]
        if existing_vars:
            desc_all_x = get_descriptives(df, existing_vars, weight_series=weight_series)
            if not desc_all_x.empty:
                all_x_std = desc_all_x.loc[existing_vars, 'Std'].values
                log_callback(f"全部 X 变量标准差（按原始顺序）: {all_x_std}")
            else:
                log_callback("警告：无法获取全部 X 变量的标准差")
        else:
            log_callback("警告：没有可用的 X 变量用于标准差提取")

    # 第7步：因子分析
    log_callback("执行因子分析...")
    if not used_x_vars:
        raise ValueError("没有可用的X变量用于因子分析")

    _, loadings, score_coeff, factor_scores, df, factor_corr, fa_extra = perform_factor_analysis(
        df, config, weight_series=weight_series, log_callback=log_callback
    )
    n_factors = fa_extra['n_factors'] if fa_extra else 0
    log_callback(f"因子分析完成，因子数: {n_factors}")

    # 第8步：描述统计
    desc_raw_y = get_descriptives(df, used_y_vars_orig, weight_series=weight_series)
    if y_var_recode:
        desc_d_y = get_descriptives(df, used_y_vars_recoded, weight_series=weight_series)
    else:
        desc_d_y = pd.DataFrame()

    desc_raw_x = get_descriptives(df, quasi_orig_x, weight_series=weight_series)

    # ---- 新增：计算重编码后的 X 变量描述统计 ----
    recoded_x_vars_list = []
    for var in x_vars_resolved:
        if var.startswith('d'):
            cand = var
        else:
            cand = 'd' + var
        if cand in df.columns:
            recoded_x_vars_list.append(cand)
    recoded_x_vars_list = list(dict.fromkeys(recoded_x_vars_list))  # 去重
    if recoded_x_vars_list:
        desc_recoded_x = get_descriptives(df, recoded_x_vars_list, weight_series=weight_series)
    else:
        desc_recoded_x = pd.DataFrame()

    desc_list = []
    if not desc_raw_y.empty:
        desc_list.append(('Original KPIs', desc_raw_y))
    if not desc_d_y.empty:
        desc_list.append(('Recoded KPIs', desc_d_y))
    if not desc_raw_x.empty:
        desc_list.append(('Original Channels', desc_raw_x))
    if not desc_recoded_x.empty:
        desc_list.append(('Recoded Channels', desc_recoded_x))

    log_callback("描述统计完成")

    # ---- 第9步：样本量核对（支持 Full_Table 和 Data check 双来源） ----
    kpi_check_df = None
    channel_check_df = None
    success = False
    source_flag = None

    # ---- 品牌筛选处理（仅用于样本量检查） ----
    brand_values = []
    brand_var = None
    if config.get('brand_filter_enabled', False):
        expr = config.get('filter_expr', '')
        if expr:
            import re
            # 先找 brand= 的变量
            brand_match = re.search(r'brand\s*=\s*([^;&]+)', expr)
            if brand_match:
                brand_values_raw = [v.strip() for v in brand_match.group(1).split(',') if v.strip()]
                brand_var = 'brand'
                log_callback(f"从表达式中提取品牌值（brand变量）：{brand_values_raw}")
            else:
                # 若没有 brand，则取第一个条件的值
                parts = re.split(r'[;&]', expr)
                for part in parts:
                    if '=' in part:
                        var, vals_str = part.split('=', 1)
                        brand_values_raw = [v.strip() for v in vals_str.split(',') if v.strip()]
                        brand_var = var.strip()
                        log_callback(f"未找到 brand 变量，使用第一个条件（{brand_var}）的值作为品牌列表：{brand_values_raw}")
                        break
            # ---- 将数值转换为品牌标签 ----
            if brand_var and brand_values_raw and brand_var in var_to_value_labels:
                label_map = var_to_value_labels[brand_var]  # {数值: 标签文本}
                converted = []
                for v in brand_values_raw:
                    # 尝试转为数值（可能是字符串形式的数字）
                    try:
                        num_val = int(v) if v.isdigit() else float(v)
                    except ValueError:
                        num_val = v  # 保留原字符串
                    if num_val in label_map:
                        converted.append(str(label_map[num_val]))
                    else:
                        # 如果未匹配，保留原值（可能是文本）
                        converted.append(v)
                brand_values = converted
                log_callback(f"品牌值转换为标签：{brand_values}")
            else:
                # 无值标签映射或品牌变量不存在，直接使用原始值
                brand_values = brand_values_raw
                if brand_var and brand_var not in var_to_value_labels:
                    log_callback(f"警告：品牌变量 '{brand_var}' 没有值标签映射，将直接使用原始值")
        else:
            log_callback("品牌筛选启用但 Filter 表达式为空，无法提取品牌值")
    else:
        log_callback("品牌筛选未启用")


    # ---- 定义解析 Data check 工作表的函数（内部） ----
    def parse_data_check_blocks(def_path, log_callback):
        """
        从 Data check 工作表中解析 KPI 和渠道样本量。
        KPI：表头找 count 列，若无则默认第5列（索引4）。
        渠道：从第一个非 base 行开始读取，跳过 net，遇到包含 'other' 的行停止。
        返回：kpi_df, channel_df_orig, channel_df_merged，每个均包含 'var_name', 'sample', 'table_name'
        """
        import re
        try:
            df_raw = pd.read_excel(def_path, sheet_name='Data_Check', header=None)
        except Exception as e:
            log_callback(f"无法读取 Data check 工作表：{e}")
            return None, None, None

        df_raw = df_raw.dropna(how='all')
        log_callback(f"Data check 工作表读取成功，有效行数：{len(df_raw)}")

        kpi_pattern = re.compile(r'\bkpi\b', re.I)
        orig_pattern = re.compile(r'channel/merge前|channel before merge', re.I)
        merged_pattern = re.compile(r'channel/merge后|channel after merge', re.I)

        marker_rows = []
        for idx, row in df_raw.iterrows():
            a_val = str(row[0]).strip() if pd.notna(row[0]) else ''
            if kpi_pattern.search(a_val):
                marker_rows.append((idx, 'kpi'))
                log_callback(f"行 {idx}: 识别为 KPI 标记 (内容: '{a_val}')")
            elif orig_pattern.search(a_val):
                marker_rows.append((idx, 'orig'))
                log_callback(f"行 {idx}: 识别为 原始渠道 标记 (内容: '{a_val}')")
            elif merged_pattern.search(a_val):
                marker_rows.append((idx, 'merged'))
                log_callback(f"行 {idx}: 识别为 合并渠道 标记 (内容: '{a_val}')")

        if not marker_rows:
            log_callback("未找到任何表头标记")
            return None, None, None

        max_idx = df_raw.index[-1]
        blocks = []
        for i, (marker_idx, block_type) in enumerate(marker_rows):
            start = marker_idx + 1
            end = max_idx if i + 1 == len(marker_rows) else marker_rows[i+1][0] - 1
            end_actual = end
            empty_run = 0
            for idx in range(start, end + 1):
                if idx not in df_raw.index:
                    continue
                row = df_raw.loc[idx]
                is_empty = row.isnull().all() or all(str(v).strip() == '' for v in row)
                if is_empty:
                    empty_run += 1
                    if empty_run >= 3:
                        end_actual = idx - 3
                        log_callback(f"  在行 {idx} 遇到连续三行空白，截断结束为 {end_actual}")
                        break
                else:
                    empty_run = 0
            if end_actual >= start:
                blocks.append({'type': block_type, 'start': start, 'end': end_actual})
                log_callback(f"块 {block_type}: 行 {start}-{end_actual}")
            else:
                log_callback(f"块 {block_type} 无有效数据，跳过")

        kpi_data = None
        channel_orig_data = None
        channel_merged_data = None

        for block in blocks:
            btype = block['type']
            start = block['start']
            end = block['end']
            log_callback(f"开始解析块 {btype} (行 {start}-{end})")

            if btype == 'kpi':
                header_row = df_raw.loc[start] if start in df_raw.index else None
                sample_col = None
                if header_row is not None:
                    for col_idx, val in enumerate(header_row):
                        if pd.isna(val):
                            continue
                        val_str = str(val).lower()
                        if any(keyword in val_str for keyword in ['count', '样本量', 'sample']):
                            sample_col = col_idx
                            log_callback(f"  KPI 表找到 count 列：索引 {col_idx}, 值 '{val}'")
                            break
                if sample_col is None:
                    sample_col = 4
                    log_callback(f"  KPI 表未找到 count 列，默认使用列索引 {sample_col} (第5列)")

                rows = []
                for idx in range(start + 1, end + 1):
                    if idx not in df_raw.index:
                        continue
                    row = df_raw.loc[idx]
                    var_name = str(row[0]).strip() if pd.notna(row[0]) else ''
                    if not var_name:
                        continue
                    sample_val = row[sample_col] if sample_col < len(row) else np.nan
                    sample = pd.to_numeric(sample_val, errors='coerce')
                    if pd.isna(sample) or sample == 0:
                        log_callback(f"  KPI {var_name} 样本量无效 ({sample_val})，跳过")
                        continue
                    rows.append({'var_name': var_name, 'sample': sample, 'table_name': 'KPI'})
                    log_callback(f"  有效 KPI: {var_name} -> {sample}")
                if rows:
                    kpi_data = pd.DataFrame(rows)
                    log_callback(f"KPI 表解析完成，共 {len(kpi_data)} 个有效 KPI")
                else:
                    kpi_data = pd.DataFrame(columns=['var_name', 'sample', 'table_name'])
                    log_callback("KPI 表无有效数据")

            elif btype in ('orig', 'merged'):
                base_idx = None
                for idx in range(start, end + 1):
                    if idx not in df_raw.index:
                        continue
                    row = df_raw.loc[idx]
                    a_val = str(row[0]).strip() if pd.notna(row[0]) else ''
                    if 'base' in a_val.lower():
                        base_idx = idx
                        log_callback(f"  找到第一个 base 行：行 {idx}, 内容 '{a_val}'")
                        break
                if base_idx is None:
                    log_callback(f"块 {btype} 未找到 base 行，跳过")
                    continue

                read_start = base_idx + 1
                for idx in range(base_idx + 1, end + 1):
                    if idx not in df_raw.index:
                        continue
                    row = df_raw.loc[idx]
                    a_val = str(row[0]).strip() if pd.notna(row[0]) else ''
                    if 'base' not in a_val.lower():
                        read_start = idx
                        log_callback(f"  从行 {idx} 开始读取（非 base）")
                        break

                table_name = 'Channel/merge前' if btype == 'orig' else 'Channel/merge后'
                rows = []
                for idx in range(read_start, end + 1):
                    if idx not in df_raw.index:
                        continue
                    row = df_raw.loc[idx]
                    var_name = str(row[0]).strip() if pd.notna(row[0]) else ''
                    if not var_name:
                        continue
                    if 'other' in var_name.lower():
                        log_callback(f"  遇到 other 行，停止：{var_name}")
                        break
                    if 'net' in var_name.lower():
                        log_callback(f"  跳过 net 行：{var_name}")
                        continue
                    sample_val = row[1] if len(row) > 1 else np.nan
                    sample = pd.to_numeric(sample_val, errors='coerce')
                    if pd.isna(sample) or sample == 0:
                        log_callback(f"  渠道 {var_name} 样本量无效 ({sample_val})，跳过")
                        continue
                    rows.append({'var_name': var_name, 'sample': sample, 'table_name': table_name})
                    log_callback(f"  有效渠道：{var_name} -> {sample}")

                if rows:
                    df_block = pd.DataFrame(rows)
                    log_callback(f"渠道表 {btype} 解析完成，共 {len(df_block)} 个有效渠道")
                    if btype == 'orig':
                        channel_orig_data = df_block
                    else:
                        channel_merged_data = df_block
                else:
                    log_callback(f"渠道表 {btype} 无有效数据")

        if kpi_data is None:
            kpi_data = pd.DataFrame(columns=['var_name', 'sample', 'table_name'])
        return kpi_data, channel_orig_data, channel_merged_data

    # ---- 主逻辑：优先 Full_Table，失败则 Data check ----
    kpi_df = None
    channel_df_orig = None
    channel_df_merged = None

    try:
        full_table_df = pd.read_excel(def_path, sheet_name='Full_Table', header=None)
        if full_table_df.empty:
            raise ValueError("Full_Table 工作表为空")

        brands, channel_df, kpi_df = extract_samples_from_full_table(
            full_table_df, config, name_label_map,
            filter_values=brand_values,   # 传入品牌列表
            var_to_value_labels=var_to_value_labels,
            filter_type='numeric',       # 不再使用，但保留以兼容
            log_callback=log_callback
        )
        source_flag = 'full_table'
        log_callback("样本量核对来源：Full_Table")
    except Exception as e:
        log_callback(f"Full_Table 提取样本量失败：{e}，尝试备用来源 Data check...")
        kpi_df, channel_df_orig, channel_df_merged = parse_data_check_blocks(def_path, log_callback)
        if kpi_df is not None and not kpi_df.empty:
            source_flag = 'data_check'
            log_callback("样本量核对来源：Data check")
        else:
            source_flag = None
            log_callback("Data check 数据不完整，跳过样本量核对")
            kpi_df = channel_df_orig = channel_df_merged = None

    # ---- 执行核对 ----
    if source_flag is not None and kpi_df is not None and not kpi_df.empty:
        try:
            # 先准备一个空的结果容器
            all_kpi_check = None
            all_channel_check = []
            all_success_flags = []

            if source_flag == 'full_table':
                # Full_Table：单次调用即可
                kpi_check_df, channel_check_df, success = perform_sample_check(
                    desc_raw_y, desc_d_y, desc_raw_x, channel_df, kpi_df, name_label_map,
                    config['y_vars'], quasi_orig_x,
                    y_var_recode=y_var_recode,
                    source=source_flag
                )
            else:  # data_check
                # ---- 先做 KPI 核对（只做一次） ----
                kpi_check_df, _, _ = perform_sample_check(
                    desc_raw_y, desc_d_y, desc_raw_x, None, kpi_df, name_label_map,
                    config['y_vars'], quasi_orig_x,
                    y_var_recode=y_var_recode,
                    source=source_flag,
                    skip_channel=True   # 新增参数，跳过渠道核对
                )
                all_kpi_check = kpi_check_df

                # ---- 渠道核对（根据 merge 状态决定几个表） ----
                channel_dfs_to_check = []
                if x_var_merge:
                    # 有 merge：两个渠道表都要核对
                    if channel_df_merged is not None and not channel_df_merged.empty:
                        channel_dfs_to_check.append(channel_df_merged)
                    if channel_df_orig is not None and not channel_df_orig.empty:
                        channel_dfs_to_check.append(channel_df_orig)
                else:
                    # 无 merge：只核对第二个表（即 orig 表）
                    if channel_df_orig is not None and not channel_df_orig.empty:
                        channel_dfs_to_check.append(channel_df_orig)

                for ch_df in channel_dfs_to_check:
                    _, ch_check, ch_success = perform_sample_check(
                        desc_raw_y, desc_d_y, desc_raw_x, ch_df, None, name_label_map,
                        config['y_vars'], quasi_orig_x,
                        y_var_recode=y_var_recode,
                        source=source_flag,
                        skip_kpi=True   # 跳过 KPI 核对
                    )
                    if ch_check is not None and not ch_check.empty:
                        all_channel_check.append(ch_check)
                    all_success_flags.append(ch_success)

                # ---- 合并结果 ----
                if all_kpi_check is not None and not all_kpi_check.empty:
                    channel_check_df = pd.concat(all_channel_check, ignore_index=True) if all_channel_check else pd.DataFrame()
                else:
                    channel_check_df = pd.concat(all_channel_check, ignore_index=True) if all_channel_check else pd.DataFrame()
                # 合并 KPI 和渠道的差异，重新判定 success
                all_diffs = []
                if all_kpi_check is not None and not all_kpi_check.empty:
                    all_diffs.extend(all_kpi_check['Diff'].dropna().tolist())
                if channel_check_df is not None and not channel_check_df.empty:
                    all_diffs.extend(channel_check_df['Diff'].dropna().tolist())
                success = all(abs(d) < 1e-6 for d in all_diffs) if all_diffs else False
                kpi_check_df = all_kpi_check if all_kpi_check is not None else pd.DataFrame()

            log_callback(f"样本量核对完成，整体校验结果：{'成功' if success else '存在差异'}")
        except Exception as e:
            log_callback(f"样本量核对过程发生异常：{e}，将跳过核对结果")
            import traceback
            log_callback(traceback.format_exc())
            kpi_check_df = pd.DataFrame()
            channel_check_df = pd.DataFrame()
            success = False
    else:
        kpi_check_df = pd.DataFrame()
        channel_check_df = pd.DataFrame()
        success = False
        log_callback("样本量核对跳过（无可用的样本量数据）")
    log_callback("样本量核对步骤结束，继续执行后续分析...")

    # 第10步：回归分析 + 收集描述统计信息
    reg_results = {}
    reg_descriptives_info = []  # 存储每个回归的描述统计和变量信息
    for idx, reg_cfg in enumerate(config.get('reg_configs', []), start=1):
        dep_orig = reg_cfg.get('dep', [])
        ind_orig = reg_cfg.get('ind', [])
        filter_cfg = reg_cfg.get('filter', None)  # {'expr': 'var=val', 'type': 'cat/num'}

        if not dep_orig or not ind_orig:
            log_callback(f"警告：回归 {idx} 配置不完整（缺因变量或自变量），跳过")
            continue

        # 解析因变量（用于回归）
        dep_vars = []
        for d in dep_orig:
            resolved = resolve_var(d)
            if resolved in df.columns:
                dep_vars.append(resolved)
            else:
                log_callback(f"警告：因变量 '{d}' 不在数据中，已忽略")
        if not dep_vars:
            log_callback(f"警告：回归 {idx} 的因变量均不存在于数据中，跳过")
            continue

        # 解析自变量（用于回归，Factor 开头的保留，其他的解析为列名）
        ind_parsed = parse_factor_indicator(ind_orig, n_factors)
        ind_vars_resolved = []
        for item in ind_parsed:
            if item.startswith('Factor'):
                ind_vars_resolved.append(item)
            else:
                resolved = resolve_var(item)
                if resolved in df.columns:
                    ind_vars_resolved.append(resolved)
                else:
                    log_callback(f"警告：回归 {idx} 的自变量 '{item}' 不存在于数据中，已忽略")
        if not ind_vars_resolved:
            log_callback(f"警告：回归 {idx} 的自变量解析后为空，跳过")
            continue

        # ---- 应用筛选（如果有） ----
        filter_desc = None
        if filter_cfg:
            expr = filter_cfg.get('expr', '')
            if expr and expr.strip():
                try:
                    df_reg = parse_filter_expr(expr, df, var_to_value_labels, log_callback=log_callback)
                    filter_desc = expr
                    log_callback(f"回归 {idx} 应用筛选：{filter_desc}，样本数：{len(df_reg)}")
                except Exception as e:
                    log_callback(f"警告：回归 {idx} 筛选失败：{e}，将使用全部数据")
                    df_reg = df
                    filter_desc = None
            else:
                df_reg = df
                filter_desc = None
        else:
            df_reg = df
            filter_desc = None

        log_callback(f"运行回归 {idx}...")
        reg_result = run_regression(
            df=df_reg,
            dep_vars=dep_vars,
            ind_vars=ind_vars_resolved,
            weight_var=weight_var,
            center=False
        )
        reg_results[f'Reg{idx}'] = reg_result

        # ---- 收集该回归的描述统计变量 ----
        # 因变量：分开原始和重编码
        y_orig_vars = []
        y_recoded_vars = []
        for dep in dep_vars:
            if dep.startswith('d'):
                orig_var = dep[1:]
                if orig_var in df.columns:
                    y_orig_vars.append(orig_var)
                y_recoded_vars.append(dep)
            else:
                y_orig_vars.append(dep)
                d_var = f'd{dep}'
                if d_var in df.columns:
                    y_recoded_vars.append(d_var)
        # 去重（保持顺序）
        y_orig_vars = list(dict.fromkeys(y_orig_vars))
        y_recoded_vars = list(dict.fromkeys(y_recoded_vars))

        # 自变量：所有 X 变量（quasi_orig_x）
        x_vars_for_desc = [v for v in quasi_orig_x if v in df_reg.columns]

        reg_descriptives_info.append({
            'reg_num': idx,
            'filter_desc': filter_desc,
            'df_used': df_reg,
            'y_orig_vars': y_orig_vars,
            'y_recoded_vars': y_recoded_vars,
            'x_vars': x_vars_for_desc,
            'dep_vars': dep_vars,
            'ind_vars': ind_vars_resolved,
            'recoded_x_vars': recoded_x_vars_list
        })

    # 第11步：Seen vs No Seen
    log_callback("计算Seen vs No Seen...")
    seen_vars_for_qd = quasi_orig_x
    y_vars_for_seen = used_y_vars_orig  # 使用解析后的原始变量列名
    if not y_vars_for_seen:
        # 如果没有原始变量，则使用重编码后的变量
        y_vars_for_seen = used_y_vars_recoded
    qd_stats, qd_desc, qd_means, channel_means_dict, diff_table, listwise_n = seen_vs_noseen(
        df=df,
        x_vars=quasi_orig_x,
        y_vars=y_vars_for_seen,
        seen_vars=seen_vars_for_qd,
        weight_series=weight_series
    )

    # 判断是否有回归使用了筛选
    has_filtered_reg = any(info['filter_desc'] is not None for info in reg_descriptives_info)

    # 第12步：保存结果
    log_callback("保存结果...")
    save_results(
        output_path=output_path,
        desc_list=desc_list,
        loadings=loadings,
        score_coeff=score_coeff,
        factor_corr=factor_corr,
        reg_results=reg_results,
        qd_stats=qd_stats,
        qd_desc=qd_desc,
        qd_means=qd_means,
        channel_means_dict=channel_means_dict,
        diff_table=diff_table,
        listwise_n=listwise_n,
        fa_extra=fa_extra,
        kpi_check_df=kpi_check_df,
        channel_check_df=channel_check_df,
        sample_check_success=success,
        factor_target=config.get('factor_target'),
        config=config,
        reg_descriptives_info=reg_descriptives_info,
        has_filtered_reg=has_filtered_reg,
        sample_check_source=source_flag
    )
    log_callback(f"分析完成！结果保存至: {output_path}")

    # 第13步：填充 workfile 模板
    workfile_output = None
    if workfile_path:
        log_callback("开始填充 workfile 模板...")
        try:
            import os
            base, ext = os.path.splitext(output_path)
            workfile_output = f"{base}_workfile_filled{ext}"
            fill_workfile(
                workfile_path=workfile_path,
                output_workfile_path=workfile_output,
                config=config,
                desc_list=desc_list,
                fa_extra=fa_extra,
                reg_results=reg_results,
                qd_stats=qd_stats,
                qd_desc=qd_desc,
                qd_means=qd_means,
                channel_means_dict=channel_means_dict,
                diff_table=diff_table,
                listwise_n=listwise_n,
                kpi_check_df=kpi_check_df,
                channel_check_df=channel_check_df,
                sample_check_success=success,
                factor_corr=factor_corr,
                log_callback=log_callback,
                reg_descriptives_info=reg_descriptives_info,
                has_filtered_reg=has_filtered_reg,
                sample_check_source=source_flag
            )
            log_callback(f"Workfile 填充完成！保存至: {workfile_output}")
        except Exception as e:
            log_callback(f"Workfile 填充过程中发生错误: {e}")
            import traceback
            log_callback(traceback.format_exc())

    # 第14步：生成 PPT（如果启用）
    if generate_ppt and workfile_output:
        log_callback("开始生成 PPT...")
        try:
            from .ppt_export import export_charts_to_ppt
            ppt_output = output_path.replace('.xlsx', '.pptx') if output_path.endswith('.xlsx') else output_path + '.pptx'
            sheet_names = ['2F.Results to CS_MIA', '2F.Results to CS_ROE SCAN']
            export_charts_to_ppt(workfile_output, ppt_output, template_path=ppt_template, sheet_names=sheet_names)
            log_callback(f"PPT 生成完成: {ppt_output}")
        except Exception as e:
            log_callback(f"生成 PPT 时出错: {e}")
            import traceback
            log_callback(traceback.format_exc())
