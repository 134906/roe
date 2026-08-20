import pandas as pd
import numpy as np
import re

def normalize(s):
    if not isinstance(s, str):
        s = str(s)
    return re.sub(r'[\s_\-]+', '', s).lower()

def find_table_start(df, keyword, log_callback=print):
    col_a = df.iloc[:, 0].astype(str).str.strip()
    for idx, val in col_a.items():
        if isinstance(val, str) and keyword.lower() in val.lower():
            log_callback(f"找到表头行：行号 {idx}，内容：{val[:50]}...")
            return idx
    log_callback(f"警告：未找到包含 '{keyword}' 的表头行")
    return None

def extract_channel_data(df, start_row, log_callback=print):
    table_name = df.iloc[start_row, 0] if start_row < len(df) else ''
    log_callback(f"表名: {table_name}")

    brand_row_idx = start_row - 1
    brand_row = df.iloc[brand_row_idx] if brand_row_idx >= 0 else None
    brands = []
    brand_col_map = {}
    if brand_row is not None:
        for col_idx, val in enumerate(brand_row.iloc[1:], start=1):
            if pd.notna(val) and str(val).strip() != '':
                brand_name = str(val).strip()
                brands.append(brand_name)
                brand_col_map[brand_name] = col_idx
    brands = list(dict.fromkeys(brands))
    log_callback(f"品牌行提取品牌：{brands[:10]}{'...' if len(brands)>10 else ''}")
    log_callback(f"品牌列映射：{brand_col_map}")

    channel_rows = []
    for i in range(start_row+1, len(df)):
        row = df.iloc[i]
        a_val = row.iloc[0]
        if pd.isna(a_val) or str(a_val).strip() == '':
            log_callback(f"跳过空行：行号 {i}")
            continue
        a_str = str(a_val).strip()
        if 'base' in a_str.lower():
            log_callback(f"跳过 Base 行：{a_str}")
            continue
        if 'others' in a_str.lower():
            log_callback(f"遇到 Others 行，停止：{a_str}")
            break
        channel_rows.append(row)
        log_callback(f"收集渠道行：{a_str}")

    if not channel_rows:
        raise ValueError("未找到任何渠道行")

    data_df = pd.DataFrame(channel_rows)
    log_callback(f"共收集到 {len(data_df)} 个渠道行")
    return data_df, table_name, brands, brand_col_map

def extract_kpi_data(df, start_row, log_callback=print):
    table_name = df.iloc[start_row, 0] if start_row < len(df) else ''
    log_callback(f"KPI 表名: {table_name}")

    brand_row_idx = start_row - 1
    brand_row = df.iloc[brand_row_idx] if brand_row_idx >= 0 else None
    if brand_row is not None:
        log_callback(f"品牌行（表头上一行）A列内容：{brand_row.iloc[0] if pd.notna(brand_row.iloc[0]) else '空'}")

    data_rows = []
    empty_count = 0
    for i in range(start_row+1, len(df)):
        row = df.iloc[i]
        is_empty = row.isnull().all() or all(str(v).strip() == '' for v in row)
        if is_empty:
            empty_count += 1
            if empty_count >= 3:
                log_callback("遇到连续三行空行，停止收集")
                break
            else:
                continue
        else:
            empty_count = 0
            data_rows.append(row)

    if not data_rows:
        raise ValueError("数据区域为空")

    data_df = pd.DataFrame(data_rows)
    log_callback(f"数据区域共 {len(data_df)} 行")
    return data_df, table_name, brand_row

def extract_samples_from_full_table(df_full, config, name_label_map,
                                    filter_values=None,
                                    var_to_value_labels=None,
                                    filter_type='numeric',
                                    log_callback=print):
    """
    从 Full_Table 中提取渠道和 KPI 的样本量。
    品牌筛选由 config['brand_filter_enabled'] 控制，品牌值列表通过 filter_values 传入。
    """
    # ============================================================
    # 准备：判断是否启用品牌筛选
    # ============================================================
    brand_filter_enabled = config.get('brand_filter_enabled', False)
    use_filtered = False
    filtered_brands = []

    log_callback("="*60)
    log_callback(f"开始样本量提取")
    log_callback(f"品牌筛选启用：{brand_filter_enabled}")
    log_callback(f"品牌值列表：{filter_values}")

    if brand_filter_enabled and filter_values:
        use_filtered = True
        filtered_brands = filter_values
        log_callback(f"将按品牌筛选：{filtered_brands}")
    else:
        log_callback("未启用品牌筛选或无品牌值，将对所有品牌求和")

    # ============================================================
    # 渠道表
    # ============================================================
    channel_keyword = config.get('sample_channel_keyword', '')
    if not channel_keyword:
        raise ValueError("配置缺少 'sample_channel_keyword'")

    log_callback(f"开始搜索渠道表，关键词：{channel_keyword}")
    start_ch = find_table_start(df_full, channel_keyword, log_callback)
    if start_ch is None:
        raise ValueError(f"未找到包含 '{channel_keyword}' 的渠道表头行")

    data_ch, table_name_ch, all_brands, brand_col_map = extract_channel_data(df_full, start_ch, log_callback)
    if data_ch is None or data_ch.empty:
        raise ValueError("渠道表数据区域为空")

    log_callback(f"渠道表所有品牌：{all_brands}")
    log_callback(f"渠道表品牌列映射：{brand_col_map}")

    target_cols = None
    if use_filtered:
        target_cols = []
        for brand in filtered_brands:
            matched = False
            for full_brand, col_idx in brand_col_map.items():
                norm_full = normalize(full_brand)
                norm_brand = normalize(brand)
                if norm_full == norm_brand or norm_brand in norm_full:
                    target_cols.append(col_idx)
                    matched = True
                    log_callback(f"匹配成功：{brand} -> 列索引 {col_idx}（品牌名：{full_brand}）")
                    break
            if not matched:
                log_callback(f"警告：品牌 {brand} 在渠道表中未匹配到任何列")
        if target_cols:
            target_cols = list(set(target_cols))
            log_callback(f"渠道表最终目标列索引：{target_cols}")
        else:
            log_callback("警告：筛选品牌在渠道表中未匹配到任何列，将使用所有品牌（原逻辑）")
            use_filtered = False
            target_cols = None

    # ---- 构建渠道行数据 ----
    x_vars = config.get('x_vars', [])
    channel_labels = [name_label_map.get(var, var) for var in x_vars]
    n_spss_channels = len(x_vars)
    log_callback(f"SPSS 渠道变量数量：{n_spss_channels}")
    log_callback(f"Full Table 渠道行数量：{len(data_ch)}")

    if len(data_ch) != n_spss_channels:
        log_callback(f"⚠️ 警告：渠道行数 ({len(data_ch)}) 与 SPSS 变量数 ({n_spss_channels}) 不一致，将按顺序取前 {min(len(data_ch), n_spss_channels)} 个对应")

    n_match = min(len(data_ch), n_spss_channels)
    channel_rows = []
    for i in range(n_match):
        row = data_ch.iloc[i]
        row_name = str(row.iloc[0]).strip()
        spss_var = x_vars[i]
        spss_label = channel_labels[i]

        is_match = (spss_label.lower() in row_name.lower()) or (row_name.lower() in spss_label.lower())
        if not is_match:
            log_callback(f"⚠️ 警告：第 {i+1} 个渠道行 '{row_name}' 与 SPSS 变量 '{spss_var}' 的 label '{spss_label}' 不匹配（包含匹配失败），但将按顺序对应")

        if use_filtered and target_cols is not None:
            log_callback(f"使用筛选逻辑，取列索引 {target_cols} 的值")
            nums = pd.to_numeric(row.iloc[target_cols], errors='coerce')
        else:
            log_callback(f"使用原逻辑，取所有列")
            nums = pd.to_numeric(row.iloc[1:], errors='coerce')
        total = nums.sum(skipna=True)
        log_callback(f"渠道行 '{row_name}' 样本量求和结果：{total}")

        channel_rows.append({
            'table_name': table_name_ch,
            'channel_label': row_name,
            'sample': total,
            'var_name': spss_var
        })
        log_callback(f"顺序匹配渠道 {i+1}: {row_name} (对应变量 {spss_var})，样本量 = {total}")

    if len(data_ch) > n_spss_channels:
        log_callback(f"注意：Full Table 中有 {len(data_ch) - n_spss_channels} 个渠道行未被使用（因为 SPSS 变量数不足）")

    channel_df = pd.DataFrame(channel_rows)
    log_callback(f"共匹配 {len(channel_df)} 个渠道")

    # ============================================================
    # KPI 表
    # ============================================================
    kpi_keywords = config.get('sample_kpi_keywords', [])
    kpi_types = config.get('sample_kpi_types', [])
    
    # 如果关键词为空，直接返回空 KPI
    if not kpi_keywords:
        log_callback("警告：未配置 KPI_Keywords，KPI 样本量提取将返回空数据框")
        return all_brands, channel_df, pd.DataFrame(columns=['table_name', 'kpi_label', 'sample'])

    # 移除强制 6 个的校验，支持任意数量
    log_callback(f"共读取到 {len(kpi_keywords)} 个 KPI 关键词")

    top2_labels = config.get('sample_top2box_label', ['Net : Top 2 Box'])
    kpi_rows = []

    # 确定匹配品牌列表
    if use_filtered and filtered_brands:
        match_brands = filtered_brands
    else:
        match_brands = all_brands

    log_callback(f"KPI 匹配品牌列表：{match_brands[:5]}{'...' if len(match_brands)>5 else ''}")

    for idx, (kw, typ) in enumerate(zip(kpi_keywords, kpi_types)):
        log_callback(f"处理第 {idx+1} 个 KPI 表，关键词：{kw}，类型：{typ}")
        start_k = find_table_start(df_full, kw, log_callback)
        if start_k is None:
            log_callback(f"警告：未找到 KPI 表头: {kw}，跳过")
            continue
        data_k, table_name_k, brand_row_k = extract_kpi_data(df_full, start_k, log_callback)

        typ = typ.lower()
        if typ == 'summary':
            log_callback("Summary 类型：遍历行，匹配行首品牌名")
            total_sample = 0
            for row in data_k.itertuples(index=False):
                row_brand = str(row[0]).strip()
                if not row_brand:
                    continue
                matched = False
                for b in match_brands:
                    if normalize(b) == normalize(row_brand) or normalize(b) in normalize(row_brand):
                        matched = True
                        log_callback(f"匹配到品牌：{b} == {row_brand}，累加值 {row[1]}")
                        break
                if matched:
                    val = pd.to_numeric(row[1], errors='coerce')
                    if not pd.isna(val):
                        total_sample += val
            kpi_rows.append({
                'table_name': table_name_k,
                'kpi_label': '',
                'sample': total_sample
            })
            log_callback(f"Summary 表 '{kw}' 最终样本量 = {total_sample}")

        elif typ == 'top2':
            log_callback("Top2 类型：查找包含 top2_label 的行，然后按品牌列求和")
            found_row = None
            for row in data_k.itertuples(index=False):
                row_name = str(row[0]).strip()
                if not row_name:
                    continue
                for label in top2_labels:
                    if normalize(label) in normalize(row_name):
                        found_row = row
                        log_callback(f"找到匹配行：{row_name}")
                        break
                if found_row is not None:
                    break
            if found_row is None:
                log_callback(f"警告：在表 '{kw}' 中未找到包含 {top2_labels} 中任一标签的行，跳过")
                continue

            if brand_row_k is not None:
                col_names = [str(c).strip() for c in brand_row_k.iloc[1:]]
            else:
                col_names = [f"col_{j}" for j in range(1, len(found_row))]
            log_callback(f"列名（品牌名）：{col_names}")

            if use_filtered and filtered_brands:
                target_indices = []
                for i, col_name in enumerate(col_names, start=1):
                    for b in filtered_brands:
                        if normalize(col_name) == normalize(b) or normalize(b) in normalize(col_name):
                            target_indices.append(i)
                            log_callback(f"匹配列：{col_name} -> 索引 {i}")
                            break
                if not target_indices:
                    log_callback(f"警告：表 '{kw}' 中未找到匹配品牌的列，将使用所有列（原逻辑）")
                    target_indices = range(1, len(found_row))
                else:
                    log_callback(f"目标列索引：{target_indices}")
            else:
                log_callback("无筛选，使用所有列（排除 total）")
                target_indices = range(1, len(found_row))

            numeric_vals = []
            for i in target_indices:
                if i < len(found_row):
                    val = found_row[i]
                    if i-1 < len(col_names) and 'total' in col_names[i-1].lower():
                        log_callback(f"跳过 total 列：{col_names[i-1]}")
                        continue
                    num = pd.to_numeric(val, errors='coerce')
                    if not pd.isna(num):
                        numeric_vals.append(num)
                        log_callback(f"取值：列索引 {i}，值 {val}")
            total_sample = sum(numeric_vals)
            kpi_rows.append({
                'table_name': table_name_k,
                'kpi_label': found_row[0],
                'sample': total_sample
            })
            log_callback(f"Top2 表 '{kw}' 最终样本量 = {total_sample}")

        else:
            log_callback(f"警告：未知的 KPI 类型: {typ}，跳过")

    kpi_df = pd.DataFrame(kpi_rows)
    log_callback(f"KPI 表提取完成，共 {len(kpi_df)} 个有效 KPI")
    return all_brands, channel_df, kpi_df