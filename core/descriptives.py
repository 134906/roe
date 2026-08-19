import pandas as pd
import numpy as np

def get_descriptives(df, var_list, weight_series=None):
    """
    计算描述统计，支持权重。
    返回的 DataFrame 最后一行包含 "Valid N (listwise)"。
    """
    exist_vars = [v for v in var_list if v in df.columns]
    if not exist_vars:
        return pd.DataFrame()

    if weight_series is not None:
        results = []
        for var in exist_vars:
            # 只使用该变量非缺失的样本
            valid = df[var].notna()
            w = weight_series[valid]
            x = df.loc[valid, var]
            total_w = w.sum()
            if total_w == 0 or len(x) == 0:
                results.append({'N': np.nan, 'Min': np.nan, 'Max': np.nan, 'Mean': np.nan, 'Std': np.nan})
                continue
            weighted_mean = np.average(x, weights=w)
            variance = np.average((x - weighted_mean) ** 2, weights=w)
            weighted_std = np.sqrt(variance)
            results.append({
                'N': total_w,
                'Min': x.min(),
                'Max': x.max(),
                'Mean': weighted_mean,
                'Std': weighted_std
            })
        desc = pd.DataFrame(results, index=exist_vars)

        # 计算 listwise N（所有变量非缺失且权重>0）
        mask = df[exist_vars].notna().all(axis=1) & (weight_series > 0)
        listwise_n = weight_series[mask].sum()
        listwise_row = pd.DataFrame({
            'N': [listwise_n],
            'Min': [np.nan],
            'Max': [np.nan],
            'Mean': [np.nan],
            'Std': [np.nan]
        }, index=['Valid N (listwise)'])
        desc = pd.concat([desc, listwise_row])
    else:
        desc = df[exist_vars].describe(percentiles=[]).T[['count', 'min', 'max', 'mean', 'std']]
        desc.columns = ['N', 'Min', 'Max', 'Mean', 'Std']
        mask = df[exist_vars].notna().all(axis=1)
        listwise_n = mask.sum()
        listwise_row = pd.DataFrame({
            'N': [listwise_n],
            'Min': [np.nan],
            'Max': [np.nan],
            'Mean': [np.nan],
            'Std': [np.nan]
        }, index=['Valid N (listwise)'])
        desc = pd.concat([desc, listwise_row])
    return desc

def perform_sample_check(desc_raw_y, desc_d_y, desc_raw_x, channel_df, kpi_df, name_label_map, y_vars, x_vars, y_var_recode=True, source='full_table', log_callback=print, skip_kpi=False, skip_channel=False):
    import re
    def normalize(s):
        if not isinstance(s, str):
            s = str(s)
        return re.sub(r'[\s_\-]+', '', s).lower()

    # 辅助函数：将 Mean * N 结果取整（四舍五入），若存在 NaN 则返回 NaN
    def _compute_calc_sample(mean_val, n_val):
        if pd.isna(mean_val) or pd.isna(n_val):
            return np.nan
        return int(round(mean_val * n_val))

    kpi_check_df = pd.DataFrame()
    channel_check_df = pd.DataFrame()
    success = False

    # ---------- KPI 核对 ----------
    if not skip_kpi and kpi_df is not None and not kpi_df.empty:
        kpi_check_rows = []
        if source == 'full_table':
            # Full_Table 原逻辑
            for i, (row_kpi, y_var) in enumerate(zip(kpi_df.itertuples(), y_vars)):
                if y_var_recode:
                    lookup_var = f'd{y_var}'
                    desc_y = desc_d_y
                else:
                    lookup_var = y_var
                    desc_y = desc_raw_y
                if lookup_var in desc_y.index:
                    mean_val = desc_y.loc[lookup_var, 'Mean']
                    n_val = desc_y.loc[lookup_var, 'N']
                else:
                    mean_val = n_val = np.nan
                    log_callback(f"警告：变量 {lookup_var} 不在 {'重编码' if y_var_recode else '原始'} KPIs 描述统计中")
                calc_sample = _compute_calc_sample(mean_val, n_val)
                full_sample = row_kpi.sample
                diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                if not pd.isna(diff) and abs(diff) < 1e-10:
                    diff = 0.0
                kpi_check_rows.append({
                    'TableName': row_kpi.table_name,
                    'KPI_Label': row_kpi.kpi_label,
                    'FullTable_Sample': full_sample,
                    'Calc_Mean*N': calc_sample,
                    'Diff': diff
                })
        else:  # data_check
            unmatched_y = list(y_vars)
            matched_indices = set()

            # 第一轮精确匹配
            for idx, row_kpi in enumerate(kpi_df.itertuples()):
                var_name = row_kpi.var_name
                vn = normalize(var_name)
                matched_var = None
                for y in y_vars:
                    if normalize(y) == vn:
                        matched_var = y
                        break
                if matched_var is not None:
                    unmatched_y.remove(matched_var)
                    matched_indices.add(idx)
                    if y_var_recode:
                        lookup_var = f'd{matched_var}'
                        desc_y = desc_d_y
                    else:
                        lookup_var = matched_var
                        desc_y = desc_raw_y
                    if lookup_var in desc_y.index:
                        mean_val = desc_y.loc[lookup_var, 'Mean']
                        n_val = desc_y.loc[lookup_var, 'N']
                    else:
                        mean_val = n_val = np.nan
                        log_callback(f"警告：变量 {lookup_var} 不在 {'重编码' if y_var_recode else '原始'} KPIs 描述统计中")
                    calc_sample = _compute_calc_sample(mean_val, n_val)
                    full_sample = row_kpi.sample
                    diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                    if not pd.isna(diff) and abs(diff) < 1e-10:
                        diff = 0.0
                    kpi_check_rows.append({
                        'TableName': row_kpi.table_name,
                        'KPI_Label': var_name,
                        'FullTable_Sample': full_sample,
                        'Calc_Mean*N': calc_sample,
                        'Diff': diff
                    })

            # 第二轮包含匹配
            unmatched_y_remaining = list(unmatched_y)
            for idx, row_kpi in enumerate(kpi_df.itertuples()):
                if idx in matched_indices:
                    continue
                var_name = row_kpi.var_name
                vn = normalize(var_name)
                matched_var = None
                for y in unmatched_y_remaining:
                    if vn in normalize(y) or normalize(y) in vn:
                        matched_var = y
                        break
                if matched_var is not None:
                    unmatched_y_remaining.remove(matched_var)
                    matched_indices.add(idx)
                    if y_var_recode:
                        lookup_var = f'd{matched_var}'
                        desc_y = desc_d_y
                    else:
                        lookup_var = matched_var
                        desc_y = desc_raw_y
                    if lookup_var in desc_y.index:
                        mean_val = desc_y.loc[lookup_var, 'Mean']
                        n_val = desc_y.loc[lookup_var, 'N']
                    else:
                        mean_val = n_val = np.nan
                        log_callback(f"警告：变量 {lookup_var} 不在 {'重编码' if y_var_recode else '原始'} KPIs 描述统计中")
                    calc_sample = _compute_calc_sample(mean_val, n_val)
                    full_sample = row_kpi.sample
                    diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                    if not pd.isna(diff) and abs(diff) < 1e-10:
                        diff = 0.0
                    kpi_check_rows.append({
                        'TableName': row_kpi.table_name,
                        'KPI_Label': var_name,
                        'FullTable_Sample': full_sample,
                        'Calc_Mean*N': calc_sample,
                        'Diff': diff
                    })

            # 第三轮顺序匹配
            unmatched_indices = [i for i in range(len(kpi_df)) if i not in matched_indices]
            for y_var, idx in zip(unmatched_y_remaining, unmatched_indices):
                row_kpi = kpi_df.iloc[idx]
                if y_var_recode:
                    lookup_var = f'd{y_var}'
                    desc_y = desc_d_y
                else:
                    lookup_var = y_var
                    desc_y = desc_raw_y
                if lookup_var in desc_y.index:
                    mean_val = desc_y.loc[lookup_var, 'Mean']
                    n_val = desc_y.loc[lookup_var, 'N']
                else:
                    mean_val = n_val = np.nan
                    log_callback(f"警告：变量 {lookup_var} 不在 {'重编码' if y_var_recode else '原始'} KPIs 描述统计中")
                calc_sample = _compute_calc_sample(mean_val, n_val)
                full_sample = row_kpi['sample']
                diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                if not pd.isna(diff) and abs(diff) < 1e-10:
                    diff = 0.0
                kpi_check_rows.append({
                    'TableName': row_kpi['table_name'],
                    'KPI_Label': row_kpi['var_name'],
                    'FullTable_Sample': full_sample,
                    'Calc_Mean*N': calc_sample,
                    'Diff': diff
                })

        kpi_check_df = pd.DataFrame(kpi_check_rows)

    # ---------- 渠道核对 ----------
    if not skip_channel and channel_df is not None and not channel_df.empty:
        channel_check_rows = []
        if source == 'full_table':
            for row_ch in channel_df.itertuples():
                x_var = row_ch.var_name
                if x_var in desc_raw_x.index:
                    mean_val = desc_raw_x.loc[x_var, 'Mean']
                    n_val = desc_raw_x.loc[x_var, 'N']
                else:
                    mean_val = n_val = np.nan
                    log_callback(f"警告：变量 {x_var} 不在原始渠道描述统计中")
                calc_sample = _compute_calc_sample(mean_val, n_val)
                full_sample = row_ch.sample
                diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                if not pd.isna(diff) and abs(diff) < 1e-10:
                    diff = 0.0
                channel_check_rows.append({
                    'TableName': row_ch.table_name,
                    'Channel': row_ch.channel_label,
                    'FullTable_Sample': full_sample,
                    'Calc_Mean*N': calc_sample,
                    'Diff': diff
                })
        else:  # data_check
            unmatched_x = list(x_vars)
            matched_indices = set()

            # 第一轮精确匹配
            for idx, row_ch in enumerate(channel_df.itertuples()):
                var_name = row_ch.var_name
                vn = normalize(var_name)
                matched_var = None
                for x in x_vars:
                    if normalize(x) == vn:
                        matched_var = x
                        break
                if matched_var is not None:
                    unmatched_x.remove(matched_var)
                    matched_indices.add(idx)
                    if matched_var in desc_raw_x.index:
                        mean_val = desc_raw_x.loc[matched_var, 'Mean']
                        n_val = desc_raw_x.loc[matched_var, 'N']
                    else:
                        mean_val = n_val = np.nan
                        log_callback(f"警告：变量 {matched_var} 不在原始渠道描述统计中")
                    calc_sample = _compute_calc_sample(mean_val, n_val)
                    full_sample = row_ch.sample
                    diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                    if not pd.isna(diff) and abs(diff) < 1e-10:
                        diff = 0.0
                    channel_check_rows.append({
                        'TableName': row_ch.table_name,
                        'Channel': var_name,
                        'FullTable_Sample': full_sample,
                        'Calc_Mean*N': calc_sample,
                        'Diff': diff
                    })

            # 第二轮包含匹配
            unmatched_x_remaining = list(unmatched_x)
            for idx, row_ch in enumerate(channel_df.itertuples()):
                if idx in matched_indices:
                    continue
                var_name = row_ch.var_name
                vn = normalize(var_name)
                matched_var = None
                for x in unmatched_x_remaining:
                    if vn in normalize(x) or normalize(x) in vn:
                        matched_var = x
                        break
                if matched_var is not None:
                    unmatched_x_remaining.remove(matched_var)
                    matched_indices.add(idx)
                    if matched_var in desc_raw_x.index:
                        mean_val = desc_raw_x.loc[matched_var, 'Mean']
                        n_val = desc_raw_x.loc[matched_var, 'N']
                    else:
                        mean_val = n_val = np.nan
                        log_callback(f"警告：变量 {matched_var} 不在原始渠道描述统计中")
                    calc_sample = _compute_calc_sample(mean_val, n_val)
                    full_sample = row_ch.sample
                    diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                    if not pd.isna(diff) and abs(diff) < 1e-10:
                        diff = 0.0
                    channel_check_rows.append({
                        'TableName': row_ch.table_name,
                        'Channel': var_name,
                        'FullTable_Sample': full_sample,
                        'Calc_Mean*N': calc_sample,
                        'Diff': diff
                    })

            # 第三轮顺序匹配
            unmatched_indices = [i for i in range(len(channel_df)) if i not in matched_indices]
            for x_var, idx in zip(unmatched_x_remaining, unmatched_indices):
                row_ch = channel_df.iloc[idx]
                if x_var in desc_raw_x.index:
                    mean_val = desc_raw_x.loc[x_var, 'Mean']
                    n_val = desc_raw_x.loc[x_var, 'N']
                else:
                    mean_val = n_val = np.nan
                    log_callback(f"警告：变量 {x_var} 不在原始渠道描述统计中")
                calc_sample = _compute_calc_sample(mean_val, n_val)
                full_sample = row_ch['sample']
                diff = calc_sample - full_sample if not pd.isna(calc_sample) and not pd.isna(full_sample) else np.nan
                if not pd.isna(diff) and abs(diff) < 1e-10:
                    diff = 0.0
                channel_check_rows.append({
                    'TableName': row_ch['table_name'],
                    'Channel': row_ch['var_name'],
                    'FullTable_Sample': full_sample,
                    'Calc_Mean*N': calc_sample,
                    'Diff': diff
                })

        channel_check_df = pd.DataFrame(channel_check_rows)

    # ---------- 汇总判定 ----------
    all_diffs = []
    if not kpi_check_df.empty:
        all_diffs.extend(kpi_check_df['Diff'].dropna().tolist())
    if not channel_check_df.empty:
        all_diffs.extend(channel_check_df['Diff'].dropna().tolist())
    if len(all_diffs) == 0:
        success = False
        log_callback("警告：没有任何差异值可检查，可能所有样本量计算失败")
    else:
        success = all(abs(d) < 1e-6 for d in all_diffs)
        if success:
            log_callback("样本量校对成功")
        else:
            log_callback("样本量校对失败，存在差异")

    return kpi_check_df, channel_check_df, success