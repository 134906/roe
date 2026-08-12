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

def perform_sample_check(desc_raw_y, desc_d_y, desc_raw_x, channel_df, kpi_df, name_label_map, y_vars, x_vars, y_var_recode=True, log_callback=print):
    # 保持不变（之前已给出）
    kpi_check_rows = []
    for i, (row_kpi, y_var) in enumerate(zip(kpi_df.itertuples(), y_vars)):
        if y_var_recode:
            if y_var == 'Level_B4':
                lookup_var = 'BUMO'
            else:
                lookup_var = f"d{y_var}"
            desc_y = desc_d_y
        else:
            lookup_var = y_var
            desc_y = desc_raw_y
        if lookup_var in desc_y.index:
            mean_val = desc_y.loc[lookup_var, 'Mean']
            n_val = desc_y.loc[lookup_var, 'N']
            calc_sample = mean_val * n_val if not pd.isna(mean_val) and not pd.isna(n_val) else np.nan
        else:
            calc_sample = np.nan
            log_callback(f"警告：变量 {lookup_var} 不在 {'recoded' if y_var_recode else 'original'} KPIs 描述统计中")
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
    kpi_check_df = pd.DataFrame(kpi_check_rows)

    channel_check_rows = []
    for row_ch in channel_df.itertuples():
        var_name = row_ch.var_name
        if var_name in desc_raw_x.index:
            mean_val = desc_raw_x.loc[var_name, 'Mean']
            n_val = desc_raw_x.loc[var_name, 'N']
            calc_sample = mean_val * n_val if not pd.isna(mean_val) and not pd.isna(n_val) else np.nan
        else:
            calc_sample = np.nan
            log_callback(f"警告：变量 {var_name} 不在原始渠道描述统计中")
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
    channel_check_df = pd.DataFrame(channel_check_rows)

    all_diffs = []
    all_diffs.extend(kpi_check_df['Diff'].dropna().tolist())
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