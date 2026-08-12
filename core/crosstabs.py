import pandas as pd
import numpy as np

def seen_vs_noseen(df, x_vars, y_vars, seen_vars=None, weight_series=None):
    """
    严格模拟 SPSS 语法：
    compute qd=0.
    if (Level_T2_1=1 or ... ) qd=1.
    使用原始变量名（与 SPSS 代码一致）。
    """
    print("\n" + "="*60)
    print("【Seen vs No Seen 调试】")
    print("="*60)

    # ---- 1. 确定用于 qd 的变量（原始列名） ----
    if seen_vars:
        qd_vars = seen_vars
    else:
        qd_vars = x_vars

    exist_qd_vars = [v for v in qd_vars if v in df.columns]
    if not exist_qd_vars:
        print(f"警告：没有 qd 变量存在于数据中。")
        empty = pd.DataFrame()
        return empty, empty, empty, {}, empty, np.nan

    print(f"用于 qd 的原始变量: {exist_qd_vars}")

    # ---- 2. 渠道变量 ----
    channel_vars = exist_qd_vars

    # ---- 3. 合成 qd（模拟 SPSS） ----
    qd = pd.Series(0, index=df.index)
    for var in exist_qd_vars:
        qd = qd | (df[var] == 1)
    qd = qd.astype(int)
    df['qd'] = qd

    print(f"qd 频数（未加权）:\n{qd.value_counts().sort_index()}")

    # ---- 4. 加权处理 ----
    if weight_series is not None:
        valid_mask = weight_series > 0
        w = weight_series[valid_mask]
        df_valid = df.loc[valid_mask]
        qd_valid = qd[valid_mask]

        total_w = w.sum()
        print(f"有效权重样本数: {valid_mask.sum()}, 总权重: {total_w:.4f}")

        # qd 描述统计
        qd_mean = np.average(qd_valid, weights=w)
        qd_var = np.average((qd_valid - qd_mean)**2, weights=w)
        qd_std = np.sqrt(qd_var)
        qd_stats = pd.DataFrame({
            'count': [total_w],
            'min': [qd_valid.min()],
            'max': [qd_valid.max()],
            'mean': [qd_mean],
            'std': [qd_std]
        }, index=['qd'])

        # qd 频数表
        qd_counts = df_valid.groupby(qd_valid).apply(lambda g: w[g.index].sum())
        qd_desc = qd_counts.to_frame('Frequency')
        qd_desc['Percent'] = (qd_counts / total_w * 100).round(2)

        # Y 变量列表
        exist_y = [y for y in y_vars if y in df.columns]
        print(f"Y 变量列表 (存在于df): {exist_y}")

        # 计算分组均值
        qd_means = pd.DataFrame()
        for y in exist_y:
            print(f"\n--- 处理 Y 变量: {y} ---")
            group_means = {}
            for group in [0, 1]:
                mask = (qd_valid == group)
                if mask.sum() == 0:
                    print(f"  qd={group}: 无样本")
                    group_means[group] = np.nan
                    continue
                w_sub = w[mask]
                x_sub = df_valid.loc[mask, y]
                valid_x = x_sub.notna()
                if valid_x.sum() == 0 or w_sub[valid_x].sum() == 0:
                    print(f"  qd={group}: 有效样本权重和为0，均值为NaN")
                    group_means[group] = np.nan
                else:
                    mean_val = np.average(x_sub[valid_x], weights=w_sub[valid_x])
                    print(f"  qd={group}: 样本数={mask.sum()}, 权重和={w_sub.sum():.4f}, 非缺失值数={valid_x.sum()}, 加权均值={mean_val:.4f}")
                    group_means[group] = mean_val
            qd_means[y] = pd.Series(group_means)
        qd_means = qd_means.T

        # 按渠道分组
        channel_means_dict = {}
        for var in channel_vars:
            if var in df_valid.columns:
                gm_list = []
                for group_val in [0, 1]:
                    group_mask = (df_valid[var] == group_val)
                    if group_mask.sum() == 0:
                        gm_row = pd.Series({y: np.nan for y in exist_y})
                    else:
                        w_sub = w[group_mask]
                        gm_row = {}
                        for y in exist_y:
                            x_sub = df_valid.loc[group_mask, y]
                            valid_x = x_sub.notna()
                            if valid_x.sum() == 0 or w_sub[valid_x].sum() == 0:
                                gm_row[y] = np.nan
                            else:
                                gm_row[y] = np.average(x_sub[valid_x], weights=w_sub[valid_x])
                        gm_row = pd.Series(gm_row)
                    gm_list.append(gm_row)
                gm_df = pd.DataFrame(gm_list, index=[0, 1])
                channel_means_dict[var] = gm_df

        # 差值表
        diff_data = {}
        if not qd_means.empty and 0 in qd_means.columns and 1 in qd_means.columns:
            diff_data['qd (Diff)'] = qd_means[1] - qd_means[0]
        for var, gm in channel_means_dict.items():
            if 0 in gm.index and 1 in gm.index:
                diff_data[f'{var} (Diff)'] = gm.loc[1] - gm.loc[0]
        diff_table = pd.DataFrame(diff_data, index=exist_y)

        # ---- Valid N (listwise) —— 与 qd 的 N 一致（总权重） ----
        listwise_n = total_w
        print(f"Valid N (listwise) = {listwise_n:.4f} (与 qd 有效样本一致)")

    else:
        # 未加权分支
        qd_stats = df['qd'].describe().to_frame().T
        qd_stats.index = ['qd']
        qd_counts = df['qd'].value_counts().sort_index()
        qd_desc = qd_counts.to_frame('Frequency')
        qd_desc['Percent'] = (qd_counts / len(df) * 100).round(2)

        exist_y = [y for y in y_vars if y in df.columns]
        if exist_y and 0 in df['qd'].unique() and 1 in df['qd'].unique():
            qd_means = df.groupby('qd')[exist_y].mean().T
        else:
            qd_means = pd.DataFrame(index=exist_y)

        channel_means_dict = {}
        for var in channel_vars:
            if var in df.columns:
                gm = df.groupby(var)[exist_y].mean()
                if 0 not in gm.index:
                    gm.loc[0] = np.nan
                if 1 not in gm.index:
                    gm.loc[1] = np.nan
                gm = gm.sort_index()
                channel_means_dict[var] = gm

        diff_data = {}
        if not qd_means.empty and 0 in qd_means.columns and 1 in qd_means.columns:
            diff_data['qd (Diff)'] = qd_means[1] - qd_means[0]
        for var, gm in channel_means_dict.items():
            if 0 in gm.index and 1 in gm.index:
                diff_data[f'{var} (Diff)'] = gm.loc[1] - gm.loc[0]
        diff_table = pd.DataFrame(diff_data, index=exist_y)

        # ---- Valid N (listwise) 等于总样本数 ----
        listwise_n = len(df)

    df.drop('qd', axis=1, inplace=True)
    print("="*60 + "\n")
    return qd_stats, qd_desc, qd_means, channel_means_dict, diff_table, listwise_n