import numpy as np
import pandas as pd

# ---------- Procrustes 旋转（与 VBA 一致） ----------
def _procrustes_vba(A, C, log_callback):
    n, k = A.shape
    InC = C.copy()

    def _procrustes_std(A, C):
        n, k = A.shape
        InC = np.zeros_like(C)
        for i in range(n):
            Hj2 = np.sum(A[i, :]**2)
            Kj2 = np.sum(C[i, :]**2)
            if Kj2 == 0:
                scale = 0
            else:
                scale = np.sqrt(Hj2 / Kj2)
            InC[i, :] = C[i, :] * scale

        ACMatrix = A.T @ InC
        S = ACMatrix.T @ ACMatrix
        eigvals, P = np.linalg.eigh(S)
        idx = np.argsort(eigvals)[::-1]
        eigvals = eigvals[idx]
        P = P[:, idx]

        D = np.zeros((k, k))
        for i in range(k):
            if eigvals[i] > 1e-12:
                D[i, i] = eigvals[i] ** (-0.5)
            else:
                D[i, i] = 0
        Alpha = P @ D @ P.T
        T = ACMatrix @ Alpha
        B = A @ T
        return B

    B = _procrustes_std(A, InC)
    return B

def get_max_component_order(rotated_loadings, orderby=True):
    n_vars = rotated_loadings.shape[0]
    if not orderby:
        return list(range(n_vars))
    max_abs = np.abs(rotated_loadings).max(axis=1)
    max_idx = np.argmax(np.abs(rotated_loadings), axis=1)
    sorted_indices = sorted(range(n_vars), key=lambda i: (max_idx[i], -max_abs[i]))
    return sorted_indices

# ================= 加权协方差计算 =================
def compute_weighted_covariance(df, vars_list, weights):
    """
    计算加权协方差矩阵，分母为 Σw（与 SPSS 一致）
    """
    X = df[vars_list].values.astype(float)
    w = weights.values.reshape(-1, 1)
    total_w = w.sum()

    weighted_mean = np.average(X, axis=0, weights=weights)
    centered = X - weighted_mean
    cov_matrix = (centered.T * (w / total_w).flatten()) @ centered
    return cov_matrix

# ---------- 主因子分析函数 ----------
def perform_factor_analysis(df, config, weight_series=None, log_callback=print):
    """
    执行因子分析（PCA + Procrustes 旋转）
    支持权重（加权协方差和加权标准差）。
    非标准化得分系数表格的分母采用全局变量（按 VariableID 升序）前 k 个的标准差，
    与 VBA 的 CorRange(j,j) 一致。
    """
    # ---- 1. 读取目标矩阵 ----
    target_df = config.get('factor_target')
    if target_df is None:
        raise ValueError("config 中缺少 'factor_target'")
    mask = (target_df != 0).any(axis=1)
    active_vars = target_df.index[mask].tolist()
    active_target = target_df.loc[active_vars]
    if len(active_vars) == 0:
        raise ValueError("目标矩阵中没有任何变量属于因子")
    n_vars = len(active_vars)
    k = target_df.shape[1]   # 因子数

    # ---- 2. 计算样本协方差和相关矩阵（支持加权） ----
    if weight_series is not None:
        valid_mask = weight_series > 0
        w = weight_series[valid_mask]
        df_valid = df.loc[valid_mask]
        cov_matrix = compute_weighted_covariance(df_valid, active_vars, w)
        std = np.sqrt(np.diag(cov_matrix))
        corr_matrix = cov_matrix / np.outer(std, std)
        log_callback(f"使用加权协方差矩阵，有效权重样本数：{valid_mask.sum()}，总权重：{w.sum()}")
    else:
        X = df[active_vars].values.astype(float)
        cov_matrix = np.cov(X, rowvar=False, ddof=1)
        std = np.sqrt(np.diag(cov_matrix))
        corr_matrix = cov_matrix / np.outer(std, std)

    # ---- 3. PCA ----
    eigvals, eigvecs = np.linalg.eigh(corr_matrix)
    idx_sort = np.argsort(eigvals)[::-1]
    eigvals = eigvals[idx_sort]
    eigvecs = eigvecs[:, idx_sort]
    full_eigvals = eigvals
    full_var_ratio = eigvals / eigvals.sum()
    full_cum_var = np.cumsum(full_var_ratio)

    eigvals_k = eigvals[:k]
    eigvecs_k = eigvecs[:, :k]
    unrot_loadings = eigvecs_k * np.sqrt(np.maximum(eigvals_k, 0))

    # ---- 4. 规则矩阵 ----
    rule_matrix = active_target.values.astype(float)

    # ---- 5. Procrustes 旋转 ----
    rotation = config.get('factor_rotation', 'procrustes').lower()
    if rotation == 'procrustes':
        rot_loadings = _procrustes_vba(unrot_loadings, rule_matrix, log_callback)
        log_callback("✅ Procrustes 旋转完成。")
    else:
        raise ValueError(f"不支持的旋转方法：'{rotation}'")

    # ---- 6. 符号调整 ----
    for j in range(k):
        mask_col = rule_matrix[:, j] != 0
        belong_vars_idx = np.where(mask_col)[0]
        if len(belong_vars_idx) == 0:
            continue
        group_load = rot_loadings[belong_vars_idx, j]
        max_abs_pos = np.argmax(np.abs(group_load))
        max_val = group_load[max_abs_pos]
        if max_val < 0:
            rot_loadings[:, j] = -rot_loadings[:, j]

    # ---- 7. 标准化得分系数 ----
    gram = rot_loadings.T @ rot_loadings
    try:
        inv_gram = np.linalg.inv(gram)
    except np.linalg.LinAlgError:
        inv_gram = np.linalg.pinv(gram)
    score_coeff_std = rot_loadings @ inv_gram

    # ---- 8. 非标准化得分系数（分母为全局变量按 VariableID 升序前 k 个的标准差） ----
    # 获取全局变量列表和对应的 VariableID（来自 config）
    all_vars = config['factor_target'].index.tolist()   # 完整变量列表（含全零行）
    var_ids = config.get('factor_var_ids', [])

    # 构建 (VariableID, VarName) 并按 ID 升序排序
    if len(var_ids) == len(all_vars) and len(var_ids) >= k:
        vars_with_ids = list(zip(var_ids, all_vars))
        # 确保 VariableID 是数值型
        vars_with_ids = [(int(vid) if pd.notna(vid) else vid, name) for vid, name in vars_with_ids]
        vars_with_ids.sort(key=lambda x: x[0])   # 按 ID 升序
        selected_vars = [name for _, name in vars_with_ids[:k]]
        log_callback(f"按 VariableID 升序取前 {k} 个变量作为分母: {selected_vars}")
    else:
        # 如果 VariableID 不完整，回退到参与变量的前 k 个（发出警告）
        log_callback("警告：VariableID 列表不完整或长度不足，将使用参与变量的前 k 个作为分母（可能与 VBA 不一致）")
        selected_vars = active_vars[:k]

    # 计算这些变量的标准差（加权或未加权）
    if weight_series is not None:
        valid_mask = weight_series > 0
        w = weight_series[valid_mask]
        df_valid = df.loc[valid_mask]
        std_global = []
        for var in selected_vars:
            if var in df_valid.columns:
                X_var = df_valid[var].values
                weighted_mean = np.average(X_var, weights=w)
                variance = np.average((X_var - weighted_mean)**2, weights=w)
                s = np.sqrt(variance) if variance > 0 else 1e-12
            else:
                log_callback(f"警告：变量 {var} 不在数据中，分母将使用 1e-12")
                s = 1e-12
            std_global.append(s)
        std_global = np.array(std_global)
    else:
        std_global = []
        for var in selected_vars:
            if var in df.columns:
                s = df[var].std(ddof=1)
                if pd.isna(s) or s == 0:
                    s = 1e-12
            else:
                log_callback(f"警告：变量 {var} 不在数据中，分母将使用 1e-12")
                s = 1e-12
            std_global.append(s)
        std_global = np.array(std_global)

    # 若不足 k 个，补足
    if len(std_global) < k:
        log_callback(f"警告：可用的分母标准差不足 {k} 个，用 1e-12 补全")
        std_global = np.pad(std_global, (0, k - len(std_global)), constant_values=1e-12)

    # 计算非标准化得分系数
    score_coeff_unstd_spss = score_coeff_std / std[:, None]          # SPSS 语法：除以各自变量标准差
    score_coeff_unstd_table = score_coeff_std / std_global           # 表格：除以全局前 k 个标准差

    # ---- 9. 因子得分 ----
    X_all = df[active_vars].values.astype(float)
    factor_scores = X_all @ score_coeff_unstd_spss
    factor_cols = [f'Factor{i+1}' for i in range(k)]
    factor_df = pd.DataFrame(factor_scores, columns=factor_cols, index=df.index)
    for col in factor_cols:
        df[col] = factor_df[col]

    # ---- 10. 因子相关系数 ----
    factor_corr = np.corrcoef(factor_scores.T)

    # ---- 11. 排序 ----
    orderby = config.get('orderby', True)
    order_indices = get_max_component_order(rot_loadings, orderby)

    # ---- 12. SPSS 语法 ----
    syntax_lines = []
    for i in range(k):
        coeffs = score_coeff_unstd_spss[:, i]
        terms = []
        for var, coef in zip(active_vars, coeffs):
            if abs(coef) < 1e-10:
                continue
            sign = '+' if coef >= 0 else '-'
            abs_coef = abs(coef)
            terms.append(f"{sign}{abs_coef:.4f}*{var}")
        expr = ''.join(terms)
        if expr.startswith('+'):
            expr = expr[1:]
        syntax_lines.append(f"Compute Factor{i+1} = {expr}.")
    spss_syntax = '\n'.join(syntax_lines)

    # ---- 13. 输出 ----
    fa_extra = {
        'eigenvals': full_eigvals,
        'var_ratio': full_var_ratio,
        'cum_var': full_cum_var,
        'unrot_loadings': unrot_loadings,
        'rot_loadings': rot_loadings,
        'score_coeff_std': score_coeff_std,
        'score_coeff_unstd': score_coeff_unstd_spss,
        'score_coeff_unstd_table': score_coeff_unstd_table,
        'target': active_target,
        'order_indices': order_indices,
        'sig_break': config.get('sig_break', 0.3),
        'used_vars': active_vars,
        'n_factors': k,
        'corr_matrix': corr_matrix,
        'spss_syntax': spss_syntax,
        'rotation': rotation,
        'factor_var_labels': config.get('factor_var_labels', {}),
        'factor_var_ids': config.get('factor_var_ids', []),
        'var_id_map': dict(zip(config['factor_target'].index, config.get('factor_var_ids', []))),
    }

    return None, rot_loadings, score_coeff_unstd_spss, factor_df, df, factor_corr, fa_extra