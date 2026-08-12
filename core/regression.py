import pandas as pd
import numpy as np
import statsmodels.api as sm
from scipy import stats

def parse_factor_indicator(ind_list, n_factors):
    parsed = []
    for item in ind_list:
        if item.startswith('Factor'):
            try:
                num = int(item.replace('Factor', ''))
                if 1 <= num <= n_factors:
                    parsed.append(f'Factor{num}')
                else:
                    print(f"警告：因子编号 {num} 超出范围，忽略")
            except ValueError:
                print(f"警告：无法解析因子名 '{item}'，忽略")
        else:
            parsed.append(item)
    return parsed

def run_regression(df, dep_vars, ind_vars, weight_var=None, center=False):
    results = {}
    avail_ind = [v for v in ind_vars if v in df.columns]
    if not avail_ind:
        print("警告：没有可用的自变量，回归跳过")
        return results

    weights = None
    valid_weight_mask = pd.Series(True, index=df.index)
    if weight_var and weight_var in df.columns:
        weights = df[weight_var].copy()
        weights = weights.fillna(0)
        invalid_mask = (weights <= 0)
        if invalid_mask.any():
            print(f"警告：权重变量 '{weight_var}' 中有 {invalid_mask.sum()} 个非正值（≤0 或缺失），这些个案将被排除")
        valid_weight_mask = weights > 0
        weights.loc[~valid_weight_mask] = 0

    print("\n" + "="*60)
    print("【回归诊断】开始回归分析")
    print("="*60)
    print(f"因变量列表: {dep_vars}")
    print(f"自变量列表: {avail_ind}")
    print(f"权重变量: {weight_var if weight_var else '无'}")
    print(f"中心化: {'启用' if center else '关闭'}")
    print(f"数据框总样本数: {len(df)}")
    print(f"有效权重样本数（>0）: {valid_weight_mask.sum() if weight_var else len(df)}")

    for dep in dep_vars:
        print(f"\n--- 处理因变量: {dep} ---")
        if dep not in df.columns:
            print(f"警告：因变量 {dep} 不在数据中，跳过")
            continue

        y = df[dep]
        X = df[avail_ind]

        if center:
            X = X.copy()
            for col in X.columns:
                X[col] = X[col] - X[col].mean()
            print("  已对自变量进行中心化")

        mask = y.notna() & X.notna().all(axis=1) & valid_weight_mask
        used_n = mask.sum()
        if used_n == 0:
            print(f"警告：因变量 {dep} 无有效样本（含权重>0），跳过")
            continue

        y_clean = y[mask]
        X_clean = X.loc[mask]
        weights_clean = weights[mask] if weights is not None else None

        X_design = sm.add_constant(X_clean)

        if weights_clean is not None:
            model = sm.WLS(y_clean, X_design, weights=weights_clean).fit()
        else:
            model = sm.OLS(y_clean, X_design).fit()

        params = model.params
        p = len(params)

        if weights_clean is not None:
            w = weights_clean
            sum_w = w.sum()
            resid = model.resid
            rss = np.sum(w * (resid ** 2))
        else:
            w = np.ones(len(y_clean))
            sum_w = len(y_clean)
            resid = model.resid
            rss = np.sum(resid ** 2)

        df_resid = sum_w - p
        if df_resid > 0:
            sigma2 = rss / df_resid
            # 使用纯 NumPy 数组避免 pandas 索引对齐问题
            X_mat = X_design.values
            w_vals = w.values if hasattr(w, 'values') else w
            XWX = X_mat.T @ np.diag(w_vals) @ X_mat
            try:
                inv_XWX = np.linalg.inv(XWX)
            except np.linalg.LinAlgError:
                inv_XWX = np.linalg.pinv(XWX)
            cov_beta = sigma2 * inv_XWX
            se = np.sqrt(np.diag(cov_beta))
            tvalues = params / se
            pvalues = 2 * (1 - stats.t.cdf(np.abs(tvalues), df_resid))
            ci_lower = params - stats.t.ppf(0.975, df_resid) * se
            ci_upper = params + stats.t.ppf(0.975, df_resid) * se
        else:
            se = np.full_like(params, np.nan)
            tvalues = np.full_like(params, np.nan)
            pvalues = np.full_like(params, np.nan)
            ci_lower = np.full_like(params, np.nan)
            ci_upper = np.full_like(params, np.nan)

        # 标准化系数 Beta
        beta = pd.Series(index=params.index, dtype=float)
        has_const = 'const' in params.index
        vars_no_const = [v for v in params.index if v != 'const'] if has_const else list(params.index)
        if weights_clean is not None:
            y_std = np.sqrt(np.average((y_clean - y_clean.mean())**2, weights=w))
            for v in vars_no_const:
                x_std = np.sqrt(np.average((X_design[v] - X_design[v].mean())**2, weights=w))
                if x_std != 0:
                    beta[v] = params[v] * (x_std / y_std)
                else:
                    beta[v] = np.nan
        else:
            y_std = y_clean.std(ddof=0)
            for v in vars_no_const:
                x_std = X_design[v].std(ddof=0)
                if x_std != 0:
                    beta[v] = params[v] * (x_std / y_std)
                else:
                    beta[v] = np.nan
        if has_const:
            beta['const'] = np.nan

        results[dep] = {
            'R_squared': model.rsquared,
            'adj_R_squared': model.rsquared_adj,
            'R': np.sqrt(model.rsquared),
            'std_error': np.sqrt(sigma2) if df_resid > 0 else np.nan,
            'fvalue': model.fvalue,
            'f_pvalue': model.f_pvalue,
            'nobs': sum_w,
            'coeff': params,
            'bse': pd.Series(se, index=params.index),
            'tvalues': pd.Series(tvalues, index=params.index),
            'pvalues': pd.Series(pvalues, index=params.index),
            'ci_lower': pd.Series(ci_lower, index=params.index),
            'ci_upper': pd.Series(ci_upper, index=params.index),
            'beta': beta,
            'var_names': list(params.index)
        }

        print(f"\n回归结果 (因变量 {dep}):")
        print(f"  R² = {model.rsquared:.4f}, 调整R² = {model.rsquared_adj:.4f}")
        print(f"  加权有效样本数 (总权重) = {sum_w:.2f}")
        print("  系数:")
        for name, val in params.items():
            se_val = se[0] if isinstance(se, np.ndarray) else se.loc[name]
            print(f"    {name}: {val:.6f} (标准误 {se_val:.6f})")
        print("-"*60)

    return results