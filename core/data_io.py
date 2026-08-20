import pyreadstat
import pandas as pd
import numpy as np
import re

# 全局映射字典（小写清理后的列名 -> 实际列名）
_lower_to_col = {}


def parse_filter_expr(expr, df, var_to_value_labels, log_callback=print):
    """
    解析筛选表达式，返回过滤后的 DataFrame。
    表达式示例：'wave=3,4; filter=5,6' 或 'wave=2 & filter=4,5'
    支持多个条件 AND 关系。
    """
    if not expr or not expr.strip():
        return df

    # 分割条件：优先 ; 或 &，若无但多个等号则按等号分割
    parts = None
    if ';' in expr or '&' in expr:
        parts = re.split(r'[;&]', expr)
    elif expr.count('=') >= 2:
        # 按等号位置分割，但需要保留变量名和值，这里简单处理：先按等号分割，再组合
        # 更可靠：用正则匹配所有 var=values 模式
        parts = re.findall(r'([^=;&]+)=([^=;&]+)', expr)
        if parts:
            parts = [f"{var}={vals}" for var, vals in parts]
    else:
        # 单条件
        parts = [expr]

    conditions = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '=' not in part:
            log_callback(f"警告：筛选条件 '{part}' 缺少 '='，跳过")
            continue
        var, vals_str = part.split('=', 1)
        var = var.strip()
        vals = [v.strip() for v in vals_str.split(',') if v.strip()]
        if not vals:
            log_callback(f"警告：变量 '{var}' 无有效值，跳过")
            continue
        conditions.append((var, vals))

    if not conditions:
        return df

    # 依次应用每个条件（AND）
    for var, vals in conditions:
        if var not in df.columns:
            log_callback(f"警告：变量 '{var}' 不在数据中，跳过该条件")
            continue
        col_dtype = df[var].dtype
        converted_vals = []

        # 若有值标签，先尝试标签匹配
        if var in var_to_value_labels:
            label_to_value = {str(v).strip(): k for k, v in var_to_value_labels[var].items()}
            for v in vals:
                if v in label_to_value:
                    converted_vals.append(label_to_value[v])
                else:
                    # 尝试转为数值
                    try:
                        if 'int' in str(col_dtype):
                            converted_vals.append(int(v))
                        elif 'float' in str(col_dtype):
                            converted_vals.append(float(v))
                        else:
                            converted_vals.append(v)
                    except ValueError:
                        converted_vals.append(v)
        else:
            # 无值标签，直接转换类型
            for v in vals:
                try:
                    if 'int' in str(col_dtype):
                        converted_vals.append(int(v))
                    elif 'float' in str(col_dtype):
                        converted_vals.append(float(v))
                    else:
                        converted_vals.append(v)
                except ValueError:
                    converted_vals.append(v)

        converted_vals = list(dict.fromkeys(converted_vals))  # 去重
        df = df[df[var].isin(converted_vals)].copy()
        log_callback(f"已过滤：保留 {var} in {converted_vals}，剩余样本数 {len(df)}")

    return df

def get_name_label_mapping(meta):
    name_list = meta.column_names
    label_list = meta.column_labels
    name_label_dict = {}
    for name, label in zip(name_list, label_list):
        label = label if (label is not None and str(label).strip()) else name
        name_label_dict[name] = label
    return name_label_dict

def build_var_to_value_labels(meta):
    if hasattr(meta, 'variable_value_labels') and meta.variable_value_labels:
        return meta.variable_value_labels
    var_to_vl = {}
    if hasattr(meta, 'column_index') and hasattr(meta, 'value_labels'):
        value_labels = meta.value_labels
        keys = list(value_labels.keys())
        if keys and all(k.startswith('labels') and k[6:].isdigit() for k in keys):
            for var, idx in meta.column_index.items():
                key1 = f'labels{idx}'
                key2 = f'labels{idx+1}'
                if key1 in value_labels:
                    var_to_vl[var] = value_labels[key1]
                elif key2 in value_labels:
                    var_to_vl[var] = value_labels[key2]
        else:
            sorted_keys = sorted(value_labels.keys())
            column_names = meta.column_names
            if len(sorted_keys) == len(column_names):
                for idx, var in enumerate(column_names):
                    var_to_vl[var] = value_labels[sorted_keys[idx]]
    return var_to_vl

def read_spss(file_path, filter_expr=None):
    """
    读取 SPSS .sav 文件，自动尝试多种编码，避免因编码问题导致读取失败。
    支持常见的 Unicode、西欧、中文（GBK/Big5）等编码。
    """
    global _lower_to_col

    # ---------- 扩展编码列表 ----------
    # 顺序按常用性和兼容性排列：优先 UTF-8 系列，其次西欧编码，再次中文编码
    encodings = [
        'utf-8',           # 标准 UTF-8
        'utf-8-sig',       # 带 BOM 的 UTF-8
        'latin1',          # ISO-8859-1（西欧）
        'cp1252',          # Windows 西欧（与 latin1 相似但稍有不同）
        'cp437',           # DOS 拉丁字母（一些旧 SPSS 文件使用）
        'cp850',           # 西欧（DOS）
        'gbk',             # 简体中文（GB2312/GBK）
        'gb2312',          # 简体中文
        'big5',            # 繁体中文
        'cp936',           # 简体中文（Windows）
        'cp950',           # 繁体中文（Windows）
        'shift-jis',       # 日文
        'euc-kr',          # 韩文
    ]

    df = None
    meta = None
    last_exception = None

    for enc in encodings:
        try:
            df, meta = pyreadstat.read_sav(file_path, encoding=enc)
            print(f"✅ SPSS 文件成功使用编码 '{enc}' 读取")
            break
        except UnicodeDecodeError as e:
            last_exception = e
            print(f"编码 '{enc}' 失败，尝试下一个...")
            continue
        except Exception as e:
            # 捕获其他异常（如文件损坏、权限等），记录下来但继续尝试其他编码
            last_exception = e
            print(f"使用编码 '{enc}' 时发生非编码异常: {e}，尝试下一个...")
            continue

    if df is None:
        # 所有编码尝试均失败，抛出明确异常
        raise UnicodeDecodeError(
            f"无法使用任何已知编码读取 SPSS 文件。最后错误：{last_exception}"
        )

    name_label_map = get_name_label_mapping(meta)
    var_to_value_labels = build_var_to_value_labels(meta)

    raw_total = meta.number_rows
    print(f"SPSS原始文件总样本量：{raw_total}")

    # 若提供筛选表达式，进行过滤
    if filter_expr and isinstance(filter_expr, str) and filter_expr.strip():
        df = parse_filter_expr(filter_expr, df, var_to_value_labels, log_callback=print)

    # 更新全局大小写映射
    _lower_to_col = {col.strip().lower(): col for col in df.columns}
    print(f"✅ 数据列名（前10个）: {df.columns[:10].tolist()}")
    return df, meta, name_label_map, var_to_value_labels

def recode_binary(series):
    return series.fillna(0).apply(lambda x: 1 if x == 1 else 0)

def recode_variables(df, var_list):
    """
    重编码变量：将原始变量转换为二分类的 d 前缀变量。
    使用全局 _lower_to_col 进行大小写不敏感匹配，并更新映射。
    注意：所有变量统一加 'd' 前缀（无特殊处理）。
    """
    global _lower_to_col
    for var in var_list:
        if not var:
            continue
        var_clean = var.strip()
        # 查找实际列名（大小写不敏感）
        actual = None
        if var_clean in df.columns:
            actual = var_clean
        else:
            lower_var = var_clean.lower()
            if lower_var in _lower_to_col:
                actual = _lower_to_col[lower_var]
        
        if actual is None:
            print(f"警告：变量 '{var_clean}' 不在数据中，跳过重编码")
            continue
        
        # 生成新变量名：统一加 'd' 前缀（若已是 d 开头则不再添加）
        if actual.startswith('d'):
            new_name = actual
        else:
            new_name = f'd{actual}'
        
        # 执行重编码（转为 0/1）
        df[new_name] = recode_binary(df[actual])
        # 更新映射
        _lower_to_col[new_name.lower()] = new_name
        print(f"✅ 重编码成功：{actual} -> {new_name}")
    
    return df

def merge_variables(df, merge_configs, log_callback=print):
    """
    根据 merge_configs 合并原始变量为新变量。
    逻辑：IF (var1=1 or var2=1 or ...) NewVar=1
    使用大小写不敏感匹配，并更新映射。
    """
    global _lower_to_col
    for cfg in merge_configs:
        new_name = cfg['new_name']
        orig_vars = cfg['orig_vars']
        # 检查哪些原始变量存在于数据中（大小写不敏感，去除空格）
        exist_vars = []
        for v in orig_vars:
            v_clean = v.strip()
            if v_clean in df.columns:
                exist_vars.append(v_clean)
            else:
                lower_v = v_clean.lower()
                if lower_v in _lower_to_col:
                    exist_vars.append(_lower_to_col[lower_v])
        if not exist_vars:
            log_callback(f"警告：合并变量 {new_name} 的所有原始变量都不在数据中，跳过")
            continue
        missing = [v.strip() for v in orig_vars if v.strip() not in exist_vars and v.strip().lower() not in _lower_to_col]
        if missing:
            log_callback(f"警告：合并变量 {new_name} 缺少以下原始变量: {missing}")
        # 合成：任意一个原始变量为1，则新变量为1
        merged = pd.Series(0, index=df.index)
        for v in exist_vars:
            merged = merged | (df[v] == 1)
        df[new_name] = merged.astype(int)
        # 更新映射
        _lower_to_col[new_name.lower()] = new_name
        log_callback(f"合并变量 {new_name} = IF ({' or '.join(f'{v}=1' for v in exist_vars)}) {new_name}=1")
    return df

def prepare_weight(df, weight_var, log_callback=print):
    if not weight_var or weight_var not in df.columns:
        return None, None

    weights = df[weight_var].copy()
    weights = weights.fillna(0)
    invalid_mask = (weights <= 0)
    if invalid_mask.any():
        log_callback(f"警告：权重变量 '{weight_var}' 中有 {invalid_mask.sum()} 个非正值（≤0 或缺失），这些个案将被排除")
    valid_mask = weights > 0
    weights.loc[~valid_mask] = 0
    return weights, valid_mask