import pyreadstat
import pandas as pd
import numpy as np

# 全局映射字典（小写清理后的列名 -> 实际列名）
_lower_to_col = {}

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

def read_spss(file_path, filter_var=None, filter_val=None, filter_type='numeric'):
    global _lower_to_col
    df, meta = pyreadstat.read_sav(file_path)
    name_label_map = get_name_label_mapping(meta)
    raw_total = meta.number_rows
    print(f"SPSS原始文件总样本量：{raw_total}")

    var_to_value_labels = build_var_to_value_labels(meta)

    if filter_var and filter_type == 'category' and filter_val is not None and filter_var in df.columns:
        value_labels = var_to_value_labels.get(filter_var, {})
        if value_labels:
            label_to_value = {str(v).strip(): k for k, v in value_labels.items()}
            def convert_val(v):
                try:
                    num = int(v) if v.isdigit() else float(v)
                    return num
                except ValueError:
                    pass
                if v in label_to_value:
                    return label_to_value[v]
                return v
            if isinstance(filter_val, list):
                filter_val = [convert_val(v) for v in filter_val]
            else:
                filter_val = convert_val(filter_val)
            print(f"转换后的筛选值：{filter_val}")
        else:
            print(f"警告：变量 '{filter_var}' 声明为分类变量，但无值标签，无法转换筛选值")

    if filter_var and filter_type == 'category' and filter_var in var_to_value_labels:
        print(f"变量 '{filter_var}' 的值标签：{var_to_value_labels[filter_var]}")
    elif filter_var and filter_type == 'category':
        print(f"变量 '{filter_var}' 没有值标签或未找到（但被声明为 category）")

    if filter_var and filter_val is not None and filter_var in df.columns:
        col_dtype = df[filter_var].dtype
        if isinstance(filter_val, list):
            converted_vals = []
            for v in filter_val:
                if isinstance(v, str):
                    try:
                        if 'int' in str(col_dtype):
                            v = int(v)
                        elif 'float' in str(col_dtype):
                            v = float(v)
                    except ValueError:
                        pass
                converted_vals.append(v)
            df = df[df[filter_var].isin(converted_vals)].copy()
            print(f"已过滤：保留 {filter_var} in {converted_vals}，过滤后剩余样本数 {len(df)}")
        else:
            if isinstance(filter_val, str):
                try:
                    if 'int' in str(col_dtype):
                        filter_val = int(filter_val)
                    elif 'float' in str(col_dtype):
                        filter_val = float(filter_val)
                except ValueError:
                    pass
            df = df[df[filter_var] == filter_val].copy()
            print(f"已过滤：保留 {filter_var}={filter_val}，过滤后剩余样本数 {len(df)}")
    else:
        if filter_var and filter_var not in df.columns:
            print(f"警告：筛选变量 '{filter_var}' 不在数据中，未进行筛选")
        elif filter_val is None:
            print("未设置筛选值，保留全部样本")

    # 初始化大小写映射（清理列名中的空格）
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