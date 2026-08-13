import pandas as pd
import re

def load_definition(def_path):
    # 读取 Config 表（前3列，用于竖排扫描和类型获取）
    config_raw = pd.read_excel(def_path, sheet_name='Config', header=None, usecols=[0, 1, 2])
    # 读取 Config 表（前2列，用于键值对读取）
    config_df = pd.read_excel(def_path, sheet_name='Config', header=None, index_col=0)
    params = config_df[1].to_dict()

    def parse_list(s):
        if isinstance(s, str):
            items = re.split(r',\s*', s)
            return [x.strip() for x in items if x.strip()]
        return []

    # ==================== 核心参数 ====================
    config = {
        'y_vars': parse_list(params.get('Y_Var', '')),
        'x_vars': parse_list(params.get('X_Var', '')),
        'weight_var': params.get('Weight', ''),
        'filter_var': params.get('Filter_Var', ''),
        'filter_val': params.get('Filter_Value', None),
        'filter_type': 'numeric',
        'filter_val_brand': False,
        'factor_rotation': params.get('Factor_Rotation', 'procrustes'),
        'sample_channel_keyword': params.get('Channel_Keyword', ''),
        'sample_top2box_label': parse_list(params.get('Top2Box_Label', 'Net : Top 2 Box')),
        'y_var_recode': True,
        'x_var_merge': False,
        'merge_configs': [],
        'merged_x_vars': [],
    }

    # ---- 从 config_raw 中扫描 C 列参数，并收集 filter 信息 ----
    filter_info = {}  # reg_num -> {'expr': 'var=val', 'type': 'cat/num'}

    for idx, row in config_raw.iterrows():
        a_col = str(row.iloc[0]).strip().replace('\n', '').replace('\r', '') if pd.notna(row.iloc[0]) else ''
        b_col = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        c_val = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''

        # ---- Filter_Var 行的 C 列：filter_type ----
        if a_col == 'Filter_Var':
            if c_val.lower() in ['category', '分类']:
                config['filter_type'] = 'category'
            else:
                config['filter_type'] = 'numeric'

        # ---- Filter_Value 行的 C 列：是否品牌筛选 ----
        if a_col == 'Filter_Value':
            if c_val.lower() == 'brand':
                config['filter_val_brand'] = True
            else:
                config['filter_val_brand'] = False

        # ---- Y_Var 行的 C 列：是否 recode Y ----
        if a_col == 'Y_Var':
            if c_val.lower() == 'recode':
                config['y_var_recode'] = True
            else:
                config['y_var_recode'] = False

        # ---- X_Var 行的 C 列：是否 merge X ----
        if a_col == 'X_Var':
            if c_val.lower() == 'merge':
                config['x_var_merge'] = True
            else:
                config['x_var_merge'] = False

        # ---- 识别 regX_filter 行，收集表达式和类型 ----
        if a_col.startswith('reg') and a_col.endswith('_filter'):
            reg_num = a_col.replace('_filter', '')
            filter_info[reg_num] = {'expr': b_col, 'type': c_val}

    print(f"Y_Var recode: {config['y_var_recode']}")
    print(f"X_Var merge: {config['x_var_merge']}")
    print(f"Filter_Value brand: {config['filter_val_brand']}")

    # ---- 读取 Merge sheet（仅在 x_var_merge=True 时） ----
    if config['x_var_merge']:
        try:
            merge_raw = pd.read_excel(def_path, sheet_name='Merge', header=None, skiprows=1, usecols=[0, 1, 2])
            merge_configs = []
            merged_names = []
            for idx, row in merge_raw.iterrows():
                new_name = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
                orig_list_str = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
                if new_name and orig_list_str:
                    orig_vars = [v.strip() for v in re.split(r',\s*', orig_list_str) if v.strip()]
                    merge_configs.append({
                        'new_name': new_name,
                        'orig_vars': orig_vars,
                    })
                    merged_names.append(new_name)
                    print(f"Merge 配置: {new_name} = IF ({' or '.join(f'{v}=1' for v in orig_vars)}) {new_name}=1")
            config['merge_configs'] = merge_configs
            config['merged_x_vars'] = merged_names
            print(f"Merge sheet 读取完成，共 {len(merge_configs)} 个合并规则")
        except Exception as e:
            print(f"读取 Merge 工作表异常：{str(e)}，将忽略 merge 操作")
            config['x_var_merge'] = False
            config['merge_configs'] = []
            config['merged_x_vars'] = []

    # ---- 动态扫描所有回归配置 ----
    reg_configs = []
    for key, value in params.items():
        str_key = str(key) if key is not None else ''
        if str_key.startswith('reg') and str_key.endswith('_dep'):
            reg_num = str_key.replace('_dep', '')
            ind_key = f'{reg_num}_ind'
            if ind_key in params:
                reg_cfg = {
                    'dep': parse_list(value),
                    'ind': parse_list(params.get(ind_key, ''))
                }
                # 如果有对应的 filter 信息，加入
                if reg_num in filter_info:
                    reg_cfg['filter'] = filter_info[reg_num]
                reg_configs.append(reg_cfg)
    config['reg_configs'] = reg_configs

    # ---- 特殊处理：多行竖排的 Sample_KPI_Keywords ----
    kpi_keywords = []
    kpi_types = []
    kpi_start_row = None

    print("===== 正在扫描Config表查找KPI_Keywords标题 =====")
    for idx, row in config_raw.iterrows():
        a_col = str(row.iloc[0]).strip().replace('\n', '').replace('\r', '') if pd.notna(row.iloc[0]) else ''
        b_col = str(row.iloc[1]).strip().replace('\n', '').replace('\r', '') if pd.notna(row.iloc[1]) else ''
        print(f"行{idx+1} | A列: {a_col} | B列: {b_col}")
        if 'KPI_Keywords' in a_col or 'KPI_Keywords' in b_col:
            kpi_start_row = idx
            print(f"✅ 找到KPI_Keywords标题行，行号：{idx+1}")
            break

    if kpi_start_row is not None:
        print("===== 开始读取KPI关键词数据 =====")
        for idx in range(kpi_start_row + 1, len(config_raw)):
            row = config_raw.iloc[idx]
            b_col_val = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            c_col_val = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
            if b_col_val == '':
                print(f"行{idx+1} B列为空，KPI数据读取结束")
                break
            kpi_keywords.append(b_col_val)
            kpi_type = c_col_val if c_col_val != '' else 'summary'
            kpi_types.append(kpi_type)
            print(f"行{idx+1} | 读取到KPI: {b_col_val} | 类型: {kpi_type}")

    if not kpi_keywords:
        raise ValueError(
            "❌ 未读取到KPI_Keywords的竖排数据！\n"
            "可能原因：\n"
            "1. 标题行名称错误，需严格包含'KPI_Keywords'；\n"
            "2. 标题行下方的B列没有有效KPI关键词；\n"
            "3. 表格列顺序错误（需A列标题、B列关键词、C列类型）"
        )

    config['sample_kpi_keywords'] = kpi_keywords
    config['sample_kpi_types'] = kpi_types
    print(f"✅ KPI数据读取完成，共读取{len(kpi_keywords)}个KPI")

    # ---- Factor_Target 读取 ----
    try:
        target_raw = pd.read_excel(
            def_path,
            sheet_name='Factor_Target',
            header=1,
            usecols=None
        )

        # ========== 🛡️ 通用防御性清理（针对整个 Sheet） ==========
        # 1. 删除所有列全为空的行（解决 Excel 末尾多余空行）
        target_raw = target_raw.dropna(how='all')
        
        # 2. 强制删除前三列（B列 VariableID, C列 Name, D列 Lable）中任意一列为空的行
        #    因为这三列缺一不可（索引0是A列可能空，所以用 1,2,3）
        target_raw = target_raw.dropna(subset=[target_raw.columns[1], target_raw.columns[2], target_raw.columns[3]])
        
        # 3. 删除 D 列（Lable）为纯空格或空字符串的行（解决看不见的空格）
        target_raw = target_raw[target_raw.iloc[:, 3].astype(str).str.strip() != '']
        # ==========================================================

        if target_raw.shape[1] < 4:   # 至少要有 A~D 四列（虽然A可能空）
            raise ValueError("Factor_Target 表至少需要4列（A列可能空，B=VariableID, C=Name, D=Lable）")
        
        factor_cols = [col for col in target_raw.columns if col.startswith('Factor')]
        if len(factor_cols) == 0:
            raise ValueError("Factor_Target工作表不存在Factor开头的列")
        
        factor_df = target_raw[factor_cols].copy()
        factor_df = factor_df.fillna(0)
        factor_df = factor_df.apply(lambda col: col.map(lambda x: 1 if x == 1 else 0))
        
        # 行索引设为 C 列
        factor_df.index = target_raw.iloc[:, 2].tolist()
        print("调试：读取到的 Name 列内容如下：", target_raw.iloc[:, 2].tolist())
        
        config['factor_target'] = factor_df
        config['extract_factor_num'] = len(factor_cols)
        
        # 变量标签：键 = C列（Name），值 = D列（Lable），供报表显示用
        # 如果 D 列有空值，可以用 C 列作为后备
        labels = {}
        for name, lable in zip(target_raw.iloc[:, 2], target_raw.iloc[:, 3]):
            labels[name] = lable if pd.notna(lable) and str(lable).strip() != '' else name
        config['factor_var_labels'] = labels
        
        # 变量 ID：用 B 列（VariableID）
        config['factor_var_ids'] = target_raw.iloc[:, 1].tolist()
        
        print(f"✅ Factor_Target读取完成，共{len(factor_cols)}个因子，{len(factor_df)}个变量")
    except Exception as e:
        print(f"读取Factor_Target工作表异常：{str(e)}")
        config['factor_target'] = None
        config['extract_factor_num'] = None
        config['factor_var_labels'] = {}
        config['factor_var_ids'] = []

    return config