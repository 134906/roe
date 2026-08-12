# -*- coding: utf-8 -*-
from .config import load_definition
from .data_io import read_spss, recode_variables, merge_variables, prepare_weight, _lower_to_col
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

    filter_var = config.get('filter_var', '')
    filter_type = config.get('filter_type', 'numeric')
    filter_val_raw = config.get('filter_val', None)
    if filter_val_raw is not None:
        filter_val_str = str(filter_val_raw).strip()
        if filter_val_str:
            filter_values = [v.strip() for v in filter_val_str.split(',') if v.strip()]
        else:
            filter_values = []
    else:
        filter_values = []
    log_callback(f"筛选参数：变量='{filter_var}'，值={filter_values}，类型='{filter_type}'")

    # 第2步：读取SPSS数据
    log_callback("读取SPSS数据...")
    filter_val_for_spss = filter_values if len(filter_values) > 1 else (filter_values[0] if filter_values else None)
    df, meta, name_label_map, var_to_value_labels = read_spss(spss_path, filter_var, filter_val_for_spss, filter_type)
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
    if config.get('filter_var'):
        config['filter_var'] = resolve_var(config['filter_var']) or config['filter_var']

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
    weight_series, weight_valid_mask = prepare_weight(df, weight_var, log_callback)
    if weight_series is not None:
        log_callback("权重变量已使用")

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

    # 第9步：样本量核对（基于全数据）
    kpi_check_df = None
    channel_check_df = None
    success = False
    try:
        log_callback("开始样本量核对...")
        full_table_df = pd.read_excel(def_path, sheet_name='Full_Table', header=None)
        if full_table_df.empty:
            log_callback("Full_Table 为空，跳过样本量检查")
        else:
            brands, channel_df, kpi_df = extract_samples_from_full_table(
                full_table_df, config, name_label_map,
                filter_values=filter_values,
                var_to_value_labels=var_to_value_labels,
                filter_type=filter_type,
                log_callback=log_callback
            )
            x_vars_for_check = config['x_vars'] if x_var_merge else quasi_orig_x
            kpi_check_df, channel_check_df, success = perform_sample_check(
                desc_raw_y, desc_d_y, desc_raw_x, channel_df, kpi_df, name_label_map,
                config['y_vars'], x_vars_for_check, y_var_recode=y_var_recode
            )
            if success:
                log_callback("样本量校对成功")
            else:
                log_callback("样本量校对失败，请检查差异")
    except Exception as e:
        log_callback(f"样本量核对过程中发生错误: {e}")
        kpi_check_df = None
        channel_check_df = None
        success = False

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
            expr = filter_cfg['expr']
            if '=' in expr:
                var_part, val_part = expr.split('=', 1)
                filter_var = var_part.strip()
                filter_vals = [v.strip() for v in val_part.split(',') if v.strip()]
                # 解析筛选变量名（不区分大小写）
                filter_var_resolved = resolve_var(filter_var)
                if filter_var_resolved not in df.columns:
                    log_callback(f"警告：回归 {idx} 的筛选变量 '{filter_var}' 不存在，将使用全部数据")
                    df_reg = df
                    filter_desc = None
                else:
                    # 尝试转换值类型
                    col_dtype = df[filter_var_resolved].dtype
                    converted_vals = []
                    for val in filter_vals:
                        try:
                            if 'int' in str(col_dtype):
                                val = int(val)
                            elif 'float' in str(col_dtype):
                                val = float(val)
                        except ValueError:
                            pass
                        converted_vals.append(val)
                    # 应用筛选
                    try:
                        if len(converted_vals) == 1:
                            df_reg = df[df[filter_var_resolved] == converted_vals[0]].copy()
                            filter_desc = f"{filter_var_resolved} == {converted_vals[0]}"
                        else:
                            df_reg = df[df[filter_var_resolved].isin(converted_vals)].copy()
                            filter_desc = f"{filter_var_resolved} in {converted_vals}"
                        log_callback(f"回归 {idx} 应用筛选：{filter_desc}，样本数：{len(df_reg)}")
                    except Exception as e:
                        log_callback(f"警告：回归 {idx} 筛选失败：{e}，将使用全部数据")
                        df_reg = df
                        filter_desc = None
            else:
                log_callback(f"警告：回归 {idx} 的筛选表达式 '{expr}' 格式无效，将使用全部数据")
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
        has_filtered_reg=has_filtered_reg
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
                has_filtered_reg=has_filtered_reg
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