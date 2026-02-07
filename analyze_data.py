# -*- coding: utf-8 -*-
import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

def analyze_reason(row):
    """根据数据特征智能分析原因"""
    customer = row['客商名称'] if pd.notna(row['客商名称']) else ''
    fin2024 = row['2024_财务金额'] if pd.notna(row['2024_财务金额']) else 0
    fin2025 = row['2025_财务金额'] if pd.notna(row['2025_财务金额']) else 0
    fin_diff = row['财务差异'] if pd.notna(row['财务差异']) else 0
    hr2024 = row['2024_人力金额'] if pd.notna(row['2024_人力金额']) else 0
    hr2025 = row['2025_人力金额'] if pd.notna(row['2025_人力金额']) else 0
    hr_diff = row['人力差异'] if pd.notna(row['人力差异']) else 0

    # 获取最大绝对金额用于判断规模
    max_amount = max(abs(fin2024), abs(fin2025), abs(hr2024), abs(hr2025))
    max_diff = max(abs(fin_diff), abs(hr_diff))

    # 1. 人力转财务判断（2025人力为0，2025财务有值）
    if hr2025 == 0 and hr2024 > 0 and fin2025 > 0:
        if max_diff > 5000:
            return "人力转财务统计，大额需核实"
        elif "人力" in customer or "劳务" in customer or "派遣" in customer:
            return "人力派遣转财务核算"
        else:
            return "人力转财务统计"

    # 2. 财务转人力判断（2025财务为0，2025人力有值）
    if fin2025 == 0 and fin2024 > 0 and hr2025 > 0:
        return "财务转人力统计"

    # 3. 新增财务供应商（2024财务为0，2025财务有值）
    if fin2024 == 0 and fin2025 > 0:
        if max_diff > 10000:
            return "新增财务供应商，大额变化"
        elif hr2025 == 0:
            return "新增财务录入，人力归零"
        else:
            return "新增财务供应商"

    # 4. 财务归零（2024财务有值，2025财务为0）
    if fin2024 > 0 and fin2025 == 0:
        if hr2025 > hr2024:
            return "财务转人力核算"
        else:
            return "停止财务统计"

    # 5. 大额变化判断（财务或人力差异>10000）
    if max_diff > 10000:
        if fin_diff > 0 and hr_diff < 0:
            return "大额变化，核算口径调整"
        elif fin_diff < 0 and hr_diff > 0:
            return "大额变化，跨期核算差异"
        else:
            return "大额变化，需重点核实"

    # 6. 极小金额（金额<10）
    if max_amount < 10:
        return "金额极小，数据核对"

    # 7. 小额金额（金额10-100）
    if max_amount < 100:
        if fin_diff > 0 and hr_diff < 0:
            return "金额较小，口径调整"
        else:
            return "金额较小，关注变化"

    # 8. 根据客商名称关键词判断
    if "人力" in customer or "劳务" in customer or "派遣" in customer:
        if hr_diff < 0 and fin_diff > 0:
            return "劳务派遣类，人力转财务"
        else:
            return "劳务派遣类，统计变化"

    if "保安" in customer:
        return "保安服务类，口径调整"

    if "运输" in customer or "物流" in customer:
        if fin_diff > 0 and hr_diff < 0:
            return "运输服务类，核算调整"
        else:
            return "运输服务类，统计变化"

    if "耐火" in customer or "材料" in customer:
        if fin_diff > 0 and hr_diff < 0:
            return "耐火材料类，核算调整"
        else:
            return "物资供应商，统计调整"

    if "环保" in customer or "科技" in customer or "技术" in customer:
        return "技术服务类，口径调整"

    if "凌源" in customer or "北票" in customer or "朝阳" in customer:
        return "本地关联企业，口径调整"

    if "钢达" in customer or "钢城" in customer or "钢联" in customer or "双鞍" in customer:
        if max_diff > 5000:
            return "集团关联企业，大额调整"
        else:
            return "关联单位，核算调整"

    # 9. 变化组合判断
    if fin_diff > 0 and hr_diff < 0:
        if abs(hr_diff) > abs(fin_diff):
            return "财务增人力减，口径调整"
        else:
            return "调整核算口径"

    if fin_diff < 0 and hr_diff > 0:
        return "跨期核算，时点差异"

    if fin_diff > 0 and hr_diff > 0:
        return "统计口径变更"

    if fin_diff < 0 and hr_diff < 0:
        return "业务规模调整"

    # 10. 默认原因
    if max_diff > 1000:
        return "核算口径调整"
    else:
        return "统计口径调整"

# 读取Excel
file_path = r'C:\Users\jintao\Desktop\123\异常数据AI分析.xlsx'
df = pd.read_excel(file_path)

print(f"开始处理 {len(df)} 条记录...")

# 逐条分析并填写原因
for idx, row in df.iterrows():
    reason = analyze_reason(row)
    df.at[idx, '原因分析'] = reason
    print(f"{idx}: {row['客商名称']} - {reason}")

# 保存回原文件
df.to_excel(file_path, index=False)
print(f"\n完成！已更新 {len(df)} 条记录到原文件")
