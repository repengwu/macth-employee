import os
import re
import math
import hashlib
import pandas as pd

# ===================== 配置区（请修改这里的路径） =====================
input_file = "/Users/bangongyi/Desktop/槽位提取/湖南槽位提取/23湖南槽位加密.xlsx"
output_file = "/Users/bangongyi/Desktop/槽位提取/湖南槽位提取/23湖南槽位加密-cs-对照组.xlsx"
sheet_name = "Sheet1"

# 如果你想继承旧版本已生成的“唯一key”，这里填旧加密结果文件路径。
# 默认写成 output_file：只要旧输出文件存在，脚本会先读取旧输出，未变更数据沿用旧唯一key。
# 如果不需要继承旧key，改成空字符串：history_file = ""
history_file = output_file

# 推荐手动指定真正稳定、能唯一定位一条数据的字段，例如：
# stable_key_columns = ["年份", "院校代码", "专业代码", "专业名称", "专业备注", "专业方向", "关键词"]
# 留空时：自动使用当前表中除“唯一key / 合并 / 序号 / Unnamed列”等之外的全部字段。
stable_key_columns = []

# 是否删除“合并”列，保持你原脚本行为：True = 输出时删除合并列；False = 保留合并列
drop_merge_column = True

key_column = "唯一key"
# =====================================================================

IGNORE_EXACT_COLUMNS = {key_column, "合并", "序号", "编号", "index", "Index"}
IGNORE_PREFIX_COLUMNS = ("Unnamed:",)
SEP = "\u241F"  # 字段分隔符，避免普通文本碰撞


def is_empty_value(value):
    """判断 Excel 空值。"""
    if value is None:
        return True
    try:
        return pd.isna(value)
    except Exception:
        return False


def normalize_value(value):
    """
    把单元格值标准化，避免 1 和 1.0、日期格式、空格差异导致同一数据生成不同 key。
    """
    if is_empty_value(value):
        return ""

    # 时间类型统一成 YYYY-MM-DD HH:MM:SS
    if isinstance(value, (pd.Timestamp,)):
        if pd.isna(value):
            return ""
        if value.time() == pd.Timestamp(value).time().replace(hour=0, minute=0, second=0, microsecond=0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")

    # 数字类型：1.0 统一成 1，避免 Excel 读入造成差异
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, ".15g")

    s = str(value).strip()
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s)
    return s


def md5_encrypt(text):
    """MD5 编码。注意：这是摘要/哈希，不是可逆加密。"""
    if text is None or text == "":
        return ""
    return hashlib.md5(str(text).encode("utf-8")).hexdigest()


def choose_stable_columns(df):
    """
    选择参与生成唯一key的字段。
    核心原则：只使用本行字段，不使用行号、排序、累计编号等会受其他行影响的内容。
    """
    if stable_key_columns:
        missing = [c for c in stable_key_columns if c not in df.columns]
        if missing:
            raise ValueError(f"stable_key_columns 中这些列在 Excel 里不存在：{missing}")
        return list(stable_key_columns)

    cols = []
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in IGNORE_EXACT_COLUMNS:
            continue
        if col_str.startswith(IGNORE_PREFIX_COLUMNS):
            continue
        cols.append(col)

    # 如果过滤后没有可用字段，再兜底使用“合并”列，避免全空。
    if not cols and "合并" in df.columns:
        cols = ["合并"]

    if not cols:
        raise ValueError("没有找到可用于生成唯一key的字段，请手动配置 stable_key_columns。")

    return cols


def build_row_signature(row, columns):
    """
    生成“本行指纹”。
    加入列名是为了避免 A=12,B=3 和 A=1,B=23 这种拼接碰撞。
    """
    parts = []
    for col in columns:
        parts.append(f"{col}={normalize_value(row.get(col, ''))}")
    return SEP.join(parts)


def build_history_map(history_path, columns):
    """
    从旧输出文件中读取历史 key。
    作用：即使你以前用的是旧算法，只要旧文件里未变更行的业务字段还在，
    本次就可以沿用旧唯一key，避免第一次切换稳定算法时整体变化。
    """
    if not history_path or not os.path.exists(history_path):
        return {}

    old_df = pd.read_excel(history_path, sheet_name=sheet_name)
    if key_column not in old_df.columns:
        print(f"⚠️ 历史文件中没有 {key_column} 列，跳过历史key复用：{history_path}")
        return {}

    missing = [c for c in columns if c not in old_df.columns]
    if missing:
        print(f"⚠️ 历史文件缺少这些参与生成key的列，无法安全复用历史key：{missing}")
        return {}

    history_map = {}
    for _, row in old_df.iterrows():
        old_key = normalize_value(row.get(key_column, ""))
        if not old_key:
            continue
        sig = build_row_signature(row, columns)
        # 同一条业务数据重复出现时，保持同样 key；如果历史里已有，则不覆盖。
        history_map.setdefault(sig, old_key)

    print(f"已读取历史key映射：{len(history_map)} 条")
    return history_map


def main():
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    source_columns = choose_stable_columns(df)
    print("参与生成唯一key的字段：", source_columns)

    history_map = build_history_map(history_file, source_columns)

    new_keys = []
    reused_count = 0
    generated_count = 0

    for _, row in df.iterrows():
        signature = build_row_signature(row, source_columns)
        if signature in history_map:
            new_keys.append(history_map[signature])
            reused_count += 1
        else:
            new_keys.append(md5_encrypt(signature))
            generated_count += 1

    df[key_column] = new_keys

    if drop_merge_column and "合并" in df.columns:
        df = df.drop(columns=["合并"])

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False, sheet_name=sheet_name)

    print("稳定版 MD5 编码完成！")
    print(f"复用历史key：{reused_count} 条")
    print(f"新生成key：{generated_count} 条")
    print(f"结果已保存至：{output_file}")


if __name__ == "__main__":
    main()
