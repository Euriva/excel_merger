# --------------------------
# 通用Excel合并工具 - 网页版
# 功能：手动输入路径、匹配规则、输出名、保存位置
# --------------------------
import streamlit as st
import pandas as pd
import os
import re

# --------------------------
# 1. 网页标题与说明
# --------------------------
st.title("📊 Excel 文件合并工具")
st.write("自动批量合并符合规则的 Excel 文件，无需安装 Python")
st.divider()  # 分割线

# --------------------------
# 2. 让用户手动输入参数
# --------------------------
st.subheader("🔧 设置合并参数")

# 输入1：要扫描的文件夹路径
wechat_folder = st.text_input(
    "1. 输入要扫描的文件夹路径",
    value=r"C:\Users\tangyy\xwechat_files\wxid_73xovvobswzt22_90df\msg\file",
    help="复制文件夹地址栏路径粘贴进来"
)

# 输入2：文件名匹配规则（支持通配符，自动转正则）
file_pattern = st.text_input(
    "2. 要合并的文件名规则（支持 * 通配符）",
    value="社区拉新结算*.xlsx",
    help="例如：社区拉新结算*.xlsx  表示匹配所有以此开头的Excel文件"
)

# 输入3：输出文件名
output_filename = st.text_input(
    "3. 合并后的文件名",
    value="社区拉新结算汇总表.xlsx"
)

# 输入4：输出保存路径
output_folder = st.text_input(
    "4. 合并后文件保存路径",
    value=os.path.join(os.path.expanduser("~"), "Desktop"),
    help="默认保存到桌面，可自行修改"
)

st.divider()

# --------------------------
# 3. 把用户输入的通配符转为正则表达式（通用匹配）
# --------------------------
def wildcard_to_regex(pattern):
    """把 *.xlsx 这种通配符转为正则表达式"""
    regex = re.escape(pattern)
    regex = regex.replace(r"\*", ".*")  # * 匹配任意字符
    regex = "^" + regex + "$"          # 完全匹配文件名
    return regex

try:
    match_regex = wildcard_to_regex(file_pattern)
    st.success(f"✅ 匹配规则：{match_regex}")
except Exception as e:
    st.error(f"规则错误：{e}")


# --------------------------
# 4. 开始合并按钮
# --------------------------
if st.button("🚀 开始合并文件"):

    # 校验路径是否存在
    if not os.path.isdir(wechat_folder):
        st.error("❌ 文件夹路径不存在，请检查！")
    elif not output_filename.endswith(".xlsx"):
        st.error("❌ 输出文件名必须以 .xlsx 结尾")
    else:
        output_path = os.path.join(output_folder, output_filename)

        # --------------------------
        # 开始扫描 + 合并
        # --------------------------
        all_data = []
        found_files = 0

        # 遍历文件夹
        for root, dirs, files in os.walk(wechat_folder):
            for file in files:
                # 匹配文件名
                if re.match(match_regex, file):
                    file_path = os.path.join(root, file)
                    try:
                        df = pd.read_excel(file_path, engine="openpyxl")
                        all_data.append(df)
                        found_files += 1
                        st.write(f"✅ 已读取：{file}")
                    except Exception as e:
                        st.warning(f"⚠️ 读取失败：{file}，原因：{str(e)}")

        # --------------------------
        # 结果处理
        # --------------------------
        if found_files == 0:
            st.error("❌ 未找到任何符合条件的文件")
        else:
            # 合并所有表格
            combined_df = pd.concat(all_data, ignore_index=True)
            # 保存文件
            combined_df.to_excel(output_path, index=False)

            # 成功提示
            st.success(f"""
                🎉 合并完成！
                共合并文件：{found_files} 个
                保存路径：{output_path}
            """)

            # 提供下载（网页版必备）
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 点击下载合并后的文件",
                    data=f,
                    file_name=output_filename
                )