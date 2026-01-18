import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 頁面設定
st.set_page_config(page_title="Global Excel Cross-Checker", layout="wide")

def find_cols(df, prod_idx, dest_idx, file_label):
    try:
        if df.empty:
            raise ValueError(f"The {file_label} file is empty.")

        # 方法 1: 使用 Index 獲取
        if len(df.columns) > max(prod_idx, dest_idx):
            new_df = df.iloc[:, [prod_idx, dest_idx]].copy()
        else:
            # 方法 2: 模糊匹配
            prod_col = [c for c in df.columns if any(k in str(c).lower() for k in ['prod', 'item', 'sku'])][0]
            dest_col = [c for c in df.columns if any(k in str(c).lower() for k in ['dest', 'country', 'loc'])][0]
            new_df = df[[prod_col, dest_col]].copy()
        
        new_df.columns = ['prod', 'dest']
        return new_df
    except Exception:
        st.error(f"❌ 無法在 {file_label} 中定位產品/目的地欄位。請檢查檔案結構。")
        return None

def process_data(ship_file, rw_file, csp_file):
    # 讀取資料
    raw_ship = pd.read_excel(ship_file)
    raw_rw = pd.read_excel(rw_file)
    raw_csp = pd.read_excel(csp_file)

    # 清理與轉換
    df_ship = find_cols(raw_ship, 0, 5, "Shipment")
    df_rw = find_cols(raw_rw, 0, 1, "RW")
    df_csp = find_cols(raw_csp, 0, 2, "CSP")

    if df_ship is None or df_rw is None or df_csp is None:
        return None

    for df in [df_ship, df_rw, df_csp]:
        df.drop_duplicates(inplace=True)
        df['prod'] = df['prod'].astype(str).str.strip()
        df['dest'] = df['dest'].astype(str).str.strip()
        df['exists'] = True

    # 合併邏輯
    master = pd.merge(df_ship, df_rw, on=['prod', 'dest'], how='outer', suffixes=('_ship', '_rw'))
    master = pd.merge(master, df_csp, on=['prod', 'dest'], how='outer')
    master.rename(columns={'exists': 'exists_csp'}, inplace=True)

    master['In_Shipment'] = master['exists_ship'].fillna(False).astype(bool)
    master['In_RW'] = master['exists_rw'].fillna(False).astype(bool)
    master['In_CSP'] = master['exists_csp'].fillna(False).astype(bool)

    final_df = master[['prod', 'dest', 'In_Shipment', 'In_RW', 'In_CSP']].copy()
    final_df.sort_values(by=['In_Shipment', 'prod'], ascending=[True, True], inplace=True)
    
    return final_df

# UI 介面
st.title("📊 Global Excel Cross-Checker")
st.info("請上傳三個 Excel 檔案來生成比對報告")

col1, col2, col3 = st.columns(3)
with col1:
    ship_file = st.file_uploader("Upload Shipment File", type=['xlsx'])
with col2:
    rw_file = st.file_uploader("Upload RW File", type=['xlsx'])
with col3:
    csp_file = st.file_uploader("Upload CSP File", type=['xlsx'])

if ship_file and rw_file and csp_file:
    if st.button("🚀 GENERATE REPORT", use_container_width=True):
        with st.spinner('正在處理數據中...'):
            result_df = process_data(ship_file, rw_file, csp_file)
            
            if result_df is not None:
                st.success("處理完成！")
                
                # 顯示預覽
                st.subheader("Data Preview (First 10 rows)")
                st.dataframe(result_df.head(10), use_container_width=True)

                # 準備 Excel 下載檔案 (使用 BytesIO)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    sheet_name = "Cross_Check_Result"
                    result_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet_name]
                    
                    # 格式化
                    yellow_format = workbook.add_format({'bg_color': '#FFFF00', 'border': 1})
                    
                    for i, col in enumerate(result_df.columns):
                        column_len = max(result_df[col].astype(str).str.len().max(), len(col)) + 2
                        worksheet.set_column(i, i, column_len)

                    last_row = len(result_df)
                    worksheet.conditional_format(1, 0, last_row, 4, {
                        'type': 'formula',
                        'criteria': '=$C2=FALSE',
                        'format': yellow_format
                    })
                
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 Download Excel Report",
                    data=processed_data,
                    file_name=f"Comparison_Report_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )