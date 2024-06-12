import streamlit as st
import pandas as pd
import time

st.title('SAP 和 WMS 檔案比對工具')

uploaded_file_sap = st.file_uploader("上傳 SAP 檔案", type=["xlsx"])
uploaded_files_wms = st.file_uploader("上傳 WMS 檔案", type=["xlsx"], accept_multiple_files=True)

if uploaded_file_sap and uploaded_files_wms:
    start_time = time.time()
    
    st.write("讀取 SAP 檔案中...")
    try:
        sap_df = pd.read_excel(uploaded_file_sap)
        sap_read_time = (time.time() - start_time) * 1000
        st.write(f"SAP 檔案讀取完成，耗時 {sap_read_time:.2f} 毫秒（{sap_read_time / 1000:.2f} 秒）。")
        st.write(f"SAP 檔案共有 {len(sap_df)} 筆資料。")
    except Exception as e:
        st.error(f"讀取 SAP 檔案時發生錯誤: {e}")
    
    st.write("讀取並合併 WMS 檔案中...")
    try:
        start_time = time.time()
        wms_dfs = [pd.read_excel(file) for file in uploaded_files_wms]
        wms_combined_df = pd.concat(wms_dfs)
        wms_read_time = (time.time() - start_time) * 1000
        st.write(f"WMS 檔案合併完成，共有 {len(wms_combined_df)} 筆資料，耗時 {wms_read_time:.2f} 毫秒（{wms_read_time / 1000:.2f} 秒）。")
    except Exception as e:
        st.error(f"讀取 WMS 檔案時發生錯誤: {e}")
    
    st.write("進行比對...")
    try:
        start_time = time.time()
        sap_dist_numbers = set(sap_df['DistNumber'])
        wms_dist_numbers = set(wms_combined_df['產品序號'])
        wms_missing_in_sap = wms_dist_numbers - sap_dist_numbers
        missing_in_sap_df = wms_combined_df[wms_combined_df['產品序號'].isin(wms_missing_in_sap)]
        compare_time = (time.time() - start_time) * 1000
        st.write(f"比對完成，共有 {len(missing_in_sap_df)} 筆資料在 WMS 上有但 SAP 沒有，耗時 {compare_time:.2f} 毫秒（{compare_time / 1000:.2f} 秒）。")
    except KeyError as e:
        st.error(f"資料框中缺少必需的欄位: {e}")
    except Exception as e:
        st.error(f"比對過程中發生錯誤: {e}")
    
    st.write("顯示比對結果...")
    try:
        st.dataframe(missing_in_sap_df)
    except Exception as e:
        st.error(f"顯示比對結果時發生錯誤: {e}")
    
    # 手動比對樣本
    st.write("進行樣本比對...")
    try:
        start_time = time.time()
        sample_sap_dist_numbers = list(sap_dist_numbers)[:5]  # 取前五個作為樣本
        sample_wms_dist_numbers = list(wms_dist_numbers)[:5]  # 取前五個作為樣本
        sap_sample_check = sap_df[sap_df['DistNumber'].isin(sample_sap_dist_numbers)]
        wms_sample_check = wms_combined_df[wms_combined_df['產品序號'].isin(sample_wms_dist_numbers)]
        sample_check_time = (time.time() - start_time) * 1000
        st.write(f"樣本比對完成，耗時 {sample_check_time:.2f} 毫秒（{sample_check_time / 1000:.2f} 秒）。")
        
        st.write("SAP 檔案中的樣本產品序號：")
        st.dataframe(sap_sample_check)
        st.write("WMS 檔案中的樣本產品序號：")
        st.dataframe(wms_sample_check)
    except KeyError as e:
        st.error(f"樣本比對過程中缺少必需的欄位: {e}")
    except Exception as e:
        st.error(f"樣本比對過程中發生錯誤: {e}")
    
    # 檢查重複的產品序號
    st.write("檢查重複產品序號...")
    try:
        start_time = time.time()
        sap_duplicates = sap_df[sap_df.duplicated(subset=['DistNumber'], keep=False)]
        wms_duplicates = wms_combined_df[wms_combined_df.duplicated(subset=['產品序號'], keep=False)]
        duplicates_check_time = (time.time() - start_time) * 1000
        st.write(f"SAP 檔案中有 {len(sap_duplicates)} 筆重複的產品序號，耗時 {duplicates_check_time:.2f} 毫秒（{duplicates_check_time / 1000:.2f} 秒）。")
        st.dataframe(sap_duplicates)
        st.write(f"WMS 檔案中有 {len(wms_duplicates)} 筆重複的產品序號，耗時 {duplicates_check_time:.2f} 毫秒（{duplicates_check_time / 1000:.2f} 秒）。")
        st.dataframe(wms_duplicates)
    except KeyError as e:
        st.error(f"檢查重複產品序號過程中缺少必需的欄位: {e}")
    except Exception as e:
        st.error(f"檢查重複產品序號過程中發生錯誤: {e}")
else:
    st.write("請上傳所有必要的檔案。")
