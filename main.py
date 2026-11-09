import pandas as pd
import datetime
import re
import os
from dataclasses import dataclass, fields

pd.set_option('display.max_rows', None)

# Path to excel files
ttos_Path = "src/ttos.xlsx"
lemon_Path = "src/lemon.xls"
bank_Path = "src/nganhang.xlsx"

# 1. ADD @dataclass DECORATOR
@dataclass
class ttos:
    quyen_hd: str
    so_hd: str
    mau_so: str
    ki_hieu: str
    so_hddt: str
    so_hddt_goc: str
    trang_thai_hd: str
    e_invoice_key: str
    ma_cqt: str
    loi_phat_hanh: str
    nph: str
    noi_dung_tt: str
    ma_kh: str
    mst: str
    loai_phi: str
    tong_tien: str
    hinh_thuc_tt: str
    da_gui_tbss: str
    ma_loi_tbss: str
    ly_do_sai_sot: str
    van_ban_dinh_kem: str
    da_phat_hanh_hddt: str
    ngay_phat_hanh: datetime
    bank_date: datetime.date = None

@dataclass
class lemon:
    so_phieu: str
    so_hd: str
    loai_nv: str
    ngay_phieu: datetime.date
    dien_giai: str
    ngay_hd: datetime.date
    ma_kh: str
    ten_kh: str
    pttt: str
    bang_gia: str
    nghiep_vu: str
    mst: str
    thanh_tien_nt_dieu_chinh: str
    tong_thanh_toan_nt: str
    tien_hang: str
    serial_number: str
    huy: int
    ttos_so_hd: str = None # This field is added to store the corresponding TTOS invoice number
    ttos_so_hd_goc: str = None # This field is added to store the corresponding original TTOS invoice number
    ttos_noi_dung: str = None # This field is added to store the corresponding TTOS payment description
    bank_date: datetime.date = None # This field is added to store the corresponding bank transaction date


@dataclass
class bank:
    ngay_gd: datetime.date
    so_tham_chieu: str
    so_tien_ghi_no: str
    so_tien_ghi_co: str
    mo_ta: str

ttos_string_cols = {
    "Số HĐ": str, "Số HĐĐT": str, "Số HĐĐT gốc": str
}

lemon_string_cols = {
    "Số hóa đơn": str
}

# Data extraction
ttos_data = pd.read_excel(ttos_Path, dtype=ttos_string_cols)
lemon_data = pd.read_excel(lemon_Path, dtype=lemon_string_cols)
bank_data = pd.read_excel(bank_Path)

# Renaming columns
ttos_data = ttos_data.rename(columns={
    "Quyển HĐ": "quyen_hd", "Số HĐ": "so_hd", "Mẫu số": "mau_so",
    "Ký hiệu": "ki_hieu", "Số HĐĐT": "so_hddt", "Số HĐĐT gốc": "so_hddt_goc",
    "Trạng thái HĐ": "trang_thai_hd", "E-invoice Fkey": "e_invoice_key",
    "Mã CQT": "ma_cqt", "Lỗi phát hành": "loi_phat_hanh",
    "Người phát hành": "nph", "Nội dung thanh toán": "noi_dung_tt",
    "Mã KH": "ma_kh", "MST": "mst", "Loại phí": "loai_phi",
    "Tổng tiền": "tong_tien", "Hình thức thanh toán": "hinh_thuc_tt",
    "Đã gửi TBSS": "da_gui_tbss", "Mã lỗi TBSS": "ma_loi_tbss",
    "Lý do sai sót": "ly_do_sai_sot", "Văn bản đính kèm": "van_ban_dinh_kem",
    "Đã phát hành HĐĐT": "da_phat_hanh_hddt", "Ngày phát hành HĐĐT": "ngay_phat_hanh"
})

lemon_data = lemon_data.rename(columns={
    "Số phiếu": "so_phieu", "Số hóa đơn": "so_hd", "Loại nghiệp vụ": "loai_nv",
    "Ngày phiếu": "ngay_phieu", "Diễn giải": "dien_giai", "Ngày hóa đơn": "ngay_hd",
    "Mã khách hàng": "ma_kh", "Tên khách hàng": "ten_kh", "PT thanh toán": "pttt",
    "Bảng giá": "bang_gia", "Nghiệp vụ": "nghiep_vu", "Mã số thuế": "mst",
    "Thành tiền NT điều chỉnh": "thanh_tien_nt_dieu_chinh",
    "Tổng thanh toán (NT)": "tong_thanh_toan_nt", "Tiền hàng": "tien_hang",
    "Số Sêri": "serial_number", "Hủy": "huy"
})

bank_data = bank_data.rename(columns={
    "Ngày giao dịch": "ngay_gd", "Số tham chiếu": "so_tham_chieu",
    "Số tiền ghi nợ": "so_tien_ghi_no", "Số tiền ghi có": "so_tien_ghi_co",
    "Mô tả": "mo_ta"
})


# 2. EXTRACT DATA TO CLASSES
def dataframe_to_dataclass(df, dataclass_type):
    """
    Safely converts a dataframe to a list of dataclass objects.
    It ignores extra columns in the Excel file that aren't defined in the dataclass.
    """
    valid_fields = {f.name for f in fields(dataclass_type)}
    records = df.to_dict(orient='records')
    
    # Filter dictionary to only include keys that exist in the dataclass
    return [dataclass_type(**{k: v for k, v in row.items() if k in valid_fields}) for row in records]


def compare_bill_id_lemon_and_ttos(lemon_objects, ttos_objects):
    """
    Compares the bill IDs between lemon and ttos objects.
    Returns a list of tuples containing (lemon_bill_id, ttos_bill_id) for matching records.
    """
    matches = []
    unmatches = []
    lemon_dict = {obj.so_hd: obj for obj in lemon_objects}  # Create a dictionary for quick lookup

    found_matches = False

    for ttos_obj in ttos_objects:
        if ttos_obj.so_hddt in lemon_dict:
            matches.append((lemon_dict[ttos_obj.so_hddt].so_hd, ttos_obj.so_hddt))
            lemon_dict[ttos_obj.so_hddt].ttos_so_hd = ttos_obj.so_hd  # Update the lemon object with the matching TTOS invoice number
            lemon_dict[ttos_obj.so_hddt].ttos_noi_dung = ttos_obj.trang_thai_hd  # Update the lemon object with the corresponding TTOS payment description
            lemon_dict[ttos_obj.so_hddt].ttos_so_hd_goc = ttos_obj.so_hddt_goc  # Update the lemon object with the corresponding original TTOS invoice number
            lemon_dict[ttos_obj.so_hddt].bank_date = ttos_obj.bank_date  # Update the lemon object with the corresponding bank transaction date
            found_matches = True

    if not found_matches:
        unmatches.append(ttos_obj.so_hddt)
    return matches, unmatches

def list_unmatched_lemon_ttos(lemon_objects, ttos_objects):
    """
    Returns a list of lemon objects that do not have a matching bill ID in ttos objects.
    """
    ttos_bill_ids = {obj.so_hddt for obj in ttos_objects}
    unmatched_lemon = [obj for obj in lemon_objects if obj.so_hd not in ttos_bill_ids]
    
    return unmatched_lemon

def compare_bill_id_ttos_and_bank(ttos_objects, bank_objects):
    matches = []
    unmatches = [] # This will now store unmatched TTOS objects
    
    for ttos_obj in ttos_objects:
        so_hd = str(ttos_obj.so_hd).strip()
        
        if not so_hd or so_hd.lower() == 'nan':
            continue
            
        pattern = rf"(?<![a-zA-Z0-9]){re.escape(so_hd)}(?![a-zA-Z0-9])"
        
        # 1. Create a flag before we start checking the bank
        found_match = False 
        
        for bank_obj in bank_objects:
            mo_ta = str(bank_obj.mo_ta)
            
            if re.search(pattern, mo_ta):
                matches.append((ttos_obj.so_hd, bank_obj.mo_ta))
                ttos_obj.bank_date = bank_obj.ngay_gd  #  
                # 2. If we find a match, flip the flag to True
                found_match = True 
                
        # 3. After checking ALL bank transactions, check the flag
        if not found_match:
            # We don't append bank_obj.mo_ta here because there is no matching bank transaction!
            unmatches.append(ttos_obj.so_hd) 
                
    return matches, unmatches

def list_unmatched_ttos_and_bank(ttos_objects, bank_objects):
    """
    Finds both TTOS invoices not found in the bank, 
    and Bank transactions that don't match any TTOS invoice.
    Returns a tuple: (unmatched_ttos_list, unmatched_bank_list)
    """
    # Use sets to keep track of what HAS been matched
    matched_ttos_ids = set()
    matched_bank_indices = set()
    
    for ttos_obj in ttos_objects:
        so_hd = str(ttos_obj.so_hd).strip()
        
        if not so_hd or so_hd.lower() == 'nan':
            continue
            
        pattern = rf"(?<![a-zA-Z0-9]){re.escape(so_hd)}(?![a-zA-Z0-9])"
        
        for index, bank_obj in enumerate(bank_objects):
            mo_ta = str(bank_obj.mo_ta)
            
            if re.search(pattern, mo_ta):
                # Record the matches
                matched_ttos_ids.add(ttos_obj.so_hd)
                matched_bank_indices.add(index) # Use index in case bank has duplicate descriptions
                
    # Filter the original lists to only keep items that are NOT in the matched sets
    unmatched_ttos = [obj for obj in ttos_objects if obj.so_hd not in matched_ttos_ids]
    unmatched_bank = [obj for i, obj in enumerate(bank_objects) if i not in matched_bank_indices]
    
    return unmatched_ttos, unmatched_bank


def compare_date_in_lemon(lemon_objects):
    matches = []
    unmatches = []
    for obj in lemon_objects:
        if obj.ngay_hd == obj.bank_date:
            matches.append(obj)
        else:
            unmatches.append(obj)

    return matches, unmatches

def sort_lemon_by_bill_id(lemon_objects):
    return sorted(lemon_objects, key=lambda x: x.so_hd, reverse=True)


# --- HÀM XỬ LÝ CHÍNH ---
def run_report(lemon_list, ttos_list, bank_list, show_matched_sum=False, sheet_name="Tổng hợp"):
    # 1. Đối chiếu Ttos và Bank
    matches_ttos_bank, unmatches_ttos_bank = compare_bill_id_ttos_and_bank(ttos_list, bank_list)
    
    # 2. Đối chiếu Lemon và Ttos
    matches_lemon_ttos, unmatches_lemon_ttos = compare_bill_id_lemon_and_ttos(lemon_list, ttos_list)
    
    # 3. Đối chiếu Ngày
    matches_date, unmatches_date = compare_date_in_lemon(lemon_list)
    
    matches_cnt = len(matches_date)
    unmatches_cnt = len(unmatches_date)
    print(f"\n➔ Tổng số hóa đơn khớp toàn bộ: {matches_cnt}")
    print(f"➔ Tổng số hóa đơn lệch: {unmatches_cnt}")

    # Tính tổng tiền của các phần KHỚP
    if show_matched_sum:
        tong_tien_khop = 0
        for obj in matches_date:
            try:
                tien_str = str(obj.tong_thanh_toan_nt).replace(',', '').strip()
                tong_tien_khop += float(tien_str) if tien_str else 0
            except ValueError:
                continue
        print(f"💰 TỔNG TIỀN CỦA CÁC HÓA ĐƠN KHỚP: {tong_tien_khop:,.0f} VNĐ\n")

    # 4. Xuất file Excel
    os.makedirs("des", exist_ok=True)

    df_matches = pd.DataFrame([obj.__dict__ for obj in matches_date])
    if not df_matches.empty:
        df_matches['Trạng thái'] = 'Khớp'

    df_unmatches = pd.DataFrame([obj.__dict__ for obj in unmatches_date])
    if not df_unmatches.empty:
        df_unmatches['Trạng thái'] = df_unmatches['bank_date'].apply(
            lambda x: f"Sao kê NH ngày {x.strftime('%Y-%m-%d')}" if pd.notnull(x) else "Không có ngày ngân hàng"
        )

    df_combined = pd.concat([df_matches, df_unmatches], ignore_index=True)

    if not df_combined.empty:
        df_combined['sheet_name_date'] = pd.to_datetime(df_combined['ngay_hd']).dt.strftime('%Y-%m-%d').fillna('Unknown')

        def append_summary(df_in):
            df_out = df_in.copy()
            col_tien = 'tong_thanh_toan_nt' 
            tien_numeric = pd.to_numeric(df_out[col_tien], errors='coerce').fillna(0)
            
            total_khop = tien_numeric[df_out['Trạng thái'] == 'Khớp'].sum()
            total_lech = tien_numeric[df_out['Trạng thái'] != 'Khớp'].sum()
            total_all = total_khop + total_lech
            
            summary_data = [
                {'so_phieu': 'TỔNG TIỀN KHỚP:', col_tien: total_khop},
                {'so_phieu': 'TỔNG TIỀN LỆCH:', col_tien: total_lech},
                {'so_phieu': 'TỔNG CỘNG:', col_tien: total_all}
            ]
            return pd.concat([df_out, pd.DataFrame(summary_data)], ignore_index=True)

        def highlight_rows(row):
            if str(row.get('so_phieu', '')).startswith('TỔNG'):
                return ['font-weight: bold; background-color: #D9EAD3'] * len(row)
            if row.get('Trạng thái') != 'Khớp' and pd.notnull(row.get('Trạng thái')):
                return ['background-color: #FFB3B3'] * len(row) 
            return [''] * len(row)

        file_path = "des/Bao_cao_tong_hop.xlsx"
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Cắt ngắn tên sheet để đảm bảo không vượt quá 31 ký tự theo giới hạn của Excel
            safe_sheet_name = str(sheet_name)[:31]
            
            # TẠO MENU SHEET
            danh_sach_ngay = df_combined['sheet_name_date'].unique()
            menu_rows = [{
                "Danh mục": "Toàn bộ dữ liệu" if sheet_name == "Tổng hợp" else safe_sheet_name, 
                "Liên kết": f'=HYPERLINK("#\'{safe_sheet_name}\'!A1", "▶{safe_sheet_name}")'
            }]
            
            for ngay in danh_sach_ngay:
                sheet_ngay = str(ngay)[:31]
                menu_rows.append({
                    "Danh mục": f"Báo cáo ngày {sheet_ngay}",
                    "Liên kết": f'=HYPERLINK("#\'{sheet_ngay}\'!A1", "▶{sheet_ngay}")'
                })
            pd.DataFrame(menu_rows).to_excel(writer, sheet_name="Mục lục", index=False)
            back_btn_formula = '=HYPERLINK("#\'Mục lục\'!A1", "🔙 TRỞ VỀ MỤC LỤC")'

            # XUẤT SHEET DỮ LIỆU TỔNG/KIỂM TRA
            clean_df = df_combined.drop(columns=['sheet_name_date'])
            styled_df = append_summary(clean_df).style.apply(highlight_rows, axis=1)
            
            styled_df.to_excel(writer, sheet_name=safe_sheet_name, index=False, startrow=1)
            writer.sheets[safe_sheet_name]['A1'] = back_btn_formula
            
            # XUẤT SHEET THEO NGÀY
            for date_str, group_df in df_combined.groupby('sheet_name_date'):
                sheet_ngay_group = str(date_str)[:31]
                
                valid_bills = group_df['so_hd'].dropna().astype(str).str.strip()
                valid_bills = valid_bills[valid_bills != '']
                if not valid_bills.empty:
                    start_bill = valid_bills.min()
                    end_bill = valid_bills.max()
                    bill_info = f"Từ HĐ: {start_bill}  ➔  Đến HĐ: {end_bill}"
                else:
                    bill_info = "Không có số hóa đơn"
                
                clean_group = group_df.drop(columns=['sheet_name_date'])
                styled_group = append_summary(clean_group).style.apply(highlight_rows, axis=1)
                
                columns_to_export = ['so_phieu', 'so_hd', 'ngay_hd', 'tong_thanh_toan_nt', 'dien_giai', 'pttt', 'huy', 'ttos_so_hd', 'ttos_so_hd_goc', 'ttos_noi_dung','Trạng thái', 'bank_date']
                styled_group.to_excel(writer, sheet_name=sheet_ngay_group, index=False, startrow=1, columns=columns_to_export)
                
                writer.sheets[sheet_ngay_group]['A1'] = back_btn_formula
                writer.sheets[sheet_ngay_group]['B1'] = bill_info
                
        print(f"✅ Đã xuất file thành công tại: {file_path}")
    else:
        print("❌ Không có dữ liệu để xuất ra Excel.")


# --- CHƯƠNG TRÌNH CHÍNH (MENU) ---
if __name__ == "__main__":
    # Trích xuất dữ liệu thô ban đầu
    ttos_objects_raw = dataframe_to_dataclass(ttos_data, ttos)
    lemon_objects_raw = dataframe_to_dataclass(lemon_data, lemon)
    bank_objects_raw = dataframe_to_dataclass(bank_data, bank)

    while True:
        print("\n" + "="*40)
        print("       HỆ THỐNG ĐỐI CHIẾU HÓA ĐƠN")
        print("="*40)
        print("1. Chạy đối chiếu TOÀN BỘ dữ liệu")
        print("2. Chạy theo KHOẢNG SỐ HÓA ĐƠN (Lemon)")
        print("3. Thoát chương trình")
        print("="*40)
        
        choice = input("Nhập lựa chọn của bạn (1/2/3): ").strip()
        
        if choice == '1':
            print("\nĐang xử lý toàn bộ dữ liệu...")
            run_report(lemon_objects_raw, ttos_objects_raw, bank_objects_raw, show_matched_sum=False)
            
        elif choice == '2':
            start_hd = input("Nhập số hóa đơn BẮT ĐẦU (VD: 0000100): ").strip()
            end_hd = input("Nhập số hóa đơn KẾT THÚC (VD: 0000150): ").strip()
            
            # Sắp xếp Lemon tăng dần theo số hóa đơn để dễ dàng lọc
            sorted_lemon = sorted(lemon_objects_raw, key=lambda x: str(x.so_hd).strip())
            
            # Đảm bảo start_hd luôn nhỏ hơn hoặc bằng end_hd
            if start_hd > end_hd:
                start_hd, end_hd = end_hd, start_hd
                
            # Lọc lấy các hóa đơn nằm trong khoảng [start, end]
            filtered_lemon = [
                obj for obj in sorted_lemon 
                if start_hd <= str(obj.so_hd).strip() <= end_hd
            ]
            
            if not filtered_lemon:
                print(f"\n⚠️ Không tìm thấy hóa đơn nào trong khoảng từ {start_hd} đến {end_hd}!")
                continue
                
            print(f"\nĐang xử lý {len(filtered_lemon)} hóa đơn (từ {start_hd} đến {end_hd})...")
            
            # Khởi tạo sheet name kiểm tra
            sheet_name = f"Kiểm tra {start_hd} - {end_hd}"
            # Chạy báo cáo
            run_report(filtered_lemon, ttos_objects_raw, bank_objects_raw, show_matched_sum=True, sheet_name=sheet_name)
            
        elif choice == '3':
            print("Đang thoát chương trình. Tạm biệt!")
            break
            
        else:
            print("❌ Lựa chọn không hợp lệ, vui lòng nhập lại!")