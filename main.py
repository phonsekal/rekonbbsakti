from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse, HTMLResponse
import pandas as pd
import openpyxl
import io
import re

app = FastAPI(title="Reconciliation System API & Web UI", version="11.0")

def clean_currency(value):
    if pd.isna(value) or value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    
    val_str = str(value).strip()
    if not val_str:
        return 0.0
    
    is_negative = False
    if val_str.startswith('(') and val_str.endswith(')'):
        is_negative = True
        val_str = val_str[1:-1]
    elif val_str.startswith('-'):
        is_negative = True
        val_str = val_str[1:]
    
    val_str = re.sub(r'[^0-9,\.]', '', val_str)
    
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif '.' in val_str:
        parts = val_str.split('.')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) != 2):
            val_str = val_str.replace('.', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')

    try:
        val = float(val_str)
        return -val if is_negative else val
    except ValueError:
        return 0.0

def format_number_clean(val: float) -> str:
    """Format angka tanpa desimal. Negatif ditulis dalam kurung (123.456)"""
    val_int = int(round(val))
    if val_int < 0:
        return f"({abs(val_int):,})".replace(",", ".")
    return f"{val_int:,}".replace(",", ".")

def normalize_period(period_val) -> str:
    """Standardisasi format periode ke YYYY-MM"""
    val_str = str(period_val or "").strip()
    if not val_str or val_str == "nan":
        return ""
    # Jika format YYYYMM (misal 202604) -> ubah ke 2026-04
    if len(val_str) == 6 and val_str.isdigit():
        return f"{val_str[:4]}-{val_str[4:]}"
    return val_str

def extract_satker_from_code(kd_bb_val, default_satker="693266"):
    """Ekstraksi kode Satker dari string seperti BEN-693266-520529242"""
    val_str = str(kd_bb_val or "").strip()
    match = re.search(r'BEN-(\d+)-', val_str)
    if match:
        return str(match.group(1)).strip()
    return str(default_satker).strip()

def parse_sakti_excel(contents: bytes):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True, read_only=True)
        ws = wb.active

        kode_akun = "-"
        nama_akun = "-"
        satker_header = "693266"
        records = []

        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if r_idx < 8:
                continue

            if r_idx == 8:
                satker_cell = str(row[4] or "").strip() if len(row) > 4 and row[4] is not None else ""
                if satker_cell:
                    satker_header = satker_cell
                continue

            if r_idx == 10:
                header_val = str(row[0] or "").strip() if len(row) > 0 and row[0] is not None else ""
                if "BUKU BESAR" in header_val:
                    parts = header_val.replace("BUKU BESAR", "").strip().split(" ", 1)
                    kode_akun = parts[0] if len(parts) > 0 else "-"
                    nama_akun = parts[1] if len(parts) > 1 else "-"
                continue

            if r_idx >= 11:
                col_a = row[0] if len(row) > 0 else None  # TGL JNL
                col_c = row[2] if len(row) > 2 else None  # KD BUKU BESAR
                col_d = row[3] if len(row) > 3 else None  # NO DOK
                col_i = row[8] if len(row) > 8 else None  # DEBET
                col_j = row[9] if len(row) > 9 else None  # KREDIT

                if str(col_a or "").startswith("BUKU BESAR") or str(col_a or "").strip() == "TGL JNL":
                    continue
                if str(col_c or "").strip() == "SALDO":
                    continue
                if col_a is None and col_c is None:
                    continue

                if col_a is not None and col_c is not None:
                    debet_val = clean_currency(col_i)
                    kredit_val = clean_currency(col_j)
                    net_nilai = debet_val - kredit_val

                    kode_satker = extract_satker_from_code(col_c, default_satker=satker_header)
                    tgl_dt = pd.to_datetime(col_a, errors='coerce')
                    periode_str = tgl_dt.strftime('%Y-%m') if pd.notna(tgl_dt) else normalize_period(col_a)
                    tgl_str = tgl_dt.strftime('%Y-%m-%d') if pd.notna(tgl_dt) else str(col_a)

                    records.append({
                        'col_kode_akun': kode_akun,
                        'col_nama_akun': nama_akun,
                        'col_kode_satker': str(kode_satker).strip(),
                        'col_tgl_jurnal': tgl_str,
                        'col_kode_periode': normalize_period(periode_str),
                        'col_no_doc': str(col_d or "").strip(),
                        'col_deskripsi': str(col_c or "").strip(),
                        'nilai_clean': net_nilai,
                        'tgl_dt': tgl_dt
                    })

        wb.close()

        if not records:
            df = pd.DataFrame(columns=[
                'col_kode_akun', 'col_nama_akun', 'col_kode_satker',
                'col_tgl_jurnal', 'col_kode_periode', 'col_no_doc',
                'col_deskripsi', 'nilai_clean', 'tgl_dt'
            ])
        else:
            df = pd.DataFrame(records)

        return df, kode_akun, nama_akun

    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal membaca file Excel SAKTI: {str(e)}")

def process_reconciliation(df: pd.DataFrame, filter_mode: str = 'ALL', target_period: str = '', target_satker: str = 'ALL', is_sakti_excel: bool = False):
    if df.empty:
        return {
            "kode_akun": "-",
            "nama_akun": "-",
            "available_satkers": [],
            "selected_satker": target_satker,
            "total_rows": 0,
            "total_unmatched": 0,
            "total_nilai_unmatched": "0",
            "main_columns": ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Deskripsi / Kode BB', 'Nilai'],
            "main_data": [],
            "has_resolved_later": False,
            "total_resolved_later": 0,
            "total_nilai_resolved_later": "0",
            "resolved_columns": [],
            "resolved_data": [],
            "has_hanging_after": False,
            "total_hanging_after": 0,
            "total_nilai_hanging_after": "0",
            "hanging_columns": [],
            "hanging_data": []
        }

    if is_sakti_excel:
        kode_akun_header = str(df['col_kode_akun'].iloc[0]) if 'col_kode_akun' in df.columns and not df.empty else "-"
        nama_akun_header = str(df['col_nama_akun'].iloc[0]) if 'col_nama_akun' in df.columns and not df.empty else "-"
        
        col_satker = 'col_kode_satker'
        col_tgl_jurnal = 'col_tgl_jurnal'
        col_kode_periode = 'col_kode_periode'
        col_no_doc = 'col_no_doc'
        col_deskripsi = 'col_deskripsi'
        
        df['satker_str'] = df[col_satker].astype(str).str.strip()
        df['periode_str'] = df[col_kode_periode].apply(normalize_period)
    else:
        if df.shape[1] < 12:
            raise HTTPException(status_code=400, detail="File CSV tidak memiliki setidaknya 12 kolom (s.d. Kolom L).")

        col_kode_akun = df.columns[2]
        col_nama_akun = df.columns[3]
        col_tgl_jurnal = df.columns[6]
        col_kode_periode = df.columns[7]
        col_no_doc = df.columns[8]
        col_deskripsi = df.columns[9]
        col_l_name = df.columns[11]

        kode_akun_header = str(df[col_kode_akun].iloc[0]).strip() if not df.empty else "-"
        nama_akun_header = str(df[col_nama_akun].iloc[0]).strip() if not df.empty else "-"

        df['satker_str'] = df[col_deskripsi].apply(lambda x: extract_satker_from_code(x, "693266"))
        df['periode_str'] = df[col_kode_periode].apply(normalize_period)
        df['nilai_clean'] = df[col_l_name].apply(clean_currency)
        df['tgl_dt'] = pd.to_datetime(df[col_tgl_jurnal], errors='coerce', dayfirst=True)

    # Ambil daftar unik Satker SEBELUM dipotong filter
    available_satkers = sorted([s for s in df['satker_str'].dropna().unique() if str(s).strip() != ""])

    # 1. Filter Satker secara aman (String normalization)
    if target_satker and str(target_satker).strip().upper() != 'ALL':
        df = df[df['satker_str'] == str(target_satker).strip()].copy()

    df['abs_val'] = df['nilai_clean'].abs()
    df = df.sort_values(by=['periode_str', 'tgl_dt']).reset_index(drop=True)
    df['row_id'] = df.index

    target_period_norm = normalize_period(target_period)

    # 2. Scope Matching (Tabel 1)
    if filter_mode == 'EXACT' and target_period_norm:
        scope_df = df[df['periode_str'] == target_period_norm].copy()
    elif filter_mode == 'UNTIL' and target_period_norm:
        scope_df = df[df['periode_str'] <= target_period_norm].copy()
    else:
        scope_df = df.copy()

    scope_df['matched_in_scope'] = False

    for abs_val, group in scope_df.groupby('abs_val'):
        if abs_val == 0:
            continue
        pos_indices = group[group['nilai_clean'] > 0].index.tolist()
        neg_indices = group[group['nilai_clean'] < 0].index.tolist()
        matched_cnt = min(len(pos_indices), len(neg_indices))

        for i in range(matched_cnt):
            scope_df.loc[pos_indices[i], 'matched_in_scope'] = True
            scope_df.loc[neg_indices[i], 'matched_in_scope'] = True

    unmatched_main = scope_df[~scope_df['matched_in_scope']].copy()

    # 3. Global Matching (Tabel 2)
    df['matched_pair_id'] = -1
    for abs_val, group in df.groupby('abs_val'):
        if abs_val == 0:
            continue
        pos_indices = group[group['nilai_clean'] > 0].index.tolist()
        neg_indices = group[group['nilai_clean'] < 0].index.tolist()
        matched_cnt = min(len(pos_indices), len(neg_indices))

        for i in range(matched_cnt):
            p_idx = pos_indices[i]
            n_idx = neg_indices[i]
            df.loc[p_idx, 'matched_pair_id'] = n_idx
            df.loc[n_idx, 'matched_pair_id'] = p_idx

    resolved_pairs_list = []
    used_as_pair_row_ids = set()

    if filter_mode == 'UNTIL' and target_period_norm:
        for idx, row in unmatched_main.iterrows():
            pair_id = df.loc[row['row_id'], 'matched_pair_id']
            if pair_id != -1:
                pair_row = df.loc[pair_id]
                if pair_row['periode_str'] > target_period_norm:
                    used_as_pair_row_ids.add(pair_row['row_id'])
                    resolved_pairs_list.append({
                        "Kode Satker": str(pair_row['satker_str']),
                        "Tanggal Jurnal": str(pair_row[col_tgl_jurnal]),
                        "Kode Periode Pasangan": str(pair_row[col_kode_periode]),
                        "Nomor Dokumen Pasangan": str(pair_row[col_no_doc]),
                        "target_doc": str(row[col_no_doc]),
                        "target_period": str(row['periode_str']),
                        "nilai_clean": pair_row['nilai_clean']
                    })

    # 4. Post-Period Check (Tabel 3)
    hanging_list = []
    if filter_mode == 'UNTIL' and target_period_norm:
        post_df = df[df['periode_str'] > target_period_norm].copy()

        if used_as_pair_row_ids:
            post_df = post_df[~post_df['row_id'].isin(used_as_pair_row_ids)].copy()

        if not post_df.empty:
            post_df['matched_post'] = False
            for abs_val, group in post_df.groupby('abs_val'):
                if abs_val == 0:
                    continue
                pos_indices = group[group['nilai_clean'] > 0].index.tolist()
                neg_indices = group[group['nilai_clean'] < 0].index.tolist()
                matched_cnt = min(len(pos_indices), len(neg_indices))

                for i in range(matched_cnt):
                    post_df.loc[pos_indices[i], 'matched_post'] = True
                    post_df.loc[neg_indices[i], 'matched_post'] = True

            post_unmatched = post_df[~post_df['matched_post']].copy()
            first_digit = kode_akun_header[0] if len(kode_akun_header) > 0 and kode_akun_header[0].isdigit() else "1"

            if first_digit in ['1', '5']:
                target_sign_df = post_unmatched[post_unmatched['nilai_clean'] < 0].copy()
                opp_sign_df = post_df[post_df['nilai_clean'] > 0].copy()
            else:
                target_sign_df = post_unmatched[post_unmatched['nilai_clean'] > 0].copy()
                opp_sign_df = post_df[post_df['nilai_clean'] < 0].copy()

            if not target_sign_df.empty and not opp_sign_df.empty:
                min_opp_date = opp_sign_df['tgl_dt'].min()
                target_sign_df = target_sign_df[target_sign_df['tgl_dt'] <= min_opp_date]

            if not target_sign_df.empty:
                for _, r in target_sign_df.iterrows():
                    hanging_list.append({
                        "Kode Satker": str(r['satker_str']),
                        "Tanggal Jurnal": str(r[col_tgl_jurnal]),
                        "Kode Periode": str(r[col_kode_periode]),
                        "Nomor Dokumen": str(r[col_no_doc]),
                        "nilai_clean": r['nilai_clean']
                    })

    # Output Formatting
    if not unmatched_main.empty:
        unmatched_main['Nilai'] = unmatched_main['nilai_clean'].apply(format_number_clean)
        unmatched_main['Kode Satker'] = unmatched_main['satker_str']
        total_main = unmatched_main['nilai_clean'].sum()

        selected_columns = ['Kode Satker', col_tgl_jurnal, col_kode_periode, col_no_doc, col_deskripsi, 'Nilai']
        main_df_final = unmatched_main[selected_columns].rename(columns={
            col_tgl_jurnal: 'Tanggal Jurnal',
            col_kode_periode: 'Kode Periode',
            col_no_doc: 'Nomor Dokumen',
            col_deskripsi: 'Deskripsi / Kode BB'
        })
    else:
        main_df_final = pd.DataFrame(columns=['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Deskripsi / Kode BB', 'Nilai'])
        total_main = 0.0

    if resolved_pairs_list:
        raw_res_df = pd.DataFrame(resolved_pairs_list)
        aggregated_res = raw_res_df.groupby(
            ['Kode Satker', 'Kode Periode Pasangan', 'Nomor Dokumen Pasangan'],
            as_index=False
        ).agg({
            'Tanggal Jurnal': 'first',
            'nilai_clean': 'sum',
            'target_doc': lambda x: ', '.join(sorted(set(x))),
            'target_period': 'first'
        })

        aggregated_res['Nilai'] = aggregated_res['nilai_clean'].apply(format_number_clean)
        aggregated_res['Keterangan Penyelesaian'] = aggregated_res.apply(
            lambda r: f"Pasangan Penihil Dok. {r['target_doc']} (Periode {r['target_period']})", axis=1
        )
        total_resolved = aggregated_res['nilai_clean'].sum()

        res_columns = ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode Pasangan', 'Nomor Dokumen Pasangan', 'Nilai', 'Keterangan Penyelesaian']
        res_df_final = aggregated_res[res_columns]
    else:
        res_df_final = pd.DataFrame()
        total_resolved = 0.0

    if hanging_list:
        raw_hang_df = pd.DataFrame(hanging_list)
        aggregated_hang = raw_hang_df.groupby(
            ['Kode Satker', 'Kode Periode', 'Nomor Dokumen'],
            as_index=False
        ).agg({
            'Tanggal Jurnal': 'first',
            'nilai_clean': 'sum'
        })

        aggregated_hang['Nilai'] = aggregated_hang['nilai_clean'].apply(format_number_clean)
        total_hanging = aggregated_hang['nilai_clean'].sum()

        hang_columns = ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Nilai']
        hang_df_final = aggregated_hang[hang_columns]
    else:
        hang_df_final = pd.DataFrame()
        total_hanging = 0.0

    return {
        "kode_akun": kode_akun_header,
        "nama_akun": nama_akun_header,
        "available_satkers": available_satkers,
        "selected_satker": target_satker,
        "total_rows": len(df),
        "total_unmatched": len(main_df_final),
        "total_nilai_unmatched": format_number_clean(total_main),
        "main_columns": list(main_df_final.columns),
        "main_data": main_df_final.fillna("").to_dict(orient='records'),
        
        "has_resolved_later": not res_df_final.empty,
        "total_resolved_later": len(res_df_final),
        "total_nilai_resolved_later": format_number_clean(total_resolved),
        "resolved_columns": list(res_df_final.columns) if not res_df_final.empty else [],
        "resolved_data": res_df_final.fillna("").to_dict(orient='records') if not res_df_final.empty else [],

        "has_hanging_after": not hang_df_final.empty,
        "total_hanging_after": len(hang_df_final),
        "total_nilai_hanging_after": format_number_clean(total_hanging),
        "hanging_columns": list(hang_df_final.columns) if not hang_df_final.empty else [],
        "hanging_data": hang_df_final.fillna("").to_dict(orient='records') if not hang_df_final.empty else []
    }
