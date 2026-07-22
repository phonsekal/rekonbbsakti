from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import JSONResponse, HTMLResponse
import pandas as pd
import openpyxl
import io
import re

app = FastAPI(title="Reconciliation System API & Web UI", version="8.1")

def clean_currency(value):
    if pd.isna(value):
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

def extract_satker_from_code(kd_bb_val, default_satker="693266"):
    """Ekstraksi kode Satker dari string seperti BEN-693266-520529242"""
    val_str = str(kd_bb_val or "").strip()
    match = re.search(r'BEN-(\d+)-', val_str)
    if match:
        return str(match.group(1)).strip()
    return str(default_satker).strip()

def parse_sakti_excel(contents: bytes):
    """Ekstraksi otomatis file Buku Besar SAKTI (.xlsx / .xls)"""
    wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
    ws = wb.active

    kode_akun = "-"
    nama_akun = "-"
    satker_header = "693266"

    # Ambil Header Satker (baris 8 col 5 jika ada)
    try:
        s_cell = str(ws.cell(8, 5).value or "").strip()
        if s_cell: satker_header = s_cell
    except Exception:
        pass
    
    # Ambil Header Akun dari Baris 10 (misal: "BUKU BESAR 111821 Kas Lainnya...")
    header_val = str(ws.cell(10, 1).value or "").strip()
    if "BUKU BESAR" in header_val:
        parts = header_val.replace("BUKU BESAR", "").strip().split(" ", 1)
        kode_akun = parts[0] if len(parts) > 0 else "-"
        nama_akun = parts[1] if len(parts) > 1 else "-"

    records = []
    
    for r in range(11, ws.max_row + 1):
        col_a = ws.cell(r, 1).value  # TGL JNL
        col_c = ws.cell(r, 3).value  # KD BUKU BESAR / DESKRIPSI
        col_d = ws.cell(r, 4).value  # NO DOK
        col_i = ws.cell(r, 9).value  # DEBET
        col_j = ws.cell(r, 10).value # KREDIT
        
        # Skip baris header / footer / total
        if str(col_a or "").startswith("BUKU BESAR") or str(col_a or "").strip() == "TGL JNL":
            continue
        if str(col_c or "").strip() == "SALDO":
            continue
        if col_a is None and col_c is None:
            continue
            
        if col_a is not None and col_c is not None:
            debet_val = clean_currency(col_i)
            kredit_val = clean_currency(col_j)
            
            # Konversi DEBET & KREDIT ke Nilai Bersih: Debet (+), Kredit (-)
            net_nilai = debet_val - kredit_val
            kode_satker = extract_satker_from_code(col_c, default_satker=satker_header)
            
            tgl_dt = pd.to_datetime(col_a, errors='coerce')
            periode_str = tgl_dt.strftime('%Y-%m') if pd.notna(tgl_dt) else ""
            tgl_str = tgl_dt.strftime('%Y-%m-%d') if pd.notna(tgl_dt) else str(col_a)

            records.append({
                'col_kode_akun': kode_akun,
                'col_nama_akun': nama_akun,
                'col_kode_satker': str(kode_satker).strip(),
                'col_tgl_jurnal': tgl_str,
                'col_kode_periode': periode_str,
                'col_no_doc': str(col_d or ""),
                'col_deskripsi': str(col_c or ""),
                'nilai_clean': net_nilai,
                'tgl_dt': tgl_dt
            })
            
    return pd.DataFrame(records), kode_akun, nama_akun

def process_reconciliation(df: pd.DataFrame, filter_mode: str = 'ALL', target_period: str = '', is_sakti_excel: bool = False):
    if is_sakti_excel:
        kode_akun_header = str(df['col_kode_akun'].iloc[0]) if not df.empty else "-"
        nama_akun_header = str(df['col_nama_akun'].iloc[0]) if not df.empty else "-"
        
        col_satker = 'col_kode_satker'
        col_tgl_jurnal = 'col_tgl_jurnal'
        col_kode_periode = 'col_kode_periode'
        col_no_doc = 'col_no_doc'
        col_deskripsi = 'col_deskripsi'
        
        df['satker_str'] = df[col_satker].astype(str).str.strip() if col_satker in df.columns else "693266"
        df['periode_str'] = df[col_kode_periode].astype(str).str.strip()
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
        df['periode_str'] = df[col_kode_periode].astype(str).str.strip()
        df['nilai_clean'] = df[col_l_name].apply(clean_currency)
        df['tgl_dt'] = pd.to_datetime(df[col_tgl_jurnal], errors='coerce', dayfirst=True)

    # Ambil list Satker unik untuk ditaruh di dropdown frontend
    available_satkers = sorted([s for s in df['satker_str'].dropna().unique() if str(s).strip() != ""])

    df['abs_val'] = df['nilai_clean'].abs()
    df = df.sort_values(by=['periode_str', 'tgl_dt']).reset_index(drop=True)
    df['row_id'] = df.index

    # 1. Scope Matching (Tabel 1: Belum Memiliki Pasangan)
    if filter_mode == 'EXACT' and target_period:
        scope_df = df[df['periode_str'] == target_period].copy()
    elif filter_mode == 'UNTIL' and target_period:
        scope_df = df[df['periode_str'] <= target_period].copy()
    else:
        scope_df = df.copy()

    scope_df['matched_in_scope'] = False

    for abs_val, group in scope_df.groupby('abs_val'):
        if abs_val == 0:
            continue
        pos_indices = group[group['nilai_clean'] > 0].index.tolist()  # Debet
        neg_indices = group[group['nilai_clean'] < 0].index.tolist()  # Kredit
        matched_cnt = min(len(pos_indices), len(neg_indices))

        for i in range(matched_cnt):
            scope_df.loc[pos_indices[i], 'matched_in_scope'] = True
            scope_df.loc[neg_indices[i], 'matched_in_scope'] = True

    unmatched_main = scope_df[~scope_df['matched_in_scope']].copy()

    # 2. Global Matching (Tabel 2: Pasangan Penihil)
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

    if filter_mode == 'UNTIL' and target_period:
        for idx, row in unmatched_main.iterrows():
            pair_id = df.loc[row['row_id'], 'matched_pair_id']
            if pair_id != -1:
                pair_row = df.loc[pair_id]
                if pair_row['periode_str'] > target_period:
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

    # 3. Post-Period Check (Tabel 3: Menggantung Pasca Periode X)
    hanging_list = []
    if filter_mode == 'UNTIL' and target_period:
        post_df = df[df['periode_str'] > target_period].copy()

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
            first_digit = kode_akun_header[0] if len(kode_akun_header) > 0 else "1"

            if first_digit in ['1', '5']:
                target_sign_df = post_unmatched[post_unmatched['nilai_clean'] < 0].copy()  # Kredit (Minus)
                opp_sign_df = post_df[post_df['nilai_clean'] > 0].copy()                  # Debet (Plus)

                if not target_sign_df.empty and not opp_sign_df.empty:
                    min_opp_date = opp_sign_df['tgl_dt'].min()
                    target_sign_df = target_sign_df[target_sign_df['tgl_dt'] <= min_opp_date]
            else:
                target_sign_df = post_unmatched[post_unmatched['nilai_clean'] > 0].copy()  # Debet (Plus)
                opp_sign_df = post_df[post_df['nilai_clean'] < 0].copy()                  # Kredit (Minus)

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

    # Formatting Output
    if not unmatched_main.empty:
        unmatched_main['Nilai_Clean'] = unmatched_main['nilai_clean']
        unmatched_main['Nilai'] = unmatched_main['nilai_clean'].apply(format_number_clean)
        unmatched_main['Kode Satker'] = unmatched_main['satker_str']
        total_main = unmatched_main['nilai_clean'].sum()

        selected_columns = ['Kode Satker', col_tgl_jurnal, col_kode_periode, col_no_doc, col_deskripsi, 'Nilai']
        main_df_final = unmatched_main[['Kode Satker', col_tgl_jurnal, col_kode_periode, col_no_doc, col_deskripsi, 'Nilai', 'Nilai_Clean']].rename(columns={
            col_tgl_jurnal: 'Tanggal Jurnal',
            col_kode_periode: 'Kode Periode',
            col_no_doc: 'Nomor Dokumen',
            col_deskripsi: 'Deskripsi / Kode BB'
        })
    else:
        main_df_final = pd.DataFrame(columns=['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Deskripsi / Kode BB', 'Nilai', 'Nilai_Clean'])
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

        aggregated_res['Nilai_Clean'] = aggregated_res['nilai_clean']
        aggregated_res['Nilai'] = aggregated_res['nilai_clean'].apply(format_number_clean)
        aggregated_res['Keterangan Penyelesaian'] = aggregated_res.apply(
            lambda r: f"Pasangan Penihil Dok. {r['target_doc']} (Periode {r['target_period']})", axis=1
        )
        total_resolved = aggregated_res['nilai_clean'].sum()

        res_columns = ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode Pasangan', 'Nomor Dokumen Pasangan', 'Nilai', 'Keterangan Penyelesaian']
        res_df_final = aggregated_res[['Kode Satker', 'Tanggal Jurnal', 'Kode Periode Pasangan', 'Nomor Dokumen Pasangan', 'Nilai', 'Keterangan Penyelesaian', 'Nilai_Clean']]
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

        aggregated_hang['Nilai_Clean'] = aggregated_hang['nilai_clean']
        aggregated_hang['Nilai'] = aggregated_hang['nilai_clean'].apply(format_number_clean)
        total_hanging = aggregated_hang['nilai_clean'].sum()

        hang_columns = ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Nilai']
        hang_df_final = aggregated_hang[['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Nilai', 'Nilai_Clean']]
    else:
        hang_df_final = pd.DataFrame()
        total_hanging = 0.0

    return {
        "kode_akun": kode_akun_header,
        "nama_akun": nama_akun_header,
        "available_satkers": available_satkers,
        "total_rows": len(df),
        "total_unmatched": len(main_df_final),
        "total_nilai_unmatched": format_number_clean(total_main),
        "main_columns": ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Deskripsi / Kode BB', 'Nilai'],
        "main_data": main_df_final.fillna("").to_dict(orient='records'),
        
        "has_resolved_later": not res_df_final.empty,
        "total_resolved_later": len(res_df_final),
        "total_nilai_resolved_later": format_number_clean(total_resolved),
        "resolved_columns": ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode Pasangan', 'Nomor Dokumen Pasangan', 'Nilai', 'Keterangan Penyelesaian'],
        "resolved_data": res_df_final.fillna("").to_dict(orient='records') if not res_df_final.empty else [],

        "has_hanging_after": not hang_df_final.empty,
        "total_hanging_after": len(hang_df_final),
        "total_nilai_hanging_after": format_number_clean(total_hanging),
        "hanging_columns": ['Kode Satker', 'Tanggal Jurnal', 'Kode Periode', 'Nomor Dokumen', 'Nilai'],
        "hanging_data": hang_df_final.fillna("").to_dict(orient='records') if not hang_df_final.empty else []
    }

@app.get("/", response_class=HTMLResponse)
async def home_ui():
    html_content = """<!DOCTYPE html>
<html lang="id" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reconciliation App — Dede Saputra</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        body { background-color: #0b0f17; color: #e2e8f0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
        .brand-cyan { color: #00d2ff; }
        .border-dark { border-color: #1e293b; }
        .bg-card { background-color: #111827; }
    </style>
</head>
<body class="min-h-screen flex flex-col justify-between">
    <header class="border-b border-dark py-4 px-6 sm:px-12 bg-[#0b0f17]/90 backdrop-blur sticky top-0 z-50">
        <div class="max-w-6xl mx-auto flex items-center justify-between">
            <div class="flex items-center space-x-3">
                <div class="flex space-x-1">
                    <span class="w-2.5 h-6 bg-[#00d2ff] rounded-sm transform -skew-x-12"></span>
                    <span class="w-2.5 h-6 bg-[#0a84ff] rounded-sm transform -skew-x-12"></span>
                </div>
                <span class="font-bold text-lg tracking-tight text-white">dedesaputra <span class="text-slate-400 font-normal">Reconcile</span></span>
            </div>
        </div>
    </header>

    <main class="max-w-6xl mx-auto px-6 py-10 w-full flex-grow">
        <div class="mb-8">
            <h1 class="text-3xl font-extrabold text-white tracking-tight mb-2">Rekonsiliasi Transaksi (Debet / Kredit)</h1>
            <p class="text-sm text-slate-400">Mendukung unggahan file CSV standar dan Laporan Buku Besar SAKTI (.xlsx / .xls).</p>
        </div>

        <div id="uploadSection" class="bg-card rounded-2xl border border-dark p-8 mb-8 shadow-xl">
            <form id="uploadForm" class="space-y-6">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#0b0f17] p-4 rounded-xl border border-dark">
                    <div>
                        <label class="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
                            <i class="fa-solid fa-filter brand-cyan mr-1"></i> Mode Filter Periode
                        </label>
                        <select id="filterMode" name="filter_mode" class="w-full bg-[#111827] border border-dark text-slate-200 text-xs rounded-lg p-2.5 focus:border-[#00d2ff] outline-none">
                            <option value="ALL">Semua Periode (Tanpa Filter)</option>
                            <option value="EXACT">Hanya Periode X</option>
                            <option value="UNTIL">Sampai Dengan (s.d.) Periode X</option>
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">
                            <i class="fa-solid fa-calendar-days brand-cyan mr-1"></i> Pilih Periode X
                        </label>
                        <select id="targetPeriod" name="target_period" disabled class="w-full bg-[#111827] border border-dark text-slate-200 text-xs rounded-lg p-2.5 focus:border-[#00d2ff] outline-none disabled:opacity-40">
                            <option value="">-- Pilih Periode --</option>
                            <option value="2026-01">2026-01</option>
                            <option value="2026-02">2026-02</option>
                            <option value="2026-03">2026-03</option>
                            <option value="2026-04">2026-04</option>
                            <option value="2026-05">2026-05</option>
                            <option value="2026-06">2026-06</option>
                            <option value="2026-07">2026-07</option>
                            <option value="2026-08">2026-08</option>
                            <option value="2026-09">2026-09</option>
                            <option value="2026-10">2026-10</option>
                            <option value="2026-11">2026-11</option>
                            <option value="2026-12">2026-12</option>
                        </select>
                    </div>
                </div>

                <div id="dropZone" class="border border-dashed border-slate-700 rounded-xl p-10 transition-all hover:border-[#00d2ff] hover:bg-[#0b0f17]/50 cursor-pointer flex flex-col items-center justify-center text-center">
                    <input type="file" id="csvFile" name="file" accept=".csv,.xlsx,.xls" class="hidden">
                    <i class="fa-solid fa-cloud-arrow-up text-3xl brand-cyan mb-3"></i>
                    <p class="text-sm font-semibold text-slate-200" id="fileLabel">Unggah File CSV / Buku Besar Excel (.xlsx)</p>
                    <p class="text-xs text-slate-500 mt-1">Klik atau seret file CSV / Excel Buku Besar SAKTI ke area ini</p>
                </div>

                <button type="submit" id="btnSubmit" disabled class="w-full bg-[#0a84ff] hover:bg-[#0071e3] disabled:bg-slate-800 disabled:text-slate-600 text-white font-semibold py-3 px-6 rounded-xl transition-all shadow-lg flex items-center justify-center space-x-2 disabled:cursor-not-allowed">
                    <i class="fa-solid fa-bolt text-xs"></i>
                    <span>Proses & Analisis Data</span>
                </button>
            </form>

            <div id="loading" class="hidden mt-6 flex flex-col items-center">
                <div class="animate-spin rounded-full h-8 w-8 border-b-2 border-[#00d2ff] mb-2"></div>
                <span class="text-xs font-medium text-slate-400">Merapikan & memproses data...</span>
            </div>
        </div>

        <div id="resultSection" class="hidden space-y-8">
            <div class="bg-card border border-dark rounded-xl p-6 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                <div>
                    <span class="text-[10px] font-bold uppercase tracking-widest text-slate-500 block mb-1">Informasi Akun</span>
                    <h2 id="displayNamaAkun" class="text-xl font-bold text-white mb-1">-</h2>
                    <p class="text-xs text-slate-400">Kode Akun: <span id="displayKodeAkun" class="font-mono brand-cyan font-semibold">-</span></p>
                </div>

                <div class="flex items-center space-x-3 w-full sm:w-auto">
                    <!-- Dropdown Filter Satker pada Tampilan Hasil -->
                    <div class="flex flex-col">
                        <label class="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1">
                            <i class="fa-solid fa-building-columns brand-cyan mr-1"></i> Filter Satker Hasil
                        </label>
                        <select id="resultSatkerFilter" class="bg-[#0b0f17] border border-dark text-slate-200 text-xs rounded-lg p-2 focus:border-[#00d2ff] outline-none">
                            <option value="ALL">Semua Satker</option>
                        </select>
                    </div>

                    <button id="btnReset" class="mt-4 sm:mt-0 text-xs font-semibold px-4 py-2 border border-dark rounded-lg hover:bg-slate-800 text-slate-300 flex items-center gap-2">
                        <i class="fa-solid fa-arrow-left"></i> Upload File Lain
                    </button>
                </div>
            </div>

            <!-- TABEL 1 -->
            <div class="bg-card rounded-xl border border-dark overflow-hidden shadow-xl">
                <div class="p-5 border-b border-dark">
                    <h3 class="font-bold text-slate-200 text-sm">Daftar Transaksi Belum Memiliki Pasangan (Debet vs Kredit)</h3>
                    <p class="text-xs text-slate-500">Nilai yang belum memiliki pasangan penihil pada kriteria periode terpilih.</p>
                </div>
                <div class="overflow-x-auto max-h-[450px]">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-[#0b0f17] text-slate-400 uppercase sticky top-0 font-semibold border-b border-dark"><tr id="mainTableHeader"></tr></thead>
                        <tbody id="mainTableBody" class="divide-y divide-dark text-slate-300"></tbody>
                        <tfoot id="mainTableFooter" class="bg-[#0b0f17] font-bold border-t border-dark text-white sticky bottom-0"></tfoot>
                    </table>
                </div>
            </div>

            <!-- TABEL 2 -->
            <div id="resolvedSection" class="hidden bg-card rounded-xl border border-dark overflow-hidden shadow-xl">
                <div class="p-5 border-b border-dark flex items-center justify-between bg-indigo-950/20">
                    <div>
                        <h3 class="font-bold text-indigo-300 text-sm">Daftar Pasangan Penihil (Muncul di Periode Selanjutnya)</h3>
                        <p class="text-xs text-slate-400">Dokumen transaksi di periode selanjutnya yang menjadi pasangan penihil.</p>
                    </div>
                    <span class="bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-[10px] px-2.5 py-1 rounded-full font-mono">Aggregated Pairs</span>
                </div>
                <div class="overflow-x-auto max-h-[450px]">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-[#0b0f17] text-slate-400 uppercase sticky top-0 font-semibold border-b border-dark"><tr id="resolvedTableHeader"></tr></thead>
                        <tbody id="resolvedTableBody" class="divide-y divide-dark text-slate-300"></tbody>
                        <tfoot id="resolvedTableFooter" class="bg-[#0b0f17] font-bold border-t border-dark text-white sticky bottom-0"></tfoot>
                    </table>
                </div>
            </div>

            <!-- TABEL 3 -->
            <div id="hangingSection" class="hidden bg-card rounded-xl border border-dark overflow-hidden shadow-xl">
                <div class="p-5 border-b border-dark flex items-center justify-between bg-rose-950/20">
                    <div>
                        <h3 class="font-bold text-rose-300 text-sm">Daftar Transaksi Menggantung Tanpa Pasangan Pasca Periode X</h3>
                        <p class="text-xs text-slate-400">Transaksi tanpa pasangan setelah periode X yang tanggalnya mendahului transaksi pembuka periode selanjutnya.</p>
                    </div>
                    <span class="bg-rose-500/10 text-rose-400 border border-rose-500/20 text-[10px] px-2.5 py-1 rounded-full font-mono">Hanging Post-Check</span>
                </div>
                <div class="overflow-x-auto max-h-[450px]">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead class="bg-[#0b0f17] text-slate-400 uppercase sticky top-0 font-semibold border-b border-dark"><tr id="hangingTableHeader"></tr></thead>
                        <tbody id="hangingTableBody" class="divide-y divide-dark text-slate-300"></tbody>
                        <tfoot id="hangingTableFooter" class="bg-[#0b0f17] font-bold border-t border-dark text-white sticky bottom-0"></tfoot>
                    </table>
                </div>
            </div>
        </div>
    </main>

    <footer class="border-t border-dark py-6 text-center text-xs text-slate-600">
        <p>Dede Saputra • © 2026 • Reconciliation System</p>
    </footer>

    <script>
        const dropZone = document.getElementById('dropZone');
        const csvFileInput = document.getElementById('csvFile');
        const fileLabel = document.getElementById('fileLabel');
        const btnSubmit = document.getElementById('btnSubmit');
        const uploadForm = document.getElementById('uploadForm');
        const loading = document.getElementById('loading');
        const uploadSection = document.getElementById('uploadSection');
        const resultSection = document.getElementById('resultSection');
        const filterMode = document.getElementById('filterMode');
        const targetPeriod = document.getElementById('targetPeriod');
        const resultSatkerFilter = document.getElementById('resultSatkerFilter');

        const resolvedSection = document.getElementById('resolvedSection');
        const hangingSection = document.getElementById('hangingSection');

        let globalResponseData = null;

        filterMode.addEventListener('change', () => {
            if (filterMode.value === 'ALL') {
                targetPeriod.disabled = true;
                targetPeriod.classList.add('opacity-40');
            } else {
                targetPeriod.disabled = false;
                targetPeriod.classList.remove('opacity-40');
            }
        });

        dropZone.addEventListener('click', () => csvFileInput.click());

        csvFileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                fileLabel.innerHTML = `File terpilih: <span class="brand-cyan font-bold">${e.target.files[0].name}</span>`;
                btnSubmit.disabled = false;
            }
        });

        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!csvFileInput.files[0]) return;

            if (filterMode.value !== 'ALL' && !targetPeriod.value) {
                alert("Silakan pilih Periode X terlebih dahulu!");
                return;
            }

            const formData = new FormData();
            formData.append('file', csvFileInput.files[0]);
            formData.append('filter_mode', filterMode.value);
            formData.append('target_period', targetPeriod.value);

            loading.classList.remove('hidden');
            btnSubmit.disabled = true;

            try {
                const response = await fetch('/reconcile-csv/', {
                    method: 'POST',
                    body: formData
                });

                const res = await response.json();
                loading.classList.add('hidden');

                if (!response.ok) {
                    alert(res.detail || "Terjadi kesalahan saat memproses file.");
                    btnSubmit.disabled = false;
                    return;
                }

                globalResponseData = res;
                populateSatkerDropdown(res.available_satkers);
                applyFilterAndDisplayResults();
            } catch (err) {
                loading.classList.add('hidden');
                btnSubmit.disabled = false;
                alert("Gagal menghubungkan ke server: " + err.message);
            }
        });

        function populateSatkerDropdown(satkers) {
            resultSatkerFilter.innerHTML = '<option value="ALL">Semua Satker</option>';
            satkers.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.innerText = `Satker: ${s}`;
                resultSatkerFilter.appendChild(opt);
            });
        }

        resultSatkerFilter.addEventListener('change', () => {
            applyFilterAndDisplayResults();
        });

        function formatNumberCleanJS(val) {
            const valInt = Math.round(val);
            if (valInt < 0) {
                return `(${Math.abs(valInt).toLocaleString('id-ID')})`;
            }
            return valInt.toLocaleString('id-ID');
        }

        function applyFilterAndDisplayResults() {
            if (!globalResponseData) return;

            const selectedSatker = resultSatkerFilter.value;
            const res = globalResponseData;

            document.getElementById('displayKodeAkun').innerText = res.kode_akun;
            document.getElementById('displayNamaAkun').innerText = res.nama_akun;

            // Filter data per tabel sesuai pilihan dropdown Satker
            const filterFn = (item) => selectedSatker === 'ALL' || String(item['Kode Satker']).trim() === selectedSatker;

            const filteredMain = res.main_data.filter(filterFn);
            const totalMain = filteredMain.reduce((acc, cur) => acc + (cur['Nilai_Clean'] || 0), 0);

            renderTable('mainTableHeader', 'mainTableBody', 'mainTableFooter', res.main_columns, filteredMain, formatNumberCleanJS(totalMain));

            if (res.has_resolved_later) {
                const filteredRes = res.resolved_data.filter(filterFn);
                const totalRes = filteredRes.reduce((acc, cur) => acc + (cur['Nilai_Clean'] || 0), 0);
                renderTable('resolvedTableHeader', 'resolvedTableBody', 'resolvedTableFooter', res.resolved_columns, filteredRes, formatNumberCleanJS(totalRes));
                resolvedSection.classList.remove('hidden');
            } else {
                resolvedSection.classList.add('hidden');
            }

            if (res.has_hanging_after) {
                const filteredHang = res.hanging_data.filter(filterFn);
                const totalHang = filteredHang.reduce((acc, cur) => acc + (cur['Nilai_Clean'] || 0), 0);
                renderTable('hangingTableHeader', 'hangingTableBody', 'hangingTableFooter', res.hanging_columns, filteredHang, formatNumberCleanJS(totalHang));
                hangingSection.classList.remove('hidden');
            } else {
                hangingSection.classList.add('hidden');
            }

            uploadSection.classList.add('hidden');
            resultSection.classList.remove('hidden');
        }

        function renderTable(headerId, bodyId, footerId, columns, data, totalFormatted) {
            const headerTr = document.getElementById(headerId);
            const bodyTb = document.getElementById(bodyId);
            const footerTf = document.getElementById(footerId);

            headerTr.innerHTML = ''; bodyTb.innerHTML = ''; footerTf.innerHTML = '';

            if (data.length === 0) {
                bodyTb.innerHTML = `<tr><td colspan="100%" class="text-center py-8 text-emerald-400 font-medium">Tidak ada data transaksi.</td></tr>`;
                return;
            }

            columns.forEach(col => {
                const th = document.createElement('th');
                th.className = "py-3.5 px-4 border-b border-dark whitespace-nowrap text-[11px] tracking-wider";
                th.innerText = col;
                headerTr.appendChild(th);
            });

            data.forEach(row => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-800/40 transition-colors';

                columns.forEach(col => {
                    const td = document.createElement('td');
                    td.className = "py-3 px-4 whitespace-nowrap border-b border-dark/50";

                    if (col === 'Nilai') {
                        td.innerHTML = `<span class="font-bold text-white font-mono">${row[col]}</span>`;
                    } else if (col === 'Kode Satker') {
                        td.innerHTML = `<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-cyan-950/80 text-cyan-300 border border-cyan-800/50 font-mono">${row[col]}</span>`;
                    } else if (col === 'Kode Periode' || col === 'Kode Periode Pasangan') {
                        td.innerHTML = `<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-slate-300 font-mono">${row[col]}</span>`;
                    } else if (col === 'Keterangan Penyelesaian') {
                        td.innerHTML = `<span class="px-2.5 py-1 rounded text-[10px] font-semibold bg-indigo-950/60 text-indigo-300 border border-indigo-800/50 font-mono"><i class="fa-solid fa-link text-indigo-400 mr-1"></i>${row[col]}</span>`;
                    } else {
                        td.innerText = row[col] !== null ? row[col] : '';
                    }
                    tr.appendChild(td);
                });
                bodyTb.appendChild(tr);
            });

            const footerTr = document.createElement('tr');
            const nilaiColIndex = columns.indexOf('Nilai');

            columns.forEach((col, idx) => {
                const td = document.createElement('td');
                td.className = "py-3.5 px-4 uppercase text-xs";

                if (idx === 0) {
                    td.innerText = "TOTAL";
                } else if (idx === nilaiColIndex) {
                    td.innerHTML = `<span class="font-mono brand-cyan font-bold text-sm">${totalFormatted}</span>`;
                } else {
                    td.innerText = "";
                }
                footerTr.appendChild(td);
            });
            footerTf.appendChild(footerTr);
        }

        document.getElementById('btnReset').addEventListener('click', () => {
            csvFileInput.value = '';
            fileLabel.innerHTML = 'Unggah File CSV / Buku Besar Excel (.xlsx)';
            btnSubmit.disabled = true;
            globalResponseData = null;
            resultSection.classList.add('hidden');
            uploadSection.classList.remove('hidden');
        });
    </script>
</body>
</html>"""
    return HTMLResponse(content=html_content)

@app.post("/reconcile-csv/")
async def reconcile_csv(
    file: UploadFile = File(...),
    filter_mode: str = Form("ALL"),
    target_period: str = Form("")
):
    filename = file.filename.lower()
    contents = await file.read()

    try:
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            df, _, _ = parse_sakti_excel(contents)
            return JSONResponse(content=process_reconciliation(df, filter_mode, target_period, is_sakti_excel=True))
        elif filename.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(contents), encoding='utf-8')
                if df.shape[1] < 5:
                    df = pd.read_csv(io.BytesIO(contents), encoding='utf-8', sep=';')
            except Exception:
                df = pd.read_csv(io.BytesIO(contents), encoding='latin1', sep=None, engine='python')

            return JSONResponse(content=process_reconciliation(df, filter_mode, target_period, is_sakti_excel=False))
        else:
            raise HTTPException(status_code=400, detail="Format file harus berupa .csv, .xlsx, atau .xls")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses file: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
