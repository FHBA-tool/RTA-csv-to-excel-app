import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import csv
import io
import re
import math
import shutil
import tempfile
import os
import datetime
import zipfile

st.set_page_config(
    page_title="RTA CSV → Excel",
    page_icon="📊",
    layout="centered"
)

st.title("📊 RTA CSV → Excel 変換ツール")
st.markdown("---")

tab_pink, tab_sweep = st.tabs(["🎵 Pink Noise", "📈 Sin Sweep"])


# ═══════════════════════════════════════════════════════════════════════════
# Sin Sweep チャート生成ロジック
# ═══════════════════════════════════════════════════════════════════════════

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

_LABEL_FREQS = [10, 30, 50, 70, 100, 300, 500, 700,
                1000, 3000, 5000, 7000, 10000, 20000]

_PALETTE = ["4472C4", "ED7D31", "70AD47", "C0504D",
            "5B9BD5", "92D050", "FFC000", "A46BBE"]


def _freq_to_label(v):
    if v >= 1000:
        k = v / 1000
        return f"{int(k)}k" if k == int(k) else f"{k:g}k"
    return str(int(v)) if v == int(v) else f"{v:g}"


def parse_sine_csv_bytes(content: bytes) -> dict:
    data = {}
    in_data = header_skipped = False
    for line in content.decode("utf-8-sig", errors="replace").splitlines():
        line = line.strip()
        if line == "[DATA]":
            in_data = True; continue
        if not in_data: continue
        if line.startswith("COMMENT="): continue
        if not header_skipped:
            header_skipped = True; continue
        parts = line.split(",")
        if len(parts) < 2: continue
        try:
            freq, spl = float(parts[0]), float(parts[1])
            if 10 <= freq <= 25000:
                data[freq] = spl
        except ValueError:
            continue
    return data


def _thin_border():
    t = Side(border_style="thin", color="000000")
    return Border(left=t, right=t, top=t, bottom=t)


def _cell(ws, r, c, v, bold=False, fill=None, size=9):
    cell = ws.cell(r, c, v)
    cell.font      = Font(name="Meiryo UI", bold=bold, size=size)
    cell.border    = _thin_border()
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    if fill:
        cell.fill = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
    return cell


def _make_sweep_sheet(wb, group_name, entries, settings):
    ws = wb.create_sheet(group_name)
    FREQ_COL = 4; DATA_COL = 5; DATA_START = 40

    date_val = settings.get("date", datetime.date.today().isoformat())
    place    = settings.get("place", "")
    car      = settings.get("car", "")
    for i, (lbl, val) in enumerate([
        ("Date",   date_val),
        ("Place",  place),
        ("Car",    car),
        ("Signal", settings.get("signal", "Sin Sweep")),
        ("Mic",    settings.get("mic", "")),
    ]):
        r = 2 + i
        ws.row_dimensions[r].height = 16
        _cell(ws, r, 15, lbl, bold=True, fill="D9D9D9")
        _cell(ws, r, 16, val)

    ws.cell(37, FREQ_COL, "sweep data")
    _cell(ws, 38, FREQ_COL, "", bold=True, fill="D9D9D9")
    _cell(ws, 39, FREQ_COL, "Freq [Hz]", bold=True, fill="D9D9D9")
    for j, e in enumerate(entries):
        _cell(ws, 38, DATA_COL + j, e["level"], bold=True, fill="D9D9D9")
        _cell(ws, 39, DATA_COL + j, "SPL [dB]", bold=True, fill="D9D9D9")
        ws.column_dimensions[get_column_letter(DATA_COL + j)].width = 9
    ws.column_dimensions[get_column_letter(FREQ_COL)].width = 11

    all_freqs = sorted(set().union(*[set(e["freq_spl"].keys()) for e in entries]))
    df = Font(name="Meiryo UI", size=8)
    for i, freq in enumerate(all_freqs):
        row = DATA_START + i
        fc = ws.cell(row, FREQ_COL, freq)
        fc.border = _thin_border(); fc.font = df; fc.number_format = "0.000"
        for j, e in enumerate(entries):
            spl = e["freq_spl"].get(freq)
            sc = ws.cell(row, DATA_COL + j, round(spl, 4) if spl else None)
            sc.border = _thin_border(); sc.font = df
            if spl: sc.number_format = "0.00"

    return ws, all_freqs


def _add_sweep_chart(wb, ws, group_name, entries):
    n_lbl = len(_LABEL_FREQS)
    cd_name = f"_CD_{group_name}"[:31]
    if cd_name in wb.sheetnames: del wb[cd_name]
    cd = wb.create_sheet(cd_name); cd.sheet_state = "hidden"

    si = sorted(entries[0]["freq_spl"].items())
    for i, (f, s) in enumerate(si):
        cd.cell(i + 2, 1, f); cd.cell(i + 2, 2, s)
    n_main = len(si)

    for i, f in enumerate(_LABEL_FREQS):
        cd.cell(i + 2, 3, f); cd.cell(i + 2, 4, 21.0)

    chart = ScatterChart()
    chart.scatterStyle = "line"
    chart.title = "ASA-10 Sine Sweep"
    chart.height = 14; chart.width = 24

    chart.y_axis.scaling.min = 20; chart.y_axis.scaling.max = 110
    chart.y_axis.majorUnit   = 10
    chart.y_axis.title       = "SPL [dB]"
    chart.y_axis.majorGridlines = ChartLines()
    chart.y_axis.delete      = False

    chart.x_axis.scaling.logBase = 10
    chart.x_axis.scaling.min     = 10
    chart.x_axis.scaling.max     = 20000
    chart.x_axis.title           = "Frequency [Hz]"
    chart.x_axis.majorGridlines  = ChartLines()
    chart.x_axis.minorGridlines  = ChartLines()
    chart.x_axis.delete          = False

    x1 = Reference(cd, min_col=1, min_row=2, max_row=n_main + 1)
    y1 = Reference(cd, min_col=2, min_row=1, max_row=n_main + 1)
    s1 = Series(y1, x1, title_from_data=True)
    s1.graphicalProperties.line.solidFill = _PALETTE[0]
    s1.graphicalProperties.line.width     = 15240
    chart.series.append(s1)

    for k, entry in enumerate(entries[1:], 1):
        si_k = sorted(entry["freq_spl"].items())
        cx, cy = 5 + (k-1)*2, 6 + (k-1)*2
        for i, (f, s) in enumerate(si_k):
            cd.cell(i+2, cx, f); cd.cell(i+2, cy, s)
        nk = len(si_k)
        sk = Series(
            Reference(cd, min_col=cy, min_row=2, max_row=nk+1),
            Reference(cd, min_col=cx, min_row=2, max_row=nk+1)
        )
        sk.graphicalProperties.line.solidFill = _PALETTE[k % len(_PALETTE)]
        sk.graphicalProperties.line.width     = 15240
        chart.series.append(sk)

    x2 = Reference(cd, min_col=3, min_row=2, max_row=n_lbl+1)
    y2 = Reference(cd, min_col=4, min_row=2, max_row=n_lbl+1)
    chart.series.append(Series(y2, x2))

    ws.add_chart(chart, "S7")


def _patch_xml(x):
    rpr = '<a:rPr sz="1100" b="0"><a:latin typeface="MS Gothic"/></a:rPr>'
    x = x.replace('<a:r><a:t>ASA-10 Sine Sweep</a:t></a:r>',
                  f'<a:r>{rpr}<a:t>ASA-10 Sine Sweep</a:t></a:r>', 1)
    x = x.replace('<a:defRPr/>',
                  '<a:defRPr sz="900" b="0"><a:latin typeface="MS Gothic"/></a:defRPr>')
    x = x.replace('</tx></title>', '</tx><overlay val="0"/></title>')
    x = x.replace('<axPos val="l"/>', '<axPos val="b"/>', 1)

    major_sp = (f'<majorGridlines><spPr><a:ln xmlns:a="{_A}" w="3810">'
                '<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>'
                '</a:ln></spPr></majorGridlines>')
    x = x.replace('<majorGridlines/>', major_sp)

    minor_sp = (f'<minorGridlines><spPr><a:ln xmlns:a="{_A}" w="2540">'
                '<a:solidFill><a:srgbClr val="EBEBEB"/></a:solidFill>'
                '</a:ln></spPr></minorGridlines>')
    x = x.replace('<minorGridlines/>', minor_sp)

    txpr_x = (f'<txPr><a:bodyPr xmlns:a="{_A}"/><a:lstStyle xmlns:a="{_A}"/>'
              f'<a:p xmlns:a="{_A}"><a:pPr>'
              f'<a:defRPr sz="2800" b="0">'
              f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
              f'<a:latin typeface="MS Gothic"/></a:defRPr>'
              f'</a:pPr></a:p></txPr>')
    x = re.sub(
        r'(<valAx>(?:(?!</valAx>).)*<axPos val="b"/>)((?:(?!</valAx>).)*)(</valAx>)',
        lambda m: m.group(1) + txpr_x + m.group(2) + m.group(3),
        x, count=1, flags=re.DOTALL
    )

    txpr_y = (f'<txPr><a:bodyPr xmlns:a="{_A}"/><a:lstStyle xmlns:a="{_A}"/>'
              f'<a:p xmlns:a="{_A}"><a:pPr>'
              f'<a:defRPr sz="800" b="0"><a:latin typeface="MS Gothic"/></a:defRPr>'
              f'</a:pPr></a:p></txPr>')
    x = re.sub(
        r'(<valAx>(?:(?!</valAx>).)*<axPos val="l"/>(?:(?!</valAx>).)*)(</valAx>)',
        lambda m: m.group(1) + txpr_y + m.group(2),
        x, count=1, flags=re.DOTALL
    )

    def make_dlbl(idx, text):
        rpr = f'<a:rPr xmlns:a="{_A}" sz="700" b="0"><a:latin typeface="MS Gothic"/></a:rPr>'
        return (f'<dLbl><idx val="{idx}"/>'
                f'<spPr><a:noFill xmlns:a="{_A}"/><a:ln xmlns:a="{_A}"><a:noFill/></a:ln></spPr>'
                f'<tx><rich>'
                f'<a:bodyPr xmlns:a="{_A}"/><a:lstStyle xmlns:a="{_A}"/>'
                f'<a:p xmlns:a="{_A}"><a:r>{rpr}<a:t>{text}</a:t></a:r></a:p>'
                f'</rich></tx>'
                f'<showLegendKey val="0"/><showVal val="0"/>'
                f'<showCatName val="0"/><showSerName val="0"/>'
                f'<showPercent val="0"/><showBubbleSize val="0"/>'
                f'<dLblPos val="b"/></dLbl>')

    dlbls = ('<dLbls>'
             + ''.join(make_dlbl(i, _freq_to_label(f)) for i, f in enumerate(_LABEL_FREQS))
             + '<showLegendKey val="0"/><showVal val="0"/>'
             '<showCatName val="0"/><showSerName val="0"/>'
             '<showPercent val="0"/><showBubbleSize val="0"/>'
             '</dLbls>')
    sers = list(re.finditer(r'<ser>.*?</ser>', x, flags=re.DOTALL))
    if sers:
        start, end = sers[-1].start(), sers[-1].end()
        new_ser = re.sub(r'<spPr>.*?</spPr>',
                         f'<spPr><a:ln xmlns:a="{_A}"><a:noFill/></a:ln></spPr>',
                         x[start:end], count=1, flags=re.DOTALL)
        new_ser = new_ser.replace('<xVal>', dlbls + '<xVal>', 1)
        x = x[:start] + new_ser + x[end:]

    x = re.sub(r'<legend>.*?</legend>', '', x, flags=re.DOTALL)

    x = x.replace('<plotArea>', '<plotArea>'
                  '<layout><manualLayout><layoutTarget val="inner"/>'
                  '<xMode val="factor"/><yMode val="factor"/>'
                  '<x val="0.06"/><y val="0.04"/><w val="0.88"/><h val="0.80"/>'
                  '</manualLayout></layout>', 1)
    x = x.replace('</plotArea>',
                  f'<spPr><a:noFill xmlns:a="{_A}"/>'
                  f'<a:ln xmlns:a="{_A}" w="6350">'
                  '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>'
                  '</a:ln></spPr></plotArea>', 1)
    x = x.replace('</chartSpace>',
                  f'<spPr><a:solidFill xmlns:a="{_A}"><a:srgbClr val="FFFFFF"/></a:solidFill>'
                  f'<a:ln xmlns:a="{_A}"><a:noFill/></a:ln></spPr></chartSpace>', 1)
    return x


def build_sweep_excel_bytes(groups, settings) -> bytes:
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        from openpyxl import Workbook
        wb = Workbook(); wb.remove(wb.active)
        for grp in groups:
            _make_sweep_sheet(wb, grp["name"], grp["entries"], settings)
        wb.save(tmp)

        wb2 = openpyxl.load_workbook(tmp)
        for grp in groups:
            _add_sweep_chart(wb2, wb2[grp["name"]], grp["name"], grp["entries"])
        wb2.save(tmp)

        # XML パッチ
        patched = tmp + ".p"
        shutil.copy2(tmp, patched)
        with zipfile.ZipFile(patched, 'r') as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r'xl/charts/chart\d+\.xml$', item.filename):
                    data = _patch_xml(data.decode('utf-8')).encode('utf-8')
                zout.writestr(item, data)
        os.remove(patched)

        with open(tmp, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


# ═══════════════════════════════════════════════════════════════════════════
# Tab 1: Pink Noise（既存）
# ═══════════════════════════════════════════════════════════════════════════

with tab_pink:
    st.header("① Excelファイル設定")

    template_file = st.file_uploader(
        "テンプレートExcelをアップロード",
        type=["xlsx", "xlsm"],
        key="template"
    )
    output_filename = st.text_input("出力ファイル名", value="出力ファイル.xlsx")
    if not output_filename.endswith(".xlsx"):
        output_filename += ".xlsx"

    st.markdown("---")
    st.header("② 測定条件")

    col1, col2 = st.columns(2)
    with col1:
        place = st.selectbox("Place", ["FF garage", "FF parking", "その他"], key="pn_place")
        if place == "その他":
            place = st.text_input("Place（入力）", key="pn_place_custom")
        car = st.selectbox("Car", ["FF B33 demo car", "FF B25 demo car", "FE B33 demo car",
                                    "FE B25 demo car", "FHP B33 demo car", "その他"], key="pn_car")
        if car == "その他":
            car = st.text_input("Car（入力）", key="pn_car_custom")
        dsp = st.text_input("DSP data", key="pn_dsp")
        filter_val = st.selectbox("filter", ["bypass", "ON"], key="pn_filter")
        gain_values = []
        v = 5.0
        while v >= -30.0:
            gain_values.append(round(v, 2)); v = round(v - 0.25, 2)
        gain = st.selectbox("gain", options=gain_values, index=gain_values.index(0.0),
                             format_func=lambda x: f"{x:.2f} dB", key="pn_gain")
    with col2:
        eq = st.selectbox("EQ", ["bypass", "ON"], key="pn_eq")
        signal = st.selectbox("Signal", ["Pink Noise", "Sin wave", "White Noise"], key="pn_signal")
        mic_pos = st.selectbox("Mic position",
                               ["FR Sheet", "FL Sheet", "RR Sheet", "RL Sheet", "その他"],
                               key="pn_mic_pos")
        if mic_pos == "その他":
            mic_pos = st.text_input("Mic position（入力）", key="pn_mic_pos_custom")
        mic = st.selectbox("Mic", ["Binaural", "1 point"], key="pn_mic")

    st.markdown("---")
    st.header("③ CSVファイル設定")
    st.subheader("フロアノイズCSV")
    floor_csv_file = st.file_uploader("フロアノイズCSVをアップロード（1ファイルのみ）",
                                       type=["csv"], key="floor_csv")
    st.subheader("測定データCSV")
    data_csv_files = st.file_uploader("測定データCSVをアップロード（複数可）",
                                       type=["csv"], accept_multiple_files=True, key="data_csv")
    if data_csv_files:
        st.write("**選択されたファイル:**")
        for i, f in enumerate(data_csv_files):
            st.write(f"　{i+1}. {f.name}")

    st.markdown("---")

    def read_csv_from_upload(uploaded_file):
        content = uploaded_file.read().decode("utf-8-sig")
        return list(csv.reader(content.splitlines()))

    def write_range(ws, start_row, start_col, data):
        if not data or not data[0]: return
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                ws.cell(row=start_row + i, column=start_col + j, value=value)

    def process_csv(ws, csv_data, csv_filename, place, car, dsp, filter_val,
                    gain, eq, signal, mic_pos, mic, is_floor_noise=False):
        rows_1to3 = csv_data[0:3]
        write_range(ws, 39, 4, list(map(list, zip(*rows_1to3))))
        if len(csv_data) >= 5:
            ws.cell(row=2, column=20, value=csv_data[4][0])
        ws.cell(row=3, column=20, value=place)
        ws.cell(row=4, column=20, value=car)
        ws.cell(row=5, column=20, value=csv_filename)
        if not is_floor_noise:
            ws.cell(row=2, column=22, value=dsp)
            ws.cell(row=3, column=22, value=filter_val)
            ws.cell(row=4, column=22, value=gain)
            ws.cell(row=5, column=22, value=eq)
            ws.cell(row=2, column=24, value=signal)
        ws.cell(row=3, column=24, value=mic_pos)
        ws.cell(row=4, column=24, value=mic)

    if st.button("🚀 実行！", use_container_width=True, type="primary", key="pn_run"):
        errors = []
        if not template_file: errors.append("テンプレートExcelをアップロードしてください")
        if not place:         errors.append("Placeを入力してください")
        if not car:           errors.append("Carを入力してください")
        if not mic_pos:       errors.append("Mic positionを入力してください")
        if not data_csv_files: errors.append("測定データCSVをアップロードしてください")
        if not floor_csv_file: errors.append("フロアノイズCSVをアップロードしてください")

        if errors:
            for e in errors: st.error(e)
        else:
            with st.spinner("⏳ ファイル生成中..."):
                try:
                    wb = openpyxl.load_workbook(io.BytesIO(template_file.read()))
                    if "テンプレート" not in wb.sheetnames:
                        st.error(f"「テンプレート」シートがありません: {wb.sheetnames}"); st.stop()
                    if "フロアノイズ" not in wb.sheetnames:
                        st.error(f"「フロアノイズ」シートがありません: {wb.sheetnames}"); st.stop()

                    wb["フロアノイズ"].title = "フロアノイズ_元"
                    template_ws = wb["テンプレート"]

                    for entry in data_csv_files:
                        sheet_name = os.path.splitext(entry.name)[0][:31]
                        entry.seek(0)
                        new_ws = wb.copy_worksheet(template_ws)
                        new_ws.title = sheet_name
                        process_csv(new_ws, read_csv_from_upload(entry), entry.name,
                                    place, car, dsp, filter_val, gain, eq,
                                    signal, mic_pos, mic)

                    floor_ws_orig = wb["フロアノイズ_元"]
                    new_floor_ws = wb.copy_worksheet(floor_ws_orig)
                    new_floor_ws.title = "フロアノイズ"
                    floor_csv_file.seek(0)
                    process_csv(new_floor_ws, read_csv_from_upload(floor_csv_file),
                                floor_csv_file.name, place, car, dsp, filter_val,
                                gain, eq, signal, mic_pos, mic, is_floor_noise=True)

                    del wb["テンプレート"], wb["フロアノイズ_元"]
                    sheets = wb.sheetnames
                    floor_idx = sheets.index("フロアノイズ")
                    wb.move_sheet("フロアノイズ", offset=len(sheets) - 1 - floor_idx)

                    output = io.BytesIO(); wb.save(output); output.seek(0)
                    st.success("✅ 完了！")
                    st.download_button("📥 Excelファイルをダウンロード",
                                       data=output.getvalue(),
                                       file_name=output_filename,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                except Exception as e:
                    import traceback
                    st.error(f"エラー:\n{traceback.format_exc()}")


# ═══════════════════════════════════════════════════════════════════════════
# Tab 2: Sin Sweep
# ═══════════════════════════════════════════════════════════════════════════

with tab_sweep:
    st.header("① CSVファイル設定")
    st.caption("ファイル名の末尾を `グループ名_レベル名.csv` の形式にすると自動認識します（例: ASA-10_1W.csv）")

    sw_csv_files = st.file_uploader(
        "スイープ測定CSVをアップロード（複数可）",
        type=["csv"], accept_multiple_files=True, key="sw_csv"
    )

    if sw_csv_files:
        st.write("**検出されたファイル:**")
        rows = []
        for f in sw_csv_files:
            name = os.path.splitext(f.name)[0]
            parts = name.rsplit("_", 1)
            group = parts[0] if len(parts) == 2 else name
            level = parts[1] if len(parts) == 2 else ""
            rows.append({"ファイル名": f.name, "グループ（シート名）": group, "レベル": level})
        st.table(rows)

    st.markdown("---")
    st.header("② 測定条件")

    col1, col2 = st.columns(2)
    with col1:
        sw_place = st.selectbox("Place", ["FF garage", "FF parking", "その他"], key="sw_place")
        if sw_place == "その他":
            sw_place = st.text_input("Place（入力）", key="sw_place_custom")
        sw_car = st.selectbox("Car", ["FF B33 demo car", "FF B25 demo car", "FE B33 demo car",
                                       "FE B25 demo car", "FHP B33 demo car", "その他"], key="sw_car")
        if sw_car == "その他":
            sw_car = st.text_input("Car（入力）", key="sw_car_custom")
    with col2:
        sw_mic_pos = st.selectbox("Mic position",
                                   ["FR Sheet", "FL Sheet", "RR Sheet", "RL Sheet", "その他"],
                                   key="sw_mic_pos")
        if sw_mic_pos == "その他":
            sw_mic_pos = st.text_input("Mic position（入力）", key="sw_mic_pos_custom")
        sw_mic = st.selectbox("Mic", ["Binaural", "1 point"], key="sw_mic")

    st.markdown("---")
    st.header("③ 出力設定")

    sw_output_name = st.text_input("出力ファイル名", value="出力結果_Sweep.xlsx", key="sw_out")
    if not sw_output_name.endswith(".xlsx"):
        sw_output_name += ".xlsx"

    st.markdown("---")

    if st.button("🚀 Excel生成！", use_container_width=True, type="primary", key="sw_run"):
        if not sw_csv_files:
            st.error("CSVファイルをアップロードしてください")
        else:
            with st.spinner("⏳ チャート生成中...（対数スケールグラフを構築しています）"):
                try:
                    groups_dict = {}
                    bad = []
                    for f in sw_csv_files:
                        f.seek(0)
                        freq_spl = parse_sine_csv_bytes(f.read())
                        if not freq_spl:
                            bad.append(f.name); continue
                        name = os.path.splitext(f.name)[0]
                        parts = name.rsplit("_", 1)
                        group = parts[0] if len(parts) == 2 else name
                        level = parts[1] if len(parts) == 2 else "1W"
                        groups_dict.setdefault(group, []).append(
                            {"level": level, "freq_spl": freq_spl})

                    if bad:
                        for b in bad: st.warning(f"スキップ（形式不明）: {b}")
                    if not groups_dict:
                        st.error("有効なCSVデータがありません"); st.stop()

                    groups = [{"name": k, "entries": v} for k, v in groups_dict.items()]
                    settings = {
                        "place": sw_place, "car": sw_car,
                        "signal": "Sin Sweep",
                        "mic": sw_mic, "mic_pos": sw_mic_pos,
                        "date": datetime.date.today().isoformat(),
                    }

                    xlsx_bytes = build_sweep_excel_bytes(groups, settings)
                    st.success("✅ 完了！")
                    st.download_button(
                        "📥 Excelファイルをダウンロード",
                        data=xlsx_bytes,
                        file_name=sw_output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    import traceback
                    st.error(f"エラー:\n{traceback.format_exc()}")
