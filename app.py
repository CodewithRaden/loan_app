from flask import Flask, render_template, request, session, send_file, redirect, url_for
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, ROUND_HALF_UP
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT

app = Flask(__name__)
app.secret_key = "random-secret"


def fmt(x: float) -> str:
    if abs(x) < 0.5:  # toleransi setengah rupiah
        x = 0
    return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def rupiah_round(x: float) -> int:
    """Pembulatan ke rupiah (0 desimal) dengan ROUND_HALF_UP (gaya Excel)."""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def build_schedule(
    pokok: float,
    bunga_tahunan: float,
    tenor: int,
    metode: str,
    start_date: datetime,
    first_due_date: datetime,
):
    i_bulanan = bunga_tahunan / 12.0

    # --- bunga prorata untuk periode pertama ---
    if start_date.day == first_due_date.day:
        # Sama tanggal → pakai 30/360 (anggap 1 bulan penuh)
        selisih_hari = 30
        bunga1_actual = pokok * bunga_tahunan * (30 / 360.0)
    else:
        # Beda tanggal → pakai selisih aktual
        selisih_hari = max(0, (first_due_date - start_date).days)
        bunga1_actual = pokok * bunga_tahunan * (selisih_hari / 360.0)

    data = []
    data.append([0, start_date.strftime("%d %b %Y"), 0, 0, 0, rupiah_round(pokok)])
    sisa = pokok
    penjelasan = ""

    # ----------------------
    # PERHITUNGAN ANGSURAN
    # ----------------------
    if metode == "efektif":
        pmt = (
            pokok
            * (i_bulanan * (1 + i_bulanan) ** tenor)
            / ((1 + i_bulanan) ** tenor - 1)
        )
        bunga1_std = sisa * i_bulanan
        pokok1 = pmt - bunga1_std
        tambahan_bunga = bunga1_actual - bunga1_std
        total1 = pmt + tambahan_bunga
        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= pokok1
        data.append([1, jatuh1, pokok1, bunga1_actual, total1, sisa])

        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            bunga = sisa * i_bulanan
            pokok_bayar = pmt - bunga
            sisa -= pokok_bayar
            data.append([bulan, jatuh, pokok_bayar, bunga, pmt, sisa])

        penjelasan = (
            "🔹 Metode Anuitas\n"
            f"i bulanan = {bunga_tahunan:.2%} / 12 = {i_bulanan:.6f}\n"
            f"PMT = P × [ i(1+i)^n / ((1+i)^n - 1) ] = {fmt(pmt)}\n\n"
            f"Periode pertama (selisih {selisih_hari} hari):\n"
            f" Bunga actual = {fmt(pokok)} × {bunga_tahunan:.2%} × ({selisih_hari}/360) = {fmt(bunga1_actual)}\n"
            f" Bunga standar = {fmt(pokok)} × {i_bulanan:.6f} = {fmt(bunga1_std)}\n"
            f" Pokok dibayar = PMT − Bunga standar = {fmt(pmt)} − {fmt(bunga1_std)} = {fmt(pokok1)}\n"
            f" Tambahan (penyesuaian hari) = {fmt(bunga1_actual)} − {fmt(bunga1_std)} = {fmt(tambahan_bunga)}\n"
            f" Total angsuran bln-1 = PMT + Tambahan = {fmt(pmt)} + {fmt(tambahan_bunga)} = {fmt(total1)}\n"
        )

    elif metode == "flat":
        cicilan_pokok = pokok / tenor
        bunga_flat_bulanan = pokok * (bunga_tahunan / 12)  # bunga per bulan tetap
        total_bunga = bunga_flat_bulanan * tenor  # total bunga

        for bulan in range(1, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            sisa -= cicilan_pokok
            bunga = bunga_flat_bulanan
            total = cicilan_pokok + bunga
            data.append([bulan, jatuh, cicilan_pokok, bunga, total, sisa])

    else:
        raise ValueError("Metode tidak dikenal")

    # ----------------------
    # DataFrame awal
    # ----------------------
    df = pd.DataFrame(
        data,
        columns=[
            "Bulan",
            "Tanggal",
            "Angsuran Pokok",
            "Bunga",
            "Total Angsuran",
            "Sisa Pokok",
        ],
    )

    # ==========================================================
    # ADJUSTMENT 1: Sisa pokok dibulatkan ke bunga terakhir
    # ==========================================================
    last_idx = len(df) - 1
    residu_pokok = df.at[last_idx, "Sisa Pokok"]
    if abs(residu_pokok) > 0.000001:
        tambah = residu_pokok
        df.at[last_idx, "Bunga"] += tambah
        df.at[last_idx, "Total Angsuran"] += tambah
        df.at[last_idx, "Sisa Pokok"] = 0.0

    # ==========================================================
    # ADJUSTMENT 2: Sesuaikan agar konsisten dengan Excel rounding
    # ==========================================================
    total_angsuran_rounded = df["Total Angsuran"].apply(rupiah_round).sum()
    total_bunga_rounded = df["Bunga"].apply(rupiah_round).sum()
    target_bunga = total_angsuran_rounded - rupiah_round(pokok)
    selisih = target_bunga - total_bunga_rounded

    if selisih != 0:
        df.at[last_idx, "Bunga"] += selisih
        df.at[last_idx, "Total Angsuran"] += selisih
        # Sisa Pokok tetap 0

    # ----------------------
    # Summary (pakai angka rounded)
    # ----------------------
    summary = {
        "total_pokok": df["Angsuran Pokok"].apply(rupiah_round).sum(),
        "total_bunga": df["Bunga"].apply(rupiah_round).sum(),
        "total_angsuran": df["Total Angsuran"].apply(rupiah_round).sum(),
    }

    return df, penjelasan, summary


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # ambil input
        pokok_str = request.form["pokok"]
        pokok = float(pokok_str.replace(".", "").replace(",", ""))

        bunga_tahunan = float(request.form["bunga"]) / 100.0
        tenor = int(request.form["tenor"])
        metode = request.form["metode"]
        namecstm = request.form.get("namecstm", "Customer")

        # tanggal
        start_date_str = request.form.get("start_date")
        start_date = (
            datetime.strptime(start_date_str, "%Y-%m-%d")
            if start_date_str
            else datetime.today()
        )

        first_due_date_str = request.form.get("first_due_date")
        first_due_date = (
            datetime.strptime(first_due_date_str, "%Y-%m-%d")
            if first_due_date_str
            else start_date + relativedelta(months=1)
        )

        session["params"] = {
            "pokok": pokok,
            "bunga_tahunan": bunga_tahunan,
            "tenor": tenor,
            "metode": metode,
            "start_date": start_date.strftime("%Y-%m-%d"),
            "first_due_date": first_due_date.strftime("%Y-%m-%d"),
            "namecstm": namecstm,
        }

        df, penjelasan, summary = build_schedule(
            pokok, bunga_tahunan, tenor, metode, start_date, first_due_date
        )

        # format angka untuk HTML
        df_show = df.copy()
        for col in ["Angsuran Pokok", "Bunga", "Total Angsuran", "Sisa Pokok"]:
            df_show[col] = df_show[col].map(fmt)

        summary_fmt = {k: fmt(v) for k, v in summary.items()}

        return render_template(
            "result.html",
            tables=[
                df_show.to_html(
                    classes="table table-striped table-bordered",
                    index=False,
                    escape=False,
                )
            ],
            title="Hasil Simulasi",
            penjelasan=penjelasan,
            summary=summary_fmt,
        )

    return render_template("index.html")


def format_number(val, is_bulan=False):
    if is_bulan:  # khusus untuk kolom Bulan
        return str(int(val))
    if isinstance(val, (int, float)):
        return f"{val:,.2f}"
    return str(val)


@app.route("/export_pdf")
def export_pdf():
    params = session.get("params")
    if not params:
        return redirect(url_for("index"))

    pokok = float(params["pokok"])
    bunga_tahunan = float(params["bunga_tahunan"])
    tenor = int(params["tenor"])
    metode = params["metode"]
    start_date = datetime.strptime(params["start_date"], "%Y-%m-%d")
    first_due_date = datetime.strptime(params["first_due_date"], "%Y-%m-%d")
    namecstm = params.get("namecstm", "Customer")

    # Ambil data schedule
    df, _, _ = build_schedule(
        pokok, bunga_tahunan, tenor, metode, start_date, first_due_date
    )

    # Rename kolom
    df = df.rename(
        columns={
            "No": "Bulan",
            "Tanggal": "Tanggal",
            "Jumlah Pokok": "Angsuran Pokok",
            "Jumlah Tingkat Pengembalian": "Bunga",
            "Jumlah Pembayaran": "Total Angsuran",
            "Saldo Pembiayaan": "Sisa Pokok",
        }
    )

    # Format isi tabel
    data = [df.columns.tolist()]
    for _, row in df.iterrows():
        formatted_row = []
        for col, v in zip(df.columns, row.values):
            if col == "Bulan":
                formatted_row.append(format_number(v, is_bulan=True))
            else:
                formatted_row.append(format_number(v))
        data.append(formatted_row)

    # Tambah baris total
    total_row = ["T o t a l", ""]
    for col in df.columns[2:]:
        if col == "Sisa Pokok":
            total_row.append("")
        elif pd.api.types.is_numeric_dtype(df[col]):
            total_row.append(format_number(df[col].sum()))
        else:
            total_row.append("")
    data.append(total_row)

    # === PDF ===
    output = io.BytesIO()

    col_widths = [40, 100, 100, 100, 100, 100]
    table_total_width = sum(col_widths)

    # Atur margin kiri supaya tabel dan header rata kiri
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        leftMargin=20,
        rightMargin=30,
        topMargin=20,
        bottomMargin=18,
    )

    elements = []
    styles = getSampleStyleSheet()
    header_style_big = ParagraphStyle(
        "header_big",
        parent=styles["Normal"],
        fontSize=12,  # ukuran font
        leading=14,  # jarak antar baris
        spaceAfter=6,  # jarak bawah tiap paragraf
        fontName="Helvetica-Bold",
    )

    header_bold = ParagraphStyle(
        "header_bold",
        parent=styles["Normal"],
        fontSize=12,
        leading=14,
        spaceAfter=6,
        fontName="Helvetica-Bold",
    )

    # HEADER pakai Table biar sejajar dengan tabel utama
    header_data = [
        [Paragraph("<b>Lampiran - II</b>", header_bold)],
        [Paragraph(f"{namecstm}", header_style_big)],
        [
            Paragraph(
                "Pembayaran Pokok Pembiayaan dan Tingkat Pengembalian", header_style_big
            )
        ],
        [Paragraph(f"Jumlah Pembiayaan : Rp. {pokok:,.2f}", header_style_big)],
        [
            Paragraph(
                f"Tingkat Pengembalian : {bunga_tahunan*100:.2f} %", header_style_big
            )
        ],
        [Paragraph(f"Periode Pembiayaan : {tenor} bulan", header_style_big)],
    ]
    header_table = Table(header_data, colWidths=[table_total_width])
    header_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Spacer(1, 12))

    # TABEL utama
    table = Table(data, repeatRows=1, colWidths=col_widths)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),  # header
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),  # header rata tengah
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),  # kolom bulan center
                ("ALIGN", (1, 0), (1, -1), "CENTER"),  # kolom tanggal kiri
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),  # angka kanan
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                # Style khusus baris total
                ("BACKGROUND", (0, -1), (-1, -1), colors.whitesmoke),
                ("SPAN", (0, -1), (1, -1)),
                ("ALIGN", (0, -1), (1, -1), "CENTER"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]
        )
    )

    elements.append(table)
    doc.build(elements)
    output.seek(0)

    safe_name = "".join(c if c.isalnum() else "_" for c in namecstm)
    return send_file(
        output,
        download_name=f"TABEL_ANGSURAN_{safe_name}.pdf",
        as_attachment=True,
        mimetype="application/pdf",
    )


if __name__ == "__main__":
    app.run(debug=True)
