from flask import Flask, render_template, request, session, send_file, redirect, url_for
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

app = Flask(__name__)
app.secret_key = "random-secret"


def fmt(x: float) -> str:
    """Format angka ke gaya Indonesia tanpa desimal."""
    return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def build_schedule(
    pokok: float,
    bunga_tahunan: float,
    tenor: int,
    metode: str,
    start_date: datetime,
    first_due_date: datetime,
):
    """
    Bangun DataFrame jadwal angsuran + string penjelasan,
    sesuai logika yang kamu pakai di halaman hasil.
    """
    i_bulanan = bunga_tahunan / 12.0
    selisih_hari = max(0, (first_due_date - start_date).days)  # guard
    bunga1_actual = pokok * bunga_tahunan * (selisih_hari / 360.0)

    data = []
    # baris bulan ke-0 (pencairan)
    data.append([0, start_date.strftime("%d %b %Y"), 0.0, 0.0, 0.0, pokok])

    sisa = pokok
    penjelasan = ""

    if metode == "anuitas":
        # PMT anuitas
        pmt = (
            pokok
            * (i_bulanan * (1 + i_bulanan) ** tenor)
            / ((1 + i_bulanan) ** tenor - 1)
        )

        # bulan-1: total pakai bunga aktual; pokok pakai bunga standar supaya sisa konsisten
        bunga1_std = sisa * i_bulanan
        pokok1 = pmt - bunga1_std
        tambahan_bunga = bunga1_actual - bunga1_std
        total1 = pmt + tambahan_bunga

        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= pokok1
        data.append([1, jatuh1, pokok1, bunga1_actual, total1, max(sisa, 0)])

        # bulan 2..n (normal)
        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            bunga = sisa * i_bulanan
            pokok_bayar = pmt - bunga
            sisa -= pokok_bayar
            data.append([bulan, jatuh, pokok_bayar, bunga, pmt, max(sisa, 0)])

        penjelasan = (
            "🔹 Metode Anuitas\n"
            f"i bulanan = {bunga_tahunan:.2%} / 12 = {i_bulanan:.6f}\n"
            f"PMT = P × [ i(1+i)^n / ((1+i)^n - 1) ] = {fmt(pmt)}\n\n"
            f"Periode pertama (selisih {selisih_hari} hari):\n"
            f"  Bunga actual = {fmt(pokok)} × {bunga_tahunan:.2%} × ({selisih_hari}/360) = {fmt(bunga1_actual)}\n"
            f"  Bunga standar = {fmt(pokok)} × {i_bulanan:.6f} = {fmt(bunga1_std)}\n"
            f"  Pokok dibayar = PMT − Bunga standar = {fmt(pmt)} − {fmt(bunga1_std)} = {fmt(pokok1)}\n"
            f"  Tambahan (penyesuaian hari) = {fmt(bunga1_actual)} − {fmt(bunga1_std)} = {fmt(tambahan_bunga)}\n"
            f"  Total angsuran bln-1 = PMT + Tambahan = {fmt(pmt)} + {fmt(tambahan_bunga)} = {fmt(total1)}\n"
        )

    elif metode == "efektif":
        cicilan_pokok = pokok / tenor

        # bulan-1: bunga actual day count
        bunga1 = sisa * bunga_tahunan * (selisih_hari / 360.0)
        total1 = cicilan_pokok + bunga1
        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= cicilan_pokok
        data.append([1, jatuh1, cicilan_pokok, bunga1, total1, max(sisa, 0)])

        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            bunga = sisa * i_bulanan
            total = cicilan_pokok + bunga
            sisa -= cicilan_pokok
            data.append([bulan, jatuh, cicilan_pokok, bunga, total, max(sisa, 0)])

        penjelasan = (
            "🔹 Metode Efektif\n"
            f"Cicilan pokok tetap = P / n = {fmt(pokok)} / {tenor} = {fmt(cicilan_pokok)}\n\n"
            f"Periode pertama (selisih {selisih_hari} hari):\n"
            f"  Bunga = {fmt(pokok)} × {bunga_tahunan:.2%} × ({selisih_hari}/360) = {fmt(bunga1)}\n"
            f"  Total angsuran bln-1 = {fmt(cicilan_pokok)} + {fmt(bunga1)} = {fmt(total1)}\n"
            f"Bulan berikutnya: bunga = sisa × {i_bulanan:.6f} dan total menurun seiring sisa pokok menurun."
        )

    elif metode == "flat":
        cicilan_pokok = pokok / tenor
        bunga_flat_bulanan = pokok * i_bulanan

        bunga1 = pokok * bunga_tahunan * (selisih_hari / 360.0)
        total1 = cicilan_pokok + bunga1
        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= cicilan_pokok
        data.append([1, jatuh1, cicilan_pokok, bunga1, total1, max(sisa, 0)])

        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            sisa -= cicilan_pokok
            total = cicilan_pokok + bunga_flat_bulanan
            data.append(
                [bulan, jatuh, cicilan_pokok, bunga_flat_bulanan, total, max(sisa, 0)]
            )

        penjelasan = (
            "🔹 Metode Flat\n"
            f"Cicilan pokok tetap = P / n = {fmt(pokok)} / {tenor} = {fmt(cicilan_pokok)}\n"
            f"Bunga bulanan tetap (mulai bln-2) = P × i_bulanan = {fmt(pokok)} × {i_bulanan:.6f} = {fmt(bunga_flat_bulanan)}\n\n"
            f"Periode pertama (selisih {selisih_hari} hari):\n"
            f"  Bunga actual = {fmt(pokok)} × {bunga_tahunan:.2%} × ({selisih_hari}/360) = {fmt(bunga1)}\n"
            f"  Total angsuran bln-1 = {fmt(cicilan_pokok)} + {fmt(bunga1)} = {fmt(total1)}"
        )

    else:
        raise ValueError("Metode tidak dikenal")

    df = pd.DataFrame(
        data,
        columns=[
            "Bulan",
            "Tanggal Jatuh Tempo",
            "Angsuran Pokok",
            "Bunga",
            "Total Angsuran",
            "Sisa Pokok",
        ],
    )
    return df, penjelasan


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # ambil input
        pokok_str = request.form["pokok"]
        # bersihkan pemisah ribuan (misal "2.300.000.000")
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

        # simpan HANYA parameter ke session (aman di cookie)
        session["params"] = {
            "pokok": pokok,
            "bunga_tahunan": bunga_tahunan,
            "tenor": tenor,
            "metode": metode,
            "start_date": start_date.strftime("%Y-%m-%d"),
            "first_due_date": first_due_date.strftime("%Y-%m-%d"),
            "namecstm": namecstm,
        }

        # bangun jadwal untuk tampilan
        df, penjelasan = build_schedule(
            pokok, bunga_tahunan, tenor, metode, start_date, first_due_date
        )

        # format angka untuk HTML (jangan simpan versi terformat ke session!)
        df_show = df.copy()
        for col in ["Angsuran Pokok", "Bunga", "Total Angsuran", "Sisa Pokok"]:
            df_show[col] = df_show[col].map(fmt)

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
        )

    # GET
    return render_template("index.html")


@app.route("/export_excel")
def export_excel():
    params = session.get("params")
    if not params:
        # kalau session hilang (serverless, tab baru, dsb), arahkan balik ke form
        return redirect(url_for("index"))

    # rebuild jadwal dari parameter
    pokok = float(params["pokok"])
    bunga_tahunan = float(params["bunga_tahunan"])
    tenor = int(params["tenor"])
    metode = params["metode"]
    start_date = datetime.strptime(params["start_date"], "%Y-%m-%d")
    first_due_date = datetime.strptime(params["first_due_date"], "%Y-%m-%d")

    df, _ = build_schedule(
        pokok, bunga_tahunan, tenor, metode, start_date, first_due_date
    )

    # tulis ke memory (BytesIO) -> aman di Vercel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Simulasi Angsuran")
    output.seek(0)

    # nama file aman
    namecstm = params.get("namecstm", "Customer")
    safe_name = "".join(c if c.isalnum() else "_" for c in namecstm)

    return send_file(
        output,
        download_name=f"TABEL_ANGSURAN_{safe_name}.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
