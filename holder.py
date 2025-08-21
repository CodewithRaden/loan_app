from flask import Flask, render_template, request, session, send_file, redirect, url_for
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal, ROUND_HALF_UP
import io

app = Flask(__name__)
app.secret_key = "random-secret"


def fmt(x: float) -> str:
    """Format angka ke gaya Indonesia tanpa desimal."""
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
    if metode == "anuitas":
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

    elif metode == "efektif":
        cicilan_pokok = pokok / tenor
        bunga1 = sisa * bunga_tahunan * (selisih_hari / 360.0)
        total1 = cicilan_pokok + bunga1
        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= cicilan_pokok
        data.append([1, jatuh1, cicilan_pokok, bunga1, total1, sisa])

        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            bunga = sisa * i_bulanan
            total = cicilan_pokok + bunga
            sisa -= cicilan_pokok
            data.append([bulan, jatuh, cicilan_pokok, bunga, total, sisa])

    elif metode == "flat":
        cicilan_pokok = pokok / tenor
        bunga_flat_bulanan = pokok * i_bulanan
        bunga1 = pokok * bunga_tahunan * (selisih_hari / 360.0)
        total1 = cicilan_pokok + bunga1
        jatuh1 = first_due_date.strftime("%d %b %Y")
        sisa -= cicilan_pokok
        data.append([1, jatuh1, cicilan_pokok, bunga1, total1, sisa])

        for bulan in range(2, tenor + 1):
            jatuh = (first_due_date + relativedelta(months=bulan - 1)).strftime(
                "%d %b %Y"
            )
            sisa -= cicilan_pokok
            total = cicilan_pokok + bunga_flat_bulanan
            data.append([bulan, jatuh, cicilan_pokok, bunga_flat_bulanan, total, sisa])

    else:
        raise ValueError("Metode tidak dikenal")

    # ----------------------
    # DataFrame awal
    # ----------------------
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
        totals = {
            "Bulan": "Total",
            "Tanggal Jatuh Tempo": "",
            "Angsuran Pokok": fmt(df["Angsuran Pokok"].apply(rupiah_round).sum()),
            "Bunga": fmt(df["Bunga"].apply(rupiah_round).sum()),
            "Total Angsuran": fmt(df["Total Angsuran"].apply(rupiah_round).sum()),
            "Sisa Pokok": "",
        }
        df_show = pd.concat([df_show, pd.DataFrame([totals])], ignore_index=True)

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


@app.route("/export_excel")
def export_excel():
    params = session.get("params")
    if not params:
        return redirect(url_for("index"))

    pokok = float(params["pokok"])
    bunga_tahunan = float(params["bunga_tahunan"])
    tenor = int(params["tenor"])
    metode = params["metode"]
    start_date = datetime.strptime(params["start_date"], "%Y-%m-%d")
    first_due_date = datetime.strptime(params["first_due_date"], "%Y-%m-%d")

    df, _, _ = build_schedule(
        pokok, bunga_tahunan, tenor, metode, start_date, first_due_date
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Simulasi Angsuran")
    output.seek(0)

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
