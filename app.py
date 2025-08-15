from flask import Flask, render_template, request, session, send_file
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

app = Flask(__name__)
app.secret_key = "random-secret"  # wajib untuk session


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        pokok = float(request.form["pokok"])
        bunga_tahunan = float(request.form["bunga"]) / 100
        tenor = int(request.form["tenor"])
        metode = request.form["metode"]

        session["namecstm"] = request.form.get("namecstm", "Customer")

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

        # suku bunga
        bunga_bulanan = bunga_tahunan / 12

        # day count: selisih aktual hari antara pencairan dan jatuh tempo 1
        selisih_hari = (first_due_date - start_date).days
        bunga_aktual_pertama = pokok * bunga_tahunan * selisih_hari / 360.0
        bunga_bulanan_standar = (
            pokok * bunga_bulanan
        )  # untuk hitung pokok bulan-1 pada anuitas
        tambahan_bunga_pertama = bunga_aktual_pertama - bunga_bulanan_standar

        data = []
        # baris 0 (pencairan)
        data.append([0, start_date.strftime("%d %b %Y"), 0.0, 0.0, 0.0, pokok])

        sisa = pokok

        if metode == "anuitas":
            # PMT tetap
            pmt = (
                pokok
                * (bunga_bulanan * (1 + bunga_bulanan) ** tenor)
                / ((1 + bunga_bulanan) ** tenor - 1)
            )

            for bulan in range(1, tenor + 1):
                jatuh_tempo = (
                    first_due_date + relativedelta(months=bulan - 1)
                ).strftime("%d %b %Y")

                if bulan == 1:
                    # porsi pokok dihitung pakai bunga bulanan standar (agar sisa pokok sama dgn jadwal normal)
                    pokok_bayar = pmt - (sisa * bunga_bulanan)
                    bunga_tampil = bunga_aktual_pertama
                    total_angsuran = pmt + tambahan_bunga_pertama
                else:
                    bunga_tampil = sisa * bunga_bulanan
                    pokok_bayar = pmt - bunga_tampil
                    total_angsuran = pmt

                sisa -= pokok_bayar
                data.append(
                    [
                        bulan,
                        jatuh_tempo,
                        pokok_bayar,
                        bunga_tampil,
                        total_angsuran,
                        max(sisa, 0),
                    ]
                )

        elif metode == "efektif":
            cicilan_pokok = pokok / tenor
            for bulan in range(1, tenor + 1):
                jatuh_tempo = (
                    first_due_date + relativedelta(months=bulan - 1)
                ).strftime("%d %b %Y")
                if bulan == 1:
                    bunga_tampil = sisa * bunga_tahunan * selisih_hari / 360.0
                else:
                    bunga_tampil = sisa * bunga_bulanan
                total_angsuran = cicilan_pokok + bunga_tampil
                sisa -= cicilan_pokok
                data.append(
                    [
                        bulan,
                        jatuh_tempo,
                        cicilan_pokok,
                        bunga_tampil,
                        total_angsuran,
                        max(sisa, 0),
                    ]
                )

        elif metode == "flat":
            cicilan_pokok = pokok / tenor
            bunga_flat_bulanan = pokok * bunga_bulanan
            for bulan in range(1, tenor + 1):
                jatuh_tempo = (
                    first_due_date + relativedelta(months=bulan - 1)
                ).strftime("%d %b %Y")
                if bulan == 1:
                    bunga_tampil = pokok * bunga_tahunan * selisih_hari / 360.0
                else:
                    bunga_tampil = bunga_flat_bulanan
                total_angsuran = cicilan_pokok + bunga_tampil
                sisa -= cicilan_pokok
                data.append(
                    [
                        bulan,
                        jatuh_tempo,
                        cicilan_pokok,
                        bunga_tampil,
                        total_angsuran,
                        max(sisa, 0),
                    ]
                )

        # DataFrame & simpan ke session
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
        session["data"] = df.to_dict(orient="list")

        # format rupiah untuk tampilan
        def format_indo(x):
            return f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

        for col in ["Angsuran Pokok", "Bunga", "Total Angsuran", "Sisa Pokok"]:
            df[col] = df[col].map(format_indo)

        return render_template(
            "result.html",
            tables=[
                df.to_html(
                    classes="table table-striped table-bordered",
                    index=False,
                    escape=False,
                )
            ],
            title="Hasil Simulasi",
        )

    return render_template("index.html")


@app.route("/export_excel")
def export_excel():
    if "data" not in session:
        return "Tidak ada data untuk diekspor", 400

    df = pd.DataFrame(session["data"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Simulasi Angsuran")
    output.seek(0)

    # ambil nama customer dari session
    namecstm = session.get("namecstm", "Customer")
    safe_name = "".join(
        c if c.isalnum() else "_" for c in namecstm
    )  # biar aman untuk filename

    return send_file(
        output,
        download_name=f"TABEL_ANGSURAN_{safe_name}.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
