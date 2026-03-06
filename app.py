from flask import Flask, render_template, request
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import io
import base64
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    table_html = None
    chart_url = None

    if request.method == "POST":
        file = request.files.get("file")

        if file:
            # đọc toàn bộ excel
            df = pd.read_excel(file, header=None)

            # bỏ 2 dòng đầu không cần
            # giữ nguyên nếu file của bạn header thật nằm ở dòng 6,7
            # xóa toàn bộ cột trống trước khi tách header/data
            df = df.dropna(axis=1, how="all")

            # lấy 2 dòng header
            header = df.iloc[6:8].fillna("")
            header.iloc[0] = header.iloc[0].replace("", pd.NA).ffill().fillna("")

            # lấy dữ liệu bảng
            data = df.iloc[8:-4].reset_index(drop=True)

            # ===== render header =====
            header_html = "<tr>"

            cols = header.shape[1]
            i = 0

            while i < cols:
                text = str(header.iloc[0, i]).strip()

                if text == "":
                    i += 1
                    continue

                colspan = 1
                j = i + 1

                while j < cols and str(header.iloc[0, j]).strip() == "":
                    colspan += 1
                    j += 1

                header_html += f"<th colspan='{colspan}'>{text}</th>"
                i = j

            header_html += "</tr><tr>"

            for col in header.iloc[1]:
                val = str(col).strip()
                header_html += f"<th>{val}</th>"

            header_html += "</tr>"

            # ===== render body =====
            body_html = ""

            for _, row in data.iterrows():
                if row.isna().all():
                    continue

                body_html += "<tr>"
                for val in row:
                    if pd.isna(val):
                        body_html += "<td></td>"
                    else:
                        body_html += f"<td>{val}</td>"
                body_html += "</tr>"

            # ===== table html =====
            table_html = f"""
            <table class="table">
                {header_html}
                {body_html}
            </table>
            """
            # =========================
            # PHẦN BIỂU ĐỒ TRÒN
            # =========================
            # Dựa trên cấu trúc file gốc:
            # cột 1 = STT
            # cột 2 = Khoa
            # cột 6 = Người bệnh vào điều trị nội trú -> Tổng số
            df_chart = df.iloc[5:].copy()

            # lấy đúng 3 cột cần dùng từ file gốc
            chart_data = df_chart.iloc[:, [0, 1, 5]].copy()
            chart_data.columns = ["STT", "Khoa", "NoiTru_TongSo"]

            # chỉ lấy các dòng khoa thật
            chart_data = chart_data[chart_data["STT"].notna()]
            chart_data = chart_data[chart_data["Khoa"].notna()]

            # chuyển số
            chart_data["NoiTru_TongSo"] = pd.to_numeric(
                chart_data["NoiTru_TongSo"], errors="coerce"
            ).fillna(0)

            # bỏ khoa có giá trị = 0
            chart_data = chart_data[chart_data["NoiTru_TongSo"] > 0]

            if not chart_data.empty:
                labels = chart_data["Khoa"].tolist()
                sizes = chart_data["NoiTru_TongSo"].tolist()

                def make_autopct(values):
                    def my_autopct(pct):
                        total = sum(values)
                        val = int(round(pct * total / 100.0))
                        return f"{val} người\n({pct:.1f}%)"
                    return my_autopct

                plt.figure(figsize=(10, 8))

                wedges, texts, autotexts = plt.pie(
                    sizes,
                    labels=None,   # bỏ tên khoa trên lát cắt cho đỡ rối
                    autopct=make_autopct(sizes),
                    startangle=90
                )

                plt.axis("equal")

                # note màu bên phải
                legend_labels = [
                    f"{khoa}: {so_nguoi} người ({so_nguoi / sum(sizes) * 100:.1f}%)"
                    for khoa, so_nguoi in zip(labels, sizes)
                ]

                plt.legend(
                    wedges,
                    legend_labels,
                    title="Chú thích",
                    loc="center left",
                    bbox_to_anchor=(1, 0.5),
                    fontsize=12,
                    title_fontsize=14
                )

                img = io.BytesIO()
                plt.savefig(img, format="png", bbox_inches="tight")
                img.seek(0)
                chart_url = base64.b64encode(img.getvalue()).decode()
                plt.close()

    return render_template("index.html", table=table_html, chart_url=chart_url)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)