from pathlib import Path

from main import app, export_school_csv, export_school_pdf, get_dashboard_data, get_territorial_data


def build_static_site():
    project_dir = Path(__file__).resolve().parent
    docs_dir = project_dir / "docs"
    docs_dir.mkdir(exist_ok=True)

    csv_path = export_school_csv()
    target_csv = docs_dir / "ceara_municipios.csv"
    target_csv.write_text(csv_path.read_text(encoding="utf-8-sig"), encoding="utf-8-sig")

    pdf_path = export_school_pdf()
    target_pdf = docs_dir / "ceara_municipios.pdf"
    target_pdf.write_bytes(pdf_path.read_bytes())

    with app.app_context():
        html = app.jinja_env.get_template("index.html").render(
            map_page=get_territorial_data(),
            scientific_page=get_dashboard_data(),
            csv_href="ceara_municipios.csv",
            pdf_href="ceara_municipios.pdf",
        )

    (docs_dir / "index.html").write_text(html, encoding="utf-8")
    (docs_dir / ".nojekyll").write_text("", encoding="utf-8")
    return docs_dir


if __name__ == "__main__":
    output_dir = build_static_site()
    print(f"Relatório estático gerado em: {output_dir}")
