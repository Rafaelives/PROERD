from pathlib import Path

from main import app, export_school_kml, get_territorial_data


def build_static_site():
    project_dir = Path(__file__).resolve().parent
    docs_dir = project_dir / "docs"
    docs_dir.mkdir(exist_ok=True)

    kml_path = export_school_kml()
    target_kml = docs_dir / "ceara_municipios.kml"
    target_kml.write_text(kml_path.read_text(encoding="utf-8"), encoding="utf-8")

    with app.app_context():
        html = app.jinja_env.get_template("index.html").render(
            map_page=get_territorial_data(),
            kml_href="ceara_municipios.kml",
        )

    (docs_dir / "index.html").write_text(html, encoding="utf-8")
    (docs_dir / ".nojekyll").write_text("", encoding="utf-8")
    return docs_dir


if __name__ == "__main__":
    output_dir = build_static_site()
    print(f"Relatório estático gerado em: {output_dir}")
