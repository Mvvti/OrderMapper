import glob
import json
import os
import sys
import threading
import traceback
from datetime import datetime

import webview

from src.pipeline_runner import run_pipeline as _run_pipeline
from src.pdf_pipeline.pipeline_runner import run_pdf_pipeline as _run_pdf_pipeline


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


class API:
    def get_defaults(self):
        base = os.path.dirname(sys.executable) if hasattr(sys, '_MEIPASS') else os.path.dirname(os.path.abspath(__file__))

        def find_first(pattern):
            results = glob.glob(os.path.join(base, pattern))
            return results[0] if results else ""

        cennik_default = find_first("*cennik*.xls*")
        placowki_default = find_first("*plac*.*")
        template_default = find_first("Template.xlsx")

        return {
            "cennik": cennik_default,
            "placowki": placowki_default,
            "template": template_default,
        }

    def pick_file(self, *args, **kwargs):
        try:
            result = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("Excel Files (*.xlsx;*.xls)",),
            )
            return result[0] if result else None
        except Exception:
            return None

    def pick_folder(self, *args, **kwargs):
        try:
            result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
            return result[0] if result else None
        except Exception:
            return None

    def validate_inputs(self, order, cennik, placowki, template, output, date):
        errors = []

        excel_fields = [
            ("order", order, "Plik zamĂłwienia"),
            ("cennik", cennik, "Plik cennika"),
            ("placowki", placowki, "Plik placĂłwek"),
            ("template", template, "Plik template"),
        ]

        for field_name, path, label in excel_fields:
            normalized_path = (path or "").strip()
            lower_path = normalized_path.lower()

            if not normalized_path:
                errors.append({"field": field_name, "message": f"{label}: pole jest wymagane."})
                continue

            if not os.path.exists(normalized_path):
                errors.append({"field": field_name, "message": f"{label}: wskazany plik nie istnieje."})

            if not (lower_path.endswith(".xlsx") or lower_path.endswith(".xls")):
                errors.append(
                    {"field": field_name, "message": f"{label}: dozwolone sÄ… tylko pliki .xlsx lub .xls."}
                )

        output_path = (output or "").strip()
        if not output_path:
            errors.append({"field": "output", "message": "Folder output: pole jest wymagane."})
        elif not os.path.exists(output_path):
            errors.append({"field": "output", "message": "Folder output: wskazany folder nie istnieje."})
        elif not os.path.isdir(output_path):
            errors.append({"field": "output", "message": "Folder output: podana Ĺ›cieĹĽka nie jest folderem."})

        date_value = (date or "").strip()
        if not date_value:
            errors.append({"field": "date", "message": "Data: pole jest wymagane (format YYYY-MM-DD)."})
        else:
            try:
                datetime.strptime(date_value, "%Y-%m-%d")
            except ValueError:
                errors.append({"field": "date", "message": "Data: nieprawidĹ‚owy format, uĹĽyj YYYY-MM-DD."})

        return errors

    def pick_pdf_file(self, *args, **kwargs):
        try:
            result = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("PDF Files (*.pdf)",),
            )
            return result[0] if result else None
        except Exception:
            return None

    def pick_pdf_files(self, *args, **kwargs):
        try:
            result = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=True,
                file_types=("PDF Files (*.pdf)",),
            )
            return list(result) if result else []
        except Exception:
            return []

    def open_url(self, url):
        import webbrowser
        webbrowser.open(url)

    def run_pipeline(self, order, cennik, placowki, template, output_dir, date):
        def worker():
            def log(msg):
                webview.windows[0].evaluate_js(f"appendLog({json.dumps(msg)})")

            try:
                result = _run_pipeline(
                    excel_file_path=order,
                    cennik_file_path=cennik,
                    placowki_file_path=placowki,
                    template_file_path=template,
                    output_dir=output_dir,
                    test_data_utworzenia=date,
                    log_callback=log,
                )
                webview.windows[0].evaluate_js(f"onPipelineDone({json.dumps(result)})")
            except Exception:
                err = traceback.format_exc()
                webview.windows[0].evaluate_js(f"onPipelineError({json.dumps(err)})")

        threading.Thread(target=worker, daemon=True).start()

    def run_pdf_pipeline(self, pdf_paths, output_folder):
        def worker():
            def log(msg):
                webview.windows[0].evaluate_js(f"appendPdfLog({json.dumps(msg)})")

            try:
                result = _run_pdf_pipeline(
                    pdf_paths=pdf_paths,
                    output_folder=output_folder,
                    log_fn=log,
                )
                webview.windows[0].evaluate_js(f"onPdfPipelineDone({json.dumps(result)})")
            except Exception:
                err = traceback.format_exc()
                webview.windows[0].evaluate_js(f"onPdfPipelineError({json.dumps(err)})")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    html_path = get_resource_path(os.path.join("frontend", "index.html"))
    html_url = f"file:///{html_path.replace(chr(92), '/')}"

    api = API()
    webview.create_window(
        title="Zamówienia",
        url=html_url,
        width=980,
        height=760,
        resizable=True,
        js_api=api,
    )
    webview.start()

