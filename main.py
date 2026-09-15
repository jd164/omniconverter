"""
Multi-Format Call Log Converter (JSON to Excel, CSV, JSON, and ICS)
Run via Command Line Interface (CLI) or launch the interactive web dashboard.
"""

import argparse
import os
import sys
import webbrowser
from converter import (
    carregar_dados,
    converter_todos,
    exportar_csv,
    exportar_excel,
    exportar_ics,
    exportar_json,
)


def gerar_ics(json_path="dados.json", ics_path="call_logs.ics", title="Call with {contact}"):
    """
    Backwards compatibility function.
    Generates .ics calendar file from JSON.
    """
    return exportar_ics(json_path, ics_path, event_title=title)


def iniciar_servidor_web(port=8000, open_browser=True):
    """
    Launches local FastAPI web server and opens browser.
    """
    try:
        from app import app
        import uvicorn

        url = f"http://127.0.0.1:{port}"
        print("=" * 60)
        print(f"🚀 Launching Interactive Web Dashboard at: {url}")
        print("Press Ctrl+C to stop the server.")
        print("=" * 60)

        if open_browser:
            import threading
            import time

            def open_url():
                time.sleep(1.2)
                webbrowser.open(url)

            threading.Thread(target=open_url, daemon=True).start()

        uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")
    except ImportError:
        import http.server
        import socketserver

        web_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")
        if not os.path.exists(web_dir):
            web_dir = os.path.dirname(os.path.abspath(__file__))

        class Handler(http.server.SimpleHTTPRequestHandler):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, directory=web_dir, **kwargs)

        print("=" * 60)
        print(f"🚀 Starting static server at: http://127.0.0.1:{port}")
        print("=" * 60)
        if open_browser:
            webbrowser.open(f"http://127.0.0.1:{port}")

        with socketserver.TCPServer(("127.0.0.1", port), Handler) as httpd:
            httpd.serve_forever()


def main():
    parser = argparse.ArgumentParser(
        description="Convert Call Logs to Excel (.xlsx), CSV, JSON, and iCalendar (.ics)",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        "-i", "--input",
        default="dados.json",
        help="Path to input JSON file (default: dados.json)"
    )
    parser.add_argument(
        "-o", "--output",
        default="call_logs",
        help="Base output filename (default: call_logs)"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["all", "xlsx", "csv", "json", "ics"],
        default="all",
        help="Target output format:\n"
             "  all  -> Generates Excel (.xlsx), CSV, JSON, and ICS (default)\n"
             "  xlsx -> Styled Excel spreadsheet with analytics sheet\n"
             "  csv  -> UTF-8 BOM CSV (Excel compatible)\n"
             "  json -> Clean structured JSON\n"
             "  ics  -> iCalendar calendar file\n"
    )
    parser.add_argument(
        "-t", "--title",
        default="Call with {contact}",
        help="Event summary/title template for ICS calendar (e.g. 'Call with {contact}')"
    )
    parser.add_argument(
        "--web",
        action="store_true",
        help="Launch the interactive web dashboard in your browser"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for local web server (default: 8000)"
    )

    args = parser.parse_args()

    if args.web:
        iniciar_servidor_web(port=args.port)
        return

    print("=" * 60)
    print("📊 OmniConverter • Multi-Format Call Log Conversion Tool")
    print("=" * 60)

    if not os.path.exists(args.input):
        print(f"❌ Error: Input file '{args.input}' not found!")
        sys.exit(1)

    base = args.output
    for ext in [".xlsx", ".csv", ".json", ".ics"]:
        if base.lower().endswith(ext):
            base = base[:-len(ext)]

    try:
        if args.format == "all":
            res = converter_todos(args.input, base_name=base, event_title=args.title)
            print(f"✅ Conversion completed successfully ({res['totalCalls']} calls | {res['totalDuration']}):")
            print(f"   • Excel:      {res['xlsx']}")
            print(f"   • CSV:        {res['csv']}")
            print(f"   • JSON:       {res['json']}")
            print(f"   • Calendar:   {res['ics']}")
        elif args.format == "xlsx":
            out = f"{base}.xlsx"
            exportar_excel(args.input, out, event_title=args.title)
            print(f"✅ Excel spreadsheet generated: {out}")
        elif args.format == "csv":
            out = f"{base}.csv"
            exportar_csv(args.input, out)
            print(f"✅ CSV file generated: {out}")
        elif args.format == "json":
            out = f"{base}.json"
            exportar_json(args.input, out)
            print(f"✅ JSON file generated: {out}")
        elif args.format == "ics":
            out = f"{base}.ics"
            exportar_ics(args.input, out, event_title=args.title)
            print(f"✅ iCalendar (.ics) file generated: {out}")

        print("\n💡 Tip: To use the visual interactive dashboard, run:")
        print("   python main.py --web")
        print("   or double-click start_dashboard.bat")
    except Exception as e:
        print(f"❌ Error during conversion: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()