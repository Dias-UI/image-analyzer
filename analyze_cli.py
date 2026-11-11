"""Headless interface for the Image Analyzer. Run with --help for usage."""
import argparse
import importlib.util
import json
import os
import sys
import tempfile
from pathlib import Path

import openpyxl

spec = importlib.util.spec_from_file_location('image_analyzer_gui', Path(__file__).with_name('image analyzer.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)


def collect_inputs(inputs):
    paths = []
    for name in inputs:
        path = Path(name).expanduser().resolve()
        if path.is_dir():
            paths.extend(sorted(p for p in path.rglob('*') if p.is_file() and p.suffix.lower() in app.SUPPORTED_EXT))
        elif path.is_file() and path.suffix.lower() in app.SUPPORTED_EXT:
            paths.append(path)
        else:
            raise ValueError(f'Unsupported or missing input: {path}')
    paths = list(dict.fromkeys(paths))
    if not paths:
        raise ValueError('No supported images or PDFs found.')
    return paths


def prepare_workbook(path, sheet):
    if path.suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ValueError('Output must be an .xlsx or .xlsm workbook.')
    if path.exists():
        wb = openpyxl.load_workbook(path, keep_vba=path.suffix.lower() == '.xlsm')
        if sheet and sheet not in wb.sheetnames:
            raise ValueError(f'Worksheet not found: {sheet}')
        ws = wb[sheet] if sheet else wb.active
    else:
        if path.suffix.lower() == '.xlsm':
            raise ValueError('A new macro-enabled workbook cannot be created; use .xlsx.')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet or 'Image Analysis'
    headers = {str(c.value).strip().casefold(): c.column for c in ws[1] if c.value is not None}
    if headers and 'filename' not in headers:
        raise ValueError('The first row must contain a Filename header, or be empty. Select the intended sheet with --sheet.')
    for name in app.EXPORT_COLS:
        if name.casefold() not in headers:
            col = max(headers.values(), default=0) + 1
            ws.cell(1, col, name)
            headers[name.casefold()] = col
    return wb, ws, headers


def save_workbook(wb, path):
    fd, temp = tempfile.mkstemp(suffix=path.suffix, dir=str(path.parent))
    os.close(fd)
    try:
        wb.save(temp)
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('inputs', nargs='+', help='Image/PDF files or folders (folders searched recursively)')
    parser.add_argument('--excel', required=True, help='Existing or new .xlsx workbook')
    parser.add_argument('--sheet', help='Existing worksheet name; for a new workbook, its sheet name')
    parser.add_argument('--provider', choices=['local', 'openrouter'], default='local')
    parser.add_argument('--local-model', choices=['qwen', 'gemma'], default='qwen', help='Configured local vision model')
    parser.add_argument('--base-url', default=None, help='Local llama.cpp OpenAI-compatible API base URL')
    parser.add_argument('--model', default=app.OPENROUTER_MODELS[0]['id'], help='OpenRouter model ID')
    parser.add_argument('--pdf-engine', choices=['cloudflare-ai', 'mistral-ocr', 'native'])
    parser.add_argument('--translate', action='store_true')
    parser.add_argument('--geocode', action='store_true', help='Use GEOCODING_KEY for GPS reverse geocoding')
    args = parser.parse_args(argv)
    try:
        args.base_url = args.base_url or app.LOCAL_MODELS[args.local_model][0]
        paths = collect_inputs(args.inputs)
        output = Path(args.excel).expanduser().resolve()
        if not output.parent.is_dir():
            raise ValueError(f'Output directory does not exist: {output.parent}')
        if args.provider == 'local' and any(p.suffix.lower() == '.pdf' for p in paths):
            raise ValueError('PDF processing requires --provider openrouter.')
        key = os.environ.get('OPENROUTER_KEY') or os.environ.get('OPENROUTER_API_KEY', '')
        if args.provider == 'openrouter' and not key:
            raise ValueError('Set OPENROUTER_KEY or OPENROUTER_API_KEY.')
        if args.geocode and not os.environ.get('GEOCODING_KEY'):
            raise ValueError('--geocode requires GEOCODING_KEY for Google Maps reverse geocoding.')
        wb, ws, headers = prepare_workbook(output, args.sheet)
        if args.provider == 'local':
            app.ensure_local_server(args.base_url, args.local_model)
        settings = {'provider': args.provider, 'lmstudio_base': args.base_url, 'model_id': args.model,
                    'pdf_engine': args.pdf_engine, 'api_key': key, 'translate': args.translate,
                    'geocoding_key': os.environ.get('GEOCODING_KEY', '') if args.geocode else ''}
        # The processing methods use only these two attributes, and do not create a Tk window.
        class Processor:
            stop_flag = False
            geocode_cache = {}
            _process_image = app.ImageAnalyzerApp._process_image
            _process_pdf = app.ImageAnalyzerApp._process_pdf
        processor = Processor()
        prompt = app.build_analysis_prompt(args.translate)
        successes, errors = [], []
        for path in paths:
            try:
                row = (processor._process_pdf(str(path), settings) if path.suffix.lower() == '.pdf'
                       else processor._process_image(str(path), settings, prompt))
                row_number = ws.max_row + 1
                for name in app.EXPORT_COLS:
                    value = app._export_value(row.get(name, ''))
                    cell = ws.cell(row_number, headers[name.casefold()], value)
                    if value.startswith('='):
                        cell.data_type = 's'
                successes.append(str(path))
            except Exception as exc:
                errors.append({'path': str(path), 'error': str(exc)})
        if successes:
            save_workbook(wb, output)
        wb.close()
        print(json.dumps({'workbook': str(output), 'sheet': ws.title, 'processed': successes, 'errors': errors}, ensure_ascii=False))
        return 1 if errors else 0
    except Exception as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
