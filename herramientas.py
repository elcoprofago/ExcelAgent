# Las herramientas que el modelo puede pedir sobre el libro abierto: SPECS (formato OpenAI), parse_args, ToolError y
# ToolBox.execute(nombre, args, cancel) -> texto.
#
# Todo corre en el hilo de trabajo, el mismo que abrio la conexion COM: cruzar objetos COM entre hilos los invalida.
# Las descripciones van en ingles (los modelos las siguen mejor); los mensajes al usuario los escribe el modelo.
import datetime
import inspect
import json
import os
import re

import excel

MAX_READ_CELLS = 2000
MAX_TEXT_CELLS = 400        # leer .Text es una llamada COM por celda
MAX_WRITE_CELLS = 20000
MAX_ATTACHMENT_CHARS = 8000
MAX_FORMAT_CALLS = 2000     # al pegar, los formatos de numero y fecha se ponen por tramos de columna
MAX_FIND_HITS = 50

# Los errores de celda llegan por Value2 como estos enteros (CVErr).
CELL_ERRORS = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
               -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!'}

# Herramientas que cambian el libro: tras una que salio bien, repetir una lectura ya no es repetir en vano.
STATE_CHANGING = {'write_range', 'fill_formula', 'format_range', 'autofit_columns', 'sort_range', 'autofilter',
                  'create_table', 'remove_duplicates', 'replace', 'clear_range', 'insert_rows_columns',
                  'delete_rows_columns', 'sheet', 'add_chart', 'conditional_format', 'data_validation_list',
                  'freeze_panes', 'save', 'paste_attachment', 'outlook_contacts'}

NUMBER_FORMATS = {'currency': '$ #,##0.00', 'number': '#,##0.00', 'integer': '#,##0', 'percent': '0.00%',
                  'date': 'dd/mm/yyyy', 'datetime': 'dd/mm/yyyy hh:mm', 'text': '@', 'general': 'General'}
COLOR_NAMES = dict(excel.COLORES, red='#FF0000', green='#00B050', yellow='#FFFF00', blue='#0070C0',
                   orange='#FFC000', purple='#7030A0', grey='#D9D9D9', gray='#D9D9D9', lightblue='#00B0F0',
                   pink='#FF99CC', black='#000000', white='#FFFFFF', blanco='#FFFFFF')
ALIGN = {'left': -4131, 'center': -4108, 'right': -4152}
CHART_TYPES = {'column': 51, 'bar': 57, 'line': 4, 'pie': 5, 'scatter': -4169, 'area': 1}
COMPARE_OPS = {'between': 1, 'not_between': 2, 'equal': 3, 'not_equal': 4, 'greater_than': 5, 'less_than': 6,
               'greater_or_equal': 7, 'less_or_equal': 8}


class ToolError(Exception):
    pass


def parse_args(raw):
    'Los argumentos de una llamada llegan como texto JSON; puede venir vacio, cortado o mal formado.'
    if raw is None or not str(raw).strip():
        return {}
    try:
        v = json.loads(raw)
    except ValueError as e:
        raise ToolError(f'arguments are not valid JSON ({e}); resend the call with valid JSON')
    if not isinstance(v, dict):
        raise ToolError('arguments must be a JSON object')
    return v


def com_text(e):
    'El texto legible de un error COM (la descripcion de Excel, no el HRESULT).'
    args = getattr(e, 'args', ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return str(e)


def color_value(color):
    'Color por nombre (espanol o ingles) o #RRGGBB, al numero BGR de Excel.'
    c = str(color).strip().lower()
    c = COLOR_NAMES.get(c, c)
    h = c.lstrip('#').lower()
    if len(h) != 6 or any(ch not in '0123456789abcdef' for ch in h):
        raise ToolError(f"unknown color '{color}'. Use #RRGGBB or one of: {', '.join(sorted(COLOR_NAMES))}")
    return excel.color_bgr('#' + h)


def grid(v):
    'Value2/Formula de un rango siempre como lista de filas.'
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) if isinstance(r, tuple) else [r] for r in v]


def show(v):
    if v is None:
        return ''
    if isinstance(v, int) and v in CELL_ERRORS:
        return CELL_ERRORS[v]
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def list_separator(app):
    'xlListSeparator de Excel (en espanol suele ser ;). Por pywin32 International llega como tupla, no como metodo.'
    intl = app.International
    try:
        return str(intl[4] if isinstance(intl, tuple) else app.International(5)) or ','
    except Exception:
        return ','


def decimal_separator(app):
    'xlDecimalSeparator de Excel: con que separador decimal se leen los numeros escritos como texto sin otra pista.'
    try:
        intl = app.International
        sep = str(intl[2] if isinstance(intl, tuple) else app.International(3))
        return sep if sep in ('.', ',') else ','
    except Exception:
        return ','


def interpret_columns(rows, decimal):
    """Los textos de cada columna convertidos en numeros y fechas con la convencion de esa columna (numeros.py).
    Solo los textos comunes: no las formulas ni lo ya marcado como numeros.Texto. Devuelve las filas nuevas y
    [(fila, columna, formato)] para las celdas que necesitan un formato."""
    import numeros
    width = max((len(r) for r in rows), default=0)
    conv = [numeros.convencion([r[i] for r in rows if i < len(r)], decimal) for i in range(width)]
    out, formats = [], []
    for ri, r in enumerate(rows):
        new = []
        for ci, v in enumerate(r):
            if type(v) is str and not v.lstrip().startswith('='):
                v, f = numeros.interpretar(v, *conv[ci])
                if f:
                    formats.append((ri, ci, f))
            new.append(v)
        out.append(new)
    return out, formats


def apply_formats(target, formats, limit=2000):
    """Pone los formatos de numero por tramos: celdas seguidas de una columna con el mismo formato van en una sola
    llamada COM. Devuelve cuantas celdas quedaron sin formato por pasar el limite de llamadas."""
    groups = {}
    for r, c, f in formats:
        groups.setdefault((c, f), []).append(r)
    calls, left = 0, 0
    for (c, f), rs in groups.items():
        rs.sort()
        start = prev = rs[0]
        for r in rs[1:] + [None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            if calls < limit:
                # Worksheet.Range: target.Range(...) se resuelve relativo a target y fuera de A1 caia en otra celda
                set_number_format(target.Worksheet.Range(target.Cells(start + 1, c + 1), target.Cells(prev + 1, c + 1)), f)
                calls += 1
            else:
                left += prev - start + 1
            if r is not None:
                start = prev = r
    return left


def set_number_format(r, code):
    """NumberFormat con un codigo en ingles. Por pywin32 la propiedad viaja con el idioma del usuario: en este Excel en
    espanol '$ #,##0.00' asignado directo se veia '$ 1234,5000'. Medido: con LCID 1033 se ve '$ 1.234,50'."""
    import pythoncom
    ole = r._oleobj_
    ole.Invoke(ole.GetIDsOfNames('NumberFormat'), 1033, pythoncom.DISPATCH_PROPERTYPUT, 0, code)


def serial(v):
    'Fecha u hora de Python como numero de serie de Excel (dias desde el 30/12/1899).'
    if isinstance(v, datetime.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    if not isinstance(v, datetime.datetime):
        v = datetime.datetime(v.year, v.month, v.day)
    return (v.replace(tzinfo=None) - datetime.datetime(1899, 12, 30)).total_seconds() / 86400


def to_cell(v):
    'Un valor ya interpretado como va por COM: fechas como numero de serie, numeros.Texto con apostrofo.'
    import numeros
    if v is None:
        return ''
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return serial(v)
    if isinstance(v, numeros.Texto):
        return "'" + v
    return v


def date_format(v):
    if isinstance(v, datetime.time):
        return 'hh:mm'
    if isinstance(v, datetime.datetime) and v.time() != datetime.time():
        return 'dd/mm/yyyy hh:mm'
    return 'dd/mm/yyyy'


def addr(r):
    return str(r.Address).replace('$', '')


class ToolBox:
    def __init__(self, sesion, confirm=None, approval='ask'):
        self.s = sesion
        self.confirm = confirm          # (kind, title, detail) -> bool; corre en este hilo y espera a la ventana
        self.approval = approval        # ask | edits | all
        self.specs = SPECS

    def execute(self, name, args, cancel=None):
        'Ejecuta una herramienta y devuelve SIEMPRE texto. Los errores vuelven como ERROR: ... para que el modelo se corrija.'
        fn = getattr(self, 'tool_' + str(name), None)
        if fn is None or name not in NAMES:
            return f"ERROR: unknown tool '{name}'. Available: {', '.join(sorted(NAMES))}"
        if not isinstance(args, dict):
            return 'ERROR: arguments must be a JSON object'
        try:
            inspect.signature(fn).bind(**args)
        except TypeError as e:
            return f'ERROR: bad arguments for {name}: {e}'
        if cancel is not None:
            self.s.detener = cancel     # lo que recorre Outlook se corta con el mismo boton Detener
        try:
            return fn(**args)
        except ToolError as e:
            return f'ERROR: {e}'
        except (TypeError, ValueError) as e:
            return f'ERROR: {e}'
        except excel.ExcelError as e:
            return f'ERROR: {e}'
        except OSError as e:
            return f'ERROR: {e}'
        except Exception as e:          # noqa: BLE001 - un error COM de Excel vuelve como texto, no rompe el bucle
            return f'ERROR: Excel refused: {com_text(e)}'

    def _ask(self, kind, title, detail):
        # kind: 'edit' (pisa datos), 'delete' (borra filas, hojas, duplicados), 'save' (pisa el archivo original)
        if self.approval == 'all' or (kind == 'edit' and self.approval == 'edits'):
            return True
        if self.confirm is None:
            return False
        return bool(self.confirm(kind, title, detail))

    def _denied(self, what):
        return (f'DENIED: the user did not allow {what}. Do not retry the same change; '
                'ask the user what they want instead.')

    # ------------------------------------------------------------ acceso al libro

    def _wb(self):
        if self.s.wb is None:
            raise ToolError("no workbook is open. Ask the user to open one with the 'Abrir Excel' button")
        return self.s.wb

    def _sheet(self, sheet=None):
        wb = self._wb()
        if not sheet:
            return wb.ActiveSheet
        names = [ws.Name for ws in wb.Worksheets]
        for n in names:
            if n.lower() == str(sheet).lower():
                return wb.Worksheets(n)
        raise ToolError(f"there is no sheet '{sheet}'. Sheets: {', '.join(names)}")

    def _range(self, h, ref):
        try:
            return h.Range(str(ref).replace('$', ''))
        except Exception:
            raise ToolError(f"'{ref}' is not a valid range (use A1 notation like B2 or A1:D10)")

    def _used(self, h):
        return h.UsedRange

    def _block(self, h, ref=None):
        'El rango pedido o, si no se indico, el area usada de la hoja.'
        r = self._range(h, ref) if ref else self._used(h)
        if r.Count == 1 and r.Value2 is None and not ref:
            raise ToolError(f"sheet '{h.Name}' is empty")
        return r

    def _filled(self, r):
        'Celdas no vacias dentro de r (hasta 10 para mostrar) y el total.'
        n = int(self.s.app.WorksheetFunction.CountA(r))
        if n == 0:
            return 0, []
        cells = []
        f = grid(r.Formula)
        for i, row in enumerate(f):
            for j, v in enumerate(row):
                if v not in (None, ''):
                    cells.append(f'{excel.col_letter(r.Column + j)}{r.Row + i}={str(v)[:40]}')
                    if len(cells) >= 10:
                        return n, cells
        return n, cells

    def _errors_in(self, r):
        out = []
        for i, row in enumerate(grid(r.Value2)):
            for j, v in enumerate(row):
                if isinstance(v, int) and v in CELL_ERRORS:
                    out.append(f'{excel.col_letter(r.Column + j)}{r.Row + i} {CELL_ERRORS[v]}')
        return out

    def _preview(self, r, rows=5, cols=8):
        g = grid(r.Value2)
        lines = []
        for i, row in enumerate(g[:rows]):
            lines.append(f'{r.Row + i}: ' + ' | '.join(show(v) for v in row[:cols]))
        return '\n'.join(lines)

    def _col_index(self, h, block, column):
        'Una columna por letra (C) o por el texto de su encabezado; devuelve el numero de columna de la hoja.'
        c = str(column).strip()
        if c.isalpha() and len(c) <= 3:
            return excel.col_num(c)
        heads = grid(h.Range(h.Cells(block.Row, block.Column), h.Cells(block.Row, block.Column + block.Columns.Count - 1)).Value2)[0]
        for j, v in enumerate(heads):
            if show(v).strip().lower() == c.lower():
                return block.Column + j
        raise ToolError(f"no column '{column}'. Headers: {', '.join(show(v) for v in heads)}")

    # ------------------------------------------------------------ lectura

    def tool_workbook_info(self):
        wb = self._wb()
        s = self.s
        lines = [f'File: {s.ruta}', f'Active sheet: {wb.ActiveSheet.Name}',
                 f"Unsaved changes: {'yes' if not wb.Saved else 'no'}"]
        if getattr(wb, 'ReadOnly', False):
            lines.append('The file is open READ-ONLY: save is not possible, use save_copy.')
        if s.respaldo:
            lines.append(f'Backup copy made when it was opened: {s.respaldo}')
        for ws in wb.Worksheets:
            u = ws.UsedRange
            empty = u.Count == 1 and u.Value2 is None
            desc = 'empty' if empty else f'used range {addr(u)} ({u.Rows.Count} rows x {u.Columns.Count} columns)'
            extra = []
            if ws.ListObjects.Count:
                extra.append('tables: ' + ', '.join(f'{t.Name} {addr(t.Range)}' for t in ws.ListObjects))
            if ws.ChartObjects().Count:
                extra.append(f'{ws.ChartObjects().Count} chart(s)')
            if ws.AutoFilterMode:
                extra.append('autofilter on')
            lines.append(f"- Sheet '{ws.Name}': {desc}" + (' · ' + ' · '.join(extra) if extra else ''))
            if not empty:
                lines.append(self._preview(u, rows=3))
        if s.adjuntos:
            lines.append('Attachments loaded by the user: ' + ', '.join(f"{i}. {a['nombre']}" for i, a in enumerate(s.adjuntos, 1)))
        return '\n'.join(lines)

    def tool_read_range(self, range=None, sheet=None, formulas=False, as_displayed=False):
        h = self._sheet(sheet)
        r = self._block(h, range)
        total = r.Rows.Count * r.Columns.Count
        limit = MAX_TEXT_CELLS if as_displayed else MAX_READ_CELLS
        note = ''
        if total > limit:
            rows = max(1, limit // r.Columns.Count)
            cols = min(r.Columns.Count, limit)
            r = h.Range(h.Cells(r.Row, r.Column), h.Cells(r.Row + rows - 1, r.Column + cols - 1))
            note = f'\n[only the first {rows} rows x {cols} columns of {total} cells; read the rest in smaller ranges]'
        if as_displayed:
            g = [[r.Cells(i, j).Text for j in range_(1, r.Columns.Count + 1)] for i in range_(1, r.Rows.Count + 1)]
        else:
            g = grid(r.Formula if formulas else r.Value2)
        first = r.Column
        head = 'row | ' + ' | '.join(excel.col_letter(first + j) for j in range_(r.Columns.Count))
        lines = [f"Sheet '{h.Name}' {addr(r)}:", head]
        for i, row in enumerate(g):
            lines.append(f'{r.Row + i} | ' + ' | '.join(show(v) for v in row))
        return '\n'.join(lines) + note

    def tool_find(self, text, sheet=None, whole_cell=False):
        sheets = [self._sheet(sheet)] if sheet else list(self._wb().Worksheets)
        hits = []
        for h in sheets:
            u = h.UsedRange
            c = u.Find(What=str(text), LookIn=-4163, LookAt=1 if whole_cell else 2, MatchCase=False)
            if c is None:
                continue
            first = c.Address
            while c is not None and len(hits) < MAX_FIND_HITS:
                hits.append(f"'{h.Name}'!{addr(c)} = {show(c.Value2)}")
                c = u.FindNext(c)
                if c is None or c.Address == first:
                    break
        if not hits:
            return f"'{text}' was not found"
        more = f'\n[stopped at {MAX_FIND_HITS} matches]' if len(hits) >= MAX_FIND_HITS else ''
        return f'{len(hits)} match(es):\n' + '\n'.join(hits) + more

    # ------------------------------------------------------------ escritura

    def tool_write_range(self, start_cell, values, sheet=None):
        if not isinstance(values, list) or not values:
            raise ToolError('values must be a non-empty list of rows, like [["Name","Total"],["Ana",10]]')
        rows = [r if isinstance(r, list) else [r] for r in values]
        width = max(len(r) for r in rows)
        if not width:
            # Caso real (Qwen3.5-9B local): [[]] respondia 'Wrote 1 x 0 cells' sin escribir nada, y el modelo lo repitio 50 veces
            raise ToolError('values has no cells: each row must list its cell values, like [["Total","=SUM(B2:B5)"]]')
        if len(rows) * width > MAX_WRITE_CELLS:
            raise ToolError(f'too many cells in one call ({len(rows) * width}); write at most {MAX_WRITE_CELLS}')
        for r in rows:
            if any(isinstance(v, (dict, list)) for v in r):
                raise ToolError('each cell must be a text, a number, a boolean or null')
        # '1.500', '$ 1.500', '15%' o '01/09/2026' como texto: por Formula Excel los lee como en EE.UU. (1,5 y el 9 de
        # enero) o los deja como texto. Se convierten antes, con la convencion de cada columna.
        rows, formats = interpret_columns(rows, decimal_separator(self.s.app))
        data = tuple(tuple(to_cell(v) for v in r + [None] * (width - len(r))) for r in rows)
        h = self._sheet(sheet)
        start = self._range(h, start_cell).Cells(1, 1)
        target = h.Range(start, h.Cells(start.Row + len(data) - 1, start.Column + width - 1))
        n, sample = self._filled(target)
        if n and not self._ask('edit', f'Pisar {n} celda(s) con datos en {addr(target)}',
                               'Van a quedar reemplazadas, entre otras: ' + ', '.join(sample)):
            return self._denied(f'overwriting {n} non-empty cell(s) in {addr(target)}')
        target.Formula = data
        apply_formats(target, formats)
        errs = self._errors_in(target)
        out = f"Wrote {len(data)} x {width} cells in '{h.Name}'!{addr(target)}.\n" + self._preview(target)
        if errs:
            out += ('\nCells with errors: ' + ', '.join(errs[:20]) +
                    '. Check the formula (English function names, comma separators) and fix it.')
        return out

    def tool_fill_formula(self, range, formula, sheet=None):
        f = str(formula).strip()
        if not f.startswith('='):
            f = '=' + f
        h = self._sheet(sheet)
        r = self._range(h, range)
        n, sample = self._filled(r)
        if n and not self._ask('edit', f'Pisar {n} celda(s) con datos en {addr(r)}',
                               f'Se llenan con la formula {f}. Hoy tienen, entre otras: ' + ', '.join(sample)):
            return self._denied(f'overwriting {n} non-empty cell(s) in {addr(r)}')
        r.Formula = f
        errs = self._errors_in(r)
        out = f"Filled '{h.Name}'!{addr(r)} with {f} (relative references adjusted per cell).\n" + self._preview(r)
        if errs:
            out += '\nCells with errors: ' + ', '.join(errs[:20]) + '. Fix the formula.'
        return out

    def tool_format_range(self, range, sheet=None, bold=None, italic=None, font_color=None, fill_color=None,
                          font_size=None, number_format=None, horizontal_align=None, wrap_text=None, borders=None,
                          merge=None, column_width=None):
        h = self._sheet(sheet)
        r = self._range(h, range)
        # Todo se valida antes de tocar la hoja: un color mal escrito no deja el formato aplicado a medias.
        no_fill = fill_color and str(fill_color).lower() in ('none', 'ninguno', 'sin color')
        fcolor = color_value(font_color) if font_color else None
        bcolor = color_value(fill_color) if fill_color and not no_fill else None
        if horizontal_align and horizontal_align not in ALIGN:
            raise ToolError('horizontal_align must be left, center or right')
        done = []
        if bold is not None:
            r.Font.Bold = bool(bold); done.append('bold' if bold else 'not bold')
        if italic is not None:
            r.Font.Italic = bool(italic); done.append('italic' if italic else 'not italic')
        if fcolor is not None:
            r.Font.Color = fcolor; done.append(f'font color {font_color}')
        if no_fill:
            r.Interior.ColorIndex = -4142; done.append('no fill')
        elif bcolor is not None:
            r.Interior.Color = bcolor; done.append(f'fill {fill_color}')
        if font_size:
            r.Font.Size = float(font_size); done.append(f'size {font_size}')
        if number_format:
            code = NUMBER_FORMATS.get(str(number_format).lower(), str(number_format))
            set_number_format(r, code); done.append(f'number format {code}')
        if horizontal_align:
            r.HorizontalAlignment = ALIGN[horizontal_align]; done.append(f'aligned {horizontal_align}')
        if wrap_text is not None:
            r.WrapText = bool(wrap_text); done.append('wrap text' if wrap_text else 'no wrap')
        if borders is not None:
            if borders:
                for edge in (7, 8, 9, 10, 11, 12):      # contorno e interiores
                    try:
                        r.Borders(edge).LineStyle = 1
                    except Exception:
                        pass                            # una sola fila no tiene borde horizontal interior
                done.append('borders')
            else:
                r.Borders.LineStyle = -4142; done.append('no borders')
        if merge is not None:
            if merge:
                r.Merge(); r.HorizontalAlignment = ALIGN['center']; done.append('merged')
            else:
                r.UnMerge(); done.append('unmerged')
        if column_width:
            r.EntireColumn.ColumnWidth = float(column_width); done.append(f'column width {column_width}')
        if not done:
            raise ToolError('nothing to do: give at least one format option')
        return f"Formatted '{h.Name}'!{addr(r)}: " + ', '.join(done)

    def tool_autofit_columns(self, range=None, sheet=None):
        h = self._sheet(sheet)
        r = self._range(h, range) if range else h.UsedRange
        r.EntireColumn.AutoFit()
        return f"Column widths adjusted in '{h.Name}' ({addr(r)})"

    def tool_sort_range(self, key_column, range=None, sheet=None, descending=False, has_header=True):
        h = self._sheet(sheet)
        r = self._block(h, range)
        if r.Rows.Count < (3 if has_header else 2):
            raise ToolError('not enough rows to sort')
        col = self._col_index(h, r, key_column)
        if not r.Column <= col < r.Column + r.Columns.Count:
            raise ToolError(f'column {excel.col_letter(col)} is outside {addr(r)}')
        key = h.Range(h.Cells(r.Row + (1 if has_header else 0), col), h.Cells(r.Row + r.Rows.Count - 1, col))
        srt = h.Sort
        srt.SortFields.Clear()
        srt.SortFields.Add(Key=key, SortOn=0, Order=2 if descending else 1)
        srt.SetRange(r)
        srt.Header = 1 if has_header else 2
        srt.Apply()
        srt.SortFields.Clear()
        return (f"Sorted '{h.Name}'!{addr(r)} by column {excel.col_letter(col)} "
                f"({'descending' if descending else 'ascending'}).\n" + self._preview(r))

    def tool_autofilter(self, on=True, range=None, sheet=None):
        h = self._sheet(sheet)
        if not on:
            h.AutoFilterMode = False
            return f"Autofilter removed from '{h.Name}'"
        if h.AutoFilterMode:
            return f"'{h.Name}' already has an autofilter"
        r = self._block(h, range)
        r.AutoFilter()
        return f"Autofilter turned on in '{h.Name}'!{addr(r)}"

    def tool_create_table(self, range=None, sheet=None, style='TableStyleMedium2'):
        h = self._sheet(sheet)
        if h.ListObjects.Count > 0:
            raise ToolError(f"'{h.Name}' already has a table ({h.ListObjects(1).Name} at {addr(h.ListObjects(1).Range)})")
        r = self._block(h, range)
        if r.Rows.Count < 2:
            r = h.UsedRange
        if r.Rows.Count < 2:
            raise ToolError('a table needs at least two rows: headers and data')
        if h.AutoFilterMode:
            h.AutoFilterMode = False        # una tabla no se puede crear sobre un autofiltro
        t = h.ListObjects.Add(1, r, None, 1)
        try:
            t.TableStyle = style
        except Exception:
            t.TableStyle = 'TableStyleMedium2'
        return f"Table {t.Name} created in '{h.Name}'!{addr(t.Range)} with {t.ListRows.Count} data rows"

    def tool_remove_duplicates(self, range=None, columns=None, sheet=None, has_header=True):
        h = self._sheet(sheet)
        r = self._block(h, range)
        if columns:
            idx = [self._col_index(h, r, c) - r.Column + 1 for c in columns]
        else:
            idx = list(range_(1, r.Columns.Count + 1))
        before = excel._filas_con_datos(h)
        what = 'all columns' if not columns else 'columns ' + ', '.join(map(str, columns))
        if not self._ask('delete', f'Quitar filas duplicadas en {addr(r)}',
                         f'Se borran las filas repetidas (comparando {what}); queda la primera de cada grupo.'):
            return self._denied('removing duplicate rows')
        app = self.s.app
        state, visible = app.WindowState, app.Visible
        try:
            # RemoveDuplicates falla con la ventana minimizada: se oculta y se pasa a normal un instante.
            if state == excel.XL_MINIMIZADO:
                app.Visible = False
                app.WindowState = excel.XL_NORMAL
            r.RemoveDuplicates(tuple(idx), 1 if has_header else 2)
        finally:
            if state == excel.XL_MINIMIZADO:
                app.WindowState = state
                app.Visible = visible
        after = excel._filas_con_datos(h)
        return f"Removed {before - after} duplicate row(s) in '{h.Name}' (rows with data: {before} -> {after})"

    def tool_replace(self, find, replace_with, range=None, sheet=None, whole_cell=False, match_case=False):
        h = self._sheet(sheet)
        r = self._block(h, range)
        count = int(self.s.app.WorksheetFunction.CountIf(r, str(find) if whole_cell else f'*{find}*'))
        if count == 0:
            return f"'{find}' was not found in '{h.Name}'!{addr(r)}; nothing replaced"
        if not self._ask('edit', f"Reemplazar '{find}' por '{replace_with}'",
                         f"En '{h.Name}'!{addr(r)}: unas {count} celda(s)."):
            return self._denied(f"replacing '{find}'")
        r.Replace(What=str(find), Replacement=str(replace_with), LookAt=1 if whole_cell else 2, MatchCase=bool(match_case))
        return f"Replaced '{find}' with '{replace_with}' in about {count} cell(s) of '{h.Name}'!{addr(r)}"

    def tool_clear_range(self, range, what='contents', sheet=None):
        h = self._sheet(sheet)
        r = self._range(h, range)
        if what not in ('contents', 'formats', 'all'):
            raise ToolError("what must be 'contents', 'formats' or 'all'")
        n, sample = self._filled(r)
        if what != 'formats' and n and not self._ask('edit', f'Borrar {n} celda(s) con datos en {addr(r)}',
                                                   'Se pierden, entre otras: ' + ', '.join(sample)):
            return self._denied(f'clearing {addr(r)}')
        {'contents': r.ClearContents, 'formats': r.ClearFormats, 'all': r.Clear}[what]()
        return f"Cleared {what} of '{h.Name}'!{addr(r)}"

    def _whole(self, h, where):
        w = str(where).replace('$', '').strip().upper()
        parts = w.split(':')
        if len(parts) == 1:
            w = w + ':' + w
            parts = [parts[0], parts[0]]
        if all(p.isdigit() for p in parts):
            return h.Range(w).EntireRow, 'row(s) ' + w
        if all(p.isalpha() for p in parts):
            return h.Range(w).EntireColumn, 'column(s) ' + w
        raise ToolError("where must be whole rows like '5' or '5:7', or whole columns like 'C' or 'C:D'")

    def tool_insert_rows_columns(self, where, sheet=None):
        h = self._sheet(sheet)
        r, what = self._whole(h, where)
        r.Insert()
        return f"Inserted empty {what} in '{h.Name}' (what was there moved down/right)"

    def tool_delete_rows_columns(self, where, sheet=None):
        h = self._sheet(sheet)
        r, what = self._whole(h, where)
        n = int(self.s.app.WorksheetFunction.CountA(r))
        if not self._ask('delete', f'Eliminar {what}', f"En la hoja '{h.Name}'. Tienen {n} celda(s) con datos."):
            return self._denied(f'deleting {what}')
        r.Delete()
        return f"Deleted {what} in '{h.Name}' (what was after moved up/left)"

    def tool_sheet(self, action, name, new_name=None):
        wb = self._wb()
        names = [ws.Name for ws in wb.Worksheets]
        if action == 'add':
            if name in names:
                raise ToolError(f"a sheet named '{name}' already exists")
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = name
            return f"Sheet '{name}' added and activated"
        h = self._sheet(name)
        if action == 'activate':
            h.Activate()
            return f"Sheet '{h.Name}' is now active"
        if action == 'rename':
            if not new_name:
                raise ToolError('rename needs new_name')
            old = h.Name
            h.Name = new_name
            return f"Sheet '{old}' renamed to '{new_name}'"
        if action == 'copy':
            h.Copy(After=wb.Worksheets(wb.Worksheets.Count))
            copy = wb.Worksheets(wb.Worksheets.Count)
            if new_name:
                copy.Name = new_name
            return f"Sheet '{h.Name}' copied as '{copy.Name}'"
        if action == 'delete':
            if len(names) == 1:
                raise ToolError('a workbook needs at least one sheet')
            n = int(self.s.app.WorksheetFunction.CountA(h.Cells))
            if not self._ask('delete', f"Eliminar la hoja '{h.Name}'", f'Tiene {n} celda(s) con datos.'):
                return self._denied(f"deleting sheet '{h.Name}'")
            h.Delete()
            return f"Sheet '{name}' deleted"
        raise ToolError('action must be add, activate, rename, copy or delete')

    def tool_add_chart(self, range, chart_type='column', title=None, sheet=None, at_cell=None):
        if chart_type not in CHART_TYPES:
            raise ToolError('chart_type must be one of: ' + ', '.join(CHART_TYPES))
        h = self._sheet(sheet)
        r = self._range(h, range)
        u = h.UsedRange
        anchor = self._range(h, at_cell) if at_cell else h.Cells(u.Row, u.Column + u.Columns.Count + 1)
        co = h.ChartObjects().Add(anchor.Left, anchor.Top, 420, 250)
        ch = co.Chart
        ch.ChartType = CHART_TYPES[chart_type]
        ch.SetSourceData(r)
        if title:
            ch.HasTitle = True
            ch.ChartTitle.Text = str(title)
        return f"{chart_type} chart of '{h.Name}'!{addr(r)} placed at {addr(anchor)}"

    def tool_conditional_format(self, range, rule, value=None, value2=None, fill_color=None, font_color=None, sheet=None):
        fill = color_value(fill_color or '#FFC7CE')     # antes de crear la regla: sin color valido no se crea
        font = color_value(font_color) if font_color else None
        h = self._sheet(sheet)
        r = self._range(h, range)
        fc = r.FormatConditions
        if rule in COMPARE_OPS:
            if value is None or (rule in ('between', 'not_between') and value2 is None):
                raise ToolError(f'rule {rule} needs value' + (' and value2' if 'between' in rule else ''))
            # Formula1 se interpreta en el idioma de Excel: los numeros viajan como numero, el texto como ="texto".
            def operand(v):
                return v if isinstance(v, (int, float)) else f'="{v}"'
            if 'between' in rule:
                cond = fc.Add(1, COMPARE_OPS[rule], operand(value), operand(value2))
            else:
                cond = fc.Add(1, COMPARE_OPS[rule], operand(value))
        elif rule == 'text_contains':
            if value is None:
                raise ToolError('text_contains needs value')
            # Con argumentos por nombre Excel contesta 'parametro no opcional': va todo por posicion (xlTextString, xlContains).
            cond = fc.Add(9, 1, '', '', str(value), 0)
        elif rule == 'duplicates':
            cond = fc.AddUniqueValues()
            cond.DupeUnique = 1
        elif rule == 'color_scale':
            fc.AddColorScale(3)
            return f"Color scale applied to '{h.Name}'!{addr(r)}"
        elif rule == 'data_bar':
            fc.AddDatabar()
            return f"Data bars applied to '{h.Name}'!{addr(r)}"
        else:
            raise ToolError('rule must be one of: ' + ', '.join(list(COMPARE_OPS) + ['text_contains', 'duplicates', 'color_scale', 'data_bar']))
        cond.Interior.Color = fill
        if font is not None:
            cond.Font.Color = font
        return f"Conditional format '{rule}' applied to '{h.Name}'!{addr(r)}"

    def tool_data_validation_list(self, range, items, sheet=None):
        if not isinstance(items, list) or not items:
            raise ToolError('items must be a non-empty list of the allowed values')
        h = self._sheet(sheet)
        r = self._range(h, range)
        sep = list_separator(self.s.app)
        v = r.Validation
        v.Delete()
        v.Add(Type=3, AlertStyle=1, Operator=1, Formula1=sep.join(str(i) for i in items))
        v.InCellDropdown = True
        return f"Drop-down list with {len(items)} option(s) set in '{h.Name}'!{addr(r)}"

    def tool_freeze_panes(self, cell=None, sheet=None):
        h = self._sheet(sheet)
        wb = self._wb()
        h.Activate()
        win = wb.Windows(1)
        win.FreezePanes = False
        if not cell:
            return f"Panes unfrozen in '{h.Name}'"
        c = self._range(h, cell).Cells(1, 1)
        win.ScrollRow = 1
        win.ScrollColumn = 1
        win.SplitRow = c.Row - 1
        win.SplitColumn = c.Column - 1
        win.FreezePanes = True
        return f"Frozen in '{h.Name}': rows above {c.Row} and columns left of {excel.col_letter(c.Column)} stay visible"

    # ------------------------------------------------------------ archivos

    def tool_save(self):
        wb = self._wb()
        if getattr(wb, 'ReadOnly', False):
            raise ToolError('the file is open read-only: it cannot be saved over. Use save_copy')
        if wb.Saved:
            return 'There are no unsaved changes'
        backup = f' A backup of the original is at {self.s.respaldo}.' if self.s.respaldo else ''
        if not self._ask('save', f'Guardar los cambios en {os.path.basename(self.s.ruta)}',
                         'Se reemplaza el archivo original.' + (f' Queda una copia de como estaba en {self.s.respaldo}' if self.s.respaldo else '')):
            return self._denied('saving over the original file')
        wb.Save()
        return f'Saved {self.s.ruta}.{backup}'

    def tool_save_copy(self, name, folder=None):
        self._wb()
        where = None
        if folder:
            where = excel.carpeta_conocida(folder) or os.path.abspath(os.path.expanduser(folder))
        dest = excel.destino_copia(self.s, str(name), where)
        return excel.op_guardar_como(self.s, dest) + ' (the open workbook is still the original)'

    def tool_export_csv(self, path=None, sheet=None):
        h = self._sheet(sheet)
        dest = path or os.path.splitext(self.s.ruta)[0] + '_' + h.Name + '.csv'
        if not dest.lower().endswith('.csv'):
            dest += '.csv'
        dest = os.path.abspath(dest)
        if os.path.exists(dest):
            raise ToolError(f'{dest} already exists: choose another name (existing files are never overwritten)')
        import csv
        with open(dest, 'w', encoding='utf-8-sig', newline='') as fh:
            w = csv.writer(fh, delimiter=';')
            for row in grid(h.UsedRange.Value2):
                w.writerow([show(v) for v in row])
        return f"Sheet '{h.Name}' exported to {dest} (separator ';', UTF-8)"

    def tool_excel_window(self, action='show'):
        self._wb()
        if action == 'show':
            return excel.op_mostrar_excel(self.s)
        if action == 'minimize':
            return excel.op_minimizar_excel(self.s)
        raise ToolError("action must be 'show' or 'minimize'")

    # ------------------------------------------------------------ adjuntos y Outlook

    def _attachment(self, index):
        if not self.s.adjuntos:
            raise ToolError("no attachments loaded. The user can add one with the clip button, above Enviar")
        i = len(self.s.adjuntos) if index is None else int(index)
        if not 1 <= i <= len(self.s.adjuntos):
            raise ToolError(f'index must be 1..{len(self.s.adjuntos)}')
        return self.s.adjuntos[i - 1]

    def _part(self, a, source):
        'La parte del adjunto a pegar (hoja, tabla): por nombre o numero; sin source, la primera con datos.'
        partes = a.get('partes') or [{'nombre': 'Texto', 'filas': a['filas'], 'columnas': a['columnas']}]
        if source in (None, ''):
            return next((p for p in partes if p['filas']), partes[0])
        src = str(source).strip()
        if src.isdigit() and 1 <= int(src) <= len(partes):
            return partes[int(src) - 1]
        for p in partes:
            if p['nombre'].lower() == src.lower():
                return p
        raise ToolError(f"no part '{src}' in {a['nombre']}. Parts: "
                        + ', '.join(f"{i}. {p['nombre']}" for i, p in enumerate(partes, 1)))

    def tool_read_attachment(self, index=None):
        a = self._attachment(index)
        text = a['texto']
        cut = f'\n[truncated: {len(text)} characters in total]' if len(text) > MAX_ATTACHMENT_CHARS else ''
        kind = {'imagen': 'image, text read by OCR', 'contactos': 'contact list, one row per contact',
                'planilla': 'spreadsheet', 'documento': 'document'}.get(a['clase'], 'text file')
        partes = a.get('partes') or []
        listado = ''
        if len(partes) > 1:
            listado = ('Parts (paste_attachment takes one by name or number in source): '
                       + '; '.join(f"{i}. {p['nombre']} ({len(p['filas'])} rows x {p['columnas']} columns)"
                                   for i, p in enumerate(partes, 1)) + '\n')
        return (f"{a['nombre']} ({kind}); detected {len(a['filas'])} rows x {a['columnas']} columns.\n"
                f"{listado}{text[:MAX_ATTACHMENT_CHARS]}{cut}")

    def tool_paste_attachment(self, start_cell='A1', index=None, sheet=None, source=None):
        a = self._attachment(index)
        parte = self._part(a, source)
        filas = parte['filas']
        if not filas:
            raise ToolError('the attachment has no usable text')
        h = self._sheet(sheet)
        width = max(1, parte['columnas'])
        if len(filas) * width > MAX_WRITE_CELLS:
            raise ToolError(f'{len(filas)} rows x {width} columns is more than {MAX_WRITE_CELLS} cells; '
                            'ask the user to open that file in Excel instead')
        start = self._range(h, start_cell).Cells(1, 1)
        target = h.Range(start, h.Cells(start.Row + len(filas) - 1, start.Column + width - 1))
        n, sample = self._filled(target)
        if n and not self._ask('edit', f'Pisar {n} celda(s) con datos en {addr(target)}',
                               f"Con el contenido de {a['nombre']}. Hoy tienen, entre otras: " + ', '.join(sample)):
            return self._denied(f'overwriting {n} non-empty cell(s) in {addr(target)}')
        # Telefonos como texto: si no, Excel les saca el '+' y los ceros, los muestra como 5,49116E+12, o toma
        # '+54 9 11 ...' por formula. Por el encabezado de la columna o por la forma del valor (una fecha con guiones
        # tambien tiene forma de telefono: se mira antes).
        # Los numeros y fechas escritos como texto se convierten (numeros.py); los de las planillas ya traen su tipo.
        # Las fechas van como numero de serie (Value2 no las acepta) con el formato puesto despues. El resto del texto
        # va con apostrofo: si no, Excel lo interpreta como en EE.UU. ('1-2' seria el 2 de enero, 'TRUE' un booleano).
        import contactos
        import numeros
        tel = [contactos.es_columna_telefono(t) for t in filas[0]] + [False] * width
        rows = []
        for r, f in enumerate(filas):
            row = []
            for i, v in enumerate(f):
                if isinstance(v, str) and v.strip() and (tel[i] or (contactos.parece_telefono(v) and not numeros.es_fecha(v))):
                    v = numeros.Texto(v.strip())
                elif r and tel[i] and isinstance(v, int) and not isinstance(v, bool):
                    v = numeros.Texto(str(v))
                row.append(v)
            rows.append(row)
        rows, formats = interpret_columns(rows, decimal_separator(self.s.app))
        con_formato = {(r, i) for r, i, _ in formats}
        for r, row in enumerate(rows):
            for i, v in enumerate(row):
                if isinstance(v, (datetime.datetime, datetime.date, datetime.time)) and (r, i) not in con_formato:
                    formats.append((r, i, date_format(v)))
                elif type(v) is str and v.strip():
                    row[i] = numeros.Texto(v)
        data = tuple(tuple([to_cell(v) for v in row] + [''] * (width - len(row))) for row in rows)
        target.Value2 = data
        extra = ''
        sin = apply_formats(target, formats, MAX_FORMAT_CALLS)
        if sin:
            extra += f'\n{sin} cells were left without their number or date format: give it with format_range.'
        otras = [p['nombre'] for p in a.get('partes') or [] if p is not parte and p['filas']]
        if otras:
            extra += f"\nOther parts of the attachment, not pasted: {', '.join(otras)} (paste them with source)."
        return (f"Pasted {a['nombre']} ({parte['nombre']}) into '{h.Name}'!{addr(target)}.\n"
                + self._preview(target) + extra)

    def tool_outlook_contacts(self, with_table=False, whole_mailbox=False):
        h = self._sheet()
        n = int(self.s.app.WorksheetFunction.CountA(h.Cells))
        if n and not self._ask('edit', f"Traer contactos de Outlook a '{h.Name}'",
                               f'Se escriben desde A1 y la hoja ya tiene {n} celda(s) con datos.'):
            return self._denied(f"writing contacts over the data of '{h.Name}'")
        return excel.op_contactos_outlook(self.s, bool(with_table), bool(whole_mailbox))


range_ = range      # los parametros llamados 'range' tapan el builtin dentro de los metodos


def _fn(name, desc, props, required):
    return {'type': 'function', 'function': {'name': name, 'description': desc,
            'parameters': {'type': 'object', 'properties': props, 'required': required}}}


_SHEET = {'type': 'string', 'description': 'Sheet name; default: the active sheet'}
_RANGE_OPT = {'type': 'string', 'description': 'A1-style range; default: the used area of the sheet'}

SPECS = [
    _fn('workbook_info', 'Overview of the open workbook: file, sheets with their used range and first rows, tables, '
        'charts, unsaved changes, attachments. Call it first when you do not know the data yet.', {}, []),
    _fn('read_range', 'Read cells. Returns a grid with row numbers and column letters. Dates come as Excel serial '
        f'numbers unless as_displayed is true. At most {MAX_READ_CELLS} cells per call.',
        {'range': _RANGE_OPT, 'sheet': _SHEET,
         'formulas': {'type': 'boolean', 'description': 'Show formulas instead of their results'},
         'as_displayed': {'type': 'boolean', 'description': f'Text exactly as shown in Excel (dates, currency); max {MAX_TEXT_CELLS} cells'}}, []),
    _fn('find', 'Find cells that contain a text, in one sheet or in all of them.',
        {'text': {'type': 'string'}, 'sheet': {'type': 'string', 'description': 'Default: all sheets'},
         'whole_cell': {'type': 'boolean', 'description': 'Match the whole cell only'}}, ['text']),
    _fn('write_range', 'Write a block of values starting at start_cell. A string starting with "=" is a formula, '
        'in ENGLISH syntax with comma separators (=SUM(B2:B9), =IF(C2>100,"High","Low")). Dates can be written as '
        'text like "15/03/2026". To keep a code like 007 as text, write it as "\'007". '
        'Asks the user before overwriting cells that have data.',
        {'start_cell': {'type': 'string', 'description': 'Top-left cell, e.g. A1'},
         'values': {'type': 'array', 'description': 'List of rows; each row is a list of cells (text, number, boolean or null)',
                    'items': {'type': 'array', 'items': {}}},
         'sheet': _SHEET}, ['start_cell', 'values']),
    _fn('fill_formula', 'Put the same formula in every cell of a range, written for the FIRST cell; relative '
        'references shift per row/column like dragging the fill handle. English syntax, comma separators.',
        {'range': {'type': 'string', 'description': 'e.g. D2:D50'}, 'formula': {'type': 'string', 'description': 'e.g. =B2*C2'},
         'sheet': _SHEET}, ['range', 'formula']),
    _fn('format_range', 'Change how cells look. Colors: #RRGGBB or names (red, green, yellow, blue, orange, grey, '
        'lightblue, pink, black, white; also in Spanish). number_format: a preset (currency, number, integer, percent, '
        'date, datetime, text, general) or an English Excel format code like "#,##0.00".',
        {'range': {'type': 'string'}, 'sheet': _SHEET, 'bold': {'type': 'boolean'}, 'italic': {'type': 'boolean'},
         'font_color': {'type': 'string'}, 'fill_color': {'type': 'string', 'description': "Background; 'none' removes it"},
         'font_size': {'type': 'number'}, 'number_format': {'type': 'string'},
         'horizontal_align': {'type': 'string', 'enum': ['left', 'center', 'right']},
         'wrap_text': {'type': 'boolean'}, 'borders': {'type': 'boolean', 'description': 'true: thin borders on every cell; false: none'},
         'merge': {'type': 'boolean', 'description': 'Merge the range into one cell (use rarely, for titles)'},
         'column_width': {'type': 'number', 'description': 'Width of the columns of the range, in characters'}}, ['range']),
    _fn('autofit_columns', 'Adjust column widths to their content.', {'range': _RANGE_OPT, 'sheet': _SHEET}, []),
    _fn('sort_range', 'Sort the rows of a block by one column.',
        {'key_column': {'type': 'string', 'description': 'Column letter (C) or header text (Importe)'},
         'range': _RANGE_OPT, 'sheet': _SHEET, 'descending': {'type': 'boolean'},
         'has_header': {'type': 'boolean', 'description': 'First row is headers (default true)'}}, ['key_column']),
    _fn('autofilter', 'Turn the filter arrows on the header row on or off.',
        {'on': {'type': 'boolean', 'description': 'Default true'}, 'range': _RANGE_OPT, 'sheet': _SHEET}, []),
    _fn('create_table', 'Convert a block with headers into an Excel table (banded rows, filters, auto-expanding).',
        {'range': _RANGE_OPT, 'sheet': _SHEET, 'style': {'type': 'string', 'description': 'e.g. TableStyleMedium2 (default), TableStyleLight9'}}, []),
    _fn('remove_duplicates', 'Delete repeated rows, keeping the first of each group. Asks the user first.',
        {'range': _RANGE_OPT, 'sheet': _SHEET, 'has_header': {'type': 'boolean'},
         'columns': {'type': 'array', 'items': {'type': 'string'}, 'description': 'Columns to compare (letters or headers); default all'}}, []),
    _fn('replace', 'Replace text in cells. Asks the user first.',
        {'find': {'type': 'string'}, 'replace_with': {'type': 'string'}, 'range': _RANGE_OPT, 'sheet': _SHEET,
         'whole_cell': {'type': 'boolean'}, 'match_case': {'type': 'boolean'}}, ['find', 'replace_with']),
    _fn('clear_range', 'Erase cells. Asks the user before erasing data.',
        {'range': {'type': 'string'}, 'what': {'type': 'string', 'enum': ['contents', 'formats', 'all']}, 'sheet': _SHEET}, ['range']),
    _fn('insert_rows_columns', 'Insert empty whole rows (where="5" or "5:7") or columns (where="C" or "C:D").',
        {'where': {'type': 'string'}, 'sheet': _SHEET}, ['where']),
    _fn('delete_rows_columns', 'Delete whole rows or columns (same where syntax as insert). Asks the user first.',
        {'where': {'type': 'string'}, 'sheet': _SHEET}, ['where']),
    _fn('sheet', 'Add, activate, rename, copy or delete a sheet (delete asks the user first).',
        {'action': {'type': 'string', 'enum': ['add', 'activate', 'rename', 'copy', 'delete']},
         'name': {'type': 'string'}, 'new_name': {'type': 'string'}}, ['action', 'name']),
    _fn('add_chart', 'Create a chart from a block (first column = categories, first row = series names).',
        {'range': {'type': 'string'}, 'chart_type': {'type': 'string', 'enum': list(CHART_TYPES)},
         'title': {'type': 'string'}, 'sheet': _SHEET,
         'at_cell': {'type': 'string', 'description': 'Top-left cell where the chart goes; default: right of the data'}}, ['range']),
    _fn('conditional_format', 'Highlight cells automatically by a rule.',
        {'range': {'type': 'string'}, 'sheet': _SHEET,
         'rule': {'type': 'string', 'enum': list(COMPARE_OPS) + ['text_contains', 'duplicates', 'color_scale', 'data_bar']},
         'value': {'description': 'Number or text to compare with'}, 'value2': {'description': 'Upper limit for between'},
         'fill_color': {'type': 'string', 'description': 'Default light red'}, 'font_color': {'type': 'string'}}, ['range', 'rule']),
    _fn('data_validation_list', 'Make cells a drop-down list that only accepts the given options.',
        {'range': {'type': 'string'}, 'items': {'type': 'array', 'items': {'type': 'string'}}, 'sheet': _SHEET}, ['range', 'items']),
    _fn('freeze_panes', 'Keep the rows above and the columns left of a cell always visible while scrolling '
        '(A2 freezes the header row). Without cell, unfreezes.', {'cell': {'type': 'string'}, 'sheet': _SHEET}, []),
    _fn('save', 'Save the workbook over its original file. Only when the user asked to save. Asks the user first.', {}, []),
    _fn('save_copy', 'Save a copy with another name, same format; the original file is not touched. Never overwrites.',
        {'name': {'type': 'string', 'description': 'File name, e.g. "Ventas 2026"'},
         'folder': {'type': 'string', 'description': "Folder path, or 'escritorio'/'documentos'; default: next to the original"}}, ['name']),
    _fn('export_csv', "Export a sheet to a CSV file (';' separator). Never overwrites.",
        {'path': {'type': 'string', 'description': 'Default: next to the workbook'}, 'sheet': _SHEET}, []),
    _fn('excel_window', 'Show the Excel window to the user, or minimize it.',
        {'action': {'type': 'string', 'enum': ['show', 'minimize']}}, []),
    _fn('read_attachment', "Read a file the user attached with the clip button: text, Markdown or CSV; a spreadsheet "
        '(Excel or OpenOffice, every sheet); a document (Word, PDF, OpenOffice or RTF, its tables apart); contacts exported '
        'as vCard .vcf or CSV; or an image such as a screenshot, read by OCR.',
        {'index': {'type': 'integer', 'description': 'Attachment number; default the last one'}}, []),
    _fn('paste_attachment', 'Paste the rows of an attachment into the sheet: one sheet of a spreadsheet, one table of a '
        'document, or the lines of a text split by tab, ; or ,. Numbers and dates are kept, also when written as text. Contacts '
        'come as Nombre, Telefono, Otros telefonos, Correo; phone numbers are kept as text.',
        {'start_cell': {'type': 'string'}, 'index': {'type': 'integer'}, 'sheet': _SHEET,
         'source': {'type': 'string', 'description': 'Which part of the attachment (sheet name, or its number as '
                    'read_attachment lists them); default the first one with data'}}, []),
    _fn('outlook_contacts', "Bring the contacts from the user's Outlook into the active sheet from A1 (if the "
        'Contacts folder is empty, collects senders from Inbox and Sent). Can take minutes.',
        {'with_table': {'type': 'boolean', 'description': 'Also turn them into a table'},
         'whole_mailbox': {'type': 'boolean', 'description': 'Scan every mail instead of the 1200 most recent per '
                           'folder; only if the user asks for all of them (about 16 minutes for 20,000 mails)'}}, []),
]
NAMES = {s['function']['name'] for s in SPECS}
