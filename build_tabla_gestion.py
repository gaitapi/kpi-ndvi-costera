import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

df = pd.read_csv(
    '/home/gaitapi/proyectos/kpi-costa/script para correr en GEE el NDVI de todos los tramos costeros 45 metros/KPI_NDVI_tendencia.csv'
)
df['d'] = df['d'].str.strip().str.lower().str.replace('san josé','san jose')

C = {
    'header_bg':   '0D2640', 'header_fg':  'EEE8DC',
    'ndvi_hdr':    '144A7A', 'tend_hdr':   '1A5C3A',
    'inst_hdr':    '6B4C1E', 'rest_hdr':   '4A1E6B',
    'alt_row':     'F5F8FC', 'white':      'FFFFFF',
    'vul1_bg':     'FCE4EC', 'vul2_bg':    'FFF9C4', 'vul3_bg': 'E8F5E9',
    'tend_mej':    'C8E6C9', 'tend_est':   'FFF9C4', 'tend_deg': 'FFCDD2',
}

DEPTS = ['colonia','san jose','montevideo','canelones','maldonado','rocha']
DEPT_NAMES = {
    'colonia':'Colonia','san jose':'San José','montevideo':'Montevideo',
    'canelones':'Canelones','maldonado':'Maldonado','rocha':'Rocha',
}
VUL_LABEL = {1:'Alta', 2:'Media', 3:'Baja'}
ACC_LABEL = {0:'0', 1:'1', 2:'2', 3:'3'}   # acciones registradas (valor numérico original)
NDVI_YEARS = [2017,2018,2019,2020,2021,2022,2023,2024]

INSTRUMENTOS = [
    'Plan de manejo costero',
    'Guía / protocolo de manejo',
    'Lineamientos de gestión',
    'Anteproyecto de obra',
    'Proyecto ejecutivo',
]
INST_OPTS  = '"No existe,En elaboración,Vigente,Desactualizado"'
ACCION_OPTS = '"No,Sí"'
YEAR_OPTS  = '"' + ','.join(str(y) for y in range(2017, 2041)) + '"'

def fill(h): return PatternFill('solid', fgColor=h)
def font(bold=False, color='1A2636', size=9, name='Arial'):
    return Font(bold=bold, color=color, size=size, name=name)
def align(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def border():
    s = Side(style='thin', color='BFCAD8')
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(ws, row, col, text, bg, fg='EEE8DC', bold=True, size=9, wrap=True, h='center'):
    c = ws.cell(row=row, column=col, value=text)
    c.fill = fill(bg); c.font = Font(bold=bold, color=fg, size=size, name='Arial')
    c.alignment = align(h=h, v='center', wrap=wrap); c.border = border()

def ndvi_color(val):
    if pd.isna(val): return 'D0D0D0'
    t = min(max(val / 0.75, 0), 1)
    r = int(183*(1-t) + 46*t); g = int(28*(1-t) + 125*t); b = int(28*(1-t) + 50*t)
    return f'{r:02X}{g:02X}{b:02X}'

wb = Workbook()
wb.remove(wb.active)

ACCIONES = [
    'Plantación / Revegetación',
    'Retiro de especies exóticas',
    'Manejo de accesos',
    'Manejo de pisoteo / senderos',
    'Obra de estabilización',
    'Monitoreo activo',
    'Otras acciones',
]

for dept_key in DEPTS:
    dept_df = df[df['d'] == dept_key].copy().reset_index(drop=True)
    if dept_df.empty: continue
    ws = wb.create_sheet(title=DEPT_NAMES[dept_key])
    ws.freeze_panes = 'A4'
    ws.sheet_view.zoomScale = 85

    # ── Definición de columnas ────────────────────────────────────────────────
    COL_TRAMO  = 1; COL_LARGO = 2; COL_VUL = 3; COL_ACC = 4
    ndvi_s = 5; ndvi_e = ndvi_s + len(NDVI_YEARS) - 1
    COL_TEND = ndvi_e+1; COL_SLOPE = ndvi_e+2; COL_MKP = ndvi_e+3
    inst_s = COL_MKP+1; inst_e = inst_s + len(INSTRUMENTOS) - 1
    rest_s = inst_e+1
    # 6 acciones normales × 2 cols + 1 acción "Otras" × 3 cols
    rest_e = rest_s + (len(ACCIONES)-1)*2 + 2
    TOTAL  = rest_e
    n_data = len(dept_df)
    last_data_row = 3 + n_data   # fila 3 = cabeceras, fila 4..N = datos

    # ── Filas 1-3: cabeceras ──────────────────────────────────────────────────
    grupos = [
        (COL_TRAMO, COL_ACC,   'IDENTIFICACIÓN',              C['header_bg']),
        (ndvi_s,    ndvi_e,    'NDVI ANUAL · Sentinel-2 · Buffer 45 m', C['ndvi_hdr']),
        (COL_TEND,  COL_MKP,   'TENDENCIA',                   C['tend_hdr']),
        (inst_s,    inst_e,    'INSTRUMENTOS DE MANEJO',      C['inst_hdr']),
        (rest_s,    rest_e,    'ACCIONES DE RESTAURACIÓN',    C['rest_hdr']),
    ]
    for c1, c2, label, bg in grupos:
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        hdr(ws, 1, c1, label, bg, size=10, bold=True)

    # Fila 2: sub-grupos
    for col in range(1, TOTAL+1):
        c = ws.cell(row=2, column=col); c.border = border()
        bg = C['header_bg']
        if ndvi_s <= col <= ndvi_e: bg = C['ndvi_hdr']
        elif COL_TEND <= col <= COL_MKP: bg = C['tend_hdr']
        elif inst_s <= col <= inst_e: bg = C['inst_hdr']
        elif rest_s <= col <= rest_e: bg = C['rest_hdr']
        c.fill = fill(bg)

    for i, inst in enumerate(INSTRUMENTOS):
        hdr(ws, 2, inst_s+i, inst, C['inst_hdr'], size=8, wrap=True)

    for i, acc in enumerate(ACCIONES[:-1]):
        c1 = rest_s + i*2
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c1+1)
        hdr(ws, 2, c1, acc, C['rest_hdr'], size=8, wrap=True)
    ob = rest_s + (len(ACCIONES)-1)*2
    ws.merge_cells(start_row=2, start_column=ob, end_row=2, end_column=ob+2)
    hdr(ws, 2, ob, 'Otras acciones', C['rest_hdr'], size=8, wrap=True)

    # Fila 3: nombres de columna
    h3 = {
        COL_TRAMO: ('Tramo', C['header_bg']),
        COL_LARGO: ('Largo (m)', C['header_bg']),
        COL_VUL:   ('Vulnerabilidad', C['header_bg']),
        COL_ACC:   ('Acc. registradas', C['header_bg']),
        COL_TEND:  ('Tendencia', C['tend_hdr']),
        COL_SLOPE: ('Slope Sen', C['tend_hdr']),
        COL_MKP:   ('MK p-val', C['tend_hdr']),
    }
    for col, (label, bg) in h3.items():
        hdr(ws, 3, col, label, bg, size=9)
    for i, yr in enumerate(NDVI_YEARS):
        hdr(ws, 3, ndvi_s+i, str(yr), C['ndvi_hdr'], size=9)
    for i in range(len(INSTRUMENTOS)):
        hdr(ws, 3, inst_s+i, 'Estado', C['inst_hdr'], size=8)
    for i in range(len(ACCIONES)-1):
        base = rest_s + i*2
        hdr(ws, 3, base,   'Realizada', C['rest_hdr'], size=8)
        hdr(ws, 3, base+1, 'Año',       C['rest_hdr'], size=8)
    hdr(ws, 3, ob,   'Realizada',   C['rest_hdr'], size=8)
    hdr(ws, 3, ob+1, 'Año',         C['rest_hdr'], size=8)
    hdr(ws, 3, ob+2, 'Descripción', C['rest_hdr'], size=8)

    # ── Datos ─────────────────────────────────────────────────────────────────
    vul_bgs = {1: C['vul1_bg'], 2: C['vul2_bg'], 3: C['vul3_bg']}
    vul_fgs = {1: '7B0000', 2: '5D4000', 3: '1B5E20'}

    for ridx, row_data in dept_df.iterrows():
        er = ridx + 4
        row_bg = C['alt_row'] if ridx % 2 == 1 else C['white']
        vul_val = int(row_data['v']) if pd.notna(row_data['v']) else 1

        def dc(col, value, fmt=None, h='center', bold=False, bg=None, fg=None):
            c = ws.cell(row=er, column=col, value=value)
            c.fill = fill(bg or row_bg)
            c.font = Font(bold=bold, color=fg or '1A2636', size=9, name='Arial')
            c.alignment = align(h=h, v='center'); c.border = border()
            if fmt: c.number_format = fmt

        dc(COL_TRAMO, int(row_data['t']), bold=True)
        dc(COL_LARGO, int(row_data['l']), fmt='#,##0')

        # Vulnerabilidad coloreada
        vc = ws.cell(row=er, column=COL_VUL, value=VUL_LABEL.get(vul_val, str(vul_val)))
        vc.fill = fill(vul_bgs.get(vul_val, C['white']))
        vc.font = Font(bold=True, color=vul_fgs.get(vul_val,'000000'), size=9, name='Arial')
        vc.alignment = align(); vc.border = border()

        acc_val = int(row_data['a']) if pd.notna(row_data['a']) else 0
        dc(COL_ACC, acc_val)

        # NDVI con color de celda
        for yi, yr in enumerate(NDVI_YEARS):
            col = ndvi_s + yi
            val = row_data.get(f'NDVI_{yr}', None)
            if pd.notna(val):
                v = float(val)
                nb = ndvi_color(v)
                lum = 0.299*int(nb[0:2],16) + 0.587*int(nb[2:4],16) + 0.114*int(nb[4:6],16)
                fg_c = 'FFFFFF' if lum < 128 else '1A2636'
                c = ws.cell(row=er, column=col, value=round(v,4))
                c.fill = fill(nb); c.font = Font(color=fg_c, size=8, name='Arial')
                c.alignment = align(); c.border = border(); c.number_format = '0.0000'
            else:
                dc(col, None)

        # Tendencia
        tend = str(row_data.get('tendencia','')).lower()
        t_bg = {'mejora':C['tend_mej'],'estable':C['tend_est'],'degradacion':C['tend_deg']}.get(tend,C['white'])
        t_fg = {'mejora':'1B5E20','estable':'5D4000','degradacion':'7B0000'}.get(tend,'000000')
        tc = ws.cell(row=er, column=COL_TEND, value=tend.capitalize())
        tc.fill=fill(t_bg); tc.font=Font(bold=True,color=t_fg,size=9,name='Arial')
        tc.alignment=align(); tc.border=border()

        sl = row_data.get('sens_slope', None)
        dc(COL_SLOPE, round(float(sl),6) if pd.notna(sl) else None, fmt='0.000000')
        mkp = row_data.get('mk_p', None)
        dc(COL_MKP, round(float(mkp),4) if pd.notna(mkp) else None, fmt='0.0000')

        # Instrumentos — celdas vacías (dropdown via validación de rango)
        for i in range(len(INSTRUMENTOS)):
            col = inst_s + i
            c = ws.cell(row=er, column=col, value=None)
            c.fill = fill(C['alt_row'] if ridx%2==1 else C['white'])
            c.font = Font(color='5A3E00', size=9, name='Arial')
            c.alignment = align(); c.border = border()

        # Acciones — celdas vacías (dropdown via validación de rango)
        for i in range(len(ACCIONES)-1):
            base = rest_s + i*2
            for off in range(2):
                c = ws.cell(row=er, column=base+off, value=None)
                c.fill = fill(C['alt_row'] if ridx%2==1 else C['white'])
                c.font = Font(color='3B1A5A', size=9, name='Arial')
                c.alignment = align(); c.border = border()
        for off in range(3):
            c = ws.cell(row=er, column=ob+off, value=None)
            c.fill = fill(C['alt_row'] if ridx%2==1 else C['white'])
            c.font = Font(color='3B1A5A', size=9, name='Arial')
            c.alignment = align(h='left' if off==2 else 'center'); c.border = border()

    # ── Data Validations sobre rangos de DATOS (filas 4 .. last_data_row) ────
    # Instrumentos: una DV por columna
    for i in range(len(INSTRUMENTOS)):
        col_l = get_column_letter(inst_s + i)
        dv = DataValidation(type='list', formula1=INST_OPTS, allow_blank=True,
                            showErrorMessage=False)
        dv.sqref = f'{col_l}4:{col_l}{last_data_row}'
        ws.add_data_validation(dv)

    # Acciones: DV para columnas "Realizada" y DV para columnas "Año"
    for i in range(len(ACCIONES)-1):
        base = rest_s + i*2
        col_r = get_column_letter(base)
        col_y = get_column_letter(base+1)
        dv_r = DataValidation(type='list', formula1=ACCION_OPTS, allow_blank=True,
                              showErrorMessage=False)
        dv_r.sqref = f'{col_r}4:{col_r}{last_data_row}'
        ws.add_data_validation(dv_r)
        dv_y = DataValidation(type='list', formula1=YEAR_OPTS, allow_blank=True,
                              showErrorMessage=False)
        dv_y.sqref = f'{col_y}4:{col_y}{last_data_row}'
        ws.add_data_validation(dv_y)

    # Otras acciones (3 cols)
    col_or = get_column_letter(ob)
    col_oy = get_column_letter(ob+1)
    dv_or = DataValidation(type='list', formula1=ACCION_OPTS, allow_blank=True,
                           showErrorMessage=False)
    dv_or.sqref = f'{col_or}4:{col_or}{last_data_row}'
    ws.add_data_validation(dv_or)
    dv_oy = DataValidation(type='list', formula1=YEAR_OPTS, allow_blank=True,
                           showErrorMessage=False)
    dv_oy.sqref = f'{col_oy}4:{col_oy}{last_data_row}'
    ws.add_data_validation(dv_oy)

    # ── Anchos de columna ─────────────────────────────────────────────────────
    ws.column_dimensions[get_column_letter(COL_TRAMO)].width  = 7
    ws.column_dimensions[get_column_letter(COL_LARGO)].width  = 9
    ws.column_dimensions[get_column_letter(COL_VUL)].width    = 13
    ws.column_dimensions[get_column_letter(COL_ACC)].width    = 14
    for i in range(len(NDVI_YEARS)):
        ws.column_dimensions[get_column_letter(ndvi_s+i)].width = 8
    ws.column_dimensions[get_column_letter(COL_TEND)].width   = 12
    ws.column_dimensions[get_column_letter(COL_SLOPE)].width  = 10
    ws.column_dimensions[get_column_letter(COL_MKP)].width    = 9
    for i in range(len(INSTRUMENTOS)):
        ws.column_dimensions[get_column_letter(inst_s+i)].width = 17
    for i in range(len(ACCIONES)-1):
        ws.column_dimensions[get_column_letter(rest_s+i*2)].width   = 13
        ws.column_dimensions[get_column_letter(rest_s+i*2+1)].width = 8
    ws.column_dimensions[get_column_letter(ob)].width   = 13
    ws.column_dimensions[get_column_letter(ob+1)].width = 8
    ws.column_dimensions[get_column_letter(ob+2)].width = 24

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 42
    ws.row_dimensions[3].height = 18
    for i in range(n_data):
        ws.row_dimensions[i+4].height = 16

    ws.auto_filter.ref = f'A3:{get_column_letter(TOTAL)}{last_data_row}'

# ── Hoja RESUMEN ──────────────────────────────────────────────────────────────
ws_r = wb.create_sheet(title='Resumen', index=0)
ws_r.sheet_view.zoomScale = 90
ws_r.merge_cells('A1:H1')
hdr(ws_r, 1, 1, 'KPI VEGETACIÓN COSTERA — RESUMEN POR DEPARTAMENTO',
    C['header_bg'], size=13, bold=True)
ws_r.row_dimensions[1].height = 30

for j, h_txt in enumerate(['Departamento','Tramos','Vul. Alta','Vul. Media','Vul. Baja',
                             'Mejora','Estable','Degradación'], 1):
    hdr(ws_r, 2, j, h_txt, C['header_bg'], size=10)

for ri, dk in enumerate(DEPTS, 1):
    dfs = df[df['d']==dk]
    row = ri+2; bg = C['alt_row'] if ri%2==0 else C['white']
    vals = [DEPT_NAMES[dk], len(dfs),
            int((dfs['v']==1).sum()), int((dfs['v']==2).sum()), int((dfs['v']==3).sum()),
            int((dfs['tendencia']=='mejora').sum()), int((dfs['tendencia']=='estable').sum()),
            int((dfs['tendencia']=='degradacion').sum())]
    for j, val in enumerate(vals, 1):
        c = ws_r.cell(row=row, column=j, value=val)
        c.fill = fill(bg)
        c.font = Font(bold=(j==1), size=10, name='Arial')
        c.alignment = align(h='left' if j==1 else 'center'); c.border = border()

tot = len(DEPTS)+3
hdr(ws_r, tot, 1, 'TOTAL', C['header_bg'], size=10)
for j in range(2,9):
    cl = get_column_letter(j)
    c = ws_r.cell(row=tot, column=j, value=f'=SUM({cl}3:{cl}{tot-1})')
    c.fill = fill(C['header_bg'])
    c.font = Font(bold=True, color='EEE8DC', size=10, name='Arial')
    c.alignment = align(); c.border = border()

for j in range(1,9):
    ws_r.column_dimensions[get_column_letter(j)].width = 18
ws_r.row_dimensions[1].height = 30

# ── Guardar ───────────────────────────────────────────────────────────────────
out = '/home/gaitapi/proyectos/kpi-ndvi-costera/data/KPI_Gestion_Costera.xlsx'
wb.save(out)
print(f'OK: {out}')
print(f'Hojas: {wb.sheetnames}')
