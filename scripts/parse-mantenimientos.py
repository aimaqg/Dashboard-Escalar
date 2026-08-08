# El problema: xlsx lee las celdas fusionadas como null en las filas siguientes
# openpyxl sí las lee correctamente. Reescribir el parser usando openpyxl directamente
# en el script de Node mediante un paso previo de Python, o mejorar el parser Node.

# Solución: usar el valor de la celda superior izquierda de cada rango fusionado
# Esto es lo que hace XLSX.js con cellMerge:true -- hay que activarlo

import openpyxl, json

def parse_hoja(ws, hdr_row_idx=2):
    """Lee una hoja respetando celdas fusionadas"""
    # Mapa de rangos fusionados: coordenada -> valor de la celda ancla
    merge_map = {}
    for merged in ws.merged_cells:
        # La celda ancla (superior izquierda) tiene el valor
        anchor = ws.cell(merged.min_row, merged.min_col)
        for row in range(merged.min_row, merged.max_row+1):
            for col in range(merged.min_col, merged.max_col+1):
                if not (row==merged.min_row and col==merged.min_col):
                    merge_map[(row,col)] = anchor.value

    def cel(row, col):
        key = (row, col)
        if key in merge_map: return merge_map[key]
        return ws.cell(row, col).value

    # Encabezados en hdr_row_idx (1-indexed)
    hdrs = [str(cel(hdr_row_idx, c) or '').strip().upper() for c in range(1, ws.max_column+1)]

    def colidx(patron):
        import re
        for i,h in enumerate(hdrs):
            if re.search(patron, h): return i+1  # 1-indexed para ws.cell
        return None

    iCopr  = colidx(r'COPROPIEDAD')
    iElem  = colidx(r'EQUIPO|ELEMENTO')
    iMes   = colidx(r'^MES$')
    iDia   = colidx(r'^DIA$|^D[ÍI]A$')
    iProv  = colidx(r'PROVEEDOR')
    iResp  = colidx(r'RESPONSABLE')
    iCosto = colidx(r'COSTO')
    iTipo  = colidx(r'TIPO')
    iObs   = colidx(r'OBSERVACI')
    iEst   = colidx(r'ESTADO')

    # Para historial
    iFecha = colidx(r'FECHA')
    iRes   = colidx(r'RESULTADO')

    rows_out = []
    for r in range(hdr_row_idx+1, ws.max_row+1):
        row_data = {
            'copr':  cel(r, iCopr)  if iCopr  else None,
            'elem':  cel(r, iElem)  if iElem  else None,
            'mes':   cel(r, iMes)   if iMes   else None,
            'dia':   cel(r, iDia)   if iDia   else None,
            'prov':  cel(r, iProv)  if iProv  else None,
            'resp':  cel(r, iResp)  if iResp  else None,
            'costo': cel(r, iCosto) if iCosto else None,
            'tipo':  cel(r, iTipo)  if iTipo  else None,
            'obs':   cel(r, iObs)   if iObs   else None,
            'estado':cel(r, iEst)   if iEst   else None,
            'fecha': cel(r, iFecha) if iFecha else None,
            'result':cel(r, iRes)   if iRes   else None,
        }
        rows_out.append(row_data)
    return rows_out

wb = openpyxl.load_workbook('/mnt/user-data/outputs/Mantenimientos_Escalar.xlsx')
print('Hojas:', wb.sheetnames)

# PROGRAMADOS
ws_prog = wb['PROGRAMADOS']
filas_prog = parse_hoja(ws_prog, hdr_row_idx=3)

MESES_FULL = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

def mes_idx(v):
    if not v: return -1
    n = str(v).strip().upper()
    if n in MESES_FULL: return MESES_FULL.index(n)
    for i,m in enumerate(MESES_FULL):
        if m.startswith(n[:3]): return i
    return -1

from datetime import date, timedelta
hoy = date.today()
anio = hoy.year

programados = []
for f in filas_prog:
    copr = str(f['copr'] or '').strip()
    elem = str(f['elem'] or '').strip()
    mes  = str(f['mes']  or '').strip()
    dia  = f['dia']
    if not mes and not copr and not elem: continue

    mi = mes_idx(mes)
    dia_n = int(dia) if dia and str(dia).strip().isdigit() else None
    prox = ''
    if mi >= 0 and dia_n:
        try:
            d = date(anio, mi+1, dia_n)
            if d < hoy: d = date(anio+1, mi+1, dia_n)
            prox = d.isoformat()
        except: pass

    tipo_v  = str(f['tipo']   or '').strip() or 'Preventivo'
    estado_v= str(f['estado'] or '').strip() or 'Pendiente'
    costo_v = None
    try: costo_v = float(str(f['costo']).replace(',','').replace('$','')) if f['costo'] else None
    except: pass

    programados.append({
        '_copropiedad': copr,
        '_elemento':    elem,
        'Copropiedad':  copr,
        'Elemento':     elem,
        'Mes':          mes.upper() if mes.upper()=='MENSUAL' else (MESES_FULL[mi] if mi>=0 else mes),
        'Dia':          str(dia_n) if dia_n else '',
        'Tipo':         tipo_v,
        'Proveedor':    str(f['prov']  or '').strip(),
        'Responsable':  str(f['resp']  or '').strip(),
        'Costo_estimado': costo_v,
        'Estado':       estado_v,
        'Observaciones': str(f['obs'] or '').strip(),
        'Proximo_vencimiento': prox,
    })

# HISTORIAL
ws_hist = wb['HISTORIAL']
filas_hist = parse_hoja(ws_hist, hdr_row_idx=2)
historial = []
for f in filas_hist:
    copr = str(f['copr'] or '').strip()
    elem = str(f['elem'] or '').strip()
    fecha_raw = f['fecha']
    if not fecha_raw and not copr and not elem: continue
    fecha_str = ''
    if fecha_raw:
        if isinstance(fecha_raw, date): fecha_str = fecha_raw.isoformat()
        elif isinstance(fecha_raw, str) and '/' in fecha_raw:
            parts = fecha_raw.split('/')
            if len(parts)==3:
                try: fecha_str = date(int(parts[2]),int(parts[1]),int(parts[0])).isoformat()
                except: fecha_str = fecha_raw

    costo_v = None
    try: costo_v = float(str(f['costo']).replace(',','').replace('$','')) if f['costo'] else None
    except: pass

    historial.append({
        '_copropiedad': copr,
        '_elemento':    elem,
        'Copropiedad':  copr,
        'Elemento':     elem,
        'Tipo':         str(f['tipo']   or '').strip(),
        'Fecha_ejecucion': fecha_str,
        'Proveedor':    str(f['prov']   or '').strip(),
        'Costo_real':   costo_v,
        'Responsable':  str(f['resp']   or '').strip(),
        'Resultado':    str(f['result'] or '').strip(),
        'Observaciones':str(f['obs']    or '').strip(),
    })

out = {
    'actualizado': date.today().isoformat(),
    'totalProgramados': len(programados),
    'totalHistorial': len(historial),
    'programados': programados,
    'historial': historial,
}

with open('/home/claude/Dashboard-Escalar/mantenimientos.json','w') as fp:
    json.dump(out, fp, ensure_ascii=False, indent=2)

print(f'Programados: {len(programados)} | Historial: {len(historial)}')
for p in programados[:8]:
    print(f"  {p['_copropiedad']:12s} | {p['_elemento']:22s} | {p['Mes']:12s} | dia={p['Dia']:3s} | {p['Estado']:12s} | {p['Proximo_vencimiento']}")
