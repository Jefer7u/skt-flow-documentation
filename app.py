import streamlit as st
import json
import pandas as pd
import io
import os
import re
from datetime import datetime
from typing import Dict, List, Any, Tuple, Optional, Set
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.worksheet.worksheet import Worksheet

st.set_page_config(page_title="Simetrik Docs Pro | PeYa", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

C: Dict[str, str] = {"red": "EA0050", "white": "FFFFFF", "grey": "F5F5F5", "dark": "1C1C1C", "border": "D8D8D8", "blue": "1565C0", "teal": "00695C", "amber": "E65100", "purple": "4A148C", "green": "1B5E20", "slate": "37474F", "rose": "880E4F"}
RT_LABEL: Dict[str, str] = {"native": "📥 Fuente", "source_union": "🔗 Unión de Fuentes", "source_group": "📊 Agrupación", "reconciliation": "⚖️ Conciliación Estándar", "advanced_reconciliation": "🔬 Conciliación Avanzada", "consolidation": "🗂️ Consolidación", "resource_join": "🔀 Join de Recursos", "cumulative_balance": "📈 Balance Acumulado"}
RT_COLOR: Dict[str, str] = {"native": C["blue"], "source_union": C["teal"], "source_group": C["amber"], "reconciliation": C["red"], "advanced_reconciliation": C["purple"], "consolidation": C["slate"], "resource_join": C["green"], "cumulative_balance": C["green"]}
RT_ORDER: Dict[str, int] = {"native": 1, "source_union": 2, "source_group": 3, "reconciliation": 4, "advanced_reconciliation": 5, "consolidation": 6, "resource_join": 7, "cumulative_balance": 8}

def mk_border() -> Border: return Border(left=Side(border_style="thin", color=C["border"]), right=Side(border_style="thin", color=C["border"]), top=Side(border_style="thin", color=C["border"]), bottom=Side(border_style="thin", color=C["border"]))

def sc(cell: Any, bg: Optional[str] = None, bold: bool = False, color: str = C["dark"], size: int = 10, ha: str = 'left', va: str = 'top', wrap: bool = True) -> None:
    cell.border, cell.alignment, cell.font = mk_border(), Alignment(horizontal=ha, vertical=va, wrap_text=wrap), Font(name='Arial', bold=bold, size=size, color=color)
    if bg: cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

def hdr(cell: Any, text: str, bg: str = C["dark"]) -> None:
    cell.value = text
    sc(cell, bg=bg, bold=True, color=C["white"], size=10, ha='center', va='center', wrap=False)

def section_title(ws: Worksheet, row: int, text: str, bg: str = C["red"], cols: int = 5) -> int:
    ws.merge_cells(f'A{row}:{chr(64+cols)}{row}')
    sc(ws.cell(row, 1, text), bg=bg, bold=True, color=C["white"], size=10, ha='left', va='center', wrap=False)
    ws.row_dimensions[row].height = 20
    return row + 1

def meta_row(ws: Worksheet, row: int, label: str, value: Any, cols: int = 5, bg_val: Optional[str] = None) -> int:
    bg_val = bg_val or C["grey"]
    sc(ws.cell(row, 1, label), bg=C["slate"], bold=True, color=C["white"], size=9, ha='left', va='center', wrap=False)
    ws.merge_cells(f'B{row}:{chr(64+cols)}{row}')
    sc(ws.cell(row, 2, str(value) if value is not None else "—"), bg=bg_val, size=9, va='center', wrap=True)
    ws.row_dimensions[row].height = 14
    return row + 1

def row_height(n_lines: int, base: int = 13) -> int: return max(14, n_lines * base)

def build_maps(data: Dict[str, Any]) -> Tuple[Dict[int, str], Dict[int, str], Dict[int, Any], Dict[int, str]]:
    res_map, col_map, seg_map, meta_map = {}, {}, {}, {}
    for r in data.get('resources', []):
        eid = r.get('export_id')
        if not eid: continue
        res_map[eid] = r.get('name', str(eid))
        for c in (r.get('columns') or []):
            if c.get('export_id'): col_map[c['export_id']] = c.get('label') or c.get('name') or str(c['export_id'])
        sg = r.get('source_group') or {}
        for c in sg.get('columns', []) + sg.get('values', []):
            if c.get('column_id') and c['column_id'] not in col_map: col_map[c['column_id']] = f"col_{c['column_id']}"
        adv = r.get('advanced_reconciliation') or {}
        for rg in adv.get('reconcilable_groups', []):
            for cs in rg.get('columns_selection', []):
                if cs.get('column_id') and cs['column_id'] not in col_map: col_map[cs['column_id']] = f"col_{cs['column_id']}"
            for m in (rg.get('segmentation_config') or {}).get('segmentation_metadata', []):
                if m.get('export_id'): meta_map[m['export_id']] = m.get('value', '?')
        for seg in (r.get('segments') or []):
            if seg.get('export_id'): seg_map[seg['export_id']] = {'name': seg.get('name', ''), 'resource': r.get('name', ''), 'rules': [rule for fset in (seg.get('segment_filter_sets') or []) for rule in (fset.get('segment_filter_rules') or [])]}
    return res_map, col_map, seg_map, meta_map

def parse_transformation_logic(col: Dict[str, Any], res_map: Dict[int, str], col_map: Dict[int, str]) -> str:
    lines = []
    v = col.get('v_lookup')
    if v:
        vs = v.get('v_lookup_set') or {}
        origin = res_map.get(vs.get('origin_source_id'), f"ID:{vs.get('origin_source_id')}")
        keys = " & ".join(f"A.{col_map.get(r.get('column_a_id'), '?')} = B.{col_map.get(r.get('column_b_id'), '?')}" for r in vs.get('rules', []))
        lines.append(f"🔍 BUSCAR V EN: {origin}")
        if keys: lines.append(f"🔑 CLAVE MATCH: {keys}")

    for t in (col.get('transformations') or []):
        t_type = str(t.get('type', '')).lower()
        if t.get('is_parent'):
            q = (t.get('query') or '').strip()
            if q and q.upper() != 'N/A': lines.append(f"⚙️ FÓRMULA: {q}")
                
        if t_type in ['duplicate', 'row_number'] or 'partition_by' in t:
            dup_label = "Booleano (Flag)" if t_type == 'duplicate' else "Numérico (Índice)"
            lines.append(f"👯 CONTROL DUPLICADOS [{dup_label}]")
            if t.get('partition_by'):
                part_names = [col_map.get(p.get('column_id') if isinstance(p, dict) else p, f"ID:{p}") for p in t['partition_by']]
                lines.append(f"   ├─ Partición: {', '.join(part_names)}")
            if t.get('order_by'):
                order_strs = [f"{col_map.get(o.get('column_id'), '?')} ({str(o.get('direction', 'ASC')).upper()})" for o in t['order_by']]
                lines.append(f"   └─ Orden: {', '.join(order_strs)}")
    return "\n".join(lines) if lines else "Dato directo / heredado"

def build_relations(resources: List[Dict], nodes: List[Dict], res_map: Dict) -> Dict[int, Dict[str, List[str]]]:
    all_ids = {r.get('export_id') for r in resources}
    rels = {r.get('export_id'): {"parents": [], "children": []} for r in resources}
    for n in nodes:
        t_id, s_val = n.get('target'), n.get('source')
        if not (t_id and s_val): continue
        for sid in (s_val if isinstance(s_val, list) else [s_val]):
            if t_id in rels: rels[t_id]["parents"].append(res_map.get(sid, str(sid)) + ("" if sid in all_ids else " ↗"))
            if sid in rels: rels[sid]["children"].append(res_map.get(t_id, str(t_id)) + ("" if t_id in all_ids else " ↗"))
    return rels

def generar_excel(data: Dict, selected_ids: Set[int]) -> io.BytesIO:
    res_map, col_map, seg_map, meta_map = build_maps(data)
    resources = [r for r in data.get('resources', []) if r.get('export_id') in selected_ids]
    resources = list({r['export_id']: r for r in resources}.values())
    resources.sort(key=lambda r: (RT_ORDER.get(r.get('resource_type', ''), 99), r.get('export_id', 0)))
    rels = build_relations(resources, data.get('nodes', []), res_map)
    map_hojas = {r['export_id']: re.sub(r'[\\/*?:\[\]]', '', str(r.get('name', ''))[:18] + "_" + str(r['export_id']))[:31] for r in resources}

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        wb = writer.book
        ws = wb.create_sheet("📚 Índice", 0)
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A5"

        ws.merge_cells('A1:H1')
        sc(ws.cell(1, 1, "SIMETRIK DOCUMENTATION PRO  ·  PeYa Finance"), bg=C["red"], bold=True, color=C["white"], size=13, ha='center', va='center')
        ws.row_dimensions[1].height = 32
        ws.merge_cells('A2:H2')
        sc(ws.cell(2, 1, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Recursos: {len(resources)}"), bg=C["dark"], color=C["white"], size=9, ha='center')
        
        for i, h in enumerate(["#", "ID", "NOMBRE", "TIPO", "PROVIENE DE", "ALIMENTA A", "ENCADENADA", "LINK"], 1): hdr(ws.cell(4, i, h), h, bg=C["dark"])

        for row_n, res in enumerate(resources, 5):
            eid, rt = res.get('export_id'), res.get('resource_type', '')
            chained = res.get('reconciliation', {}).get('is_chained', False) or res.get('advanced_reconciliation', {}).get('is_chained', False)
            bg = C["grey"] if row_n % 2 == 0 else C["white"]
            
            vals = [row_n - 4, eid, res.get('name', ''), RT_LABEL.get(rt, rt), ", ".join(rels[eid]["parents"]) or "—", ", ".join(rels[eid]["children"]) or "—", "Sí" if chained else "No"]
            for col_n, val in enumerate(vals, 1):
                c = ws.cell(row_n, col_n, val)
                sc(c, bg=bg, size=9, va='center', wrap=False)
                if col_n == 4: c.font = Font(name='Arial', bold=True, size=9, color=RT_COLOR.get(rt, C["dark"]))

            lnk = ws.cell(row_n, 8, "Ver →")
            lnk.hyperlink = f"#'{map_hojas[eid]}'!A1"
            lnk.font = Font(name='Arial', color="0D47A1", underline="single", size=9)
            lnk.border = mk_border()

        for col_n, w in enumerate([6,11,44,26,36,36,12,8], 1): ws.column_dimensions[chr(64+col_n)].width = w
        ws.auto_filter.ref = f"A4:H{len(resources)+4}"

        for res in resources:
            eid, rt = res.get('export_id'), res.get('resource_type', '')
            ws = wb.create_sheet(map_hojas[eid])
            ws.sheet_view.showGridLines = False
            
            ws.merge_cells('A1:E1')
            sc(ws.cell(1, 1, f"{RT_LABEL.get(rt, '')}  ·  {res.get('name', '')}"), bg=RT_COLOR.get(rt, C["dark"]), bold=True, color=C["white"], size=12, ha='left', va='center')
            ws.row_dimensions[1].height = 30
            
            row = meta_row(ws, 2, "ID Recurso", eid)
            row = meta_row(ws, row, "Tipo", RT_LABEL.get(rt, rt))
            row = meta_row(ws, row, "Proviene de", ", ".join(rels[eid]["parents"]) or "Origen")
            row = meta_row(ws, row, "Alimenta a", ", ".join(rels[eid]["children"]) or "Fin de flujo")
            ws.freeze_panes = f"A{row+1}"
            row += 1

            columns = sorted(res.get('columns') or [], key=lambda x: x.get('position', 0))
            if columns:
                row = section_title(ws, row, "📋  CONFIGURACIÓN DE COLUMNAS", bg=C["blue"])
                for col_n, h in enumerate(["LABEL / NOMBRE", "TIPO DATO", "TIPO COL.", "LÓGICA · FÓRMULA · BUSCAR V"], 1): hdr(ws.cell(row, col_n, h), h, bg=C["blue"])
                ws.merge_cells(f'D{row}:E{row}')
                ws.row_dimensions[row].height = 18
                ws.auto_filter.ref = f"A{row}:E{row+len(columns)}"
                row += 1
                
                for i, col in enumerate(columns):
                    bg = C["grey"] if i % 2 == 0 else "FFFFFF"
                    logic = parse_transformation_logic(col, res_map, col_map)
                    c1 = ws.cell(row, 1, col.get('label') or col.get('name', ''))
                    c2 = ws.cell(row, 2, col.get('data_format', ''))
                    c3 = ws.cell(row, 3, (col.get('column_type') or '').replace('_', ' ').upper())
                    ws.merge_cells(f'D{row}:E{row}')
                    c4 = ws.cell(row, 4, logic)
                    for c, al in [(c1,'left'),(c2,'center'),(c3,'center'),(c4,'left')]: sc(c, bg=bg, size=9, va='top', wrap=True, ha=al)
                    ws.row_dimensions[row].height = row_height(logic.count('\n') + 1)
                    row += 1

            for col_n, w in enumerate([22,30,28,40,22], 1): ws.column_dimensions[chr(64+col_n)].width = w

        if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])
    output.seek(0)
    return output

# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI 
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div style='background:linear-gradient(135deg,#EA0050 0%,#B0003A 100%); padding:28px 36px;border-radius:16px; box-shadow:0 6px 24px rgba(234,0,80,0.25);margin-bottom:24px'>
    <h1 style='color:white;margin:0;font-family:Arial,sans-serif; font-size:2.2rem;letter-spacing:-0.5px;font-weight:700'>
        📊 Simetrik Docs Pro
    </h1>
    <p style='color:rgba(255,255,255,0.88);margin:8px 0 0; font-size:1.05rem;font-family:Arial'>
        PeYa Finance Operations &amp; Control &nbsp;·&nbsp; Generador Automático de Documentación
    </p>
</div>""", unsafe_allow_html=True)

up = st.file_uploader("📂 Arrastrá o subí el JSON exportado de tu flujo de Simetrik", type=['json'])

if not up:
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.info("**1. En Simetrik**\n\nVe a tu Flujo de trabajo y haz clic en ⚙️ **Configuración**.")
    c2.info("**2. Exportar**\n\nBusca la opción **Exportar JSON** en el menú de la derecha.")
    c3.success("**3. Generar**\n\nSube el archivo aquí y obtén tu Excel corporativo al instante.")
    st.stop()

try:
    data = json.load(up)
    all_resources = data.get('resources', [])
    resources_unique = list({r.get('export_id'): r for r in all_resources if r.get('export_id')}.values())
    resources_unique.sort(key=lambda r: (RT_ORDER.get(r.get('resource_type', ''), 99), r.get('export_id', 0)))
    res_map, col_map, seg_map, meta_map = build_maps(data)
    rels_all = build_relations(resources_unique, data.get('nodes', []), res_map)
except Exception as e:
    st.error(f"Error al analizar el JSON: {e}")
    st.stop()

if 'sel' not in st.session_state:
    st.session_state.sel = {r.get('export_id'): True for r in resources_unique}

st.markdown("### 📈 Resumen del Flujo")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Recursos", len(resources_unique))
m2.metric("Fuentes Originales", len([r for r in resources_unique if r.get('resource_type') == 'native']))
m3.metric("Reglas de Conciliación", len([r for r in resources_unique if 'reconciliation' in r.get('resource_type', '')]))
m4.metric("Nodos de Conexión", len(data.get('nodes', [])))
st.markdown("---")

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/PedidosYa_logo.svg/512px-PedidosYa_logo.svg.png", width=150)
    st.markdown("### ⚙️ Configuración")
    
    all_types = sorted({r.get('resource_type', '') for r in resources_unique}, key=lambda x: RT_ORDER.get(x, 99))
    filtro_tipo = st.multiselect("Filtrar por tipo:", options=all_types, format_func=lambda x: RT_LABEL.get(x, x), default=all_types)
    resources_visible = [r for r in resources_unique if r.get('resource_type', '') in filtro_tipo]
    
    st.markdown("#### Selección Rápida")
    c_btn1, c_btn2 = st.columns(2)
    if c_btn1.button("✅ Todos", use_container_width=True):
        for r in resources_visible: st.session_state.sel[r.get('export_id')] = True
    if c_btn2.button("⬜ Ninguno", use_container_width=True):
        for r in resources_visible: st.session_state.sel[r.get('export_id')] = False

st.markdown("### 1️⃣  Seleccioná los recursos a documentar")
tipo_groups = {}
for r in resources_visible: tipo_groups.setdefault(r.get('resource_type', ''), []).append(r)

selected_ids = set()
for rt in sorted(tipo_groups.keys(), key=lambda x: RT_ORDER.get(x, 99)):
    grupo = tipo_groups[rt]
    with st.expander(f"{RT_LABEL.get(rt, rt)} ({len(grupo)} recursos)", expanded=False):
        for r in grupo:
            eid = r.get('export_id')
            ca, cb, cc, cd = st.columns([0.5, 4, 3, 3])
            checked = ca.checkbox("", value=st.session_state.sel.get(eid, True), key=f"chk_{eid}")
            st.session_state.sel[eid] = checked
            cb.markdown(f"**{r.get('name','')}** <span style='color:gray;font-size:0.8em'>`ID:{eid}`</span>", unsafe_allow_html=True)
            cc.caption(f"⬅️ {', '.join(rels_all[eid]['parents']) or '—'}")
            cd.caption(f"➡️ {', '.join(rels_all[eid]['children']) or '—'}")
            if checked: selected_ids.add(eid)

st.markdown("---")

n_sel = len(selected_ids)
if not selected_ids:
    st.warning("⚠️ Seleccioná al menos un recurso arriba.")
    st.stop()

if st.button(f"🚀  GENERAR EXCEL DE DOCUMENTACIÓN ({n_sel} recursos)", type="primary", use_container_width=True):
    with st.status("Preparando documento corporativo...", expanded=True) as status:
        try:
            st.write("Resolviendo reglas de Duplicados, Particiones y Búsquedas V...")
            excel_bytes = generar_excel(data, selected_ids)
            status.update(label=f"✅ ¡Completado! Documentación generada con éxito.", state="complete", expanded=False)
            
            st.download_button(
                label="📥  DESCARGAR REPORTE EXCEL",
                data=excel_bytes,
                file_name=f"DOC_Simetrik_{datetime.now().strftime('%Y-%m-%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            st.balloons()
        except Exception as e:
            status.update(label="❌ Error durante la generación", state="error")
            st.error(str(e))
