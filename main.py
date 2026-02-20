#!/usr/bin/env python3
"""
Pulguitas — Servidor local
Ejecutá: python app.py
Luego abrí: http://localhost:5173
"""

import os, re, json, sys, subprocess, tempfile, threading, webbrowser
from collections import defaultdict
from pathlib import Path

# ── Auto-instalar dependencias ────────────────────────────────────────────────
for pkg in ("pdfplumber", "openpyxl", "flask", "pymupdf"):
    try:
        __import__(pkg if pkg != "flask" else "flask")
    except ImportError:
        print(f"Instalando {pkg}...")
        subprocess.run([sys.executable, "-m", "pip", "install", pkg,
                        "--break-system-packages", "-q"], check=True)

import pdfplumber, openpyxl, fitz  # fitz = pymupdf
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, request, jsonify, send_file, send_from_directory

app = Flask(__name__, static_folder=".")

# ── Rutas de archivos ─────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
PRODUCTOS_TXT = BASE_DIR / "productos.txt"
OUTPUT_XLSX   = BASE_DIR / "resumen_pedidos.xlsx"

# ══════════════════════════════════════════════════════════════════════════════
# CATÁLOGO
# ══════════════════════════════════════════════════════════════════════════════
TALLE_LABELS = {"S":"Talle S","M":"Talle M","L":"Talle L","XL":"Talle XL",
                "XS":"Talle XS","SM":"Talle S/M","LXL":"Talle L/XL","U":"Talle Único"}
TALLE_SIZES  = {
    "VERANO":{"S":"50x50","M":"70x70","L":"90x90"},
    "ANTIESTRES":{"M":"70x70","L":"90x90"},
    "INVIERNO":{"S":"50x50","M":"70x70","L":"90x90"},
    "DECO":{"S":"S/M","M":"M/L","L":"L/XL"},
    "ESCALERA":{"M":"30x38","L":"40x38"},
    "NORDICA":{"M":"60x60","L":"80x80","XL":"90x90"},
    "ROPITA":{"XS":"XS","SM":"S/M","LXL":"L/XL"},
    "MANTA":{"U":"70x70"},
}
CAT_COLORS = {
    "VERANO":("29B6F6","E1F5FE"),"ANTIESTRES":("66BB6A","E8F5E9"),
    "INVIERNO":("FFA726","FFF3E0"),"DECO":("AB47BC","F3E5F5"),
    "ESCALERA":("78909C","ECEFF1"),"NORDICA":("26A69A","E0F2F1"),
    "ROPITA":("EC407A","FCE4EC"),"MANTA":("FFCA28","FFFDE7"),
}

def cargar_catalogo(texto):
    catalogo_dict = {}; order = []
    for linea in texto.split("\n"):
        linea = linea.strip()
        if not linea or linea.startswith("#"): continue
        partes = [p.strip() for p in linea.split("|")]
        if len(partes) < 3: continue
        cat = partes[0].upper(); modelo = partes[1]
        color  = partes[2] if len(partes) > 2 else ""
        talles = [t.strip() for t in partes[3].split(",")] if len(partes) > 3 else []
        if cat not in catalogo_dict:
            catalogo_dict[cat] = {"talle_cols": talles, "filas": []}; order.append(cat)
        else:
            for t in talles:
                if t not in catalogo_dict[cat]["talle_cols"]:
                    catalogo_dict[cat]["talle_cols"].append(t)
        catalogo_dict[cat]["filas"].append((modelo, color))
    catalogo = []
    for cat in order:
        info = catalogo_dict[cat]; tc = info["talle_cols"]
        sizes = TALLE_SIZES.get(cat, {})
        headers = ["Modelo","Color"] + [
            f"{TALLE_LABELS.get(t,t)} ({sizes[t]})" if sizes.get(t) else TALLE_LABELS.get(t,t)
            for t in tc
        ]
        while len(headers) < 5: headers.append("")
        catalogo.append({"cat":cat,"headers":headers[:5],"talle_cols":tc,"filas":info["filas"]})
    return catalogo

# ══════════════════════════════════════════════════════════════════════════════
# MAPA PDF → (cat, modelo, color, talle)
# ══════════════════════════════════════════════════════════════════════════════
MAPA_PRODUCTOS = {
    "gatito verano (talla s":("VERANO","Gatito","Beige","S"),
    "gatito verano (talla m":("VERANO","Gatito","Beige","M"),
    "gatito verano (talla l":("VERANO","Gatito","Beige","L"),
    "gatito verano":("VERANO","Gatito","Beige","S"),
    "cama pancho antiestrés - ergonomica (talla m":("ANTIESTRES","Pancho","","M"),
    "cama pancho antiestrés - ergonomica (talla l":("ANTIESTRES","Pancho","","L"),
    "cama pancho antiestres - ergonomica (talla m":("ANTIESTRES","Pancho","","M"),
    "cama pancho antiestres - ergonomica (talla l":("ANTIESTRES","Pancho","","L"),
    "cama pancho antiestrés - ergonomica":("ANTIESTRES","Pancho","","M"),
    "cama pancho antiestres - ergonomica":("ANTIESTRES","Pancho","","M"),
    "cama garra - antiestres, ergonomica (talla m":("ANTIESTRES","Garra","Gris","M"),
    "cama garra - antiestres, ergonomica (talla l":("ANTIESTRES","Garra","Gris","L"),
    "cama garra - antiestres":("ANTIESTRES","Garra","Gris","M"),
    "cama nordica lavable (talla m":("NORDICA","Nórdica","Gris","M"),
    "cama nordica lavable (talla l":("NORDICA","Nórdica","Gris","L"),
    "cama nordica lavable (talla xl":("NORDICA","Nórdica","Gris","XL"),
    "cama nordica lavable":("NORDICA","Nórdica","Gris","M"),
    "cama bahia - ortopedico (talla l 70x95 cm hasta 50 kilos, gris":("DECO","Bahía","Gris","L"),
    "cama bahia - ortopedico (talla l 70x95 cm hasta 50 kilos, rosa":("DECO","Bahía","Rosa","L"),
    "cama bahia - ortopedico (talla l 70x95 cm hasta 50 kilos, salmon":("DECO","Bahía","Salmón","L"),
    "cama bahia - ortopedico (talla l 70x95 cm hasta 50 kilos, mostaza":("DECO","Bahía","Mostaza","L"),
    "cama bahia - ortopedico (talla m 45x60 cm hasta 20 kilos, gris":("DECO","Bahía","Gris","M"),
    "cama bahia - ortopedico (talla m 45x60 cm hasta 20 kilos, mostaza":("DECO","Bahía","Mostaza","M"),
    "cama bahia - ortopedico (talla l 70x95 cm hasta 50":("DECO","Bahía","Gris","L"),
    "cama bahia - ortopedico (talla m 45x60 cm hasta 20":("DECO","Bahía","Mostaza","M"),
    "mini sofa - ortopedico (gris/oscuro, talla l":("DECO","Mini Sofá","Gris Oscuro","L"),
    "mini sofa - ortopedico (gris/oscuro, talla m":("DECO","Mini Sofá","Gris Oscuro","M"),
    "mini sofa - ortopedico (gris/claro, talla m":("DECO","Mini Sofá","Gris Claro","M"),
    "mini sofa - ortopedico (gris/claro, talla l":("DECO","Mini Sofá","Gris Claro","L"),
    "mini sofa - ortopedico (mostaza, talla m":("DECO","Mini Sofá","Mostaza","M"),
    "mini sofa - ortopedico (mostaza, talla l":("DECO","Mini Sofá","Mostaza","L"),
    "mini sofa - ortopedico (rosa, talla m":("DECO","Mini Sofá","Rosa","M"),
    "mini sofa - ortopedico (rosa, talla l":("DECO","Mini Sofá","Rosa","L"),
    "mini sofa - ortopedico (gris/oscuro":("DECO","Mini Sofá","Gris Oscuro","M"),
    "mini sofa - ortopedico (gris/claro":("DECO","Mini Sofá","Gris Claro","M"),
    "mini sofa - ortopedico (mostaza":("DECO","Mini Sofá","Mostaza","M"),
    "mini sofa - ortopedico (rosa":("DECO","Mini Sofá","Rosa","M"),
    "sofa cama - ortopedico (gris/oscuro, talla l":("DECO","Sofá Cama","Gris","L"),
    "sofa cama - ortopedico (gris/oscuro, talla m":("DECO","Sofá Cama","Gris","M"),
    "sofa cama - ortopedico (salmon, talla l":("DECO","Sofá Cama","Salmón","L"),
    "sofa cama - ortopedico (salmon, talla m":("DECO","Sofá Cama","Salmón","M"),
    "sofa cama - ortopedico (mostaza, talla l":("DECO","Sofá Cama","Mostaza","L"),
    "sofa cama - ortopedico (mostaza, talla m":("DECO","Sofá Cama","Mostaza","M"),
    "timoteo (talla s 40x60 cm hasta 6 kilos, mostaza":("DECO","Timoteo","Mostaza","S"),
    "timoteo (talla s 40x60 cm hasta 6 kilos, gris":("DECO","Timoteo","Gris","S"),
    "timoteo (talla s 40x60 cm hasta 6 kilos, rosa":("DECO","Timoteo","Rosa","S"),
    "timoteo (talla m 60x80 cm hasta 15 kilos, mostaza":("DECO","Timoteo","Mostaza","M"),
    "timoteo (talla m 60x80 cm hasta 15 kilos, gris":("DECO","Timoteo","Gris","M"),
    "timoteo (talla m 60x80 cm hasta 15 kilos, rosa":("DECO","Timoteo","Rosa","M"),
    "escaleras ortopedica (talla l":("ESCALERA","Escalera","Gris","L"),
    "escaleras ortopedica (talla m":("ESCALERA","Escalera","Gris","M"),
        # Productos de invierno
    "gatito invierno (talla s": ("INVIERNO", "Gatito", "Beige/Marrón", "S"),
    "gatito invierno (talla m": ("INVIERNO", "Gatito", "Beige/Marrón", "M"),
    "gatito invierno (talla l": ("INVIERNO", "Gatito", "Beige/Marrón", "L"),
    
    
    # Mantitas
    "mantitas doble faz (gris/blanco": ("MANTA", "Manta Doble Faz", "Gris/Blanco", "U"),
    "mantitas doble faz (beige/blanco": ("MANTA", "Manta Doble Faz", "Beige/Blanco", "U"),
    "mantitas doble faz (gris": ("MANTA", "Manta Doble Faz", "Gris/Blanco", "U"),
    "mantitas doble faz (beige": ("MANTA", "Manta Doble Faz", "Beige/Blanco", "U"),
    "remeras deportivas (argentina":("ROPITA","Ropita","Argentina","U"),
    "buzo panda":("ROPITA","Ropita","Panda","U"),
}

def resolver(nombre):
    """
    Resuelve un nombre de producto buscando en MAPA_PRODUCTOS
    Prioriza entradas que contienen color específico
    """
    key = nombre.lower().strip()
    
    # Primero buscar entradas que tengan color específico
    mejor_con_color = None
    mejor_len_color = 0
    
    for p, info in MAPA_PRODUCTOS.items():
        # Verificar si esta entrada tiene un color definido (no es vacío)
        if info[2] and info[2].strip():  # info[2] es el color
            if key.startswith(p) and len(p) > mejor_len_color:
                mejor_con_color = info
                mejor_len_color = len(p)
    
    if mejor_con_color:
        return mejor_con_color
    
    # Si no hay con color, buscar entradas genéricas
    mejor, n = None, 0
    for p, info in MAPA_PRODUCTOS.items():
        if key.startswith(p) and len(p) > n:
            mejor = info
            n = len(p)
    
    return mejor

# ══════════════════════════════════════════════════════════════════════════════
# PDF EXTRACTION (usando pymupdf)
# ══════════════════════════════════════════════════════════════════════════════
def extraer_ordenes_con_fitz(pdf_path):
    """
    Extrae órdenes y productos usando pymupdf
    """
    doc = fitz.open(pdf_path)
    texto_completo = ""
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        texto_completo += page.get_text() + "\n"
    
    doc.close()
    
    lineas = texto_completo.split('\n')
    ordenes = defaultdict(list)
    i = 0
    
    print("📄 Procesando páginas...")
    
    while i < len(lineas):
        linea = lineas[i].strip()
        
        # Detectar inicio de orden
        if linea.startswith('Orden #'):
            match = re.search(r'Orden #(\d+)', linea)
            if match:
                num_orden = match.group(1)
                print(f"\n🔍 Procesando Orden #{num_orden}")
                
                # Recolectar SOLO las líneas de esta orden hasta la próxima "Orden #"
                lineas_orden = []
                j = i + 1
                while j < len(lineas):
                    siguiente_linea = lineas[j].strip()
                    # Si encontramos otra orden, terminamos
                    if siguiente_linea.startswith('Orden #'):
                        break
                    lineas_orden.append(lineas[j])
                    j += 1
                
                # Procesar SOLO las líneas de esta orden
                k = 0
                productos_en_orden = 0
                
                while k < len(lineas_orden):
                    linea_actual = lineas_orden[k].strip()
                    
                    # Detectar inicio de producto
                    palabras_clave = ["cama", "sofa", "mini", "escalera", "manta", "gatito", 
                                    "nordica", "pancho", "garra", "timoteo", "mantitas", 
                                    "remeras", "buzo", "mantita", "huella", "corona", 
                                    "hamburguesa", "ballena", "cactus", "panda", "palta",
                                    "argentina", "boca", "river", "inter", "mami"]
                    
                    if any(linea_actual.lower().startswith(p) for p in palabras_clave):
                        productos_en_orden += 1
                        print(f"\n  Producto #{productos_en_orden}:")
                        print(f"    L{k+1}: {linea_actual[:60]}...")
                        
                        # Capturar nombre del producto
                        nombre_producto = linea_actual
                        
                        m = k + 1
                        lineas_capturadas = 1
                        
                        # Seguir capturando hasta encontrar la cantidad
                        while m < len(lineas_orden):
                            linea_m = lineas_orden[m].strip()
                            
                            # Si encontramos un número solo, es la cantidad
                            if linea_m.isdigit():
                                cantidad = int(linea_m)
                                print(f"    Cantidad encontrada: {cantidad}")
                                m += 1
                                break
                            
                            # Si encontramos "Subtotal", significa que no hay número antes
                            if "Subtotal" in linea_m:
                                cantidad = 1
                                print(f"    Cantidad: {cantidad} (por defecto)")
                                break
                            
                            # Si encontramos otra palabra clave, es un NUEVO producto
                            if any(linea_m.lower().startswith(p) for p in palabras_clave):
                                print(f"    → Siguiente producto detectado")
                                cantidad = 1
                                break
                            
                            # Si no, es parte del mismo producto
                            nombre_producto += " " + linea_m
                            lineas_capturadas += 1
                            print(f"    L{lineas_capturadas}: {linea_m[:60]}...")
                            m += 1
                        
                        # Si salimos sin cantidad, establecer 1
                        if 'cantidad' not in locals():
                            cantidad = 1
                        
                        # Resolver producto
                        print(f"    Nombre completo: {nombre_producto[:100]}...")
                        info = resolver(nombre_producto)
                        if info:
                            cat, modelo, color, talle = info
                            print(f"    → {modelo} {talle} {color} x{cantidad}")
                            ordenes[num_orden].append((info, cantidad))
                        else:
                            print(f"    ⚠ No resuelto: {nombre_producto[:80]}...")
                        
                        k = m
                    else:
                        k += 1
                
                print(f"  Total productos en orden #{num_orden}: {productos_en_orden}")
                i = j
            else:
                i += 1
        else:
            i += 1
    
    # Agrupar productos por orden
    ordenes_agrupadas = {}
    for num_orden, productos in ordenes.items():
        print(f"\n📊 Orden #{num_orden} - productos sin agrupar: {len(productos)}")
        grupos = defaultdict(int)
        for info, cant in productos:
            grupos[info] += cant
        ordenes_agrupadas[num_orden] = [(info, cant) for info, cant in grupos.items()]
        print(f"   Productos agrupados: {len(ordenes_agrupadas[num_orden])}")
        for info, cant in ordenes_agrupadas[num_orden]:
            print(f"     → {info[1]} {info[3]} {info[2]} x{cant}")
    
    return ordenes_agrupadas
# ══════════════════════════════════════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════════════════════════════════════
NC = 5
def fc(h): return PatternFill("solid",fgColor=h)
def bd(c='BBBBBB',s='thin'): x=Side(style=s,color=c); return Border(left=x,right=x,top=x,bottom=x)
C=Alignment(horizontal='center',vertical='center',wrap_text=True)
L=Alignment(horizontal='left',  vertical='center',wrap_text=True)

def build_excel(ordenes, catalogo, out_path):
    """
    Genera el Excel con catálogo y pedidos
    """
    print("\n📊 DEBUG - Productos recibidos para Excel:")
    print("-" * 50)
    
    det = defaultdict(int)
    
    for num_orden, productos in ordenes.items():
        print(f"\nOrden #{num_orden}:")
        for info, cant in productos:
            cat, modelo, color, talle = info
            print(f"  → {cat} | {modelo} | {color} | {talle} x{cant}")
            det[(cat, modelo, color, talle)] += cant
    
    # Crear workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catálogo + Pedidos"
    
    # Hoja principal: Catálogo + Pedidos
    r = 1
    for sec in catalogo:
        cat = sec["cat"]
        hc, bc = CAT_COLORS.get(cat, ("607D8B", "ECEFF1"))
        
        # Título de categoría
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        c = ws.cell(row=r, column=1, value=cat)
        c.font = Font(name='Arial', bold=True, size=12, color='FFFFFF')
        c.fill = fc(hc)
        c.alignment = C
        c.border = bd('888888', 'medium')
        for col in range(2, NC+1):
            ws.cell(row=r, column=col).border = bd('888888', 'medium')
        ws.row_dimensions[r].height = 22
        r += 1
        
        # Encabezados
        for i, h in enumerate(sec["headers"]):
            c = ws.cell(row=r, column=1+i, value=h)
            c.font = Font(name='Arial', bold=True, size=9, color='333333')
            c.fill = fc(bc)
            c.alignment = C
            c.border = bd()
        ws.row_dimensions[r].height = 30
        r += 1
        
        # Filas de productos
        for i, (modelo, color) in enumerate(sec["filas"]):
            bg = 'F9F9F9' if i % 2 else 'FFFFFF'
            cm = {}
            for ti, tk in enumerate(sec["talle_cols"]):
                v = det.get((cat, modelo, color, tk), 0)
                if v:
                    cm[2+ti] = v
            
            # Escribir modelo y color
            for j in range(NC):
                v = [modelo, color, '', '', ''][j]
                c = ws.cell(row=r, column=1+j, value=v)
                c.font = Font(name='Arial', size=9, bold=(j == 0))
                c.fill = fc(bg)
                c.alignment = L if j <= 1 else C
                c.border = bd()
            
            # Escribir cantidades
            for ci, val in cm.items():
                if val:
                    c = ws.cell(row=r, column=1+ci, value=val)
                    c.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
                    c.fill = fc(hc)
                    c.alignment = C
                    c.border = bd()
            
            ws.row_dimensions[r].height = 16
            r += 1
        
        ws.row_dimensions[r].height = 8
        r += 1
    
    # Ancho de columnas
    for col, w in zip('ABCDE', [22, 18, 17, 17, 17]):
        ws.column_dimensions[col].width = w
    
    # =========================================================
    # Hoja de Resumen Productos
    # =========================================================
    ws_resumen = wb.create_sheet("Resumen Productos")
    
    # Encabezados
    headers_resumen = ["Producto", "Modelo", "Color", "Talle", "Cantidad Total"]
    for i, header in enumerate(headers_resumen, 1):
        cell = ws_resumen.cell(row=1, column=i, value=header)
        cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        cell.fill = PatternFill("solid", fgColor="4A5568")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bd()
    
    # Ancho de columnas
    ws_resumen.column_dimensions['A'].width = 40  # Producto
    ws_resumen.column_dimensions['B'].width = 20  # Modelo
    ws_resumen.column_dimensions['C'].width = 20  # Color
    ws_resumen.column_dimensions['D'].width = 15  # Talle
    ws_resumen.column_dimensions['E'].width = 15  # Cantidad
    
    # Agrupar todos los productos de todas las órdenes
    resumen = defaultdict(int)
    productos_detalle = []  # Guardar detalle para mostrarlo en la tabla
    
    for num_orden, productos in ordenes.items():
        for info, cant in productos:
            cat, modelo, color, talle = info
            # Crear una clave única para el producto
            clave = (modelo, color, talle)
            resumen[clave] += cant
            # Guardar detalle (usamos el primer producto como referencia)
            if clave not in [p[0] for p in productos_detalle]:
                # Crear nombre del producto
                nombre_producto = f"{modelo} {color} {talle}".strip()
                productos_detalle.append((clave, modelo, color, talle, nombre_producto))
    
    # Ordenar por cantidad (de mayor a menor)
    productos_ordenados = sorted(
        [(clave, modelo, color, talle, nombre, resumen[clave]) 
         for (clave, modelo, color, talle, nombre) in productos_detalle],
        key=lambda x: -x[5]  # Ordenar por cantidad descendente
    )
    
    # Escribir datos
    for row, (clave, modelo, color, talle, nombre, cantidad) in enumerate(productos_ordenados, 2):
        # Columna A: Nombre del producto
        cell_a = ws_resumen.cell(row=row, column=1, value=nombre)
        cell_a.font = Font(name='Arial', size=10)
        cell_a.alignment = Alignment(horizontal='left', vertical='center')
        cell_a.border = bd()
        
        # Columna B: Modelo
        cell_b = ws_resumen.cell(row=row, column=2, value=modelo)
        cell_b.font = Font(name='Arial', size=10)
        cell_b.alignment = Alignment(horizontal='left', vertical='center')
        cell_b.border = bd()
        
        # Columna C: Color
        cell_c = ws_resumen.cell(row=row, column=3, value=color)
        cell_c.font = Font(name='Arial', size=10)
        cell_c.alignment = Alignment(horizontal='left', vertical='center')
        cell_c.border = bd()
        
        # Columna D: Talle
        cell_d = ws_resumen.cell(row=row, column=4, value=talle)
        cell_d.font = Font(name='Arial', size=10)
        cell_d.alignment = Alignment(horizontal='center', vertical='center')
        cell_d.border = bd()
        
        # Columna E: Cantidad (con color de fondo)
        cell_e = ws_resumen.cell(row=row, column=5, value=cantidad)
        cell_e.font = Font(name='Arial', bold=True, size=11)
        cell_e.fill = PatternFill("solid", fgColor="E9F0FA")
        cell_e.alignment = Alignment(horizontal='center', vertical='center')
        cell_e.border = bd()
        
        # Color de fondo alternado para mejor legibilidad
        if row % 2 == 0:
            for col in range(1, 6):
                ws_resumen.cell(row=row, column=col).fill = PatternFill("solid", fgColor="F9F9F9")
    
    # Fila de total
    total_row = len(productos_ordenados) + 2
    cell_total = ws_resumen.cell(row=total_row, column=4, value="TOTAL:")
    cell_total.font = Font(name='Arial', bold=True, size=11)
    cell_total.alignment = Alignment(horizontal='right', vertical='center')
    cell_total.border = bd()
    
    cell_total_num = ws_resumen.cell(row=total_row, column=5, value=f"=SUM(E2:E{total_row-1})")
    cell_total_num.font = Font(name='Arial', bold=True, size=11)
    cell_total_num.fill = PatternFill("solid", fgColor="E2E8F0")
    cell_total_num.alignment = Alignment(horizontal='center', vertical='center')
    cell_total_num.border = bd()
    
    # =========================================================
    # Guardar Excel
    # =========================================================
    wb.save(out_path)
    return {}
def agregar_hoja_resumen(wb, ordenes):
    """
    Agrega una hoja con el resumen de todos los productos extraídos
    """
    ws = wb.create_sheet("Resumen Productos")
    
    # Encabezados
    headers = ["Producto", "Modelo", "Color", "Talle", "Cantidad Total"]
    for i, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=header)
        cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        cell.fill = PatternFill("solid", fgColor="4A5568")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bd()
    
    # Ancho de columnas
    ws.column_dimensions['A'].width = 40  # Producto
    ws.column_dimensions['B'].width = 20  # Modelo
    ws.column_dimensions['C'].width = 20  # Color
    ws.column_dimensions['D'].width = 15  # Talle
    ws.column_dimensions['E'].width = 15  # Cantidad
    
    # Agrupar todos los productos de todas las órdenes
    resumen = defaultdict(int)
    productos_detalle = []  # Guardar detalle para mostrarlo en la tabla
    
    for num_orden, productos in ordenes.items():
        for info, cant in productos:
            cat, modelo, color, talle = info
            # Crear una clave única para el producto
            clave = (modelo, color, talle)
            resumen[clave] += cant
            # Guardar detalle (usamos el primer producto como referencia)
            if clave not in [p[0] for p in productos_detalle]:
                # Buscar un nombre de producto de ejemplo (opcional)
                nombre_producto = f"{modelo} {color} {talle}".strip()
                productos_detalle.append((clave, modelo, color, talle, nombre_producto))
    
    # Ordenar por cantidad (de mayor a menor)
    productos_ordenados = sorted(
        [(clave, modelo, color, talle, nombre, resumen[clave]) 
         for (clave, modelo, color, talle, nombre) in productos_detalle],
        key=lambda x: -x[5]  # Ordenar por cantidad descendente
    )
    
    # Escribir datos
    for row, (clave, modelo, color, talle, nombre, cantidad) in enumerate(productos_ordenados, 2):
        # Columna A: Nombre del producto
        cell_a = ws.cell(row=row, column=1, value=nombre)
        cell_a.font = Font(name='Arial', size=10)
        cell_a.alignment = Alignment(horizontal='left', vertical='center')
        cell_a.border = bd()
        
        # Columna B: Modelo
        cell_b = ws.cell(row=row, column=2, value=modelo)
        cell_b.font = Font(name='Arial', size=10)
        cell_b.alignment = Alignment(horizontal='left', vertical='center')
        cell_b.border = bd()
        
        # Columna C: Color
        cell_c = ws.cell(row=row, column=3, value=color)
        cell_c.font = Font(name='Arial', size=10)
        cell_c.alignment = Alignment(horizontal='left', vertical='center')
        cell_c.border = bd()
        
        # Columna D: Talle
        cell_d = ws.cell(row=row, column=4, value=talle)
        cell_d.font = Font(name='Arial', size=10)
        cell_d.alignment = Alignment(horizontal='center', vertical='center')
        cell_d.border = bd()
        
        # Columna E: Cantidad (con color de fondo)
        cell_e = ws.cell(row=row, column=5, value=cantidad)
        cell_e.font = Font(name='Arial', bold=True, size=11)
        cell_e.fill = PatternFill("solid", fgColor="E9F0FA")
        cell_e.alignment = Alignment(horizontal='center', vertical='center')
        cell_e.border = bd()
        
        # Color de fondo alternado para mejor legibilidad
        if row % 2 == 0:
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="F9F9F9")
    
    # Fila de total
    total_row = len(productos_ordenados) + 2
    cell_total = ws.cell(row=total_row, column=4, value="TOTAL:")
    cell_total.font = Font(name='Arial', bold=True, size=11)
    cell_total.alignment = Alignment(horizontal='right', vertical='center')
    cell_total.border = bd()
    
    cell_total_num = ws.cell(row=total_row, column=5, value=f"=SUM(E2:E{total_row-1})")
    cell_total_num.font = Font(name='Arial', bold=True, size=11)
    cell_total_num.fill = PatternFill("solid", fgColor="E2E8F0")
    cell_total_num.alignment = Alignment(horizontal='center', vertical='center')
    cell_total_num.border = bd()
    
    return ws
# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES PARA ANOTAR PDF DE ETIQUETAS
# ══════════════════════════════════════════════════════════════════════════════
def formatear_productos_orden(productos, resolver_func):
    """Agrupa productos iguales y devuelve líneas de texto en formato personalizado"""
    grupos = defaultdict(int)
    for info, cant in productos:
        cat, modelo, color, talle = info
        key = (cat, modelo, talle, color)
        grupos[key] += cant
    
    lineas = []
    for (cat, modelo, talle, color), cant in grupos.items():
        # Caso especial: MANTA
        if cat == "MANTA":
            color_simple = color.split('/')[0].lower()
            linea = f"Manta {color_simple} x{cant}"
        
        # Caso especial: GATITO INVIERNO
        elif cat == "INVIERNO" and modelo == "Gatito":
            linea = f"Gatito invierno {talle} x{cant}"
        
        # Formato normal para el resto
        else:
            if color and color.strip():
                linea = f"{modelo} {talle} {color} x{cant}"
            else:
                linea = f"{modelo} {talle} x{cant}"
        
        lineas.append(linea)
    
    return lineas
def anotar_pdf_con_productos(pdf_etiquetas_path, pdf_pedidos_path, output_path):
    """Añade texto con productos justo debajo del segundo número de seguimiento o del texto IMPORTANTE"""
    # Extraer órdenes del PDF de pedidos
    ordenes = extraer_ordenes_con_fitz(pdf_pedidos_path)
    
    # Abrir PDF de etiquetas
    doc = fitz.open(pdf_etiquetas_path)
    
    # 1 cm en puntos (1 cm = 28.35 puntos)
    UN_CM = 28.35
    
    for pagina in doc:
        text = pagina.get_text()
        # Buscar número de orden en la página
        match = re.search(r"#(\d+)", text)
        if match:
            orden = match.group(1)
            if orden in ordenes:
                productos = ordenes[orden]
                lineas = formatear_productos_orden(productos, resolver)
                
                # Buscar todas las palabras
                palabras = pagina.get_text("words")
                
                # Encontrar TODAS las ocurrencias de "seguimiento"
                seguimientos = []
                for w in palabras:
                    if "seguimiento" in w[4].lower():
                        seguimientos.append(w)
                
                # Ordenar por posición Y (de arriba a abajo)
                seguimientos.sort(key=lambda w: w[3])
                
                y_pos = None
                
                # CASO 1: HAY 2 O MÁS SEGUIMIENTOS - usar el segundo
                if len(seguimientos) >= 2:
                    segundo_seguimiento = seguimientos[1]  # el segundo en orden vertical
                    y_pos = segundo_seguimiento[3] + UN_CM
                    print(f"Orden #{orden}: usando SEGUNDO seguimiento en Y={segundo_seguimiento[3]}")
                
                # CASO 2: HAY 1 SOLO SEGUIMIENTO - usar IMPORTANTE
                elif len(seguimientos) == 1:
                    # Buscar "IMPORTANTE:"
                    importantes = []
                    for w in palabras:
                        if "importante" in w[4].lower():
                            importantes.append(w)
                    
                    if importantes:
                        # Tomar el último "IMPORTANTE" (el de más abajo)
                        importantes.sort(key=lambda w: w[3])
                        ultimo_importante = importantes[-1]
                        y_pos = ultimo_importante[3] + UN_CM + 50  # 50 puntos extra para quedar debajo del texto
                        print(f"Orden #{orden}: usando IMPORTANTE en Y={ultimo_importante[3]}")
                
                # CASO 3: NO HAY SEGUIMIENTOS - fallback a números largos
                if y_pos is None:
                    numeros = []
                    for w in palabras:
                        if re.search(r'\d{10,}', w[4]):  # números largos (10+ dígitos)
                            numeros.append(w)
                    
                    if numeros:
                        numeros.sort(key=lambda w: w[3])
                        ultimo_numero = numeros[-1]
                        y_pos = ultimo_numero[3] + UN_CM
                        print(f"Orden #{orden}: usando último número en Y={ultimo_numero[3]}")
                
                # CASO 4: FALLBACK FINAL
                if y_pos is None:
                    y_pos = pagina.rect.height - 80
                    print(f"Orden #{orden}: usando fallback")
                
                # Configurar texto
                tam_fuente = 20
                
                # Escribir cada línea
                for i, linea in enumerate(lineas):
                    x_pos = 20
                    punto = fitz.Point(x_pos, y_pos + (i * 18))
                    pagina.insert_text(punto, linea, fontsize=tam_fuente, 
                                      fontname="helv", color=(0,0,0))
    
    # Guardar sobreescribiendo el original
    doc.save(output_path)
    doc.close()
    return True

# ══════════════════════════════════════════════════════════════════════════════
# FLASK ROUTES
# ══════════════════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "pulguitas_ui.html")

@app.route("/productos", methods=["GET"])
def get_productos():
    txt = PRODUCTOS_TXT.read_text(encoding="utf-8") if PRODUCTOS_TXT.exists() else ""
    return jsonify({"content": txt})

@app.route("/productos", methods=["POST"])
def save_productos():
    data = request.get_json()
    PRODUCTOS_TXT.write_text(data["content"], encoding="utf-8")
    return jsonify({"ok": True})

@app.route("/analizar", methods=["POST"])
def analizar():
    if "pdf" not in request.files:
        return jsonify({"error": "No se recibió PDF"}), 400
    pdf_file = request.files["pdf"]
    prod_txt = request.form.get("productos", "")

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        catalogo = cargar_catalogo(prod_txt)
        ordenes = extraer_ordenes_con_fitz(tmp_path)
        sin_mapear = build_excel(ordenes, catalogo, str(OUTPUT_XLSX))
        
        rows = []
        total_unidades = 0
        for num_orden, productos in ordenes.items():
            for info, cant in productos:
                cat, modelo, color, talle = info
                total_unidades += cant
                rows.append({
                    "producto": f"Orden #{num_orden}",
                    "cantidad": cant,
                    "categoria": cat,
                    "modelo": modelo,
                    "color": color,
                    "talle": talle,
                })
        
        return jsonify({
            "ok": True,
            "total_tipos": len(ordenes),
            "total_unidades": total_unidades,
            "sin_mapear": 0,
            "sin_mapear_list": [],
            "rows": rows,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)

@app.route("/anotar", methods=["POST"])
def anotar():
    """Recibe PDF de pedidos y PDF de etiquetas, devuelve PDF anotado"""
    if "pedidos" not in request.files or "etiquetas" not in request.files:
        return jsonify({"error": "Se requieren dos archivos: pedidos PDF y etiquetas PDF"}), 400
    
    pedidos_file = request.files["pedidos"]
    etiquetas_file = request.files["etiquetas"]
    
    with tempfile.NamedTemporaryFile(suffix="_pedidos.pdf", delete=False) as tmp_pedidos:
        pedidos_file.save(tmp_pedidos.name)
        tmp_pedidos_path = tmp_pedidos.name
    
    with tempfile.NamedTemporaryFile(suffix="_etiquetas.pdf", delete=False) as tmp_etiquetas:
        etiquetas_file.save(tmp_etiquetas.name)
        tmp_etiquetas_path = tmp_etiquetas.name
    
    output_path = tmp_etiquetas_path.replace("_etiquetas", "_anotado")
    
    try:
        anotar_pdf_con_productos(tmp_etiquetas_path, tmp_pedidos_path, output_path)
        return send_file(output_path, as_attachment=True, 
                        download_name=etiquetas_file.filename,
                        mimetype="application/pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # Limpiar archivos temporales
        for p in [tmp_pedidos_path, tmp_etiquetas_path, output_path]:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except:
                pass

@app.route("/descargar")
def descargar():
    if not OUTPUT_XLSX.exists():
        return "No hay Excel generado aún", 404
    return send_file(str(OUTPUT_XLSX), as_attachment=True,
                     download_name="resumen_pedidos.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    port = 5173
    url  = f"http://localhost:{port}"
    print(f"\n🐾 Pulguitas App corriendo en {url}")
    print("   Ctrl+C para detener\n")
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(port=port, debug=False)
