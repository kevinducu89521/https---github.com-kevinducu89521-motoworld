import os
import sys
# pyrefly: ignore [missing-import]
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="CCCCCC", sz="4", val="single"):
    tblPr = table._element.xpath('w:tblPr')
    if tblPr:
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr[0].append(borders)

def add_callout(doc, text, title="NOTA IMPORTANTE", border_color="D72638", bg_color="FBF0F1"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"📌 {title}: ")
    run_t.bold = True
    run_t.font.name = "Arial"
    run_t.font.size = Pt(10)
    run_t.font.color.rgb = RGBColor(215, 38, 56)
    
    run_b = p.add_run(text)
    run_b.font.name = "Arial"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = RGBColor(50, 50, 50)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_code_snippet(doc, code_str, caption="Código fuente"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "14161F")
    set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
    
    tcPr = cell._element.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="18" w:space="0" w:color="D72638"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    
    r_cap = p.add_run(f"/* {caption} */\n")
    r_cap.font.name = "Consolas"
    r_cap.font.size = Pt(8.5)
    r_cap.font.color.rgb = RGBColor(249, 168, 38)
    r_cap.bold = True
    
    r_code = p.add_run(code_str)
    r_code.font.name = "Consolas"
    r_code.font.size = Pt(8.5)
    r_code.font.color.rgb = RGBColor(220, 225, 235)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(4)

def build_document(output_path):
    doc = docx.Document()
    
    # Page setup (margins: 1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Header / Footer setup
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_ft = p_ft.add_run("MotoWorld | Documento Técnico y Explicativo de Desarrollo Web")
        r_ft.font.name = "Arial"
        r_ft.font.size = Pt(8)
        r_ft.font.color.rgb = RGBColor(140, 140, 140)
    
    # -------------------------------------------------------------
    # PORTADA / HEADER
    # -------------------------------------------------------------
    p_tag = doc.add_paragraph()
    p_tag.paragraph_format.space_before = Pt(10)
    p_tag.paragraph_format.space_after = Pt(4)
    r_tag = p_tag.add_run("PROYECTO DE DESARROLLO WEB · DOCUMENTACIÓN TÉCNICA")
    r_tag.font.name = "Arial"
    r_tag.font.size = Pt(10)
    r_tag.bold = True
    r_tag.font.color.rgb = RGBColor(215, 38, 56)
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(6)
    r_title = p_title.add_run("MOTOWORLD: REVISTA DIGITAL DE MOTOCICLISMO")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(24)
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(17, 24, 39)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(18)
    r_sub = p_sub.add_run("Guía Detallada de Arquitectura, Estructura Semántica HTML5, Sistema de Diseño CSS3, Componentes e Integraciones")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(100, 110, 125)
    
    # Metadata Box Table
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    col_widths = [Inches(2.0), Inches(4.5)]
    data = [
        ("Proyecto / Aplicación:", "MotoWorld - Revista Digital Interactiva de Motociclismo"),
        ("Tecnologías Utilizadas:", "HTML5 Semántico, CSS3 Moderno (Flexbox, Grid, Custom Properties, Glassmorphism)"),
        ("Estructura de Secciones:", "1. Inicio (Hero) | 2. Historia (Línea de Tiempo) | 3. Motos Legendarias (Cards) | 4. Contacto (Footer y Redes Sociales)"),
        ("Documento Elaborado Para:", "Sustentación de Proyecto / Manual de Arquitectura Front-End")
    ]
    
    for i, (k, v) in enumerate(data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        
        set_cell_background(c0, "F4F5F7")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.name = "Arial"
        r0.font.size = Pt(9.5)
        r0.bold = True
        r0.font.color.rgb = RGBColor(30, 41, 59)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.name = "Arial"
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(71, 85, 105)
        
    set_table_borders(meta_table, color="D1D5DB", sz="4")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    
    def add_section_h1(title, num):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        
        r_num = p.add_run(f"{num}. ")
        r_num.font.name = "Arial"
        r_num.font.size = Pt(15)
        r_num.bold = True
        r_num.font.color.rgb = RGBColor(215, 38, 56)
        
        r_t = p.add_run(title)
        r_t.font.name = "Arial"
        r_t.font.size = Pt(15)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(17, 24, 39)
        
    def add_section_h2(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        
        r = p.add_run(title)
        r.font.name = "Arial"
        r.font.size = Pt(12)
        r.bold = True
        r.font.color.rgb = RGBColor(30, 41, 59)

    def add_body_p(text, bold_prefix=None, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Arial"
            r_pre.font.size = Pt(10)
            r_pre.bold = True
            r_pre.font.color.rgb = RGBColor(17, 24, 39)
            
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(55, 65, 81)
        return p

    def add_bullet(title, desc):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        
        r_t = p.add_run(f"{title}: ")
        r_t.font.name = "Arial"
        r_t.font.size = Pt(9.5)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(20, 25, 35)
        
        r_d = p.add_run(desc)
        r_d.font.name = "Arial"
        r_d.font.size = Pt(9.5)
        r_d.font.color.rgb = RGBColor(70, 80, 95)

    # -------------------------------------------------------------
    # 1. RESUMEN EJECUTIVO Y OBJETIVO DEL PROYECTO
    # -------------------------------------------------------------
    add_section_h1("Resumen Ejecutivo y Objetivo del Proyecto", "1")
    add_body_p(
        "El proyecto MotoWorld es un sitio web de categoría revista digital interactiva, concebido para brindar a los entusiastas del motociclismo una experiencia visual inmersiva, moderna y de alto rendimiento. La web reúne la historia evolutiva de las dos ruedas, una exhibición técnica de motocicletas legendarias y canales de contacto directo a través de redes sociales."
    )
    add_body_p(
        "El desarrollo fue planificado bajo estándares de calidad profesional, aplicando principios de diseño editorial moderno, modo oscuro (Dark Theme), tipografías de gran carácter y una arquitectura de código limpio, modular y 100% responsiva (adaptada a dispositivos móviles, tablets y computadoras de escritorio)."
    )
    
    add_callout(
        doc,
        "El proyecto no utiliza librerías externas pesadas ni frameworks invasivos (como Bootstrap o Tailwind); todo el sistema visual fue construido desde cero utilizando HTML5 y CSS3 nativo, asegurando un control absoluto del rendimiento, accesibilidad y fidelidad gráfica.",
        title="FILOSOFÍA DE DESARROLLO"
    )

    # -------------------------------------------------------------
    # 2. ARQUITECTURA DE ARCHIVOS Y CARPETAS
    # -------------------------------------------------------------
    add_section_h1("Arquitectura de Archivos y Organización del Proyecto", "2")
    add_body_p(
        "Para garantizar la mantenibilidad, escalabilidad y orden en el proyecto, la estructura de carpetas sigue las mejores prácticas de desarrollo front-end estándar:"
    )

    dir_table = doc.add_table(rows=6, cols=3)
    dir_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    dir_table.autofit = False
    
    col_w = [Inches(1.8), Inches(1.2), Inches(3.5)]
    headers = ["Directorio / Archivo", "Tipo", "Descripción y Propósito"]
    
    # Header Row
    for j, h in enumerate(headers):
        cell = dir_table.rows[0].cells[j]
        cell.width = col_w[j]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    rows_data = [
        ("index.html", "Archivo HTML5", "Documento raíz que contiene toda la estructura semántica, marcado de contenido y enlaces de navegación."),
        ("css/style.css", "Hoja de Estilos", "Archivo central con más de 700 líneas de CSS3: variables globales, resets, componentes y media queries."),
        ("img/social/", "Carpeta de Iconos", "Aloja las imágenes oficiales de redes sociales (whatsapp.png, instagram.png, facebook.png) utilizadas en el footer."),
        ("img/hero/ & img/legendary/", "Recursos Multimedia", "Imágenes temáticas locales de portada y modelos históricos (ej. vespa.jpg, hero.png)."),
        ("js/script..js", "Capa JavaScript", "Módulo de scripts reservado para interactividad futura y animaciones dinámicas.")
    ]
    
    for i, row_data in enumerate(rows_data, start=1):
        row = dir_table.rows[i]
        bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
        for j, val in enumerate(row_data):
            cell = row.cells[j]
            cell.width = col_w[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(9)
            if j == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(215, 38, 56)
            elif j == 1:
                r.font.color.rgb = RGBColor(100, 116, 139)
            else:
                r.font.color.rgb = RGBColor(51, 65, 85)
                
    set_table_borders(dir_table, color="CBD5E1", sz="4")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 3. ESTRUCTURA SEMÁNTICA HTML5 Y SECCIONES
    # -------------------------------------------------------------
    add_section_h1("Estructura Semántica y Desglose de Secciones (HTML5)", "3")
    add_body_p(
        "El documento HTML está estructurado rigurosamente mediante etiquetas semánticas estándar de HTML5 (<header>, <nav>, <main>, <section>, <article>, <footer>), garantizando una óptima indexación para motores de búsqueda (SEO) y accesibilidad para lectores de pantalla."
    )

    add_section_h2("3.1 Cabecera Técnica (<head>) y Metadatos")
    add_bullet("Codificación de Caracteres", "Uso de <meta charset=\"UTF-8\"> para soporte completo de caracteres especiales y acentos en español.")
    add_bullet("Control de Viewport", "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"> que asegura que la interfaz se adapte perfectamente a cualquier resolución de pantalla móvil o de escritorio.")
    add_bullet("Metaetiquetas SEO", "<meta name=\"description\"> descriptiva y <title> optimizado para posicionamiento web.")
    add_bullet("Google Fonts", "Conexión pre-conectada (<link rel=\"preconnect\">) a Google Fonts para cargar las familias tipográficas Oswald y Montserrat.")

    add_section_h2("3.2 Barra de Navegación Fija (<header> y <nav>)")
    add_body_p(
        "La cabecera permanece fija en la parte superior de la ventana del navegador con un elegante fondo semitransparente con desenfoque de cristal (Glassmorphism). Contiene:"
    )
    add_bullet("Logotipo de la Marca", "Texto estilizado con tipografía Oswald donde la palabra 'WORLD' destaca en color rojo racing mediante la clase .logo-highlight.")
    add_bullet("Menú de Navegación", "Lista no ordenada (<ul>) con 4 enlaces ancla: Inicio (#hero), Historia (#historia), Motos Legendarias (#legendarias) y Contacto (#contacto).")
    add_bullet("Microinteracción Hover", "Los enlaces cuentan con una línea indicadora animada que se expande horizontalmente de 0% a 100% al pasar el cursor (hover) mediante pseudo-elementos ::after.")

    add_section_h2("3.3 Sección Hero / Portada Principal (<section id=\"hero\">)")
    add_body_p(
        "Es la primera sección de impacto visual que recibe al usuario al entrar al sitio. Incluye:"
    )
    add_bullet("Título Principal (h1)", "'MÁQUINAS QUE HICIERON HISTORIA', con acento tipográfico rojo y sombra luminosa (text-glow).")
    add_bullet("Texto de Bienvenida", "Breve párrafo explicativo que sintetiza el espíritu y propósito de la revista digital.")
    add_bullet("Botón de Acción (Call to Action)", "Botón .btn .btn-primary 'Ver Motos Legendarias' con icono de flecha animada y salto directo a la sección de catálogo.")
    add_bullet("Fondo Cinematográfico", "Capa compuesta por un gradiente oscuro radial superpuesto a una imagen de alta resolución con efecto de fijación visual (background-attachment: fixed).")

    add_section_h2("3.4 Sección Historia / Línea de Tiempo (<section id=\"historia\">)")
    add_body_p(
        "Presenta un recorrido cronológico por los grandes hitos de la industria de las dos ruedas. La interfaz simula una línea de tiempo (Timeline) con:"
    )
    add_bullet("Línea Central Iluminada", "Línea vertical continua con degradado de rojo a ámbar (var(--primary) a var(--secondary)).")
    add_bullet("Nodos Luminosos (.timeline-dot)", "Puntos circulares con halo brillante que marcan cada época.")
    add_bullet("Distribución Alternada (Zig-Zag)", "En pantallas de escritorio, los elementos impares se alinean a la derecha y los pares a la izquierda utilizando selectores avanzados :nth-child(odd) y :nth-child(even).")
    
    # Hitos table
    hist_table = doc.add_table(rows=7, cols=3)
    hist_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hist_table.autofit = False
    
    h_col_w = [Inches(1.3), Inches(2.2), Inches(3.0)]
    h_headers = ["Año / Época", "Hito / Modelo", "Relevancia Histórica"]
    
    for j, h in enumerate(h_headers):
        cell = hist_table.rows[0].cells[j]
        cell.width = h_col_w[j]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    h_data = [
        ("1885", "Daimler Reitwagen", "Primer vehículo de dos ruedas con motor de combustión interna creado en Alemania."),
        ("1903 - 1920", "La Era Harley e Indian", "Nacimiento de los gigantes estadounidenses y los motores bicilíndricos en V (V-Twin)."),
        ("1946 - 1950", "Vespa & Cafe Racer", "Movilidad accesible de posguerra en Italia y cultura juvenil británica de carreras."),
        ("1969", "Honda CB750 Four", "La primera superbike moderna con 4 cilindros y freno de disco accesible."),
        ("1990 - 2000", "Ducati 916 y Yamaha R1", "La era dorada de la velocidad, diseño aerodinámico y máxima relación peso-potencia."),
        ("Siglo XXI", "Electrónica y Cero Emisiones", "Llegada de la sobrealimentación (Ninja H2R), ABS en curva y motores 100% eléctricos.")
    ]
    
    for i, (yr, hit, desc) in enumerate(h_data, start=1):
        row = hist_table.rows[i]
        bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
        for j, val in enumerate([yr, hit, desc]):
            cell = row.cells[j]
            cell.width = h_col_w[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=90, right=90)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if j == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(249, 168, 38)
            elif j == 1:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            else:
                r.font.color.rgb = RGBColor(71, 85, 105)
                
    set_table_borders(hist_table, color="CBD5E1", sz="4")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_section_h2("3.5 Sección Motos Legendarias (<section id=\"legendarias\">)")
    add_body_p(
        "Muestra una colección de 6 tarjetas (cards) interactivas dispuestas en una cuadrícula CSS Grid responsive. Cada tarjeta está construida con la etiqueta semántica <article class=\"card\"> y contiene:"
    )
    add_bullet("Contenedor de Imagen (.card-img-container)", "Imagen de alta fidelidad con efecto de zoom suave al pasar el cursor.")
    add_bullet("Badge Flotante (.card-badge)", "Etiqueta con color semántico según la categoría (badge-classic, badge-superbike, badge-custom, badge-racing).")
    add_bullet("Metadatos y Categoría", "Indica la categoría de la máquina y su año de lanzamiento.")
    add_bullet("Título y Descripción", "Nombre oficial del modelo y un texto conciso sobre su legado cultural e ingenieril.")
    add_bullet("Ficha de Especificaciones (.card-specs)", "Cuadrícula técnica con tres métricas fundamentales: Cilindrada del Motor, Potencia (CV/HP) y Velocidad Máxima (km/h).")
    add_bullet("Enlace de Acción (.card-link)", "Botón con flecha indicativa que redirige hacia la sección de contacto.")

    add_section_h2("3.6 Pie de Página y Redes Sociales (<footer id=\"contacto\">)")
    add_body_p(
        "El pie de página cierra la experiencia del usuario con una composición en dos columnas principales y una barra de copyright:"
    )
    add_bullet("Columna de Marca (.brand-col)", "Logotipo distintivo MotoWorld y contenedor de redes sociales (.social-links).")
    add_bullet("Integración de Iconos Oficiales (.social-btn)", "Se colocaron las imágenes oficiales de Instagram, Facebook y WhatsApp en formato PNG de alta resolución dentro de botones circulares con enlaces externos seguros (target=\"_blank\" rel=\"noopener noreferrer\").")
    add_bullet("Columna de Navegación Rápida", "Menú vertical con enlaces a todas las secciones de la página.")
    add_bullet("Barra de Cierre (.footer-bottom)", "Texto de copyright y autoría del proyecto educativo.")

    add_code_snippet(
        doc,
        '''<div class="social-links">
    <a href="https://instagram.com/" target="_blank" rel="noopener noreferrer" class="social-btn" title="Instagram">
        <img src="img/social/instagram.png" alt="Instagram">
    </a>
    <a href="https://facebook.com/" target="_blank" rel="noopener noreferrer" class="social-btn" title="Facebook">
        <img src="img/social/facebook.png" alt="Facebook">
    </a>
    <a href="https://wa.me/" target="_blank" rel="noopener noreferrer" class="social-btn" title="WhatsApp">
        <img src="img/social/whatsapp.png" alt="WhatsApp">
    </a>
</div>''',
        caption="Estructura HTML5 de los Iconos de Redes Sociales en el Footer"
    )

    # -------------------------------------------------------------
    # 4. SISTEMA DE DISEÑO Y CSS3
    # -------------------------------------------------------------
    add_section_h1("Sistema de Diseño, CSS3 Moderno y Estilos Visuales", "4")
    add_body_p(
        "Toda la identidad visual está centralizada mediante Variables CSS (Custom Properties) en la pseudoclase :root. Esto permite consistencia absoluta en colores, tipografías, bordes, sombras y tiempos de animación."
    )

    add_section_h2("4.1 Paleta de Colores y Tokens de Diseño (:root)")
    
    pal_table = doc.add_table(rows=7, cols=3)
    pal_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    pal_table.autofit = False
    
    p_col_w = [Inches(1.8), Inches(1.4), Inches(3.3)]
    p_headers = ["Variable CSS", "Valor Hex / RGBA", "Aplicación en la Interfaz"]
    
    for j, h in enumerate(p_headers):
        cell = pal_table.rows[0].cells[j]
        cell.width = p_col_w[j]
        set_cell_background(cell, "1E293B")
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    p_data = [
        ("--primary", "#D72638", "Rojo Racing. Color primario para botones, bordes activos, títulos y badges."),
        ("--secondary", "#F9A826", "Ámbar / Dorado. Utilizado para fechas de la historia, subtítulos y acentos."),
        ("--accent-blue", "#00B4D8", "Azul Neón. Empleado en badges de superbikes y detalles tecnológicos."),
        ("--bg-main", "#0B0C10", "Fondo base ultra-oscuro de la revista digital."),
        ("--bg-card / --bg-surface", "#1A1D28 / #14161F", "Superficies de tarjetas, cajas técnicas y secciones intermedias."),
        ("--bg-glass", "rgba(15, 16, 22, 0.85)", "Fondo con efecto de cristal translúcido para la barra de navegación superior.")
    ]
    
    for i, (var_name, val_code, desc) in enumerate(p_data, start=1):
        row = pal_table.rows[i]
        bg = "F8FAFC" if i % 2 == 1 else "FFFFFF"
        for j, val in enumerate([var_name, val_code, desc]):
            cell = row.cells[j]
            cell.width = p_col_w[j]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=90, right=90)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if j == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(215, 38, 56)
            elif j == 1:
                r.font.color.rgb = RGBColor(30, 41, 59)
            else:
                r.font.color.rgb = RGBColor(71, 85, 105)
                
    set_table_borders(pal_table, color="CBD5E1", sz="4")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_section_h2("4.2 Tipografías y Jerarquía Visual")
    add_bullet("Tipografía de Títulos (--font-heading: 'Oswald')", "Fuente geométrica y condensada que transmite fuerza, deportividad y velocidad; utilizada en títulos (h1, h2, h3), botones y logotipo.")
    add_bullet("Tipografía de Cuerpo (--font-body: 'Montserrat')", "Fuente sans-serif moderna, limpia y altamente legible para párrafos, fichas técnicas y enlaces.")

    add_section_h2("4.3 Técnicas de Maquetación: Flexbox y CSS Grid")
    add_bullet("CSS Grid con Ajuste Automático", "La galería de motos utiliza grid-template-columns: repeat(auto-fit, minmax(340px, 1fr)), lo que calcula automáticamente el número de columnas según el ancho de pantalla disponible sin desbordamientos.")
    add_bullet("Flexbox para Distribución Interna", "Utilizado en la barra de navegación (.navbar), botones (.btn), fichas técnicas (.spec-box), alineación de tarjetas y el pie de página (.footer-container).")
    add_bullet("Glassmorphism & Backdrop Filter", "Implementación de backdrop-filter: blur(14px) que produce un difuminado elegante sobre los elementos que pasan por detrás de la barra de navegación.")

    add_section_h2("4.4 Microinteracciones y Efectos Visuales (Hover States)")
    add_bullet("Elevación de Tarjetas (Card Hover Lift)", "Al posar el puntero, las tarjetas se elevan con transform: translateY(-8px) y proyectan un resplandor perimetral rojo (box-shadow con tinte --primary-glow).")
    add_bullet("Zoom Fotográfico Suave", "Las imágenes de las motos crecen sutilmente con transform: scale(1.08) bajo una curva de transición fluida cubic-bezier(0.25, 1, 0.5, 1).")
    add_bullet("Botones Sociales Circulares", "Los iconos de WhatsApp, Instagram y Facebook disponen de un marco circular con borde sutil (border: 1px solid rgba(255, 255, 255, 0.15)), efecto de elevación (translateY(-4px)), aumento de brillo y sombra con halo dinámico.")

    # -------------------------------------------------------------
    # 5. DISEÑO RESPONSIVO (MEDIA QUERIES)
    # -------------------------------------------------------------
    add_section_h1("Estrategia de Diseño Responsivo (Responsive Web Design)", "5")
    add_body_p(
        "Para garantizar una visualización impecable en cualquier dispositivo, se configuraron puntos de quiebre (Media Queries) adaptativos:"
    )
    
    add_bullet("Dispositivos Grandes y Laptops (<= 1024px)", "Reducción proporcional de títulos masivos (.hero-title a 3.8rem) para evitar cortes de línea indeseados.")
    add_bullet("Tablets y Dispositivos Medianos (<= 768px)", "Transformación estructural completa: el menú se alinea en columna vertical, la línea de tiempo pasa de alternada a una columna única orientada a la izquierda con padding adaptado, y el footer reorganiza sus columnas verticalmente.")
    add_bullet("Smartphones y Pantallas Pequeñas (<= 480px)", "Ajuste de títulos a 2.2rem, cuadrícula forzada a 1 columna (1fr) y botones en ancho completo para una pulsación táctil ergonómica.")

    # -------------------------------------------------------------
    # 6. ACCESIBILIDAD, SEO Y BUENAS PRÁCTICAS
    # -------------------------------------------------------------
    add_section_h1("Accesibilidad, Rendimiento y Buenas Prácticas Web", "6")
    add_bullet("Atributos ALT Descriptivos", "Todas las imágenes cuentan con textos alternativos contextuales que describen el modelo de la motocicleta.")
    add_bullet("Atributos ARIA y Títulos", "Los botones de redes sociales cuentan con title y aria-label para navegación asistida.")
    add_bullet("Carga Diferida (Lazy Loading)", "Las imágenes de tarjetas tienen loading=\"lazy\" para reducir el tiempo de carga inicial y ahorrar ancho de banda.")
    add_bullet("Seguridad en Enlaces Externos", "Todos los enlaces que abren nuevas pestañas poseen rel=\"noopener noreferrer\" para mitigar vulnerabilidades de phishing (Tabnabbing).")
    add_bullet("Contraste de Color", "El contraste entre el fondo oscuro y el texto blanco/ámbar supera los ratios mínimos establecidos por la WCAG (Web Content Accessibility Guidelines).")

    # -------------------------------------------------------------
    # 7. CONCLUSIONES Y PROPUESTAS DE AMPLIACIÓN
    # -------------------------------------------------------------
    add_section_h1("Conclusiones y Futuras Mejoras", "7")
    add_body_p(
        "El desarrollo de MotoWorld demuestra una aplicación integral de los fundamentos del desarrollo web moderno. El resultado es un producto digital pulido, con estética visual de nivel editorial, código limpio y estructurado, y una experiencia de usuario fluida."
    )
    add_body_p(
        "Como siguientes fases de desarrollo, se recomienda implementar:"
    )
    add_bullet("Interactividad con JavaScript", "Sistema de filtrado dinámico en la sección de motos (por ejemplo: filtrar por década, país o tipo de motor).")
    add_bullet("Animaciones al Hacer Scroll (AOS / ScrollReveal)", "Aparición progresiva de los hitos de la historia conforme el usuario se desplaza por la página.")
    add_bullet("Modal de Galería Fotográfica", "Ventana emergente interactiva para ver las motocicletas en 360 grados o galería detallada.")

    doc.save(output_path)
    print(f"Documento guardado con éxito en: {output_path}")

if __name__ == "__main__":
    out_file = r"c:\Users\kevin\Documents\KEVIN\MotoWorld\Documento_Explicativo_MotoWorld.docx"
    build_document(out_file)
