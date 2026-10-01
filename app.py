import streamlit as st
import pandas as pd
import os
import time
import json
import base64
import gspread
from google.oauth2.service_account import Credentials

# 1. CONFIGURACIÓN DE PÁGINA (WIDE MODE)
st.set_page_config(page_title="Copa President - Live Leaderboard", page_icon="🏆", layout="wide")

# 2. CSS ESTILO TV PROFESIONAL (MODO OSCURO COMPLETO Y TABLAS LIMPIAS)
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .tv-header { background: linear-gradient(135deg, #001122, #002244); padding: 15px; border-radius: 8px; text-align: center; border-bottom: 4px solid #c5a059; margin-bottom: 20px; }
    .tv-title { color: white; font-family: 'Arial Black', sans-serif; font-size: 1.8rem; margin: 0; text-transform: uppercase; letter-spacing: 1px; }
    .tv-subtitle { color: #c5a059; font-size: 1rem; margin: 5px 0 0 0; font-weight: bold; }

    /* CONTENEDOR TIPO TABLA TV */
    .tv-match-table { display: table; width: 100%; border-collapse: collapse; background: #ffffff; border-radius: 6px; overflow: hidden; box-shadow: 0 3px 8px rgba(0,0,0,0.12); margin-bottom: 8px; font-family: 'Arial', sans-serif; }
    
    .tv-row { display: table-row; }
    
    /* CELDA LEONES NORMAL (BLANCA) */
    .cell-leones-normal { display: table-cell; width: 42%; background-color: #ffffff; padding: 10px 14px; vertical-align: middle; color: #222; font-weight: bold; font-size: 0.95rem; border-bottom: 1px solid #e0e0e0; border-left: 4px solid #cc0000; }
    /* CELDA LEONES GANANDO (ROJA COMPLETA) */
    .cell-leones-win { display: table-cell; width: 42%; background-color: #cc0000; padding: 10px 14px; vertical-align: middle; color: #ffffff; font-weight: bold; font-size: 0.95rem; border-bottom: 1px solid #b00000; border-left: 4px solid #800000; }

    /* CELDA TOROS NORMAL (BLANCA) */
    .cell-toros-normal { display: table-cell; width: 42%; background-color: #ffffff; padding: 10px 14px; vertical-align: middle; color: #222; font-weight: bold; font-size: 0.95rem; text-align: right; border-bottom: 1px solid #e0e0e0; border-right: 4px solid #0044cc; }
    /* CELDA TOROS GANANDO (AZUL COMPLETA) */
    .cell-toros-win { display: table-cell; width: 42%; background-color: #0044cc; padding: 10px 14px; vertical-align: middle; color: #ffffff; font-weight: bold; font-size: 0.95rem; text-align: right; border-bottom: 1px solid #003399; border-right: 4px solid #002080; }

    /* CENTRO (SOLO EL HOYO ACTUAL) */
    .tv-cell-status { display: table-cell; width: 16%; background-color: #111111; color: #ffffff; text-align: center; vertical-align: middle; font-family: 'Arial Black', sans-serif; font-size: 1.2rem; padding: 10px 6px; border-bottom: 1px solid #000; letter-spacing: 1px; }

    /* ETIQUETAS DE VENTAJA ESTILO TV (FLECHA ARRIBA) */
    .badge-advantage-left { background: rgba(0,0,0,0.25); color: #ffeb3b; padding: 2px 6px; border-radius: 4px; font-size: 0.85rem; font-family: 'Arial Black', sans-serif; margin-right: 8px; float: left; }
    .badge-advantage-right { background: rgba(0,0,0,0.25); color: #ffeb3b; padding: 2px 6px; border-radius: 4px; font-size: 0.85rem; font-family: 'Arial Black', sans-serif; margin-left: 8px; float: right; }

    .player-sub { font-size: 0.75rem; color: #666; font-weight: normal; float: right; background: rgba(0,0,0,0.06); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }
    .player-sub-win { font-size: 0.75rem; color: #ffeb3b; font-weight: normal; float: right; background: rgba(0,0,0,0.2); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }

    .player-sub-l { font-size: 0.75rem; color: #666; font-weight: normal; float: left; background: rgba(0,0,0,0.06); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }
    .player-sub-l-win { font-size: 0.75rem; color: #ffeb3b; font-weight: normal; float: left; background: rgba(0,0,0,0.2); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }

    .tv-footer-bar { display: flex; margin-top: 25px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.25); font-family: 'Arial Black', sans-serif; }
    .tv-footer-left { width: 50%; background-color: #cc0000; color: white; padding: 15px; font-size: 1.5rem; display: flex; justify-content: space-between; align-items: center; }
    .tv-footer-right { width: 50%; background-color: #0044cc; color: white; padding: 15px; font-size: 1.5rem; display: flex; justify-content: space-between; align-items: center; flex-direction: row-reverse; }
    </style>
""", unsafe_allow_html=True)

# 3. LECTURA DE DATOS LOCALES (JUGADORES)
@st.cache_data
def cargar_datos():
    nombres_posibles = ['THE PLAYERS.xlsm', 'THE PLAYERS.xlsx', 'The Players.xlsm', 'the players.xlsm', 'THE_PLAYERS.xlsm']
    archivo_encontrado = None
    for nombre in nombres_posibles:
        if os.path.exists(nombre):
            archivo_encontrado = nombre
            break
    if not archivo_encontrado:
        for f in os.listdir('.'):
            if f.endswith(('.xlsm', '.xlsx', '.xls')) and 'live' not in f.lower() and 'temp' not in f.lower():
                archivo_encontrado = f
                break
    if not archivo_encontrado:
        return pd.DataFrame(), "No se encontró el archivo de jugadores."
    try:
        df = pd.read_excel(archivo_encontrado, sheet_name='Indices', header=5)
        df.columns = df.columns.str.strip()
        df_validos = df.dropna(subset=['Nombre', 'Apellidos']).copy()
        df_validos['Handicap_Juego'] = pd.to_numeric(df_validos['Handicap_Juego'], errors='coerce').fillna(0)
        return df_validos, None
    except Exception as e:
        return pd.DataFrame(), f"Error: {e}"

df_jugadores, error_msj = cargar_datos()

# CONEXIÓN A GOOGLE SHEETS MEDIANTE DECODIFICACIÓN BASE64
def conectar_google_sheets():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        b64_string = st.secrets["gcp_base64_json"]
        json_bytes = base64.b64decode(b64_string)
        creds_dict = json.loads(json_bytes.decode('utf-8'))
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open("Copa President Live").sheet1
        return sheet
    except Exception as e:
        st.error(f"⚠ Detalle técnico del error de Google: {e}")
        return None

# 4. CREACIÓN DE PESTAÑAS
tab_sorteo, tab_leaderboard = st.tabs(["🎲 SORTEO Y EMPAREJAMIENTO", "📱 LEADERBOARD EN VIVO"])

# ==========================================
# PESTAÑA 1: EMPAREJAMIENTO
# ==========================================
with tab_sorteo:
    with st.sidebar:
        st.header("⚙️ Configuración")
        tipo_jornada = st.selectbox(
            "Modalidad:",
            (
                "Viernes: Four Ball (Selección de Parejas)", 
                "Sábado: Match Play Individual (100% HCP)"
            )
        )

    # ----------------------------------------------------
    # VIERNES: FOUR BALL (SELECCIÓN DE PAREJAS)
    # ----------------------------------------------------
    if tipo_jornada == "Viernes: Four Ball (Selección de Parejas)":
        st.markdown("<h3 style='text-align: center;'>⛳ Viernes - Four Ball (Emparejamiento por Parejas)</h3>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #888;'>Selecciona las parejas de Leones y Toros para cada salida y envíalas a Google Sheets.</p><br>", unsafe_allow_html=True)

        if len(df_jugadores) >= 44:
            df_leones = df_jugadores.iloc[:22].copy()
            df_toros = df_jugadores.iloc[22:44].copy()
            
            dict_leones = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_leones.iterrows()}
            dict_toros = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_toros.iterrows()}
            nombres_leones = list(dict_leones.keys())
            nombres_toros = list(dict_toros.keys())

            hoyo_salidas = ["Hoyo 1 (Match 1)", "Hoyo 2A (Match 2)", "Hoyo 2B (Match 3)", "Hoyo 3 (Match 4)", 
                            "Hoyo 4 (Match 5)", "Hoyo 5 (Match 6)", "Hoyo 6 (Match 7)", "Hoyo 7 (Match 8)", 
                            "Hoyo 8A (Match 9)", "Hoyo 8B (Match 10)", "Hoyo 9 (Match 11)"]

            datos_exportar_viernes = []
            form_viernes = st.form(key="form_viernes")
            with form_viernes:
                for i, salida in enumerate(hoyo_salidas):
                    st.markdown(f"**{salida}**")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("🔴 **Pareja Leones**")
                        jl1 = st.selectbox(f"León 1 - {salida}", nombres_leones, index=(i*2) % len(nombres_leones), key=f"v_l1_{i}")
                        jl2 = st.selectbox(f"León 2 - {salida}", nombres_leones, index=(i*2 + 1) % len(nombres_leones), key=f"v_l2_{i}")
                    with c2:
                        st.markdown("🔵 **Pareja Toros**")
                        jt1 = st.selectbox(f"Toro 1 - {salida}", nombres_toros, index=(i*2) % len(nombres_toros), key=f"v_t1_{i}")
                        jt2 = st.selectbox(f"Toro 2 - {salida}", nombres_toros, index=(i*2 + 1) % len(nombres_toros), key=f"v_t2_{i}")
                    
                    j_l1 = dict_leones[jl1]
                    j_l2 = dict_leones[jl2]
                    j_t1 = dict_toros[jt1]
                    j_t2 = dict_toros[jt2]
                    
                    datos_exportar_viernes.append({
                        "Match": salida, "Tipo": "Four Ball",
                        "Participante_Leones": f"{j_l1['Nombre']} {j_l1['Apellidos']} & {j_l2['Nombre']} {j_l2['Apellidos']}", 
                        "HCP_L": f"L1:{int(j_l1['Handicap_Juego'])}, L2:{int(j_l2['Handicap_Juego'])}",
                        "Participante_Toros": f"{j_t1['Nombre']} {j_t1['Apellidos']} & {j_t2['Nombre']} {j_t2['Apellidos']}", 
                        "HCP_T": f"T1:{int(j_t1['Handicap_Juego'])}, T2:{int(j_t2['Handicap_Juego'])}",
                        "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""
                    })
                    st.markdown("---")
                
                guardar_viernes = st.form_submit_button("🚀 Enviar Partidos del Viernes a Google Sheets", type="primary", use_container_width=True)
                
            if guardar_viernes:
                df_export = pd.DataFrame(datos_exportar_viernes)
                sheet = conectar_google_sheets()
                if sheet:
                    sheet.clear()
                    sheet.update([df_export.columns.values.tolist()] + df_export.values.tolist())
                    st.success("✅ ¡Partidos y parejas del Viernes sincronizados en Google Sheets correctamente!")
        else:
            st.warning("⚠️ No se encontraron suficientes jugadores en el archivo local.")

    # ----------------------------------------------------
    # SÁBADO: MATCH PLAY INDIVIDUAL
    # ----------------------------------------------------
    else:
        st.markdown("<h3 style='text-align: center;'>🤝 Sábado - Match Play Individual (100% Hándicap)</h3>", unsafe_allow_html=True)
        if len(df_jugadores) >= 44:
            df_leones = df_jugadores.iloc[:22].copy()
            df_toros = df_jugadores.iloc[22:44].copy()
            dict_leones = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_leones.iterrows()}
            dict_toros = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_toros.iterrows()}
            nombres_leones, nombres_toros = list(dict_leones.keys()), list(dict_toros.keys())
            
            datos_exportar_sabado = []
            form_sab = st.form(key="form_sabado")
            with form_sab:
                for i in range(11):
                    st.markdown(f"**Match Individual {i+1}**")
                    c1, c2 = st.columns(2)
                    with c1:
                        jl1 = st.selectbox(f"León - Match {i+1}", nombres_leones, index=i % len(nombres_leones), key=f"sab_l_{i}")
                    with c2:
                        jt1 = st.selectbox(f"Toro - Match {i+1}", nombres_toros, index=i % len(nombres_toros), key=f"sab_t_{i}")
                    
                    j_l1 = dict_leones[jl1]
                    j_t1 = dict_toros[jt1]
                    
                    datos_exportar_sabado.append({
                        "Match": f"Match {i+1} (Sábado)", "Tipo": "Individual",
                        "Participante_Leones": f"{j_l1['Nombre']} {j_l1['Apellidos']}", "HCP_L": int(j_l1['Handicap_Juego']),
                        "Participante_Toros": f"{j_t1['Nombre']} {j_t1['Apellidos']}", "HCP_T": int(j_t1['Handicap_Juego']),
                        "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""
                    })
                    st.markdown("---")
                
                guardar_sab = st.form_submit_button("🚀 Enviar Partidos del Sábado a Google Sheets", type="primary", use_container_width=True)
                
            if guardar_sab:
                df_export = pd.DataFrame(datos_exportar_sabado)
                sheet = conectar_google_sheets()
                if sheet:
                    sheet.clear()
                    sheet.update([df_export.columns.values.tolist()] + df_export.values.tolist())
                    st.success("✅ ¡Partidos del Sábado sincronizados en Google Sheets!")
        else:
            st.warning("⚠️ No se encontraron suficientes jugadores.")

# ==========================================
# PESTAÑA 2: LEADERBOARD EN VIVO (ESTILO TV PROFESIONAL)
# ==========================================
with tab_leaderboard:
    st.markdown("<br>", unsafe_allow_html=True)
    
    c_ctrl1, c_ctrl2, c_ctrl3 = st.columns([2, 2, 2])
    with c_ctrl1:
        auto_update = st.toggle("🚀 Piloto Automático (8s)")
    with c_ctrl2:
        pts_previo_leones = st.number_input("Pts Previos - Leones", min_value=0.0, max_value=50.0, value=0.0, step=0.5)
    with c_ctrl3:
        pts_previo_toros = st.number_input("Pts Previos - Toros", min_value=0.0, max_value=50.0, value=0.0, step=0.5)
        
    if not auto_update:
        st.button("🔄 ACTUALIZAR MARCADOR", type="primary", use_container_width=True)
        
    sheet = conectar_google_sheets()
    
    if sheet:
        try:
            data = sheet.get_all_records()
            df_live = pd.DataFrame(data)
            
            if not df_live.empty:
                puntos_jornada_leones = 0.0
                puntos_jornada_toros = 0.0
                tv_tables_html = ""
                
                for idx, row in df_live.iterrows():
                    l6 = pd.to_numeric(row.get('Leones_H6'), errors='coerce')
                    t6 = pd.to_numeric(row.get('Toros_H6'), errors='coerce')
                    l12 = pd.to_numeric(row.get('Leones_H12'), errors='coerce')
                    t12 = pd.to_numeric(row.get('Toros_H12'), errors='coerce')
                    l18 = pd.to_numeric(row.get('Leones_H18'), errors='coerce')
                    t18 = pd.to_numeric(row.get('Toros_H18'), errors='coerce')
                    
                    hoyo_display = "1"
                    pt_l = 0.0
                    pt_t = 0.0
                    
                    leones_gana = False
                    toros_gana = False
                    badge_l_html = ""
                    badge_t_html = ""
                    
                    d6 = (l6 - t6) if pd.notna(l6) and pd.notna(t6) else None
                    d12 = (l12 - t12) if pd.notna(l12) and pd.notna(t12) else None
                    d18 = (l18 - t18) if pd.notna(l18) and pd.notna(t18) else None
                    
                    if d18 is not None:
                        hoyo_display = "F"
                        if d18 > 0: 
                            pt_l = 1.0; leones_gana = True; badge_l_html = f"<span class='badge-advantage-left'>{int(d18)} ▲</span>"
                        elif d18 < 0: 
                            pt_t = 1.0; toros_gana = True; badge_t_html = f"<span class='badge-advantage-right'>{int(abs(d18))} ▲</span>"
                        else: 
                            pt_l = 0.5; pt_t = 0.5
                    elif d12 is not None:
                        hoyo_display = "12"
                        if d12 > 6: 
                            pt_l = 1.0; leones_gana = True; badge_l_html = f"<span class='badge-advantage-left'>{int(d12)}&6</span>"
                        elif d12 < -6: 
                            pt_t = 1.0; toros_gana = True; badge_t_html = f"<span class='badge-advantage-right'>{int(abs(d12))}&6</span>"
                        elif d12 > 0: 
                            leones_gana = True; badge_l_html = f"<span class='badge-advantage-left'>{int(d12)} ▲</span>"
                        elif d12 < 0: 
                            toros_gana = True; badge_t_html = f"<span class='badge-advantage-right'>{int(abs(d12))} ▲</span>"
                    elif d6 is not None:
                        hoyo_display = "6"
                        if d6 > 0: 
                            leones_gana = True; badge_l_html = f"<span class='badge-advantage-left'>{int(d6)} ▲</span>"
                        elif d6 < 0: 
                            toros_gana = True; badge_t_html = f"<span class='badge-advantage-right'>{int(abs(d6))} ▲</span>"
                    else:
                        hoyo_display = "1"
                    
                    puntos_jornada_leones += pt_l
                    puntos_jornada_toros += pt_t
                    
                    match_name = row['Match']
                    jugador_leones = row['Participante_Leones']
                    hcp_l = row['HCP_L']
                    jugador_toros = row['Participante_Toros']
                    hcp_t = row['HCP_T']
                    
                    # Clases CSS dinámicas
                    if leones_gana:
                        class_l = "cell-leones-win"
                        class_t = "cell-toros-normal"
                        sub_l_class = "player-sub-l-win"
                        sub_t_class = "player-sub"
                    elif toros_gana:
                        class_l = "cell-leones-normal"
                        class_t = "cell-toros-win"
                        sub_l_class = "player-sub-l"
                        sub_t_class = "player-sub-win"
                    else:
                        class_l = "cell-leones-normal"
                        class_t = "cell-toros-normal"
                        sub_l_class = "player-sub-l"
                        sub_t_class = "player-sub"
                    
                    tv_tables_html += f"""
                    <div style="font-size: 0.75rem; color: #555; font-weight: bold; margin-bottom: 2px; text-transform: uppercase; letter-spacing: 0.5px;">{match_name}</div>
                    <div class='tv-match-table'>
                        <div class='tv-row'>
                            <div class='{class_l}'>
                                {badge_l_html} {jugador_leones}
                                <span class='{sub_l_class}'>HCP: {hcp_l}</span>
                            </div>
                            <div class='tv-cell-status'>
                                {hoyo_display}
                            </div>
                            <div class='{class_t}'>
                                {jugador_toros} {badge_t_html}
                                <span class='{sub_t_class}'>HCP: {hcp_t}</span>
                            </div>
                        </div>
                    </div>
                    """
                
                total_leones = pts_previo_leones + puntos_jornada_leones
                total_toros = pts_previo_toros + puntos_jornada_toros
                
                st.markdown(f"""
                <div class='tv-header'>
                    <div class='tv-title'>🏆 Presidents Cup</div>
                    <div class='tv-subtitle'>Leaderboard Oficial en Vivo</div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown(tv_tables_html, unsafe_allow_html=True)
                
                st.markdown(f"""
                <div class='tv-footer-bar'>
                    <div class='tv-footer-left'><span>LEONES</span><span>{f"{total_leones:.1f}".replace('.0', '')}</span></div>
                    <div class='tv-footer-right'><span>TOROS</span><span>{f"{total_toros:.1f}".replace('.0', '')}</span></div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("📌 La hoja de Google Sheets está vacía. Carga las partidas en la pestaña anterior.")
        except Exception as e:
            st.warning(f"⚠ Esperando datos de Google Sheets...")
        
    if auto_update:
        time.sleep(8)
        st.rerun()