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

# 2. CSS ESTILO TV PROFESIONAL (MATCHES EN FORMATO TABLA / SCORECARD)
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .tv-header { background: linear-gradient(135deg, #001122, #002244); padding: 15px; border-radius: 8px; text-align: center; border-bottom: 4px solid #c5a059; margin-bottom: 20px; }
    .tv-title { color: white; font-family: 'Arial Black', sans-serif; font-size: 1.8rem; margin: 0; text-transform: uppercase; letter-spacing: 1px; }
    .tv-subtitle { color: #c5a059; font-size: 1rem; margin: 5px 0 0 0; font-weight: bold; }

    /* CONTENEDOR TIPO TABLA TV */
    .tv-match-table { display: table; width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.15); margin-bottom: 8px; font-family: 'Arial', sans-serif; }
    
    .tv-row { display: table-row; }
    
    /* LADO LEONES (ROJO) */
    .tv-cell-leones { display: table-cell; width: 42%; background-color: #fdf2f2; border-left: 6px solid #cc0000; padding: 10px 14px; vertical-align: middle; color: #222; font-weight: bold; font-size: 0.95rem; border-bottom: 1px solid #e0e0e0; }
    
    /* CENTRO (ESTADO / HOYO) */
    .tv-cell-status { display: table-cell; width: 16%; background-color: #002244; color: white; text-align: center; vertical-align: middle; font-family: 'Courier New', monospace; font-weight: bold; font-size: 1rem; padding: 10px 6px; border-bottom: 1px solid #001122; letter-spacing: 1px; }
    
    /* LADO TOROS (AZUL) */
    .tv-cell-toros { display: table-cell; width: 42%; background-color: #f0f5ff; border-right: 6px solid #0044cc; padding: 10px 14px; vertical-align: middle; color: #222; font-weight: bold; font-size: 0.95rem; text-align: right; border-bottom: 1px solid #e0e0e0; }

    .player-sub { font-size: 0.75rem; color: #666; font-weight: normal; float: right; background: rgba(0,0,0,0.06); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }
    .player-sub-l { font-size: 0.75rem; color: #666; font-weight: normal; float: left; background: rgba(0,0,0,0.06); padding: 1px 5px; border-radius: 4px; margin-top: 2px; }

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
                "Viernes: Four Ball Oficial (Según PDF Capitanes)", 
                "Sábado: Match Play Individual (100% HCP)"
            )
        )

    # ----------------------------------------------------
    # VIERNES: FOUR BALL OFICIAL
    # ----------------------------------------------------
    if tipo_jornada == "Viernes: Four Ball Oficial (Según PDF Capitanes)":
        st.markdown("<h3 style='text-align: center;'>⛳ Viernes - Four Ball (Salidas Oficiales)</h3>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #666;'>Sincronización automática con Google Sheets.</p><br>", unsafe_allow_html=True)

        partidas_viernes = [
            ("Hoyo 1 (Match 1)", "Moreno Vera, Edgar Fernando", 8, "Rincon Ramirez, Jose Oswaldo", 9, "Gutierrez Beltran, Wilson", 8, "Franco Rueda, Cesar Augusto", 8),
            ("Hoyo 2A (Match 2)", "Moreno Vera, Luis Alberto", 7, "Peñaranda Canal, Miguel Enrique", 8, "Consuegra Morales, Juan Carlos", 3, "Hernadez, German", 6),
            ("Hoyo 2B (Match 3)", "Peñaranda Gomez, Jairo", 9, "Peñaranda Arango, Juan Camilo", 14, "Parra Gomez, Luis Eduardo", 11, "Sanchez Reyes, Gladys Edelmira", 15),
            ("Hoyo 3 (Match 4)", "Rubiano Perez, Laura Nadmiye", 11, "Bautista Ramirez, Jairo Jose", 12, "Vargas Gonzalez, Robiel Amed", 11, "Garcia Herreros, DuplatMiguel", 13),
            ("Hoyo 4 (Match 5)", "Zerpa Albarran, Edgardo Jose", 13, "Carrillo Sepulveda, Victor Manuel", 15, "Contreras Gamboa, Jorge", 13, "Ortega Meneses, Roberto", 14),
            ("Hoyo 5 (Match 6)", "Vargas Caceres, Elias Jesus", 18, "Alvarado Rodriguez, Lina Maria", 18, "Monsalve Andres, Andres", 13, "Suarez Castrillon, Zulma Janeth", 24),
            ("Hoyo 6 (Match 7)", "Llanes Gomez, Juan pablo", 18, "Ballesteros Castellanos, Kelly Johanna", 23, "Schloeter Rebolledo, Johann Karl", 14, "Ardila Serrano, Fausto", 20),
            ("Hoyo 7 (Match 8)", "Ramirez Vasquez, Jorge Eliecer", 20, "Porras Liendo, Eliana Marcela", 25, "Rangel Vera, Jaime", 20, "Goyeneche Montoya, Carlos", 22),
            ("Hoyo 8A (Match 9)", "Suarez Castrillón, Fabio Orlando", 11, "Yañez Arellano, Oscar", 13, "Robledo Assaf, Cesar Eduardo", 11, "Giatsidakis Olivares, Juan Carlos", 12),
            ("Hoyo 8B (Match 10)", "Cely Soler, Libardo Del Carmen", 11, "Mancera Basto, Edulfo Antonio", 15, "Santos Padilla, Jose Luis", 12, "Landazabal Molina, Sergio Alfonso", 15),
            ("Hoyo 9 (Match 11)", "Ardila Reyes, Delmer", 11, "Mantilla Duran, Miguel Fabian", 12, "Florez Serrano, Elkin Gregorio", 8, "Galavis Correa, Sergio Andres", 9)
        ]

        datos_exportar_viernes = []
        for salida_hoyo, l1_nom, l1_hcp, l2_nom, l2_hcp, t1_nom, t1_hcp, t2_nom, t2_hcp in partidas_viernes:
            datos_exportar_viernes.append({
                "Match": salida_hoyo, "Tipo": "Four Ball",
                "Participante_Leones": f"{l1_nom} & {l2_nom}", "HCP_L": f"L1:{l1_hcp}, L2:{l2_hcp}",
                "Participante_Toros": f"{t1_nom} & {t2_nom}", "HCP_T": f"T1:{t1_hcp}, T2:{t2_hcp}",
                "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""
            })
            st.markdown(f"""
            <div style='background:white; padding:10px; border-radius:8px; margin-bottom:8px; border:1px solid #ddd;'>
                <b>{salida_hoyo}</b> | 🔴 Leones: {l1_nom} & {l2_nom} vs 🔵 Toros: {t1_nom} & {t2_nom}
            </div>
            """, unsafe_allow_html=True)

        if st.button("🚀 Enviar Salidas del Viernes a Google Sheets", type="primary", use_container_width=True):
            df_export = pd.DataFrame(datos_exportar_viernes)
            sheet = conectar_google_sheets()
            if sheet:
                sheet.clear()
                sheet.update([df_export.columns.values.tolist()] + df_export.values.tolist())
                st.success("✅ ¡Partidos del Viernes sincronizados en Google Sheets correctamente!")

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
                    
                    status = "AS"
                    pt_l = 0.0
                    pt_t = 0.0
                    
                    d6 = (l6 - t6) if pd.notna(l6) and pd.notna(t6) else None
                    d12 = (l12 - t12) if pd.notna(l12) and pd.notna(t12) else None
                    d18 = (l18 - t18) if pd.notna(l18) and pd.notna(t18) else None
                    
                    if d18 is not None:
                        if d18 > 0: status = "L GANA"; pt_l = 1.0
                        elif d18 < 0: status = "T GANA"; pt_t = 1.0
                        else: status = "AS"; pt_l = 0.5; pt_t = 0.5
                    elif d12 is not None:
                        if d12 > 6: status = "L GANA"; pt_l = 1.0
                        elif d12 < -6: status = "T GANA"; pt_t = 1.0
                        elif d12 > 0: status = f"{int(d12)} UP"
                        elif d12 < 0: status = f"{int(abs(d12))} UP"
                        else: status = "AS"
                    elif d6 is not None:
                        if d6 > 0: status = f"{int(d6)} UP"
                        elif d6 < 0: status = f"{int(abs(d6))} UP"
                        else: status = "AS"
                    
                    puntos_jornada_leones += pt_l
                    puntos_jornada_toros += pt_t
                    
                    match_name = row['Match']
                    jugador_leones = row['Participante_Leones']
                    hcp_l = row['HCP_L']
                    jugador_toros = row['Participante_Toros']
                    hcp_t = row['HCP_T']
                    
                    # Estructura de tabla profesional estilo TV
                    tv_tables_html += f"""
                    <div style="font-size: 0.75rem; color: #555; font-weight: bold; margin-bottom: 2px; text-transform: uppercase; letter-spacing: 0.5px;">{match_name}</div>
                    <div class='tv-match-table'>
                        <div class='tv-row'>
                            <div class='tv-cell-leones'>
                                {jugador_leones}
                                <span class='player-sub-l'>HCP: {hcp_l}</span>
                            </div>
                            <div class='tv-cell-status'>
                                {status}
                            </div>
                            <div class='tv-cell-toros'>
                                {jugador_toros}
                                <span class='player-sub'>HCP: {hcp_t}</span>
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