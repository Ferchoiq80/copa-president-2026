import streamlit as st
import pandas as pd
import urllib.parse
import os
import time
import shutil

# 1. CONFIGURACIÓN DE PÁGINA (WIDE MODE PARA TV)
st.set_page_config(page_title="Copa President - TV Broadcast Leaderboard", page_icon="🏆", layout="wide")

# 2. CSS ESTILO TELEVISIÓN / BROADCAST OFICIAL
st.markdown("""
    <style>
    /* Ocultar elementos de Streamlit para modo TV limpio */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .tv-header { background: linear-gradient(135deg, #001122, #002244); padding: 15px; border-radius: 10px; text-align: center; border-bottom: 4px solid #c5a059; margin-bottom: 25px; }
    .tv-title { color: white; font-family: 'Arial Black', sans-serif; font-size: 2.5rem; margin: 0; text-transform: uppercase; letter-spacing: 2px; }
    .tv-subtitle { color: #c5a059; font-size: 1.2rem; margin: 5px 0 0 0; font-weight: bold; letter-spacing: 1px; }

    /* Estilo de la Fila TV (Match Card) */
    .tv-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; font-family: 'Helvetica Neue', sans-serif; box-shadow: 0 2px 5px rgba(0,0,0,0.15); border-radius: 6px; overflow: hidden; }
    
    .tv-team-left { width: 38%; background-color: #cc0000; color: white; padding: 12px 20px; font-weight: bold; font-size: 1.2rem; display: flex; justify-content: space-between; align-items: center; border-left: 6px solid #880000; }
    .tv-team-right { width: 38%; background-color: #0044cc; color: white; padding: 12px 20px; font-weight: bold; font-size: 1.2rem; display: flex; justify-content: space-between; align-items: center; border-right: 6px solid #002288; text-align: right; flex-direction: row-reverse; }
    
    .tv-center { width: 22%; background-color: #f4f4f4; color: #002244; padding: 12px 10px; text-align: center; font-weight: 900; font-size: 1.3rem; border-left: 2px solid #ddd; border-right: 2px solid #ddd; font-family: 'Courier New', monospace; }
    
    /* Barra Inferior Acumulada estilo TV */
    .tv-footer-bar { display: flex; margin-top: 30px; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.3); font-family: 'Arial Black', sans-serif; }
    .tv-footer-left { width: 50%; background-color: #cc0000; color: white; padding: 20px; font-size: 2rem; display: flex; justify-content: space-between; align-items: center; }
    .tv-footer-right { width: 50%; background-color: #0044cc; color: white; padding: 20px; font-size: 2rem; display: flex; justify-content: space-between; align-items: center; flex-direction: row-reverse; }
    
    .badge-hcp { background: rgba(0,0,0,0.2); padding: 2px 8px; border-radius: 4px; font-size: 0.9rem; }
    </style>
""", unsafe_allow_html=True)

# 3. LECTURA DE DATOS
@st.cache_data
def cargar_datos():
    try:
        df = pd.read_excel('THE PLAYERS.xlsm', sheet_name='Indices', header=5)
        df.columns = df.columns.str.strip()
        df_validos = df.dropna(subset=['Nombre', 'Apellidos']).copy()
        df_validos['Handicap_Juego'] = pd.to_numeric(df_validos['Handicap_Juego'], errors='coerce').fillna(0)
        return df_validos.head(44)
    except:
        return pd.DataFrame()

df_jugadores = cargar_datos()

# 4. CREACIÓN DE PESTAÑAS
tab_sorteo, tab_leaderboard = st.tabs(["🎲 SORTEO Y EMPAREJAMIENTO", "📺 TV BROADCAST LEADERBOARD"])

# ==========================================
# PESTAÑA 1: EMPAREJAMIENTO (JORNADA 1 O 2)
# ==========================================
with tab_sorteo:
    if len(df_jugadores) < 44:
        st.warning(f"⚠️ Se necesitan 44 jugadores. Solo se encontraron {len(df_jugadores)} en la lista.")
    else:
        df_leones = df_jugadores.iloc[:22].copy()
        df_toros = df_jugadores.iloc[22:44].copy()

        with st.sidebar:
            st.header("⚙️ Configuración del Torneo")
            tipo_jornada = st.selectbox(
                "Selecciona la Jornada / Modalidad:",
                (
                    "1ª Jornada: Four Ball (Pareja 70% Promedio)",
                    "2ª Jornada: Match Play Individual (Careo / 22 Puntos)"
                )
            )
            
            if "1ª Jornada" in tipo_jornada:
                metodo_sorteo = st.radio("Método:", ("Por estricto orden de Hándicap", "Al azar (Balanceado)"))
                ejecutar = st.button("🎲 Generar Partidos 1ª Jornada", type="primary", use_container_width=True)
            else:
                st.info("ℹ️ Define los 4 jugadores de cada Foursome para disputar los 2 partidos individuales.")
                ejecutar = True 

        def calcular_hcp_individual(hcp_base):
            return int((hcp_base * 0.7) + 0.5)

        def calcular_hcp_pareja(j1, j2):
            promedio = (j1['Handicap_Juego'] + j2['Handicap_Juego']) / 2
            return int((promedio * 0.7) + 0.5)

        if "1ª Jornada" in tipo_jornada:
            if ejecutar:
                parejas_leones = []
                parejas_toros = []

                if metodo_sorteo == "Por estricto orden de Hándicap":
                    leones_ordenados = df_leones.sort_values(by='Handicap_Juego', ascending=True).to_dict('records')
                    toros_ordenados = df_toros.sort_values(by='Handicap_Juego', ascending=True).to_dict('records')
                    for i in range(0, 22, 2):
                        parejas_leones.append({"j1": leones_ordenados[i], "j2": leones_ordenados[i+1], "hcp_equipo": calcular_hcp_pareja(leones_ordenados[i], leones_ordenados[i+1])})
                        parejas_toros.append({"j1": toros_ordenados[i], "j2": toros_ordenados[i+1], "hcp_equipo": calcular_hcp_pareja(toros_ordenados[i], toros_ordenados[i+1])})
                else: 
                    leones_mezclados = df_leones.sample(frac=1).to_dict('records')
                    toros_mezclados = df_toros.sample(frac=1).to_dict('records')
                    for i in range(0, 22, 2):
                        parejas_leones.append({"j1": leones_mezclados[i], "j2": leones_mezclados[i+1], "hcp_equipo": calcular_hcp_pareja(leones_mezclados[i], leones_mezclados[i+1])})
                        parejas_toros.append({"j1": toros_mezclados[i], "j2": toros_mezclados[i+1], "hcp_equipo": calcular_hcp_pareja(toros_mezclados[i], toros_mezclados[i+1])})
                    parejas_leones = sorted(parejas_leones, key=lambda x: x['hcp_equipo'])
                    parejas_toros = sorted(parejas_toros, key=lambda x: x['hcp_equipo'])

                st.markdown(f"<h2 style='text-align: center; color: #002244;'>⛳ 1ª Jornada - Four Ball</h2><br>", unsafe_allow_html=True)
                datos_exportar = []
                
                for i in range(11):
                    p_l = parejas_leones[i]
                    p_t = parejas_toros[i]
                    
                    datos_exportar.append({
                        "Match": f"Match {i+1}",
                        "Tipo": "Four Ball",
                        "Participante_Leones": f"{p_l['j1']['Nombre']} & {p_l['j2']['Nombre']}",
                        "HCP_L": p_l['hcp_equipo'],
                        "Participante_Toros": f"{p_t['j1']['Nombre']} & {p_t['j2']['Nombre']}",
                        "HCP_T": p_t['hcp_equipo'],
                        "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""
                    })
                    
                    col_L, col_VS, col_T = st.columns([4, 1, 4])
                    with col_L: st.markdown(f"<div class='tv-team-left' style='width:100%;'><span>🔴 {p_l['j1']['Nombre']} / {p_l['j2']['Nombre']}</span><span class='badge-hcp'>HCP: {p_l['hcp_equipo']}</span></div>", unsafe_allow_html=True)
                    with col_VS: st.markdown(f"<div style='text-align:center; font-weight:bold; font-size:1.2rem; color:#555;'>M{i+1}</div>", unsafe_allow_html=True)
                    with col_T: st.markdown(f"<div class='tv-team-right' style='width:100%;'><span>🔵 {p_t['j1']['Nombre']} / {p_t['j2']['Nombre']}</span><span class='badge-hcp'>HCP: {p_t['hcp_equipo']}</span></div>", unsafe_allow_html=True)
                    st.markdown("<hr>", unsafe_allow_html=True)

                df_export = pd.DataFrame(datos_exportar)
                archivo_digitador = "Live_Scoring_President.xlsx"
                try:
                    df_export.to_excel(archivo_digitador, index=False)
                    st.success(f"✅ ¡Archivo **{archivo_digitador}** generado exitosamente!")
                except Exception as e:
                    st.error(f"❌ Error al guardar archivo: {e}")

        else: # 2ª JORNADA: 22 PARTIDOS INDIVIDUALES
            st.markdown("<h2 style='text-align: center; color: #002244;'>🤝 2ª Jornada - Careo Individual (22 Puntos)</h2>", unsafe_allow_html=True)
            
            dict_leones = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_leones.iterrows()}
            dict_toros = {f"{row['Nombre']} {row['Apellidos']} (HCP: {int(row['Handicap_Juego'])})": row for _, row in df_toros.iterrows()}
            nombres_leones = list(dict_leones.keys())
            nombres_toros = list(dict_toros.keys())
            
            datos_exportar_j2 = []
            form_j2 = st.form(key="form_jornada_2")
            with form_j2:
                for i in range(11):
                    st.markdown(f"### Foursome {i+1}")
                    c1, c2 = st.columns(2)
                    with c1:
                        jl1 = st.selectbox(f"León 1 - F{i+1}", nombres_leones, index=(i*2) % len(nombres_leones), key=f"j2_l1_{i}")
                        jl2 = st.selectbox(f"León 2 - F{i+1}", nombres_leones, index=(i*2+1) % len(nombres_leones), key=f"j2_l2_{i}")
                    with c2:
                        jt1 = st.selectbox(f"Toro 1 - F{i+1}", nombres_toros, index=(i*2) % len(nombres_toros), key=f"j2_t1_{i}")
                        jt2 = st.selectbox(f"Toro 2 - F{i+1}", nombres_toros, index=(i*2+1) % len(nombres_toros), key=f"j2_t2_{i}")
                        
                    j_l1, j_l2 = dict_leones[jl1], dict_leones[jl2]
                    j_t1, j_t2 = dict_toros[jt1], dict_toros[jt2]
                    
                    hcp_jl1, hcp_jt1 = calcular_hcp_individual(j_l1['Handicap_Juego']), calcular_hcp_individual(j_t1['Handicap_Juego'])
                    hcp_jl2, hcp_jt2 = calcular_hcp_individual(j_l2['Handicap_Juego']), calcular_hcp_individual(j_t2['Handicap_Juego'])
                    
                    datos_exportar_j2.append({"Match": f"Match {i*2 + 1} (F{i+1}-A)", "Tipo": "Individual", "Participante_Leones": f"{j_l1['Nombre']} {j_l1['Apellidos']}", "HCP_L": hcp_jl1, "Participante_Toros": f"{j_t1['Nombre']} {j_t1['Apellidos']}", "HCP_T": hcp_jt1, "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""})
                    datos_exportar_j2.append({"Match": f"Match {i*2 + 2} (F{i+1}-B)", "Tipo": "Individual", "Participante_Leones": f"{j_l2['Nombre']} {j_l2['Apellidos']}", "HCP_L": hcp_jl2, "Participante_Toros": f"{j_t2['Nombre']} {j_t2['Apellidos']}", "HCP_T": hcp_jt2, "Leones_H6": "", "Toros_H6": "", "Leones_H12": "", "Toros_H12": "", "Leones_H18": "", "Toros_H18": ""})
                    st.markdown("---")
                    
                guardar_j2 = st.form_submit_button("💾 Guardar Careo e Iniciar 22 Partidos", type="primary", use_container_width=True)
                
            if guardar_j2:
                df_export = pd.DataFrame(datos_exportar_j2)
                archivo_digitador = "Live_Scoring_President.xlsx"
                try:
                    df_export.to_excel(archivo_digitador, index=False)
                    st.success(f"✅ ¡Careo guardado! Se crearon los 22 partidos individuales.")
                except Exception as e:
                    st.error(f"❌ Error al guardar: {e}")

# ==========================================
# PESTAÑA 2: TV BROADCAST LEADERBOARD
# ==========================================
with tab_leaderboard:
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Controles superiores (Piloto automático y acumulado inicial)
    c_ctrl1, c_ctrl2, c_ctrl3 = st.columns([2, 2, 2])
    with c_ctrl1:
        auto_update = st.toggle("🚀 Piloto Automático TV (Recarga 8s)")
    with c_ctrl2:
        pts_previo_leones = st.number_input("Acumulado Día Anterior - Leones", min_value=0.0, max_value=50.0, value=0.0, step=0.5)
    with c_ctrl3:
        pts_previo_toros = st.number_input("Acumulado Día Anterior - Toros", min_value=0.0, max_value=50.0, value=0.0, step=0.5)
        
    if not auto_update:
        st.button("🔄 ACTUALIZAR TV LEADERBOARD", type="primary", use_container_width=True)
        
    archivo_digitador = "Live_Scoring_President.xlsx"
    archivo_temp = "temp_live_scoring.xlsx"
    
    if os.path.exists(archivo_digitador):
        try:
            shutil.copy(archivo_digitador, archivo_temp)
            df_live = pd.read_excel(archivo_temp)
            
            puntos_jornada_leones = 0.0
            puntos_jornada_toros = 0.0
            tv_rows_html = ""
            
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
                
                eq_l = row['Participante_Leones']
                eq_t = row['Participante_Toros']
                
                tv_rows_html += f"""
                <div class='tv-row'>
                    <div class='tv-team-left'>
                        <span>{eq_l}</span>
                        <span class='badge-hcp'>{row['HCP_L']}</span>
                    </div>
                    <div class='tv-center'>{status}</div>
                    <div class='tv-team-right'>
                        <span>{eq_t}</span>
                        <span class='badge-hcp'>{row['HCP_T']}</span>
                    </div>
                </div>
                """
            
            # Totales Acumulados Oficiales TV
            total_leones = pts_previo_leones + puntos_jornada_leones
            total_toros = pts_previo_toros + puntos_jornada_toros
            
            tot_L_str = f"{total_leones:.1f}".replace('.0', '')
            tot_T_str = f"{total_toros:.1f}".replace('.0', '')
            
            # Renderizado del Header TV
            st.markdown(f"""
            <div class='tv-header'>
                <div class='tv-title'>🏆 Presidents Cup - Live Broadcast</div>
                <div class='tv-subtitle'>Resultados en Tiempo Real / Club Tennis Cúcuta</div>
            </div>
            """, unsafe_allow_html=True)
            
            # Renderizado de las filas estilo TV
            st.markdown(tv_rows_html, unsafe_allow_html=True)
            
            # Barra Inferior Acumulada TV
            st.markdown(f"""
            <div class='tv-footer-bar'>
                <div class='tv-footer-left'>
                    <span>LEONES</span>
                    <span>{tot_L_str}</span>
                </div>
                <div class='tv-footer-right'>
                    <span>TOROS</span>
                    <span>{tot_T_str}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
                
        except Exception as e:
            st.warning("⚠ El digitador está guardando cambios. Actualizando pantalla en breve...")
    else:
        st.info("📌 Genera primero los emparejamientos en la pestaña 'Sorteo y Emparejamiento'.")
        
    if auto_update:
        time.sleep(8)
        st.rerun()