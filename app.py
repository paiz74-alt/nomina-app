import streamlit as st
import pandas as pd
import io
import os
import re
from pypdf import PdfReader
import datetime
import plotly.express as px
import plotly.graph_objects as go

# ----------------- CONFIGURACIÓN DE LA PÁGINA -----------------
st.set_page_config(
    page_title="Nómina Inteligente",
    page_icon="📈",
    layout="wide"
)

st.markdown("""
    <style>
    .main-title { color: #1F497D; font-family: 'Segoe UI', sans-serif; font-weight: 700; margin-bottom: 2px; font-size: 24px; }
    .subtitle { color: #595959; font-family: 'Segoe UI', sans-serif; font-style: italic; margin-bottom: 20px; font-size: 14px; }
    
    .metric-card { 
        background-color: #F8F9FA; 
        padding: 12px; 
        border-radius: 10px; 
        border-top: 5px solid #1F497D; 
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .metric-val { font-size: 20px; font-weight: bold; color: #1F497D; }
    .metric-label { font-size: 11px; color: #595959; font-weight: 600; text-transform: uppercase; }
    
    .annual-card {
        background-color: #F0F4F8;
        padding: 10px;
        border-radius: 8px;
        border-left: 5px solid #1F497D;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 8px;
    }
    .annual-val { font-size: 16px; font-weight: bold; color: #1F497D; }
    .annual-label { font-size: 10px; color: #4A5568; font-weight: 700; text-transform: uppercase; }

    .stButton>button { border-radius: 6px; font-weight: 600; width: 100%; }
    hr { margin-top: 1rem; margin-bottom: 1rem; border-color: #E2E8F0; }
    </style>
""", unsafe_allow_html=True)

DB_FILE = "nomina_db.xlsx"

def load_database():
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_excel(DB_FILE)
            if "Fecha Inicio" in df.columns:
                df["Fecha Inicio"] = df["Fecha Inicio"].astype(str)
            if "Fecha Fin" in df.columns:
                df["Fecha Fin"] = df["Fecha Fin"].astype(str)
            return df
        except Exception:
            pass
            
    default_data = [
        {"Fecha Inicio": "2026-06-08", "Fecha Fin": "2026-06-14", "Semana": "06/08 al 06/14", "Hrs REG": 10.0, "Hrs OT": 0.0, "Días NTP": 2, "Tarifa REG": 16.0, "Sueldo REG": 160.0, "Sueldo OT": 0.0, "Viáticos NTP": 250.0, "Pago Bruto (Gross)": 410.0, "Impuestos": 12.24, "Deducción 401-K": 0.0, "Neto Recibido": 397.76},
        {"Fecha Inicio": "2026-06-15", "Fecha Fin": "2026-06-21", "Semana": "06/15 al 06/21", "Hrs REG": 40.0, "Hrs OT": 6.0, "Días NTP": 7, "Tarifa REG": 17.0, "Sueldo REG": 680.0, "Sueldo OT": 153.0, "Viáticos NTP": 875.0, "Pago Bruto (Gross)": 1708.0, "Impuestos": 97.34, "Deducción 401-K": 0.0, "Neto Recibido": 1622.90},
        {"Fecha Inicio": "2026-06-22", "Fecha Fin": "2026-06-28", "Semana": "06/22 al 06/28", "Hrs REG": 40.0, "Hrs OT": 25.0, "Días NTP": 7, "Tarifa REG": 17.0, "Sueldo REG": 680.0, "Sueldo OT": 637.5, "Viáticos NTP": 875.0, "Pago Bruto (Gross)": 2192.50, "Impuestos": 175.04, "Deducción 401-K": 0.0, "Neto Recibido": 2017.46},
        {"Fecha Inicio": "2026-06-29", "Fecha Fin": "2026-07-05", "Semana": "06/29 al 07/05", "Hrs REG": 40.0, "Hrs OT": 1.0, "Días NTP": 5, "Tarifa REG": 17.0, "Sueldo REG": 680.0, "Sueldo OT": 25.5, "Viáticos NTP": 625.0, "Pago Bruto (Gross)": 1330.50, "Impuestos": 59.77, "Deducción 401-K": 28.22, "Neto Recibido": 1242.51}
    ]
    df_default = pd.DataFrame(default_data)
    df_default.to_excel(DB_FILE, index=False)
    return df_default

def save_database(df):
    try:
        if not df.empty and "Fecha Inicio" in df.columns:
            df = df[df["Fecha Inicio"] != "TOTALES"]
            df = df.drop_duplicates(subset=["Fecha Inicio"], keep="last")
            df = df.sort_values(by="Fecha Inicio").reset_index(drop=True)
        df.to_excel(DB_FILE, index=False)
    except Exception as e:
        st.error(f"Error al guardar: {e}")

if 'historical_db' not in st.session_state:
    st.session_state.historical_db = load_database()

if 'selected_index' not in st.session_state:
    st.session_state.selected_index = 0

def parse_payroll_file(text):
    data = {
        "start_date": datetime.date(2026, 7, 6),
        "reg_hrs": 0.0,
        "ot_hrs": 0.0,
        "ntp_days": 0,
        "rate_reg": 17.0,
        "fits": 0.0,
        "social_security": 0.0,
        "medicare": 0.0,
        "deduc_401k": 0.0
    }
    
    if "CURRENT HOURS & EARNINGS" in text or "Check Number" in text:
        period_match = re.search(r"Pay Period:\s*([\d/]+ - [\d/]+|[\d/]+-[\d/]+)", text, re.IGNORECASE)
        if period_match:
            raw_p = period_match.group(1)
            parts = raw_p.split("-")
            if len(parts) == 2:
                try: data["start_date"] = pd.to_datetime(parts[0].strip()).date()
                except: pass
        
        reg_match = re.search(r"R\s+(\d+)\s+\$([\d\.]+)", text)
        if reg_match:
            data["reg_hrs"] = float(reg_match.group(1))
            data["rate_reg"] = float(reg_match.group(2))
            
        ot_match = re.search(r"O\s+(\d+)\s+\$([\d\.]+)", text)
        if ot_match: data["ot_hrs"] = float(ot_match.group(1))
            
        ntp_match = re.search(r"NTP\s+\$([\d\.,]+)", text)
        if ntp_match:
            amount = float(ntp_match.group(1).replace(",", ""))
            data["ntp_days"] = int(amount / 125.0)

        fit_match = re.search(r"Federal\s+Income Tax\s+\|\s+\$([\d\.]+)", text, re.IGNORECASE)
        if fit_match: data["fits"] = float(fit_match.group(1))

        ss_match = re.search(r"Social Security\s+\|\s+\$([\d\.]+)", text, re.IGNORECASE)
        if ss_match: data["social_security"] = float(ss_match.group(1))

        med_match = re.search(r"Medicare\s+\|\s+\$([\d\.]+)", text, re.IGNORECASE)
        if med_match: data["medicare"] = float(med_match.group(1))

        k_match = re.search(r"401K\s+DEDUCTION\s+\|\s+\$([\d\.]+)", text, re.IGNORECASE)
        if k_match: data["deduc_401k"] = float(k_match.group(1))

    elif "TIMECARD HISTORY" in text:
        all_dates_raw = re.findall(r"\b\d{1,2}/\d{1,2}/\d{4}\b", text)
        if all_dates_raw:
            sorted_dates = sorted(list(set(all_dates_raw)), key=lambda d: pd.to_datetime(d))
            try: data["start_date"] = pd.to_datetime(sorted_dates[0]).date()
            except: pass

        lines = text.split("\n")
        for line in lines:
            if "NON TAX PD" in line or "NON TAX" in line:
                data["ntp_days"] += 1
            elif "REG" in line:
                hrs = re.search(r"(\d+\.\d+)", line)
                if hrs: data["reg_hrs"] += float(hrs.group(1))
            elif "OT" in line:
                normalized_line = line.replace(",", ".")
                hrs = re.search(r"(\d+\.\d+)", normalized_line)
                if hrs: data["ot_hrs"] += float(hrs.group(1))

    return data

st.markdown("<h1 class='main-title'>Nómina con Autoguardado</h1>", unsafe_allow_html=True)

st.markdown("### 📥 Paso 1: Cargar PDF")
uploaded_file = st.file_uploader("Sube tu Paycheck / Timecard", type=["pdf"])

extracted_vals = None
if uploaded_file is not None:
    try:
        reader = PdfReader(uploaded_file)
        full_text = "".join([page.extract_text() + "\n" for page in reader.pages])
        extracted_vals = parse_payroll_file(full_text)
        st.success(f"¡Detectado! Inicio: **{extracted_vals['start_date']}**")
    except Exception as e:
        st.error(f"Error: {e}")

default_start_date = extracted_vals["start_date"] if extracted_vals else datetime.date(2026, 7, 6)
default_reg = extracted_vals["reg_hrs"] if extracted_vals else 40.0
default_ot = extracted_vals["ot_hrs"] if extracted_vals else 16.0
default_ntp = extracted_vals["ntp_days"] if extracted_vals else 7
default_rate = extracted_vals["rate_reg"] if extracted_vals else 17.0

with st.expander("📝 Paso 2: Registrar / Editar Datos Manuales", expanded=True):
    col_f1, col_f2, col_f3 = st.columns([1.5, 1, 1])
    with col_f1:
        fecha_inicio = st.date_input("Fecha Inicio", value=default_start_date)
        fecha_fin = fecha_inicio + datetime.timedelta(days=6)
        semana_formateada = f"{fecha_inicio.strftime('%m/%d')} al {fecha_fin.strftime('%m/%d')}"
    with col_f2:
        reg_input = st.number_input("Hrs REG", min_value=0.0, value=default_reg, step=0.5)
    with col_f3:
        ot_input = st.number_input("Hrs OT", min_value=0.0, value=default_ot, step=0.5)

    col_f4, col_f5 = st.columns(2)
    with col_f4:
        ntp_input = st.number_input("Días Viáticos (NTP)", min_value=0, value=int(default_ntp), step=1)
    with col_f5:
        tarifa_input = st.number_input("Tarifa $/hr", min_value=0.0, value=default_rate, step=0.5)

sueldo_reg = reg_input * tarifa_input
sueldo_ot = ot_input * (tarifa_input * 1.5)
viaticos_ntp = ntp_input * 125.0
gross_total = sueldo_reg + sueldo_ot + viaticos_ntp

if extracted_vals and extracted_vals["fits"] > 0:
    fits = extracted_vals["fits"]
    ss = extracted_vals["social_security"]
    med = extracted_vals["medicare"]
    deduc_401k = extracted_vals["deduc_401k"]
else:
    base_imponible = sueldo_reg + sueldo_ot
    fits = base_imponible * 0.0085
    ss = base_imponible * 0.0620
    med = base_imponible * 0.0145
    deduc_401k = base_imponible * 0.0400

total_impuestos = fits + ss + med
net_pay_calc = gross_total - total_impuestos - deduc_401k

if st.button("💾 Guardar Semana"):
    new_data = {
        "Fecha Inicio": str(fecha_inicio),
        "Fecha Fin": str(fecha_fin),
        "Semana": semana_formateada,
        "Hrs REG": reg_input,
        "Hrs OT": ot_input,
        "Días NTP": int(ntp_input),
        "Tarifa REG": tarifa_input,
        "Sueldo REG": sueldo_reg,
        "Sueldo OT": sueldo_ot,
        "Viáticos NTP": viaticos_ntp,
        "Pago Bruto (Gross)": gross_total,
        "Impuestos": total_impuestos,
        "Deducción 401-K": deduc_401k,
        "Neto Recibido": net_pay_calc
    }
    df_temp = st.session_state.historical_db.copy()
    df_temp = df_temp[df_temp["Fecha Inicio"] != str(fecha_inicio)]
    df_final = pd.concat([df_temp, pd.DataFrame([new_data])], ignore_index=True)
    df_final = df_final.sort_values(by="Fecha Inicio").reset_index(drop=True)
    
    st.session_state.historical_db = df_final
    save_database(df_final)
    st.success("¡Semana guardada!")

st.markdown("---")
db_clean_for_pills = st.session_state.historical_db[st.session_state.historical_db["Fecha Inicio"] != "TOTALES"]

if not db_clean_for_pills.empty:
    db_ordenada = db_clean_for_pills.sort_values("Fecha Inicio").reset_index(drop=True)
    semanas_lista = list(db_ordenada["Semana"].unique())
    max_idx = len(semanas_lista) - 1
    
    if st.session_state.selected_index > max_idx:
        st.session_state.selected_index = max_idx

    col_nav1, col_nav2, col_nav3 = st.columns([1.5, 5, 1.5])
    with col_nav1:
        if st.button("◀ Ant", disabled=(st.session_state.selected_index == 0)):
            st.session_state.selected_index -= 1
            st.rerun()
    with col_nav2:
        semana_seleccionada = st.selectbox(
            "Semana Activa:",
            options=semanas_lista,
            index=st.session_state.selected_index
        )
        st.session_state.selected_index = semanas_lista.index(semana_seleccionada)
    with col_nav3:
        if st.button("Sig ▶", disabled=(st.session_state.selected_index == max_idx)):
            st.session_state.selected_index += 1
            st.rerun()

    datos_semana = db_ordenada[db_ordenada["Semana"] == semana_seleccionada].iloc[0]
    
    st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">BRUTO (GROSS)</div>
            <div class="metric-val">${datos_semana['Pago Bruto (Gross)']:,.2f}</div>
        </div>
        <div class="metric-card" style="border-top-color: #2E7D32;">
            <div class="metric-label" style="color:#2E7D32;">NETO RECIBIDO</div>
            <div class="metric-val" style="color:#2E7D32;">${datos_semana['Neto Recibido']:,.2f}</div>
        </div>
    """, unsafe_allow_html=True)
else:
    st.info("Sin registros.")

st.markdown("---")
tab1, tab2 = st.tabs(["🗃️ Datos", "📊 Gráficos"])

with tab1:
    if not st.session_state.historical_db.empty:
        df_db = st.session_state.historical_db.copy()
        df_db = df_db[df_db["Fecha Inicio"] != "TOTALES"]

        tot_hrs_reg = df_db["Hrs REG"].sum()
        tot_hrs_ot = df_db["Hrs OT"].sum()
        tot_viaticos = df_db["Viáticos NTP"].sum()
        tot_gross = df_db["Pago Bruto (Gross)"].sum()
        tot_impuestos = df_db["Impuestos"].sum()
        tot_401k = df_db["Deducción 401-K"].sum()
        tot_neto = df_db["Neto Recibido"].sum()
        
        st.markdown(f"""
            <div class="annual-card">
                <div class="annual-label">Gross total anual</div>
                <div class="annual-val">${tot_gross:,.2f}</div>
            </div>
            <div class="annual-card" style="border-left-color: #2E7D32;">
                <div class="annual-label" style="color:#2E7D32;">Neto total anual</div>
                <div class="annual-val" style="color:#2E7D32;">${tot_neto:,.2f}</div>
            </div>
        """, unsafe_allow_html=True)

        total_row = pd.DataFrame([{
            "Fecha Inicio": "TOTALES", "Fecha Fin": "---", "Semana": "TOTALES",
            "Hrs REG": tot_hrs_reg, "Hrs OT": tot_hrs_ot, "Días NTP": int(df_db["Días NTP"].sum()),
            "Tarifa REG": None, "Sueldo REG": df_db["Sueldo REG"].sum(), "Sueldo OT": df_db["Sueldo OT"].sum(),
            "Viáticos NTP": tot_viaticos, "Pago Bruto (Gross)": tot_gross, "Impuestos": tot_impuestos,
            "Deducción 401-K": tot_401k, "Neto Recibido": tot_neto
        }])

        df_for_editor = pd.concat([df_db, total_row], ignore_index=True)

        column_configuration = {
            "Fecha Inicio": st.column_config.TextColumn("Fecha", disabled=True),
            "Fecha Fin": st.column_config.TextColumn("Fin", disabled=True),
            "Semana": st.column_config.TextColumn("Semana", disabled=True),
            "Hrs REG": st.column_config.NumberColumn("REG", format="%.1f h"),
            "Hrs OT": st.column_config.NumberColumn("OT", format="%.1f h"),
            "Días NTP": st.column_config.NumberColumn("NTP", format="%d d"),
            "Tarifa REG": st.column_config.NumberColumn("Tarifa", format="$%.2f"),
            "Sueldo REG": st.column_config.NumberColumn("Sueldo R.", format="$%.2f", disabled=True),
            "Sueldo OT": st.column_config.NumberColumn("Sueldo O.", format="$%.2f", disabled=True),
            "Viáticos NTP": st.column_config.NumberColumn("Viát.", format="$%.2f", disabled=True),
            "Pago Bruto (Gross)": st.column_config.NumberColumn("Gross", format="$%.2f", disabled=True),
            "Impuestos": st.column_config.NumberColumn("Tax", format="$%.2f"),
            "Deducción 401-K": st.column_config.NumberColumn("401-K", format="$%.2f"),
            "Neto Recibido": st.column_config.NumberColumn("Net", format="$%.2f", disabled=True),
        }

        edited_db = st.data_editor(
            df_for_editor,
            column_config=column_configuration,
            num_rows="dynamic",
            key="db_editor",
            hide_index=True,
            use_container_width=True
        )
        
        clean_edited = edited_db[edited_db["Fecha Inicio"] != "TOTALES"]
        
        if not clean_edited.equals(df_db):
            for idx, row in clean_edited.iterrows():
                tarifa = row["Tarifa REG"] if pd.notnull(row["Tarifa REG"]) else 17.0
                clean_edited.at[idx, "Sueldo REG"] = row["Hrs REG"] * tarifa
                clean_edited.at[idx, "Sueldo OT"] = row["Hrs OT"] * (tarifa * 1.5)
                clean_edited.at[idx, "Viáticos NTP"] = row["Días NTP"] * 125.0
                nuevo_gross = clean_edited.at[idx, "Sueldo REG"] + clean_edited.at[idx, "Sueldo OT"] + clean_edited.at[idx, "Viáticos NTP"]
                clean_edited.at[idx, "Pago Bruto (Gross)"] = nuevo_gross
                impuestos_actuales = row["Impuestos"] if pd.notnull(row["Impuestos"]) else (nuevo_gross * 0.08)
                deduc_401k_actual = row["Deducción 401-K"] if pd.notnull(row["Deducción 401-K"]) else 0.0
                clean_edited.at[idx, "Neto Recibido"] = nuevo_gross - impuestos_actuales - deduc_401k_actual

            st.session_state.historical_db = clean_edited
            save_database(clean_edited)
            st.toast("💾 ¡Actualizado!")
            st.rerun()
            
    col_b1, col_b2 = st.columns([1, 1])
    with col_b1:
        if st.button("🗑️ Limpiar"):
            st.session_state.historical_db = pd.DataFrame(columns=st.session_state.historical_db.columns)
            save_database(st.session_state.historical_db)
            st.rerun()
            
    with col_b2:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            clean_download = st.session_state.historical_db[st.session_state.historical_db["Fecha Inicio"] != "TOTALES"]
            clean_download.to_excel(writer, index=False, sheet_name="Consolidado")
        
        st.download_button(
            label="📥 Descargar Excel",
            data=output.getvalue(),
            file_name="nomina_historial.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab2:
    if not db_clean_for_pills.empty:
        df_chart = db_clean_for_pills.copy().sort_values("Fecha Inicio")
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=df_chart["Semana"], y=df_chart["Pago Bruto (Gross)"],
            mode='lines+markers', name='Gross', line=dict(color='#1F497D', width=2)
        ))
        fig_trend.add_trace(go.Scatter(
            x=df_chart["Semana"], y=df_chart["Neto Recibido"],
            mode='lines+markers', name='Net', line=dict(color='#2E7D32', width=2)
        ))
        fig_trend.update_layout(margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_trend, use_container_width=True)
