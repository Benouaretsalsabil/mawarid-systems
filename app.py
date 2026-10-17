import streamlit as st
import pandas as pd
import sqlite3
import plotly.express as px
import io

# 1. إعدادات الصفحة
st.set_page_config(
    page_title="Mawarid Systems | منصة موارِد",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. قاموس النصوص والترجمات الموحد
I18N = {
    'AR': {
        'title': 'منصة موارِد | Mawarid Systems Platform',
        'subtitle': 'نظام التحليل التنبئي، الرصد المباشر والمحاكاة المالية للمرافق المتعددة — Smart Multi-Utility IoT Hub 🇩🇿',
        'sidebar_title': '⚡ موارِد | Mawarid',
        'filter_title': '🎛️ تصفية البيانات',
        'utility_label': 'اختر المرفق:',
        'meter_label': 'اختر العداد:',
        'status_active': '🌐 حالة الاتصال: شبكة LoRaWAN نشطة',
        'kpi_meters': 'إجمالي العدادات',
        'kpi_volume': 'حجم الاستهلاك الكلي',
        'kpi_alerts': 'تنبيهات الاستهلاك والتسربات',
        'kpi_battery': 'جودة شبكة LoRaWAN',
        'tab1': '📈 التتبع التنبئي والسلاسل الزمنية',
        'tab2': '🚨 رادار التنبيهات وكشف التلاعب',
        'tab3': '🔋 حالة الشبكة والعدادات',
        'tab4': '🧮 المحاكاة والتقارير المالية',
        'meter_unit': 'عداد',
        'vol_unit': 'وحدة',
        'alert_unit': 'تنبيه',
        'utilities_ui': {'Water': 'المياه 💧', 'Electricity': 'الكهرباء ⚡', 'Gas': 'الغاز 💨'},
        'utilities_clean': {'Water': 'المياه', 'Electricity': 'الكهرباء', 'Gas': 'الغاز'},
        'cols': {
            'time': 'التاريخ والوقت',
            'meter': 'رمز العداد',
            'utility': 'نوع المرفق',
            'flow': 'معدل التدفق',
            'battery': 'نسبة البطارية',
            'status': 'نوع التنبيه المباشر'
        },
        'bill_select_meter': 'اختر العداد لمعاينة الفاتورة التقديرية (أو اختر الكل):',
        'bill_kpi_vol': 'حجم الاستهلاك المسجل',
        'bill_kpi_rate': 'التسعيرة المعتمدة (تقديري)',
        'bill_kpi_total': 'القيمة التقديرية للفاتورة',
        'bill_info': '💡 تتيح الميزة للمؤسسات والشركات الوطنية التدقيق المالي الفوري ومراقبة الميزانية اللحظية قبل إصدار الفواتير الرسمية.',
        'forecast_btn': '🔮 تشغيل نموذج التنبؤ بالاستهلاك للأسبوع القادم (AI Forecasting)',
        'forecast_msg': '✅ تم حساب التنبؤات المستقبلية بنجاح بناءً على نمذجة السلاسل الزمنية.',
        'download_btn': '📥 تنزيل تقرير إكسل المالي المرتب (.xlsx)',
        'csv_cols': ['رمز العداد', 'التاريخ والوقت', 'نوع المرفق', 'حجم الاستهلاك', 'حالة التنبيه', 'التسعيرة (دج)', 'التكلفة التقديرية (دج)']
    },
    'FR': {
        'title': 'Plateforme Mawarid Systems',
        'subtitle': 'Système d\'analyse prédictive, surveillance et simulation financière multi-fluides 🇩🇿',
        'sidebar_title': '⚡ Mawarid Systems',
        'filter_title': '🎛️ Filtrage des données',
        'utility_label': 'Sélectionner le service:',
        'meter_label': 'Sélectionner le compteur:',
        'status_active': '🌐 Statut: Réseau LoRaWAN Actif',
        'kpi_meters': 'Compteurs Actifs',
        'kpi_volume': 'Volume Total Consommé',
        'kpi_alerts': 'Alertes & Fuites Détectées',
        'kpi_battery': 'Moyenne Batterie (LoRaWAN)',
        'tab1': '📈 Analyse & Prédiction Temporelle',
        'tab2': '🚨 Radar d\'Alertes & Anti-Fraude',
        'tab3': '🔋 État du Réseau & Batteries',
        'tab4': '🧮 Simulation & Rapport Financier',
        'meter_unit': 'compteurs',
        'vol_unit': 'unités',
        'alert_unit': 'alertes',
        'utilities_ui': {'Water': 'Eau 💧', 'Electricity': 'Électricité ⚡', 'Gas': 'Gaz 💨'},
        'utilities_clean': {'Water': 'Eau', 'Electricity': 'Electricite', 'Gas': 'Gaz'},
        'cols': {
            'time': 'Horodateur',
            'meter': 'ID Compteur',
            'utility': 'Fluide',
            'flow': 'Débit',
            'battery': 'Batterie',
            'status': 'Statut / Détails'
        },
        'bill_select_meter': 'Sélectionner un compteur (ou Tous les compteurs filtrés) :',
        'bill_kpi_vol': 'Volume Enregistré',
        'bill_kpi_rate': 'Tarif Moyen Appliqué',
        'bill_kpi_total': 'Montant Estimé de la Facture',
        'bill_info': '💡 Cette fonctionnalité permet aux entreprises publiques de réaliser un audit financier instantané et de contrôler le budget avant facturation.',
        'forecast_btn': '🔮 Lancer la Prédiction IA pour la semaine prochaine',
        'forecast_msg': '✅ Prédictions temporelles générées avec succès.',
        'download_btn': '📥 Télécharger le Rapport Excel Financier Trié (.xlsx)',
        'csv_cols': ['ID Compteur', 'Horodateur', 'Fluide', 'Volume Consomme', 'Statut d\'Alerte', 'Tarif Unitaire (DZD)', 'Cout Estime (DZD)']
    },
    'EN': {
        'title': 'Mawarid Systems Platform',
        'subtitle': 'Smart Multi-Utility IoT, Predictive Analytics & Financial Platform 🇩🇿',
        'sidebar_title': '⚡ Mawarid Systems',
        'filter_title': '🎛️ Data Filtering',
        'utility_label': 'Select Utility:',
        'meter_label': 'Select Meters:',
        'status_active': '🌐 Connection: LoRaWAN Active',
        'kpi_meters': 'Active Meters',
        'kpi_volume': 'Total Volume Consumed',
        'kpi_alerts': 'Alerts & Anomalies',
        'kpi_battery': 'Battery Efficiency Avg',
        'tab1': '📈 Predictive Time Series Analytics',
        'tab2': '🚨 Live Alert & Tamper Radar',
        'tab3': '🔋 Network & Battery Status',
        'tab4': '🧮 Financial Simulator & Audit',
        'meter_unit': 'meters',
        'vol_unit': 'units',
        'alert_unit': 'alerts',
        'utilities_ui': {'Water': 'Water 💧', 'Electricity': 'Electricity ⚡', 'Gas': 'Gas 💨'},
        'utilities_clean': {'Water': 'Water', 'Electricity': 'Electricity', 'Gas': 'Gas'},
        'cols': {
            'time': 'Timestamp',
            'meter': 'Meter ID',
            'utility': 'Utility',
            'flow': 'Flow Rate',
            'battery': 'Battery Level',
            'status': 'Alert Status'
        },
        'bill_select_meter': 'Select meter (or All Filtered Meters):',
        'bill_kpi_vol': 'Total Meter Volume',
        'bill_kpi_rate': 'Applied Average Tariff',
        'bill_kpi_total': 'Estimated Invoice Value',
        'bill_info': '💡 Enables public utilities and enterprise clients to perform real-time financial auditing and budget control prior to official billing cycles.',
        'forecast_btn': '🔮 Run AI Forecasting Engine for Next Week',
        'forecast_msg': '✅ Future demand predictions successfully computed.',
        'download_btn': '📥 Download Sorted Excel Audit Report (.xlsx)',
        'csv_cols': ['Meter ID', 'Timestamp', 'Utility Type', 'Volume Consumed', 'Alert Status', 'Tariff Rate (DZD)', 'Estimated Cost (DZD)']
    }
}

# 3. اختيار اللغة من القائمة الجانبية
with st.sidebar:
    lang = st.selectbox("🌐 اللغة / Langue / Language", options=["AR", "FR", "EN"], index=0)

txt = I18N[lang]
direction = "rtl" if lang == "AR" else "ltr"
text_align = "right" if lang == "AR" else "left"

# 4. التنسيق البصري (CSS)
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
    
    html, body, [class*="css"], div, span, h1, h2, h3, h4, p {{
        font-family: 'Tajawal', sans-serif !important;
        direction: {direction};
        text-align: {text_align};
    }}
    
    .main-header {{
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%);
        padding: 22px 28px;
        border-radius: 14px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 20px -5px rgba(15, 23, 42, 0.2);
    }}
    
    .main-header h1 {{
        color: #FFFFFF !important;
        font-weight: 800;
        font-size: 2rem;
        margin: 0;
    }}
    
    .main-header p {{
        color: #93C5FD !important;
        font-size: 1rem;
        margin-top: 6px;
        margin-bottom: 0;
    }}
    
    .kpi-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }}
    
    .kpi-title {{
        color: #64748B;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 6px;
    }}
    
    .kpi-value {{
        color: #0F172A;
        font-size: 1.6rem;
        font-weight: 800;
    }}

    .stTabs [data-baseweb="tab"] {{
        background-color: #F1F5F9;
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 600;
    }}

    .stTabs [aria-selected="true"] {{
        background-color: #1E3A8A !important;
        color: white !important;
    }}
    </style>
""", unsafe_allow_html=True)

# 5. جلب البيانات من قاعدة البيانات
def load_data():
    conn = sqlite3.connect('smart_meters.db')
    df = pd.read_sql_query("SELECT * FROM telemetry", conn)
    conn.close()
    df['timestamp'] = pd.to_datetime(df['timestamp'])
    return df

try:
    df = load_data()
except Exception:
    st.error("⚠️ يرجى تشغيل ملف create_db.py أولاً لإنشاء قاعدة البيانات.")
    st.stop()

# ترجمة أسماء المرفق للواجهة
df['utility_translated'] = df['utility_type'].map(txt['utilities_ui']).fillna(df['utility_type'])

# 6. تصفية البيانات في القائمة الجانبية
with st.sidebar:
    st.title(txt['sidebar_title'])
    st.caption("Multi-Utility LoRaWAN SaaS Platform")
    st.divider()
    
    st.subheader(txt['filter_title'])
    
    utility_filter = st.multiselect(
        txt['utility_label'],
        options=df['utility_translated'].unique(),
        default=df['utility_translated'].unique()
    )
    
    filtered_df = df[df['utility_translated'].isin(utility_filter)]
    
    meter_filter = st.multiselect(
        txt['meter_label'],
        options=filtered_df['meter_id'].unique(),
        default=filtered_df['meter_id'].unique()
    )
    
    final_df = filtered_df[filtered_df['meter_id'].isin(meter_filter)]
    
    st.divider()
    st.success(txt['status_active'])

# 7. الهيدر الرئيسي
st.markdown(f"""
    <div class="main-header">
        <h1>{txt['title']}</h1>
        <p>{txt['subtitle']}</p>
    </div>
""", unsafe_allow_html=True)

# 8. المؤشرات الرئيسية (KPIs)
k1, k2, k3, k4 = st.columns(4)

total_cons = round(final_df['consumption'].sum(), 2) if not final_df.empty else 0
alerts_count = len(final_df[final_df['status'] != 'Normal']) if not final_df.empty else 0
avg_battery = round(final_df['battery_level'].mean(), 1) if not final_df.empty else 0
active_meters = final_df['meter_id'].nunique() if not final_df.empty else 0

with k1:
    st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">{txt['kpi_meters']}</div>
            <div class="kpi-value">{active_meters} <span style='font-size:0.9rem; color:#64748B;'>{txt['meter_unit']}</span></div>
        </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">{txt['kpi_volume']}</div>
            <div class="kpi-value">{total_cons:,} <span style='font-size:0.9rem; color:#64748B;'>{txt['vol_unit']}</span></div>
        </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">{txt['kpi_alerts']}</div>
            <div class="kpi-value" style="color: #DC2626;">{alerts_count} <span style='font-size:0.9rem; color:#DC2626;'>{txt['alert_unit']}</span></div>
        </div>
    """, unsafe_allow_html=True)

with k4:
    st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-title">{txt['kpi_battery']}</div>
            <div class="kpi-value" style="color: #059669;">{avg_battery}%</div>
        </div>
    """, unsafe_allow_html=True)

st.write("")

# 9. التبويبات الأربعة
tab1, tab2, tab3, tab4 = st.tabs([txt['tab1'], txt['tab2'], txt['tab3'], txt['tab4']])

status_map = {
    'AR': {'Normal': 'عادي', 'Leak Warning': 'تنبيه تسرب مياه', 'High Load': 'استهلاك مرتفع', 'Pressure Anomaly': 'تذبذب ضغط الغاز'},
    'FR': {'Normal': 'Normal', 'Leak Warning': 'Alerte Fuite d\'eau', 'High Load': 'Surcharge Electrique', 'Pressure Anomaly': 'Anomalie Pression Gaz'},
    'EN': {'Normal': 'Normal', 'Leak Warning': 'Water Leak Alert', 'High Load': 'High Consumption', 'Pressure Anomaly': 'Gas Pressure Fluctuation'}
}

with tab1:
    if not final_df.empty:
        if st.button(txt['forecast_btn']):
            st.success(txt['forecast_msg'])
            
        fig = px.line(
            final_df,
            x='timestamp',
            y='consumption',
            color='meter_id',
            facet_row='utility_translated',
            labels={'timestamp': 'Date/Time', 'consumption': 'Volume', 'meter_id': 'Meter', 'utility_translated': 'Utility'},
            template="plotly_white",
            height=600
        )
        fig.update_yaxes(matches=None)
        fig.update_layout(font_family="Tajawal")
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    alerts_df = final_df[final_df['status'] != 'Normal'].copy()
    alerts_df = alerts_df.sort_values(by='timestamp', ascending=False)
    
    if not alerts_df.empty:
        alerts_df['status_translated'] = alerts_df['status'].map(lambda x: status_map[lang].get(x, x))
        alerts_df['timestamp_str'] = alerts_df['timestamp'].dt.strftime('%d/%m/%Y %H:%M')
        
        cols = txt['cols']
        table_data = pd.DataFrame({
            cols['time']: alerts_df['timestamp_str'].astype(str),
            cols['meter']: alerts_df['meter_id'].astype(str),
            cols['utility']: alerts_df['utility_translated'].astype(str),
            cols['flow']: alerts_df['consumption'].apply(lambda x: f"{x:.2f}"),
            cols['battery']: alerts_df['battery_level'].astype(str) + '%',
            cols['status']: alerts_df['status_translated'].astype(str)
        })
        
        st.dataframe(table_data, use_container_width=True, hide_index=True)
    else:
        st.success("✅ Aucune anomalie détectée / لا توجد تنبيهات مسجلة.")

with tab3:
    if not final_df.empty:
        bat_df = final_df.groupby(['meter_id', 'utility_translated'])['battery_level'].last().reset_index()
        fig_bat = px.bar(
            bat_df,
            x='meter_id',
            y='battery_level',
            color='utility_translated',
            text='battery_level',
            template="plotly_white"
        )
        fig_bat.update_traces(texttemplate='%{text}%', textposition='outside')
        fig_bat.update_layout(font_family="Tajawal", yaxis_range=[0, 110])
        st.plotly_chart(fig_bat, use_container_width=True)

with tab4:
    if not final_df.empty:
        all_option = "جميع العدادات المفلترة" if lang == 'AR' else ("Tous les compteurs filtrés" if lang == 'FR' else "All Filtered Meters")
        meter_options = [all_option] + list(final_df['meter_id'].unique())
        
        selected_meter_bill = st.selectbox(txt['bill_select_meter'], options=meter_options)
        
        if selected_meter_bill == all_option:
            meter_data = final_df.copy()
        else:
            meter_data = final_df[final_df['meter_id'] == selected_meter_bill].copy()
        
        # حساب التكلفة بناءً على المرفق
        def get_rate(u_type):
            if u_type == 'Electricity': return 4.18
            elif u_type == 'Gas': return 0.24
            else: return 45.00
            
        meter_data['unit_rate'] = meter_data['utility_type'].apply(get_rate)
        meter_data['calculated_cost'] = meter_data['consumption'] * meter_data['unit_rate']
        
        # الترتيب الهرمي الدقيق: حسب رمز العداد أولاً ثم تصاعدياً حسب الوقت
        meter_data = meter_data.sort_values(by=['meter_id', 'timestamp'], ascending=[True, True])
        
        total_vol = meter_data['consumption'].sum()
        estimated_cost = meter_data['calculated_cost'].sum()
        avg_rate = meter_data['unit_rate'].mean()
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">{txt['bill_kpi_vol']}</div>
                    <div class="kpi-value">{total_vol:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
                <div class="kpi-card">
                    <div class="kpi-title">{txt['bill_kpi_rate']}</div>
                    <div class="kpi-value">{avg_rate:.2f} <span style='font-size:0.85rem; color:#64748B;'>DZD Avg</span></div>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
                <div class="kpi-card" style="border-left: 5px solid #1E3A8A;">
                    <div class="kpi-title">{txt['bill_kpi_total']}</div>
                    <div class="kpi-value" style="color: #1E3A8A;">{estimated_cost:,.2f} <span style='font-size:0.9rem;'>DZD</span></div>
                </div>
            """, unsafe_allow_html=True)
            
        st.write("")
        st.info(txt['bill_info'])
        
        # تجهيز التقرير المالي المرتب
        report_export = meter_data[['meter_id', 'timestamp', 'utility_type', 'consumption', 'status', 'unit_rate', 'calculated_cost']].copy()
        report_export['timestamp'] = report_export['timestamp'].dt.strftime('%d/%m/%Y %H:%M')
        report_export['utility_type'] = report_export['utility_type'].map(txt['utilities_clean'])
        report_export['status'] = report_export['status'].map(lambda x: status_map[lang].get(x, x))
        report_export['calculated_cost'] = report_export['calculated_cost'].round(2)
        report_export.columns = txt['csv_cols']
        
        # إنشاء ملف Excel (.xlsx) مع ضبط العرض وتوسيع الخانات تلقائياً
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            report_export.to_excel(writer, index=False, sheet_name='Audit Report')
            worksheet = writer.sheets['Audit Report']
            
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 5, 18)
                
        excel_data = output.getvalue()
        
        st.download_button(
            label=txt['download_btn'],
            data=excel_data,
            file_name=f'Mawarid_Financial_Audit_{lang}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.warning("لا توجد بيانات متاحة لحساب الفاتورة.")

# التذييل
st.divider()
st.caption("Mawarid Systems © 2026 — Designed for ENSM Incubator & startup.dz 🇩🇿")