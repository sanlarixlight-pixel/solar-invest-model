import streamlit as st
import pandas as pd
import glob
import os
import math
import re
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import io

# Налаштування сторінки
st.set_page_config(page_title="Інвестиційний Дашборд", layout="wide")

# Функції для математики
def calc_irr(cash_flows):
    low, high = -0.99, 10.0
    for _ in range(100):
        r = (low + high) / 2
        npv = sum(cf / (1 + r)**t for t, cf in enumerate(cash_flows))
        if npv > 0: low = r
        else: high = r
    return r

def calc_npv(rate, cash_flows):
    return sum(cf / (1 + rate)**t for t, cf in enumerate(cash_flows))

# --- ЛІВА ПАНЕЛЬ ---
st.sidebar.title("⚙️ Налаштування проєкту")

capex_usd = st.sidebar.number_input("Вартість проєкту (USD)", value=50000.0, step=1000.0)
usd_rate = st.sidebar.number_input("Курс Долара", value=43.0, step=0.5)
capacity = st.sidebar.number_input("Ємність АКБ (кВт*год)", value=215.0, step=10.0)
power_inv = st.sidebar.number_input("Потужність Інвертора (кВт)", value=50.0, step=5.0)
power_pv = st.sidebar.number_input("Потужність Сонця PV (кВт)", value=80.0, step=5.0)

st.sidebar.markdown("---")
st.sidebar.subheader("Тарифи та Витрати")
grid_tariff = st.sidebar.number_input("Тариф на Розподіл (грн/МВт)", value=1826.0, step=100.0)
market_tariff = st.sidebar.number_input("Тариф оператора (грн/МВт.год)", value=5.86, step=1.0)
trader_fee = st.sidebar.number_input("Комісія Трейдера (в долях)", value=0.07, step=0.01)
opex_rate = st.sidebar.number_input("Операційні витрати (в долях)", value=0.08, step=0.01)
tax_rate = st.sidebar.number_input("Податок на прибуток", value=0.20, step=0.01)
growth_rate = st.sidebar.number_input("Річний ріст цін по РДН", value=0.08, step=0.01)

st.sidebar.markdown("---")
uploaded_gen_file = st.sidebar.file_uploader("Завантажити новий файл генерації (Excel)", type=["xlsx"])

# --- ГОЛОВНИЙ ЕКРАН ---
st.title("📊 Інвестиційний Дашборд BESS + Solar")

if st.button("🚀 РОЗРАХУВАТИ ТА ЗГЕНЕРУВАТИ ЗВІТИ", type="primary"):
    with st.spinner('Створюємо фінансову модель та генеруємо PDF...'):
        
        CAPEX_UAH = capex_usd * usd_rate
        CONF = {
            "POWER": power_inv, "EFF": 0.92, "BESS_DEGRAD": 0.97**(1/12),
            "PV_DEGRAD_YEARLY": 0.005, "SOILING_LOSS": 0.02, "UPTIME": 0.98,
            "DISCOUNT_RATE": 0.10, "FIXED_COSTS": 0
        }

        try:
            if uploaded_gen_file is not None:
                gen_raw = pd.read_excel(uploaded_gen_file, header=None)
            else:
                gen_raw = pd.read_excel("generation.xlsx", header=None)
            
            header_text = str(gen_raw.iloc[0, 0]) + " " + str(gen_raw.iloc[1, 0])
            match = re.search(r'(\d+[\.,]?\d*)\s*(кВт|kW|kw)', header_text, flags=re.IGNORECASE)
            base_pv = float(match.group(1).replace(',', '.')) if match else 950.0
            solar_scale = power_pv / base_pv if power_pv > 0 else 1.0
            
            start_row = gen_raw[gen_raw.apply(lambda r: r.astype(str).str.contains('Jan', case=False).any(), axis=1)].index[0]
            gen_data = gen_raw.iloc[start_row:start_row+12, 1:25].apply(pd.to_numeric).fillna(0)
            
            price_files = sorted(glob.glob("monthly_prices_*.xlsx"))
            if not price_files:
                st.error("❌ Не знайдено файли цін у папці ")
                st.stop()
        except Exception as e:
            st.error(f"❌ Помилка з файлами: {e}")
            st.stop()

        monthly_rows = []
        yearly_cash_flows = [-CAPEX_UAH]
        balance = -CAPEX_UAH
        payback_month = None
        months_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

        for year in range(1, 11):
            p_growth = (1 + growth_rate) ** (year - 1)
            year_profit = 0
            pv_efficiency = (1 - CONF["SOILING_LOSS"]) * ((1 - CONF["PV_DEGRAD_YEARLY"]) ** (year - 1))
            
            for m_idx, m_name in enumerate(months_labels):
                if m_idx >= len(price_files): continue
                
                cur_cap = capacity * (CONF["BESS_DEGRAD"] ** ((year-1)*12 + m_idx))
                prices_df = pd.read_excel(price_files[m_idx], skiprows=3)
                g_day = gen_data.iloc[m_idx].values * solar_scale * pv_efficiency
                
                soc, m_rev, m_buy, m_sold_kwh, m_grid_fee = 0, 0, 0, 0, 0
                
                for d_idx in range(len(prices_df)):
                    p_day = pd.to_numeric(prices_df.iloc[d_idx, 1:25], errors='coerce').fillna(0).values / 1000 * p_growth
                    h_win = max(1, math.ceil(cur_cap / CONF["POWER"])) if CONF["POWER"] > 0 else 1
                    sorted_p = sorted(p_day)
                    
                    low_p_avg = sum(sorted_p[:h_win]) / h_win
                    high_p_avg = sum(sorted_p[-h_win:]) / h_win
                    
                    cost_per_kwh = (low_p_avg + grid_tariff/1000) / CONF["EFF"]
                    rev_per_kwh = high_p_avg * (1 - trader_fee - opex_rate) - market_tariff/1000
                    
                    arbitrage_active = rev_per_kwh > cost_per_kwh
                    low_p_thresh = sorted_p[h_win-1]
                    high_p_thresh = sorted_p[-h_win]
                    
                    for h in range(24):
                        p, g = p_day[h], g_day[h]
                        ch_pv = min(g, CONF["POWER"], (cur_cap - soc)/CONF["EFF"])
                        soc += ch_pv * CONF["EFF"]
                        
                        if arbitrage_active and p <= low_p_thresh and soc < cur_cap:
                            grid_in = min(CONF["POWER"] - ch_pv, (cur_cap - soc)/CONF["EFF"])
                            m_buy += grid_in * p
                            m_grid_fee += grid_in * (grid_tariff/1000)
                            soc += grid_in * CONF["EFF"]
                        
                        ex_pv = g - ch_pv
                        dis = min(soc, CONF["POWER"]) if p >= high_p_thresh else 0
                        soc -= dis
                        
                        m_rev += (ex_pv + dis) * p
                        m_sold_kwh += (ex_pv + dis)

                m_rev *= CONF["UPTIME"]
                m_buy *= CONF["UPTIME"]
                m_grid_fee *= CONF["UPTIME"]
                m_sold_kwh *= CONF["UPTIME"]

                m_op_fee = (m_sold_kwh / 1000) * market_tariff
                m_trader = m_rev * trader_fee
                m_opex = m_rev * opex_rate
                
                ebit = m_rev - m_buy - m_grid_fee - m_op_fee - m_trader - m_opex - CONF["FIXED_COSTS"]
                tax = max(0, ebit * tax_rate)
                net_profit = ebit - tax
                
                year_profit += net_profit
                balance += net_profit
                
                if balance >= 0 and payback_month is None:
                    payback_month = (year-1)*12 + m_idx + 1

                monthly_rows.append({
                    "Рік": year, "Місяць": m_name, "Валовий Дохід": m_rev, "Закупівля Енергії": m_buy, 
                    "Розподіл": m_grid_fee, "Трейдер": m_trader, "Оператор": m_op_fee, "OPEX": m_opex, 
                    "Податок": tax, "Чистий Прибуток": net_profit, "Баланс": balance,
                    "ROI (%)": ((balance + CAPEX_UAH) / CAPEX_UAH) * 100
                })
                
            yearly_cash_flows.append(year_profit)

        project_irr = calc_irr(yearly_cash_flows) * 100
        project_npv = calc_npv(CONF["DISCOUNT_RATE"], yearly_cash_flows)
        
        df_m = pd.DataFrame(monthly_rows)
        df_y = df_m.groupby("Рік").agg({
            "Валовий Дохід": "sum", "Закупівля Енергії": "sum", "Розподіл": "sum",
            "Трейдер": "sum", "Оператор": "sum", "OPEX": "sum", "Податок": "sum",
            "Чистий Прибуток": "sum", "Баланс": "last", "ROI (%)": "last"
        }).reset_index()

        # --- ГЕНЕРАЦІЯ EXCEL У ПАМ'ЯТЬ ---
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_y.to_excel(writer, sheet_name="Річний Звіт (млн)", index=False)
            df_m.to_excel(writer, sheet_name="Деталізація по Місяцях", index=False)
        excel_data = excel_buffer.getvalue()

        # --- ГЕНЕРАЦІЯ PDF У ПАМ'ЯТЬ (ПРЕМІУМ ДИЗАЙН) ---
        pdf_buffer = io.BytesIO()
        fig_pdf = plt.figure(figsize=(14, 12)) 
        gs = fig_pdf.add_gridspec(2, 1, height_ratios=[1.8, 1.2], hspace=0.25)
        
        ax_chart = fig_pdf.add_subplot(gs[0])
        ax_table = fig_pdf.add_subplot(gs[1])
        
        color_roi = '#2980b9'
        ax_chart.set_xlabel('Роки експлуатації', fontsize=11, fontweight='bold')
        ax_chart.set_ylabel('ROI (%)', color=color_roi, fontsize=11, fontweight='bold')
        ax_chart.plot(df_y["Рік"], df_y["ROI (%)"], color=color_roi, marker='s', linewidth=3)
        ax_chart.tick_params(axis='y', labelcolor=color_roi)
        ax_chart.set_xticks(df_y["Рік"])
        ax_chart.grid(True, linestyle='--', alpha=0.4)

        ax_chart2 = ax_chart.twinx()
        color_bal = '#27ae60'
        ax_chart2.set_ylabel('Баланс (млн грн)', color=color_bal, fontsize=11, fontweight='bold')
        ax_chart2.bar(df_y["Рік"], df_y["Баланс"] / 1e6, color=color_bal, alpha=0.3)
        ax_chart2.plot(df_y["Рік"], df_y["Баланс"] / 1e6, color=color_bal, marker='o', linewidth=2)
        ax_chart2.tick_params(axis='y', labelcolor=color_bal)
        ax_chart2.axhline(0, color='darkred', linestyle='-', linewidth=2)
        
        if payback_month:
            payback_year = payback_month / 12
            ax_chart2.scatter(payback_year, 0, color='red', s=250, zorder=10, edgecolors='white', linewidths=2)
            ax_chart2.annotate(f" ОКУПНІСТЬ:\n {payback_month} міс.", (payback_year, 0), 
                         textcoords="offset points", xytext=(0, 20), ha='center', 
                         fontsize=10, fontweight='bold', color='white',
                         bbox=dict(boxstyle="round,pad=0.4", edgecolor="darkred", facecolor="red"))

        info_text = (
            f"ГЛОБАЛЬНІ МЕТРИКИ (10 РОКІВ):\n"
            f"-----------------------------\n"
            f"CAPEX: {CAPEX_UAH / 1e6:.2f} млн грн\n"
            f"Чистий Прибуток: {df_y['Чистий Прибуток'].sum() / 1e6:.2f} млн грн\n"
            f"IRR: {project_irr:.1f}%\n"
            f"NPV (10%): {project_npv / 1e6:.2f} млн грн\n\n"
            f"ТЕХНІЧНІ ПАРАМЕТРИ:\n"
            f"-----------------------------\n"
            f"АКБ Деградація: 3% / рік\n"
            f"СЕС Деградація: 0.5% / рік\n"
            f"Soiling Losses: 2.0%\n"
            f"System Uptime: 98.0%"
        )
        ax_chart.text(0.02, 0.96, info_text, transform=ax_chart.transAxes, fontsize=10,
                 verticalalignment='top', bbox=dict(boxstyle='round', facecolor='#f8f9fa', alpha=0.95, edgecolor='#ced4da'))
        ax_chart.set_title(f"ФІНАНСОВИЙ ПРОГНОЗ: Інвестиційний Проєкт", fontsize=15, fontweight='bold', pad=15)

        ax_table.axis('off') 
        table_data = [["Період\n(Рік)", "Валовий Дохід\n(млн грн)", "Чистий Прибуток\n(млн грн)", "Накопичений Баланс\n(млн грн)", "ROI\n(%)"]]
        for _, row in df_y.iterrows():
            table_data.append([
                f"Рік {int(row['Рік'])}", f"{row['Валовий Дохід']/1e6:.2f}", f"{row['Чистий Прибуток']/1e6:.2f}",
                f"{row['Баланс']/1e6:.2f}", f"{row['ROI (%)']:.1f}%"
            ])
        
        table = ax_table.table(cellText=table_data, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
        table.auto_set_font_size(False)
        table.set_fontsize(11)
        for (i, j), cell in table.get_celld().items():
            cell.set_edgecolor('#bdc3c7')
            if i == 0:
                cell.set_facecolor('#2c3e50')
                cell.set_text_props(weight='bold', color='white')
            else:
                cell.set_facecolor('#f9f9f9' if i % 2 == 0 else 'white')

        with PdfPages(pdf_buffer) as pdf:
            pdf.savefig(fig_pdf, bbox_inches='tight')
        pdf_data = pdf_buffer.getvalue()
        plt.close(fig_pdf)

        # --- ВІДОБРАЖЕННЯ НА САЙТІ ---
        st.success("✅ Розрахунок та генерація звітів успішно завершені!")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("CAPEX", f"{CAPEX_UAH / 1e6:.2f} млн ₴")
        col2.metric("Прибуток (10р)", f"{df_y['Чистий Прибуток'].sum() / 1e6:.2f} млн ₴")
        col3.metric("IRR (Рентабельність)", f"{project_irr:.1f} %")
        col4.metric("Окупність", f"{payback_month if payback_month else '>120'} міс.")
        
        st.markdown("### 📥 Завантаження детальних звітів")
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="📄 Скачати PDF-презентацію (Графік + Деталі)",
                data=pdf_data,
                file_name="Bankable_Report.pdf",
                mime="application/pdf",
                type="primary"
            )
        with dl_col2:
            st.download_button(
                label="📊 Скачати Excel-модель (Деталізація по місяцях)",
                data=excel_data,
                file_name="Investment_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        st.markdown("---")
        st.markdown("### 📈 Інтерактивний огляд")
        
        # Спрощений графік для вебу
        fig_web, ax1_w = plt.subplots(figsize=(10, 4))
        ax1_w.plot(df_y["Рік"], df_y["ROI (%)"], color=color_roi, marker='s', linewidth=2)
        ax1_w.set_ylabel('ROI (%)', color=color_roi)
        ax2_w = ax1_w.twinx()
        ax2_w.bar(df_y["Рік"], df_y["Баланс"] / 1e6, color=color_bal, alpha=0.3)
        ax2_w.axhline(0, color='darkred', linestyle='-', linewidth=1)
        st.pyplot(fig_web)

        display_df = df_y[["Рік", "Валовий Дохід", "Чистий Прибуток", "Баланс", "ROI (%)"]].copy()
        for col in ["Валовий Дохід", "Чистий Прибуток", "Баланс"]:
            display_df[col] = (display_df[col] / 1e6).round(2)

        st.dataframe(display_df, use_container_width=True)
