import streamlit as st
import pandas as pd
import os
import script

st.set_page_config(page_title="Анализатор Рейсов", page_icon="🚚", layout="wide")

BASE_FILE = "base.xlsx"

st.title("🚚 Автоматический расчет рейсов ТС")

st.markdown("""
Загрузите новую выгрузку из системы видеонаблюдения для автоматической кластеризации рейсов, очистки OCR ошибок номеров и генерации итогового отчета.
""")

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. База / Справочник (Опционально)")
    st.markdown("Если нужно подтягивать транспортное средство и группу.")
    
    base_info = st.empty()
    if os.path.exists(BASE_FILE):
        base_info.success("✅ Справочник загружен и активен на сервере.")
    else:
        base_info.warning("⚠️ Справочник отсутствует (строки не будут обогащены).")

    with st.expander("Обновить справочник"):
        uploaded_base = st.file_uploader("Загрузить новый base.xlsx", type=['xlsx', 'xls'], key="base_uploader")
        if uploaded_base is not None:
            with open(BASE_FILE, "wb") as f:
                f.write(uploaded_base.getbuffer())
            st.success("Новая база успешно сохранена! Пожалуйста, обновите страницу.")

with col2:
    st.subheader("2. Загрузка Выгрузки")
    uploaded_raw = st.file_uploader("Загрузите файл выгрузки (report_plates...)", type=['xlsx', 'xls'])
    
st.divider()

if uploaded_raw is not None:
    if st.button("🚀 ЗАПУСТИТЬ РАСЧЕТ", use_container_width=True, type="primary"):
        with st.spinner("Анализирую рейсы, сшиваю номера и формирую отчет..."):
            temp_raw = "temp_" + uploaded_raw.name
            try:
                # Временное сохранение для pandas
                with open(temp_raw, "wb") as f:
                    f.write(uploaded_raw.getbuffer())
                    
                # Запуск основного пайплайна
                output_file, df_summary = script.main(temp_raw, BASE_FILE)
                
                st.success("Расчет успешно завершен!")
                
                st.subheader("📊 Краткая сводка")
                st.dataframe(df_summary, use_container_width=True, hide_index=True)
                
                # Читаем результат для скачивания
                with open(output_file, "rb") as f:
                    excel_data = f.read()
                    
                st.download_button(
                    label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЕТ",
                    data=excel_data,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            except Exception as e:
                st.error(f"Произошла ошибка при обработке: {e}")
            finally:
                # Очистка временных файлов
                if os.path.exists(temp_raw):
                    os.remove(temp_raw)
                # Если файл отчета тоже нужно удалять с сервера после генерации кнопки
                # (Кнопка загрузит его в память браузера, так что сам файл можно удалить)
                if 'output_file' in locals() and os.path.exists(output_file):
                    os.remove(output_file)
