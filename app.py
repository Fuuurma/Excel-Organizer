import datetime
import io

import pandas as pd
import streamlit as st


def load_excel(file, header_row):
    # Use the specified header row (subtract 1 because pandas uses 0-based indexing)
    df = pd.read_excel(file, header=header_row - 1)
    return df


def get_excel_download_link(df, filename):
    """
    Generates a download link for a dataframe without saving to disk
    """
    # Use BytesIO to avoid saving files to disk
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

    # Reset the buffer position to the beginning
    output.seek(0)
    return output.getvalue(), filename


def main():
    st.title("Organizador de Excel")

    # Upload section with improved styling
    st.markdown("### 📁 Selecciona tu archivo Excel")
    uploaded_file = st.file_uploader("", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Header configuration section
        st.markdown("### ⚙️ Configuración de cabeceras")

        # Step 1: Let the user specify the header row
        header_row = st.number_input(
            "Introduce el número de línea en que empiezan las cabeceras",
            min_value=1,
            value=1,
            step=1,
            help="Si tus cabeceras no están en la primera fila, indica aquí el número de fila correcto",
        )

        # Load the Excel file with the specified header row
        df = load_excel(uploaded_file, header_row)

        # Step 2: Display the headers and ask the user if they are correct
        st.write("**Cabeceras Detectadas:**")
        headers = df.columns.tolist()
        st.write(", ".join(headers))

        # Ask the user to confirm the headers with a more pleasant UI
        confirm_headers = st.checkbox("Confirmar cabeceras", value=False)

        if not confirm_headers:
            st.info(
                "Por favor, confirma las cabeceras o ajusta el número de fila para continuar."
            )
        else:
            # Show preview without using an expander
            st.markdown("### 👁️ Vista previa de los datos")
            st.dataframe(df.head(), use_container_width=True)

            # Tab-based interface for different operations
            tab1, tab2 = st.tabs(["🔍 Filtrar", "🔢 Ordenar"])

            with tab1:
                # Filtering section
                selected_column = st.selectbox(
                    "Selecciona una columna sobre la que filtrar:", headers
                )

                # Get unique values from the selected column for filtering
                if selected_column:
                    unique_values = df[selected_column].unique()
                    selected_value = st.selectbox(
                        f"Selecciona un valor de '{selected_column}'",
                        unique_values,
                    )

                    # Filter the DataFrame based on the selected value
                    if selected_value:
                        filtered_df = df[df[selected_column] == selected_value]
                        st.write(
                            f"**Datos filtrados por:** {selected_column} = {selected_value}"
                        )
                        st.dataframe(filtered_df, use_container_width=True)

                        # Secondary filtering option
                        add_second_filter = st.checkbox("Añadir segundo filtro")
                        if add_second_filter:
                            remaining_headers = [
                                h for h in headers if h != selected_column
                            ]
                            second_column = st.selectbox(
                                "Selecciona una segunda columna para filtrar:",
                                remaining_headers,
                            )
                            if second_column:
                                second_unique_values = filtered_df[
                                    second_column
                                ].unique()
                                second_selected_value = st.selectbox(
                                    f"Selecciona un valor de '{second_column}'",
                                    second_unique_values,
                                )
                                if second_selected_value:
                                    filtered_df = filtered_df[
                                        filtered_df[second_column]
                                        == second_selected_value
                                    ]
                                    st.write(
                                        f"**Datos filtrados por:** {selected_column} = {selected_value} y {second_column} = {second_selected_value}"
                                    )
                                    st.dataframe(filtered_df, use_container_width=True)

                        # Sorting the filtered data
                        st.write("---")
                        st.write("### Ordenar los datos filtrados")
                        sort_column = st.selectbox("Ordenar por:", headers)
                        sort_order = st.radio(
                            "Orden:",
                            options=["Ascendente", "Descendente"],
                            horizontal=True,
                        )

                        if sort_column:
                            ascending = sort_order == "Ascendente"
                            sorted_filtered_df = filtered_df.sort_values(
                                by=sort_column, ascending=ascending
                            )
                            st.write(
                                f"**Datos ordenados por:** {sort_column} ({sort_order.lower()})"
                            )
                            st.dataframe(sorted_filtered_df, use_container_width=True)

                            # Display statistics directly (no expander)
                            st.markdown("### 📊 Estadísticas")
                            st.write(
                                f"**Total de registros:** {len(sorted_filtered_df)}"
                            )
                            if pd.api.types.is_numeric_dtype(
                                sorted_filtered_df[sort_column]
                            ):
                                st.write(
                                    f"**Suma:** {sorted_filtered_df[sort_column].sum()}"
                                )
                                st.write(
                                    f"**Media:** {sorted_filtered_df[sort_column].mean():.2f}"
                                )
                                st.write(
                                    f"**Mínimo:** {sorted_filtered_df[sort_column].min()}"
                                )
                                st.write(
                                    f"**Máximo:** {sorted_filtered_df[sort_column].max()}"
                                )

                            # Generate download link
                            st.markdown("### 📥 Descargar datos")
                            today = datetime.datetime.today().strftime("%d_%m_%Y")
                            output_filename = f"facturas_{selected_column}_{today}.xlsx"

                            # Use custom filename if provided
                            custom_filename = st.text_input(
                                "Nombre personalizado del archivo (opcional):",
                                value="",
                                placeholder="Deja en blanco para usar el nombre por defecto",
                            )

                            if custom_filename:
                                # Add .xlsx extension if not present
                                if not custom_filename.endswith(".xlsx"):
                                    custom_filename += ".xlsx"
                                output_filename = custom_filename

                            # Generate download button without saving to disk
                            excel_data, file_name = get_excel_download_link(
                                sorted_filtered_df, output_filename
                            )
                            st.download_button(
                                label="📥 Descargar datos filtrados y ordenados",
                                data=excel_data,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

            with tab2:
                # Sorting section
                selected_column = st.selectbox(
                    "Selecciona una columna para ordenar:",
                    headers,
                    key="sort_tab_column",
                )

                # Let the user choose ascending or descending order
                sort_order = st.radio(
                    "Orden:",
                    options=["Ascendente", "Descendente"],
                    horizontal=True,
                    key="sort_tab_order",
                )

                # Sort the DataFrame
                if selected_column:
                    ascending = sort_order == "Ascendente"
                    sorted_df = df.sort_values(by=selected_column, ascending=ascending)
                    st.write(
                        f"**Datos ordenados por:** {selected_column} ({sort_order.lower()})"
                    )
                    st.dataframe(sorted_df, use_container_width=True)

                    # Display statistics directly (no expander)
                    st.markdown("### 📊 Estadísticas")
                    st.write(f"**Total de registros:** {len(sorted_df)}")
                    if pd.api.types.is_numeric_dtype(sorted_df[selected_column]):
                        st.write(f"**Suma:** {sorted_df[selected_column].sum()}")
                        st.write(f"**Media:** {sorted_df[selected_column].mean():.2f}")
                        st.write(f"**Mínimo:** {sorted_df[selected_column].min()}")
                        st.write(f"**Máximo:** {sorted_df[selected_column].max()}")

                    # Generate download link
                    st.markdown("### 📥 Descargar datos")
                    today = datetime.datetime.today().strftime("%d_%m_%Y")
                    output_filename = (
                        f"facturas_ordenadas_{selected_column}_{today}.xlsx"
                    )

                    # Use custom filename if provided
                    custom_filename = st.text_input(
                        "Nombre personalizado del archivo (opcional):",
                        value="",
                        placeholder="Deja en blanco para usar el nombre por defecto",
                        key="sort_tab_filename",
                    )

                    if custom_filename:
                        # Add .xlsx extension if not present
                        if not custom_filename.endswith(".xlsx"):
                            custom_filename += ".xlsx"
                        output_filename = custom_filename

                    # Generate download button without saving to disk
                    excel_data, file_name = get_excel_download_link(
                        sorted_df, output_filename
                    )
                    st.download_button(
                        label="📥 Descargar datos ordenados",
                        data=excel_data,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    st.markdown("---")
    st.markdown("📊 Organizador de Excel v2.0")


if __name__ == "__main__":
    main()
