import datetime

import pandas as pd
import streamlit as st


def load_excel(file, header_row):
    # Use the specified header row (subtract 1 because pandas uses 0-based indexing)
    df = pd.read_excel(file, header=header_row - 1)
    return df


def main():
    st.title("Organizador de Excel")

    uploaded_file = st.file_uploader("Añade tu fichero Excel", type=["xlsx", "xls"])
    if uploaded_file is not None:
        # Step 1: Let the user specify the header row
        header_row = st.number_input(
            "Introduce el número de línea en que empiezan las cabeceras",
            min_value=1,
            value=1,
            step=1,
        )

        # Load the Excel file with the specified header row
        df = load_excel(uploaded_file, header_row)

        # Step 2: Display the headers and ask the user if they are correct
        st.write("### Cabeceras Detectadas:")
        headers = df.columns.tolist()
        st.write(headers)

        # Ask the user to confirm the headers
        confirm_headers = st.radio(
            "¿Son correctas estas cabeceras?",
            options=["Sí", "No"],
            index=1,
        )

        if confirm_headers == "No":
            st.warning(
                "Por favor, ajusta arriba el número correcto de fila de las cabeceras."
            )
            return

        # Step 3: Proceed with filtering if headers are confirmed
        st.write("### Vista previa de los datos:")
        st.write(df.head())

        # Step 4: Let the user choose between filtering or sorting
        action = st.radio(
            "¿Qué acción deseas realizar?",
            options=["Filtrar por un valor", "Ordenar por una columna"],
            index=0,  # Default to "Filtrar por un valor"
        )

        if action == "Filtrar por un valor":
            # Step 5: Let the user select a column to filter
            selected_column = st.selectbox(
                "Selecciona una columna sobre la que filtrar:", headers
            )

            # Step 6: Get unique values from the selected column for filtering
            if selected_column:
                unique_values = df[selected_column].unique()
                selected_value = st.selectbox(
                    f"Selecciona un valor sobre el que filtrar de '{selected_column}'",
                    unique_values,
                )

                # Step 7: Filter the DataFrame based on the selected value
                if selected_value:
                    filtered_df = df[df[selected_column] == selected_value]
                    st.write(
                        f"### Datos de {selected_column} filtrados por: {selected_value}"
                    )
                    st.write(filtered_df)

                    # Step 8: Add sorting functionality to the filtered dataset
                    st.write("### Ordenar los datos filtrados")
                    sort_column = st.selectbox(
                        "Selecciona una columna para ordenar los datos filtrados:",
                        headers,
                    )
                    sort_order = st.radio(
                        "¿En qué orden deseas ordenar?",
                        options=["Ascendente", "Descendente"],
                        index=0,  # Default to "Ascendente"
                    )

                    if sort_column:
                        ascending = sort_order == "Ascendente"
                        sorted_filtered_df = filtered_df.sort_values(
                            by=sort_column, ascending=ascending
                        )
                        st.write(
                            f"### Datos filtrados ordenados por {sort_column} ({sort_order.lower()})"
                        )
                        st.write(sorted_filtered_df)

                        # Step 9: Ask the user for the output file name
                        today = datetime.datetime.today().strftime(
                            "%d_%m_%Y"
                        )  # Format: DD_MM_YYYY
                        output_file = st.text_input(
                            "Introduce el nombre del archivo de salida (sin extensión):",
                            value=f"facturas_{selected_column}_{today}",  # Default name with date
                        )
                        output_file = f"{output_file}.xlsx"

                        # Step 10: Provide a downloadable Excel file with the filtered and sorted data
                        st.write("### Descargar datos filtrados y ordenados")
                        sorted_filtered_df.to_excel(output_file, index=False)
                        with open(output_file, "rb") as file:
                            st.download_button(
                                label="Descargar Excel con datos filtrados y ordenados",
                                data=file,
                                file_name=output_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

        elif action == "Ordenar por una columna":
            # Step 5: Let the user select a column to sort
            selected_column = st.selectbox(
                "Selecciona una columna para ordenar:", headers
            )

            # Step 6: Let the user choose ascending or descending order
            sort_order = st.radio(
                "¿En qué orden deseas ordenar?",
                options=["Ascendente", "Descendente"],
                index=0,  # Default to "Ascendente"
            )

            # Step 7: Sort the DataFrame
            if selected_column:
                ascending = sort_order == "Ascendente"
                sorted_df = df.sort_values(by=selected_column, ascending=ascending)
                st.write(
                    f"### Datos ordenados por {selected_column} ({sort_order.lower()})"
                )
                st.write(sorted_df)

                # Step 8: Ask the user for the output file name
                today = datetime.datetime.today().strftime(
                    "%d_%m_%Y"
                )  # Format: DD_MM_YYYY
                output_file = st.text_input(
                    "Introduce el nombre del archivo de salida (sin extensión):",
                    value=f"facturas_ordenadas_{selected_column}_{today}",  # Default name with date
                )
                output_file = f"{output_file}.xlsx"

                # Step 9: Provide a downloadable Excel file with the sorted data
                st.write("### Descargar datos ordenados")
                sorted_df.to_excel(output_file, index=False)
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Descargar Excel con datos ordenados",
                        data=file,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()
