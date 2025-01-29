import pandas as pd
import streamlit as st


def load_excel(file):
    df = pd.read_excel(file)
    return df


def main():
    st.title("Excel Organizer")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    if uploaded_file is not None:
        df = load_excel(uploaded_file)
        st.write("### Excel File Headers")
        headers = df.columns.tolist()
        st.write(headers)

        st.write("### Preview of the Data")
        st.write(df.head())

        # Step 1: Let the user select a column to filter
        selected_column = st.selectbox("Select a column to filter", headers)

        # Step 2: Get unique values from the selected column for filtering
        if selected_column:
            unique_values = df[selected_column].unique()
            selected_value = st.selectbox(
                f"Select a value to filter by in '{selected_column}'", unique_values
            )

            # Step 3: Filter the DataFrame based on the selected value
            if selected_value:
                filtered_df = df[df[selected_column] == selected_value]
                st.write(f"### Filtered Data for {selected_column} = {selected_value}")
                st.write(filtered_df)

                # Step 4: Provide a downloadable Excel file with the filtered data
                st.write("### Download Filtered Data")
                output_file = "filtered_data.xlsx"
                filtered_df.to_excel(output_file, index=False)
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Filtered Excel File",
                        data=file,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()
