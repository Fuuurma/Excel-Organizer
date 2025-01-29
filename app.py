import pandas as pd
import streamlit as st


def load_excel(file, header_row):
    # Use the specified header row (subtract 1 because pandas uses 0-based indexing)
    df = pd.read_excel(file, header=header_row - 1)
    return df


def main():
    st.title("Excel Organizer")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    if uploaded_file is not None:
        # Step 1: Let the user specify the header row
        header_row = st.number_input(
            "Enter the row number where headers start (e.g., 3)",
            min_value=1,
            value=1,  # Default to row 1
            step=1,
        )

        # Load the Excel file with the specified header row
        df = load_excel(uploaded_file, header_row)

        # Step 2: Display the headers and ask the user if they are correct
        st.write("### Detected Headers")
        headers = df.columns.tolist()
        st.write(headers)

        # Ask the user to confirm the headers
        confirm_headers = st.radio(
            "Are these the correct headers?",
            options=["Yes", "No"],
            index=0,  # Default to "Yes"
        )

        if confirm_headers == "No":
            st.warning("Please adjust the header row number above and reload the file.")
            return  # Stop execution if headers are not confirmed

        # Step 3: Proceed with filtering if headers are confirmed
        st.write("### Preview of the Data")
        st.write(df.head())

        # Step 4: Let the user select a column to filter
        selected_column = st.selectbox("Select a column to filter", headers)

        # Step 5: Get unique values from the selected column for filtering
        if selected_column:
            unique_values = df[selected_column].unique()
            selected_value = st.selectbox(
                f"Select a value to filter by in '{selected_column}'", unique_values
            )

            # Step 6: Filter the DataFrame based on the selected value
            if selected_value:
                filtered_df = df[df[selected_column] == selected_value]
                st.write(f"### Filtered Data for {selected_column} = {selected_value}")
                st.write(filtered_df)

                # Step 7: Provide a downloadable Excel file with the filtered data
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
