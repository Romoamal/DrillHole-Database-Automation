import pandas as pd

def update_database(database_file, daily_data_file):
    # Load database and daily data files
    database = pd.read_excel(database_file)
    daily_data = pd.read_excel(daily_data_file, header=None)  # Read without assuming a header row

    # Identify the header row dynamically
    header_row_idx = daily_data[daily_data.astype(str).apply(lambda row: row.str.contains("Hole ID", na=False)).any(axis=1)].index[0]
    
    # Extract the actual data starting from the next row
    daily_data.columns = daily_data.iloc[header_row_idx]  # Set proper column names
    daily_data = daily_data.iloc[header_row_idx + 1:].reset_index(drop=True)  # Remove header row from data

    # Rename columns based on standard database format (if necessary)
    column_mapping = {
        "Date from": "Date Logging",
        "Hole ID": "Hole ID",
        "FROM": "From",
        "TO": "To",
        "INTERVAL (M)": "Length",
        "ACT CORE (M)": "Actual Core",
        "RECOVERY (%)": "Recovery percentage",
        "GENERAL LITHOLOGY": "Material Code",
        "SUB GEN LITHOLOGY": "Layer Code",
        "ROCK CODE": "Rock Code",
        "GRAIN SIZE": "Grain",
        "SERPENTINIZED": "Serpent",
        "WEATHERING": "Weath",
        "PRIMARY": "Colour",
        "SECONDARY": "Structure Pri",
        "TERTIARY": "Structure Sec",
        "PRIMARY.1": "Minerals Pri",
        "SECONDARY.1": "Minerals Sec",
        "TERTIARY.1": "Minerals Ter",
        "DENSITY": "No Fract",
        "REMARKS": "Comment"
    }
    daily_data.rename(columns=column_mapping, inplace=True)

    # Ensure all required columns exist in daily_data
    for col in database.columns:
        if col not in daily_data.columns:
            daily_data[col] = None  # Fill missing columns with empty values

    # Append cleaned data to the existing database
    updated_database = pd.concat([database, daily_data], ignore_index=True)

    # Save updates back to the same database file
    updated_database.to_excel(database_file, index=False)

    print(f"Database successfully updated in {database_file}.")

# Example usage:
update_database("drilling_database.xlsx", "C06-090.xlsx")
