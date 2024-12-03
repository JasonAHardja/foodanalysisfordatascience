import pandas as pd

def process_food_data(file_path, category):

    # Read the CSV file with semicolon delimiter, skipping 2 rows, and handling encoding
    df = pd.read_csv(file_path, engine='python', encoding='ISO-8859-1', sep=";", skiprows=2, on_bad_lines='skip')

    # Normalize column names: lowercase and strip whitespace
    df.columns = df.columns.str.strip().str.lower()

    # Define the expected column names
    menu_category_col = "menu category"
    menu_col = "menu"
    qty_col = "qty"
    total_col = "total"
    nett_sales_col = "nett sales"

    # Ensure the expected columns are present
    if menu_category_col not in df.columns or menu_col not in df.columns or qty_col not in df.columns or total_col not in df.columns or nett_sales_col not in df.columns:
        raise ValueError("The required columns 'Menu Category', 'Menu', 'Qty', 'Total', or 'Nett Sales' were not found in the CSV file.")

    # Filter data based on the given Menu Category and create a copy
    df_filtered = df.loc[df[menu_category_col].str.upper() == category].copy()

    # Store the original "Menu" names
    df_filtered["original_menu"] = df_filtered[menu_col]

    # Convert "Total" and "Nett Sales" columns to numeric, handling the European format
    df_filtered[total_col] = pd.to_numeric(df_filtered[total_col].str.replace(".", "").str.replace(",", "."), errors='coerce').fillna(0)
    df_filtered[nett_sales_col] = pd.to_numeric(df_filtered[nett_sales_col].str.replace(".", "").str.replace(",", "."), errors='coerce').fillna(0)

    # Calculate the Profit as Total - Nett Sales
    df_filtered["profit"] = df_filtered[total_col] - df_filtered[nett_sales_col]

    # Format the Profit column in the same format as "Total" and "Nett Sales"
    df_filtered["formatted_profit"] = df_filtered["profit"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Group by the original menu name and sum the "Qty" and "Profit" columns
    grouped_df = df_filtered.groupby("original_menu")[[qty_col, "profit"]].sum().reset_index()

    # Format the summed profit
    grouped_df["profit"] = grouped_df["profit"].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Sort by "Qty" in descending order and get only the top 3 results
    result = grouped_df.sort_values(by=qty_col, ascending=False).head(3).reset_index(drop=True)

    # Rename the columns as requested
    result = result.rename(columns={
        "original_menu": "menu",
        qty_col: "order amount",
    })

    # Return only the requested columns: "menu", "order amount", and "profit"
    return result[["menu", "order amount", "profit"]]

def save_to_csv(dataframe, output_file):
    
    dataframe.to_csv(output_file, index=False)
    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    # Hardcode the file path
    #windows or mac, make sure you change it to ur designated file path. example = your username
    file_path = "/Users/jasonhardjawidjaja/Desktop/4 Sales Recapitulation Detail Report april 2024.csv"

    try:
        # Process the data for "MAKANAN" and "MINUMAN" categories
        top_foods_summary = process_food_data(file_path, "MAKANAN")
        top_drinks_summary = process_food_data(file_path, "MINUMAN")

        # Define output CSV file names
        makanan_csv = "Top_3_Makanan.csv"
        minuman_csv = "Top_3_Minuman.csv"

        # Save the results to CSV files
        save_to_csv(top_foods_summary, makanan_csv)
        save_to_csv(top_drinks_summary, minuman_csv)

        # Display a success message
        print(f"\nTop 3 MAKANAN results saved to: {makanan_csv}")
        print(f"Top 3 MINUMAN results saved to: {minuman_csv}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")
