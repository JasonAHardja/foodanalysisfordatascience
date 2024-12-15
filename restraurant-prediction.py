import pandas as pd
import numpy as np
from datetime import datetime

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

    # Format the Profit column
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

    return result[["menu", "order amount", "profit"]], df_filtered

def predict_monthly_stock(df):
    """
    Predicts needed stock based on monthly trends.
    Adds 25% safety stock to the prediction.
    """
    # Get the date column (assuming it's named 'transaction date' or similar)
    date_col = next(col for col in df.columns if 'date' in col.lower())
    
    # Convert to datetime with dayfirst=True
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    
    # Extract month and group by month and menu
    df['month'] = df[date_col].dt.month
    monthly_sales = df.groupby(['month', 'original_menu'])['qty'].sum().reset_index()
    
    # Calculate predictions for each menu item
    predictions = []
    for menu in monthly_sales['original_menu'].unique():
        menu_data = monthly_sales[monthly_sales['original_menu'] == menu].sort_values('month')
        
        if len(menu_data) > 0:
            quantities = menu_data['qty'].tolist()
            changes = [0]  # First month has no change
            
            # Calculate monthly changes
            for i in range(1, len(quantities)):
                change = quantities[i] - quantities[i-1]
                changes.append(change)
            
            # Calculate needed stock (last month's quantity + 25% safety stock)
            last_month_qty = quantities[-1]
            needed_stock = int(last_month_qty * 1.25)
            
            predictions.append({
                'menu': menu,
                'monthly_quantities': quantities,
                'monthly_changes': changes,
                'last_month_quantity': last_month_qty,
                'recommended_stock': needed_stock
            })
    
    return pd.DataFrame(predictions)


def save_to_csv(dataframe, output_file):
    dataframe.to_csv(output_file, index=False)
    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    # File path
    file_path = "/Users/jasonhardjawidjaja/Desktop/4 Sales Recapitulation Detail Report april 2024.csv"

    try:
        # Process top 3 items and get filtered dataframes
        top_foods_summary, foods_df = process_food_data(file_path, "MAKANAN")
        top_drinks_summary, drinks_df = process_food_data(file_path, "MINUMAN")

        # Save top 3 summaries
        save_to_csv(top_foods_summary, "Top_3_Makanan.csv")
        save_to_csv(top_drinks_summary, "Top_3_Minuman.csv")

        # Generate and save predictions
        food_predictions = predict_monthly_stock(foods_df)
        drink_predictions = predict_monthly_stock(drinks_df)
        
        save_to_csv(food_predictions, "Makanan_Predictions.csv")
        save_to_csv(drink_predictions, "Minuman_Predictions.csv")

        print("\nAnalysis complete! Generated files:")
        print("1. Top_3_Makanan.csv - Top 3 food items")
        print("2. Top_3_Minuman.csv - Top 3 drink items")
        print("3. Makanan_Predictions.csv - Monthly trends and stock recommendations")
        print("4. Minuman_Predictions.csv - Monthly trends and stock recommendations")

    except Exception as e:
        print(f"An error occurred: {str(e)}")