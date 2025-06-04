import pandas as pd
import os

# Define the path to the data folder and Excel file
current_dir = os.path.dirname(os.path.abspath(__file__))
data_folder = os.path.join(current_dir, 'data')
excel_file = os.path.join(data_folder, 'SAA-DBL-MergeData.xlsx')
# Output location for processed data
output_file = os.path.join(current_dir, 'temp', 'processed_data.pkl')

def load_excel_data():
    """
    Load data from the Excel file into a pandas DataFrame
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        # Display basic information about the DataFrame
        print(f"Successfully loaded data from {excel_file}")
        print(f"DataFrame shape: {df.shape}")
        print("DataFrame columns:")
        for col in df.columns:
            print(f"- {col}")
        
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

if __name__ == "__main__":
    # Make temp directory if it doesn't exist
    os.makedirs(os.path.join(current_dir, 'temp'), exist_ok=True)
    
    # Execute when the script is run directly
    df = load_excel_data()
    
    if df is not None:
        # Display the first 5 rows of the DataFrame
        print("\nFirst 5 rows of data:")
        print(df.head())
        
        # Save the processed data to a pickle file
        df.to_pickle(output_file)
        print(f"Processed data saved to {output_file}")