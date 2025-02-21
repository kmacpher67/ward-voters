#!/usr/bin/env python3
import pandas as pd
import math

def main():
    # Input CSV file
    input_file = "CityOfWarren2025-02-06-target-googlemaps.csv"
    chunk_size = 2000

    # Read the CSV file into a DataFrame
    df = pd.read_csv(input_file)

    # Group the DataFrame by the "WARD" column
    for ward, ward_df in df.groupby('WARD'):
        # Reset index so that slicing is straightforward
        ward_df = ward_df.reset_index(drop=True)
        total_rows = len(ward_df)
        # Calculate the number of files needed for this ward
        num_files = math.ceil(total_rows / chunk_size)
        
        for i in range(num_files):
            start_index = i * chunk_size
            end_index = (i + 1) * chunk_size
            # Slice the DataFrame for the current chunk
            sub_df = ward_df.iloc[start_index:end_index]
            
            # Calculate row numbers for naming (1-indexed)
            row_start = start_index + 1
            row_end = min(end_index, total_rows)
            
            # Construct the output file name
            output_file = (
                f"CityOfWarren2025-02-06-target-googlemaps-{ward}"
                f"-Rows{row_start}-{row_end}.csv"
            )
            
            # Save the chunk to a new CSV file without the index
            sub_df.to_csv(output_file, index=False)
            print(f"Saved {output_file}")

if __name__ == '__main__':
    main()
