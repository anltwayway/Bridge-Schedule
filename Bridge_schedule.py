import pandas as pd
import math
import tkinter as tk
from tkinter import simpledialog

# Step 1: Read excavation data
def read_excavation_data(filename):
    df = pd.read_excel(filename)
    return df

# Step 2: Find bridge sections
def find_bridge_sections(df, threshold):
    bridge_sections = []
    current_bridge_start = None
    current_bridge_end = None
    
    for index, row in df.iterrows():
        if row['填挖高'] > threshold:
            if current_bridge_start is None:
                current_bridge_start = row['里程']
            current_bridge_end = row['里程']
        else:
            if current_bridge_start is not None:
                bridge_sections.append((current_bridge_start, current_bridge_end))
                current_bridge_start = None
                current_bridge_end = None
    
    return bridge_sections

# Step 3: Generate the bridge table
def generate_bridge_table(bridge_sections):
    bridge_table = pd.DataFrame(columns=['Start Chainage', 'Chainage', 'End Chainage', 'Length (30m units)'])

    for start, end in bridge_sections:
        start_int = int(start)
        end_int = int(end)
        
        # Calculate the center chainage
        center_chainage = (start_int + end_int) / 2
        
        # Adjust the start and end chainage to be multiples of 30m
        adjusted_start_chainage = math.floor(start_int / 30) * 30
        adjusted_end_chainage = math.ceil(end_int / 30) * 30
        
        # If the difference between end and start chainage is not a multiple of 30m, add 1 bridge
        num_bridges = math.ceil((adjusted_end_chainage - adjusted_start_chainage) / 30)
        
        # Update center chainage
        center_chainage = (adjusted_start_chainage + adjusted_end_chainage) / 2
        
        bridge_data = {
            'Start Chainage': [adjusted_start_chainage],
            'Chainage': [center_chainage],
            'End Chainage': [adjusted_end_chainage],
            'Length (30m units)': [num_bridges]
        }
        bridge_table = pd.concat([bridge_table, pd.DataFrame(bridge_data)], ignore_index=True)
    
    # Remove rows where Length (30m units) is 0
    bridge_table = bridge_table[bridge_table['Length (30m units)'] != 0]
    
    return bridge_table

# Input filename
excavation_filename = "填挖高度表.xlsx"

# Create Tkinter root window
root = tk.Tk()
root.withdraw() # Hide the root window

# Prompt user to input threshold for bridge excavation height
threshold_input = simpledialog.askfloat("桥梁填挖高度的阈值", "请输入桥梁填挖高度的阈值（单位：米，默认为12）：", initialvalue=12)

if threshold_input is not None:
    threshold = threshold_input
else:
    threshold = 12

# Read excavation data
excavation_data = read_excavation_data(excavation_filename)

# Find bridge sections
bridge_sections = find_bridge_sections(excavation_data, threshold)

# Generate bridge table
bridge_table = generate_bridge_table(bridge_sections)

# Export bridge table to Excel file
bridge_table.to_excel("桥梁表.xlsx", index=False)
