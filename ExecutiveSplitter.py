import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create the main application window
root = tk.Tk()
root.title("Data Analysis App")

# Define grouped and customer_df as global variables
grouped = None
customer_df = None

# Function to open and analyze the Excel file with filtering
def analyze_data():
    global grouped, customer_df  # Declare grouped and customer_df as global variables
    excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if excel_file:
        # Read the Excel file into a DataFrame and perform analysis
        df = pd.read_excel(excel_file, engine='openpyxl')
        
        # Get the filter condition from the user
        filter_text = filter_entry.get()
        
        # Get the date range from the user
        start_date = pd.to_datetime(start_date_entry.get(), format='%Y-%m-%d', errors='coerce')
        end_date = pd.to_datetime(end_date_entry.get(), format='%Y-%m-%d', errors='coerce')
        
        # Filter the data based on the 'Ramo' column and date range
        filtered_df = df[(df['Ramo'] == filter_text) & 
                          (df['POL_FC'] >= start_date) & (df['POL_FC'] <= end_date)]
        
        # Group the filtered data by 'EJE_Nombre_Asignado' and concatenate customer information
        grouped = filtered_df.groupby('EJE_Nombre_Asignado')[['POL_Nombre Completo', 'POL_PNActual', 
                                                              'Ramo', 'POL_FC']].agg(list)
        
        # Create a list to store individual customer records
        customer_records = []
        
        for seller, data in grouped.iterrows():
            result_text.insert(tk.END, f"Seller: {seller}\n")
            for customer, prima, ramo, fecha in zip(data['POL_Nombre Completo'], data['POL_PNActual'], 
                                                   data['Ramo'], data['POL_FC']):
                formatted_line = f"  {customer} | Prima: {prima}  |  Ramo: {ramo}  |  Fecha: {fecha} \n"
                result_text.insert(tk.END, formatted_line)
                
                # Append the customer record to the list
                customer_records.append([customer, prima, ramo, fecha])
            
            total_prima = sum(data['POL_PNActual'])
            result_text.insert(tk.END, f" \n  Total Prima for Seller {seller}: {total_prima}\n \n")
        
        # Create a DataFrame from the list of customer records
        customer_df = pd.DataFrame(customer_records, columns=['Customer', 'Prima', 'Ramo', 'Fecha'])

# Function to save the analyzed data as an Excel file
# Function to save the analyzed data as an Excel file
def save_as_excel():
    if grouped is not None:
        save_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_file:
            # Create an ExcelWriter object to write to the same file
            with pd.ExcelWriter(save_file, engine='xlsxwriter') as writer:
                # Iterate through each seller's data and save it to a separate worksheet
                for seller, data in grouped.iterrows():
                    data_df = pd.DataFrame(list(zip(data['POL_Nombre Completo'], data['POL_PNActual'], 
                                                     data['Ramo'], data['POL_FC'])),
                                            columns=['Customer', 'Prima', 'Ramo', 'Fecha'])
                    
                    # Write the data to a separate worksheet with the seller's name as the sheet name
                    data_df.to_excel(writer, sheet_name=seller, index=False)



# Create a label and an entry for the filter condition
filter_label = tk.Label(root, text="Filter by Ramo:")
filter_label.pack()
filter_entry = tk.Entry(root)
filter_entry.pack()

# Create labels and entries for the date range
start_date_label = tk.Label(root, text="Start Date (YYYY-MM-DD):")
start_date_label.pack()
start_date_entry = tk.Entry(root)
start_date_entry.pack()

end_date_label = tk.Label(root, text="End Date (YYYY-MM-DD):")
end_date_label.pack()
end_date_entry = tk.Entry(root)
end_date_entry.pack()

# Create a button to trigger the data analysis with filtering
analyze_button = tk.Button(root, text="Analyze Data with Filter", command=analyze_data)
analyze_button.pack()

# Create a button to save the data as an Excel file
save_button = tk.Button(root, text="Save as Excel", command=save_as_excel)
save_button.pack()

# Create a text widget to display the results with a larger font
result_text = tk.Text(root, font=("Helvetica", 12))  # Adjust font size here
result_text.pack()

# Start the GUI event loop
root.mainloop()
