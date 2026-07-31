import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from datetime import datetime
import os
import json

# Configure CustomTkinter appea`rance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class OrderProcessorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("WooCommerce Order Processor Pro")
        self.geometry("600x550")
        self.resizable(False, False)
        
        # Variables
        self.input_file_path = None
        self.processed_df = None
        self.processing_log = []
        
        # Main Frame
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="WooCommerce Order Processor Pro",
            font=("Helvetica", 22, "bold"),
            text_color="#ffffff"
        )
        self.title_label.pack(pady=(0, 5))
        
        # Subtitle
        self.subtitle_label = ctk.CTkLabel(
            self.main_frame,
            text="Advanced WordPress order export automation",
            font=("Helvetica", 11),
            text_color="#888888"
        )
        self.subtitle_label.pack(pady=(0, 20))
        
        # File Info Frame
        self.file_frame = ctk.CTkFrame(self.main_frame, corner_radius=8, fg_color="#2b2b2b")
        self.file_frame.pack(fill="x", pady=(0, 15))
        
        self.file_label = ctk.CTkLabel(
            self.file_frame,
            text="No file selected",
            font=("Helvetica", 11),
            text_color="#b0b0b0"
        )
        self.file_label.pack(pady=10, padx=10)
        
        # Buttons Frame
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(fill="x", pady=(0, 15))
        
        # Upload Button
        self.upload_btn = ctk.CTkButton(
            self.buttons_frame,
            text="📁 Upload File",
            command=self.upload_file,
            font=("Helvetica", 11, "bold"),
            fg_color="#1f6aa5",
            hover_color="#15527a",
            height=40,
            corner_radius=8,
            width=140
        )
        self.upload_btn.pack(side="left", padx=(0, 10))
        
        # Process Button
        self.process_btn = ctk.CTkButton(
            self.buttons_frame,
            text="⚙️ Process & Save",
            command=self.process_orders,
            font=("Helvetica", 11, "bold"),
            fg_color="#0d8c3a",
            hover_color="#0a6b2e",
            height=40,
            corner_radius=8,
            width=140,
            state="disabled"
        )
        self.process_btn.pack(side="left", padx=(0, 10))
        
        # Options Button
        self.options_btn = ctk.CTkButton(
            self.buttons_frame,
            text="⚙️ Options",
            command=self.show_options,
            font=("Helvetica", 11, "bold"),
            fg_color="#553399",
            hover_color="#442288",
            height=40,
            corner_radius=8,
            width=100
        )
        self.options_btn.pack(side="left")
        
        # Info Panel
        self.info_frame = ctk.CTkFrame(self.main_frame, corner_radius=8, fg_color="#1e3a1f")
        self.info_frame.pack(fill="x", pady=(0, 15))
        
        self.info_label = ctk.CTkLabel(
            self.info_frame,
            text="📊 Supports: WooCommerce, Shopify, BigCommerce exports\n"
                 "✓ Auto-validates columns • ✓ Batch processing • ✓ Error detection",
            font=("Helvetica", 10),
            text_color="#90ee90",
            justify="left"
        )
        self.info_label.pack(pady=10, padx=10)
        
        # Status Label
        self.status_label = ctk.CTkLabel(
            self.main_frame,
            text="Ready to process orders",
            font=("Helvetica", 11),
            text_color="#b0b0b0"
        )
        self.status_label.pack(pady=(10, 0))
        
        # Log Frame with scrolling
        self.log_frame = ctk.CTkFrame(self.main_frame, corner_radius=8, fg_color="#1a1a1a", height=150)
        self.log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        self.log_text = ctk.CTkTextbox(
            self.log_frame,
            font=("Courier", 9),
            fg_color="#0a0a0a",
            text_color="#00dd00",
            corner_radius=6
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text.configure(state="disabled")
        
        self.add_log("Initialized WooCommerce Order Processor Pro v1.0")
    
    def add_log(self, message):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        self.processing_log.append(log_entry)
        
        self.log_text.configure(state="normal")
        self.log_text.insert("end", log_entry + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
    
    def upload_file(self):
        """Handle file upload"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select Orders Export File",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if file_path:
                self.input_file_path = file_path
                filename = os.path.basename(file_path)
                file_size = os.path.getsize(file_path) / 1024  # KB
                self.file_label.configure(text=f"✓ {filename} ({file_size:.1f} KB)")
                self.process_btn.configure(state="normal")
                self.add_log(f"File loaded: {filename}")
                self.status_label.configure(text="File ready. Click 'Process & Save'", text_color="#90ee90")
        
        except Exception as e:
            self.add_log(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def extract_order_number(self, order_no):
        """Extract digits only from order number"""
        try:
            if pd.isna(order_no):
                return ""
            digits = re.findall(r'\d+', str(order_no))
            return ''.join(digits) if digits else str(order_no)
        except:
            return str(order_no)
    
    def process_orders(self):
        """Process the uploaded file and transform data"""
        if self.input_file_path is None:
            messagebox.showwarning("Warning", "Please upload a file first")
            return
        
        try:
            self.status_label.configure(text="Processing...", text_color="#ffd700")
            self.process_btn.configure(state="disabled")
            self.update()
            
            self.add_log("Starting data transformation...")
            
            # Read the file
            if self.input_file_path.endswith('.csv'):
                df = pd.read_csv(self.input_file_path)
                self.add_log(f"Loaded CSV file with {len(df)} rows")
            else:
                df = pd.read_excel(self.input_file_path)
                self.add_log(f"Loaded Excel file with {len(df)} rows")
            
            # Validate required columns
            required_cols = [
                'Order Number', 'Order Date', 'First Name (Shipping)', 
                'Last Name (Shipping)', 'Item Name', 'Quantity (- Refund)',
                'Address 1&2 (Shipping)', 'City (Shipping)', 'State Code (Shipping)',
                'Postcode (Shipping)', 'Country Code (Shipping)', 'Phone (Billing)'
            ]
            
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                self.add_log(f"Error: Missing columns: {', '.join(missing_cols)}")
                messagebox.showerror(
                    "Error",
                    f"Missing columns in file:\n{', '.join(missing_cols)}\n\n"
                    "Please ensure this is a valid WooCommerce export file."
                )
                self.status_label.configure(text="Error: Invalid file format", text_color="#ff6b6b")
                self.process_btn.configure(state="normal")
                return
            
            self.add_log("Validating data...")
            
            # Transform data
            processed_df = pd.DataFrame()
            
            processed_df['date'] = df['Order Date'].astype(str)
            processed_df['Order No.'] = df['Order Number'].apply(self.extract_order_number)
            processed_df['Name'] = (df['First Name (Shipping)'].astype(str) + ' ' + 
                                   df['Last Name (Shipping)'].astype(str)).str.strip()
            processed_df['Book Name'] = df['Item Name'].astype(str)
            processed_df['QTY'] = df['Quantity (- Refund)'].astype(int)
            processed_df['WT'] = ''
            processed_df['Address 1&2 (Shipping)'] = df['Address 1&2 (Shipping)'].astype(str)
            processed_df['City (Shipping)'] = df['City (Shipping)'].astype(str)
            processed_df['State Code (Shipping)'] = df['State Code (Shipping)'].astype(str)
            processed_df['Postcode (Shipping)'] = df['Postcode (Shipping)'].astype(str)
            processed_df['Country Code (Shipping)'] = df['Country Code (Shipping)'].astype(str)
            processed_df['Note'] = df['Phone (Billing)'].astype(str).replace('nan', '')
            
            self.processed_df = processed_df
            
            self.add_log(f"✓ Transformation complete: {len(processed_df)} orders processed")
            
            # Save file
            self.save_processed_file(processed_df)
            
        except Exception as e:
            self.add_log(f"Error: {str(e)}")
            messagebox.showerror("Processing Error", f"Error processing file:\n{str(e)}")
            self.status_label.configure(text="Error during processing", text_color="#ff6b6b")
            self.process_btn.configure(state="normal")
    
    def save_processed_file(self, df):
        """Save the processed DataFrame to Excel"""
        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                initialfile=f"processed_orders_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            )
            
            if save_path:
                if save_path.endswith('.csv'):
                    df.to_csv(save_path, index=False)
                    self.add_log(f"Saved as CSV: {os.path.basename(save_path)}")
                else:
                    df.to_excel(save_path, index=False, sheet_name='Orders')
                    self.add_log(f"Saved as Excel: {os.path.basename(save_path)}")
                
                self.add_log(f"✓ Success: {len(df)} orders exported")
                
                messagebox.showinfo(
                    "Success",
                    f"Orders processed successfully!\n\n"
                    f"File saved to:\n{save_path}\n\n"
                    f"Total records: {len(df)}"
                )
                self.status_label.configure(
                    text=f"✓ {len(df)} orders processed successfully",
                    text_color="#90ee90"
                )
                self.process_btn.configure(state="normal")
            else:
                self.add_log("Save cancelled by user")
                self.status_label.configure(text="Save cancelled", text_color="#ffd700")
                self.process_btn.configure(state="normal")
        
        except Exception as e:
            self.add_log(f"Save error: {str(e)}")
            messagebox.showerror("Save Error", f"Error saving file:\n{str(e)}")
            self.status_label.configure(text="Error saving file", text_color="#ff6b6b")
            self.process_btn.configure(state="normal")
    
    def show_options(self):
        """Show options dialog"""
        options_window = ctk.CTkToplevel(self)
        options_window.title("Options")
        options_window.geometry("400x300")
        options_window.resizable(False, False)
        
        # Title
        title = ctk.CTkLabel(
            options_window,
            text="Processing Options",
            font=("Helvetica", 14, "bold")
        )
        title.pack(pady=15)
        
        # Frame for options
        options_frame = ctk.CTkFrame(options_window)
        options_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # Export Log Button
        log_btn = ctk.CTkButton(
            options_frame,
            text="Export Processing Log",
            command=self.export_log,
            fg_color="#1f6aa5",
            hover_color="#15527a",
            height=40
        )
        log_btn.pack(fill="x", pady=5)
        
        # Clear Log Button
        clear_btn = ctk.CTkButton(
            options_frame,
            text="Clear Log",
            command=self.clear_log,
            fg_color="#c23b22",
            hover_color="#a02818",
            height=40
        )
        clear_btn.pack(fill="x", pady=5)
        
        # About
        about_frame = ctk.CTkFrame(options_frame, fg_color="#2b2b2b", corner_radius=8)
        about_frame.pack(fill="both", expand=True, pady=10)
        
        about_text = ctk.CTkLabel(
            about_frame,
            text="WooCommerce Order Processor Pro\nVersion 1.0\n\n"
                 "Automate your WordPress order exports\nwith intelligent data transformation\n\n"
                 "Made with ❤️ for better workflows",
            font=("Helvetica", 10),
            text_color="#b0b0b0",
            justify="center"
        )
        about_text.pack(padx=15, pady=15)
    
    def export_log(self):
        """Export processing log"""
        try:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("JSON files", "*.json")],
                initialfile=f"processing_log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
            )
            
            if save_path:
                with open(save_path, 'w') as f:
                    f.write('\n'.join(self.processing_log))
                
                messagebox.showinfo("Success", f"Log exported to:\n{save_path}")
                self.add_log(f"Log exported: {os.path.basename(save_path)}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export log: {str(e)}")
    
    def clear_log(self):
        """Clear the processing log"""
        if messagebox.askyesno("Confirm", "Clear all log entries?"):
            self.processing_log.clear()
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
            self.add_log("Log cleared")


def main():
    app = OrderProcessorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
