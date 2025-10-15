import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from datetime import datetime
import hashlib
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTk
import numpy as np

class InventoryManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Management System")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize data files
        self.init_data_files()
        
        # Current user
        self.current_user = None
        
        # Show login screen
        self.show_login()
    
    def init_data_files(self):
        """Initialize Excel files if they don't exist"""
        # Products file
        if not os.path.exists('products.xlsx'):
            df = pd.DataFrame(columns=['SKU', 'Product_Name', 'Category', 'Price', 'Cost', 'Quantity', 'Supplier', 'Min_Stock'])
            df.to_excel('products.xlsx', index=False)
        
        # Invoices file
        if not os.path.exists('invoices.xlsx'):
            df = pd.DataFrame(columns=['Invoice_ID', 'Date', 'Customer_Name', 'Items', 'Total_Amount', 'Payment_Type'])
            df.to_excel('invoices.xlsx', index=False)
        
        # Users file
        if not os.path.exists('users.xlsx'):
            # Create default admin user
            password_hash = hashlib.sha256('admin123'.encode()).hexdigest()
            df = pd.DataFrame([['admin', password_hash, 'Admin']], columns=['Username', 'Password', 'Role'])
            df.to_excel('users.xlsx', index=False)
        
        # Create folders for images and barcodes
        os.makedirs('images', exist_ok=True)
        os.makedirs('barcodes', exist_ok=True)
    
    def show_login(self):
        """Display login screen"""
        self.clear_screen()
        
        login_frame = tk.Frame(self.root, bg='white', padx=50, pady=50)
        login_frame.place(relx=0.5, rely=0.5, anchor='center')
        
        tk.Label(login_frame, text="Inventory Management System", font=('Arial', 24, 'bold'), bg='white').pack(pady=20)
        
        tk.Label(login_frame, text="Username:", font=('Arial', 12), bg='white').pack(pady=5)
        self.username_entry = tk.Entry(login_frame, font=('Arial', 12), width=20)
        self.username_entry.pack(pady=5)
        
        tk.Label(login_frame, text="Password:", font=('Arial', 12), bg='white').pack(pady=5)
        self.password_entry = tk.Entry(login_frame, font=('Arial', 12), width=20, show='*')
        self.password_entry.pack(pady=5)
        
        tk.Button(login_frame, text="Login", command=self.login, bg='#4CAF50', fg='white', 
                 font=('Arial', 12), width=15, pady=5).pack(pady=20)
        
        # Bind Enter key to login
        self.root.bind('<Return>', lambda e: self.login())
    
    def login(self):
        """Handle user login"""
        username = self.username_entry.get()
        password = self.password_entry.get()
        
        if not username or not password:
            messagebox.showerror("Error", "Please enter both username and password")
            return
        
        try:
            users_df = pd.read_excel('users.xlsx')
            password_hash = hashlib.sha256(password.encode()).hexdigest()
            
            user = users_df[(users_df['Username'] == username) & (users_df['Password'] == password_hash)]
            
            if not user.empty:
                self.current_user = {'username': username, 'role': user.iloc[0]['Role']}
                self.show_main_interface()
            else:
                messagebox.showerror("Error", "Invalid username or password")
        except Exception as e:
            messagebox.showerror("Error", f"Login failed: {str(e)}")
    
    def show_main_interface(self):
        """Display main application interface"""
        self.clear_screen()
        
        # Create main frame
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True)
        
        # Title
        title_frame = tk.Frame(main_frame, bg='#2196F3', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        tk.Label(title_frame, text=f"Inventory Management System - Welcome {self.current_user['username']}", 
                font=('Arial', 18, 'bold'), bg='#2196F3', fg='white').pack(pady=15)
        
        # Logout button
        tk.Button(title_frame, text="Logout", command=self.show_login, bg='#f44336', fg='white',
                 font=('Arial', 10)).place(relx=0.95, rely=0.5, anchor='center')
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_dashboard_tab()
        self.create_products_tab()
        self.create_stock_tab()
        self.create_billing_tab()
        self.create_invoices_tab()
        self.create_barcodes_tab()
    
    def create_dashboard_tab(self):
        """Create dashboard tab with summary statistics"""
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text="Dashboard")
        
        # Statistics frame
        stats_frame = tk.Frame(dashboard_frame, bg='white', relief='raised', bd=2)
        stats_frame.pack(fill='x', padx=10, pady=10)
        
        try:
            products_df = pd.read_excel('products.xlsx')
            invoices_df = pd.read_excel('invoices.xlsx')
            
            total_products = len(products_df)
            total_stock = products_df['Quantity'].sum() if not products_df.empty else 0
            low_stock_items = len(products_df[products_df['Quantity'] <= products_df['Min_Stock']]) if not products_df.empty else 0
            total_invoices = len(invoices_df)
            
            # Create statistics display
            stats_data = [
                ("Total Products", total_products, "#4CAF50"),
                ("Total Stock", int(total_stock), "#2196F3"),
                ("Low Stock Items", low_stock_items, "#FF9800"),
                ("Total Invoices", total_invoices, "#9C27B0")
            ]
            
            for i, (label, value, color) in enumerate(stats_data):
                stat_frame = tk.Frame(stats_frame, bg=color, width=200, height=100)
                stat_frame.grid(row=0, column=i, padx=20, pady=20)
                stat_frame.pack_propagate(False)
                
                tk.Label(stat_frame, text=str(value), font=('Arial', 24, 'bold'), 
                        bg=color, fg='white').pack(pady=10)
                tk.Label(stat_frame, text=label, font=('Arial', 12), 
                        bg=color, fg='white').pack()
            
            # Low stock alerts
            if low_stock_items > 0:
                alert_frame = tk.Frame(dashboard_frame, bg='#ffebee', relief='raised', bd=2)
                alert_frame.pack(fill='x', padx=10, pady=10)
                
                tk.Label(alert_frame, text="⚠️ Low Stock Alerts", font=('Arial', 14, 'bold'), 
                        bg='#ffebee', fg='#d32f2f').pack(pady=5)
                
                low_stock_df = products_df[products_df['Quantity'] <= products_df['Min_Stock']]
                for _, row in low_stock_df.iterrows():
                    tk.Label(alert_frame, text=f"{row['Product_Name']} - Only {row['Quantity']} left", 
                            font=('Arial', 10), bg='#ffebee', fg='#d32f2f').pack()
        
        except Exception as e:
            tk.Label(dashboard_frame, text=f"Error loading dashboard: {str(e)}", 
                    font=('Arial', 12), fg='red').pack(pady=20)
    
    def create_products_tab(self):
        """Create products management tab"""
        products_frame = ttk.Frame(self.notebook)
        self.notebook.add(products_frame, text="Products")
        
        # Add product form
        form_frame = tk.Frame(products_frame, bg='white', relief='raised', bd=2)
        form_frame.pack(fill='x', padx=10, pady=10)
        
        tk.Label(form_frame, text="Add/Edit Product", font=('Arial', 14, 'bold'), bg='white').grid(row=0, column=0, columnspan=4, pady=10)
        
        # Form fields
        fields = ['SKU', 'Product Name', 'Category', 'Price', 'Cost', 'Quantity', 'Supplier', 'Min Stock']
        self.product_entries = {}
        
        for i, field in enumerate(fields):
            row = (i // 2) + 1
            col = (i % 2) * 2
            
            tk.Label(form_frame, text=f"{field}:", bg='white').grid(row=row, column=col, padx=10, pady=5, sticky='e')
            entry = tk.Entry(form_frame, width=20)
            entry.grid(row=row, column=col+1, padx=10, pady=5)
            self.product_entries[field.lower().replace(' ', '_')] = entry
        
        # Buttons
        button_frame = tk.Frame(form_frame, bg='white')
        button_frame.grid(row=5, column=0, columnspan=4, pady=10)
        
        tk.Button(button_frame, text="Add Product", command=self.add_product, 
                 bg='#4CAF50', fg='white').pack(side='left', padx=5)
        tk.Button(button_frame, text="Update Product", command=self.update_product, 
                 bg='#2196F3', fg='white').pack(side='left', padx=5)
        tk.Button(button_frame, text="Delete Product", command=self.delete_product, 
                 bg='#f44336', fg='white').pack(side='left', padx=5)
        tk.Button(button_frame, text="Clear Fields", command=self.clear_product_fields, 
                 bg='#FF9800', fg='white').pack(side='left', padx=5)
        
        # Products list
        list_frame = tk.Frame(products_frame)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview for products
        columns = ('SKU', 'Product Name', 'Category', 'Price', 'Cost', 'Quantity', 'Supplier')
        self.products_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.products_tree.heading(col, text=col)
            self.products_tree.column(col, width=100)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.products_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.products_tree.xview)
        self.products_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.products_tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # Bind selection event
        self.products_tree.bind('<<TreeviewSelect>>', self.on_product_select)
        
        # Load products
        self.load_products()
    
    def create_stock_tab(self):
        """Create stock management tab"""
        stock_frame = ttk.Frame(self.notebook)
        self.notebook.add(stock_frame, text="Stock Management")
        
        # Stock adjustment form
        form_frame = tk.Frame(stock_frame, bg='white', relief='raised', bd=2)
        form_frame.pack(fill='x', padx=10, pady=10)
        
        tk.Label(form_frame, text="Stock Adjustment", font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        # SKU selection
        tk.Label(form_frame, text="Select Product (SKU):", bg='white').pack()
        self.stock_sku_var = tk.StringVar()
        self.stock_sku_combo = ttk.Combobox(form_frame, textvariable=self.stock_sku_var, width=30)
        self.stock_sku_combo.pack(pady=5)
        
        # Quantity adjustment
        tk.Label(form_frame, text="Quantity Change (+ for stock in, - for stock out):", bg='white').pack()
        self.stock_qty_entry = tk.Entry(form_frame, width=20)
        self.stock_qty_entry.pack(pady=5)
        
        # Reason
        tk.Label(form_frame, text="Reason:", bg='white').pack()
        self.stock_reason_entry = tk.Entry(form_frame, width=40)
        self.stock_reason_entry.pack(pady=5)
        
        # Buttons
        button_frame = tk.Frame(form_frame, bg='white')
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Stock In", command=lambda: self.adjust_stock('in'), 
                 bg='#4CAF50', fg='white').pack(side='left', padx=5)
        tk.Button(button_frame, text="Stock Out", command=lambda: self.adjust_stock('out'), 
                 bg='#f44336', fg='white').pack(side='left', padx=5)
        
        # Stock history (simplified)
        history_frame = tk.Frame(stock_frame)
        history_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(history_frame, text="Current Stock Levels", font=('Arial', 14, 'bold')).pack()
        
        # Stock tree
        stock_columns = ('SKU', 'Product Name', 'Current Stock', 'Min Stock', 'Status')
        self.stock_tree = ttk.Treeview(history_frame, columns=stock_columns, show='headings', height=15)
        
        for col in stock_columns:
            self.stock_tree.heading(col, text=col)
            self.stock_tree.column(col, width=120)
        
        self.stock_tree.pack(fill='both', expand=True)
        
        # Load stock data
        self.load_stock_data()
        self.update_stock_combo()
    
    def create_billing_tab(self):
        """Create billing/POS tab"""
        billing_frame = ttk.Frame(self.notebook)
        self.notebook.add(billing_frame, text="Billing/POS")
        
        # Create two main sections
        left_frame = tk.Frame(billing_frame, bg='white', relief='raised', bd=2)
        left_frame.pack(side='left', fill='both', expand=True, padx=5, pady=10)
        
        right_frame = tk.Frame(billing_frame, bg='white', relief='raised', bd=2, width=400)
        right_frame.pack(side='right', fill='y', padx=5, pady=10)
        right_frame.pack_propagate(False)
        
        # Left side - Product selection
        tk.Label(left_frame, text="Product Selection", font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        # Search product
        search_frame = tk.Frame(left_frame, bg='white')
        search_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(search_frame, text="Search Product:", bg='white').pack(side='left')
        self.product_search_entry = tk.Entry(search_frame, width=30)
        self.product_search_entry.pack(side='left', padx=5)
        tk.Button(search_frame, text="Search", command=self.search_products).pack(side='left', padx=5)
        
        # Products list for billing
        self.billing_products_tree = ttk.Treeview(left_frame, columns=('SKU', 'Product', 'Price', 'Stock'), 
                                                 show='headings', height=20)
        
        for col in ('SKU', 'Product', 'Price', 'Stock'):
            self.billing_products_tree.heading(col, text=col)
            self.billing_products_tree.column(col, width=100)
        
        self.billing_products_tree.pack(fill='both', expand=True, padx=10, pady=5)
        self.billing_products_tree.bind('<Double-1>', self.add_to_cart)
        
        # Right side - Cart and billing
        tk.Label(right_frame, text="Shopping Cart", font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        # Customer info
        customer_frame = tk.Frame(right_frame, bg='white')
        customer_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(customer_frame, text="Customer:", bg='white').pack(side='left')
        self.customer_entry = tk.Entry(customer_frame, width=20)
        self.customer_entry.pack(side='left', padx=5)
        
        # Cart items
        self.cart_tree = ttk.Treeview(right_frame, columns=('Item', 'Qty', 'Price', 'Total'), 
                                     show='headings', height=10)
        
        for col in ('Item', 'Qty', 'Price', 'Total'):
            self.cart_tree.heading(col, text=col)
            self.cart_tree.column(col, width=80)
        
        self.cart_tree.pack(fill='x', padx=10, pady=5)
        
        # Cart total
        self.total_label = tk.Label(right_frame, text="Total: $0.00", font=('Arial', 16, 'bold'), bg='white')
        self.total_label.pack(pady=10)
        
        # Payment section
        payment_frame = tk.Frame(right_frame, bg='white')
        payment_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(payment_frame, text="Payment Type:", bg='white').pack()
        self.payment_var = tk.StringVar(value="Cash")
        payment_combo = ttk.Combobox(payment_frame, textvariable=self.payment_var, 
                                   values=["Cash", "Card", "Check"], width=15)
        payment_combo.pack(pady=5)
        
        tk.Label(payment_frame, text="Amount Received:", bg='white').pack()
        self.received_entry = tk.Entry(payment_frame, width=15)
        self.received_entry.pack(pady=5)
        
        self.change_label = tk.Label(payment_frame, text="Change: $0.00", font=('Arial', 12, 'bold'), bg='white')
        self.change_label.pack(pady=5)
        
        # Buttons
        button_frame = tk.Frame(right_frame, bg='white')
        button_frame.pack(fill='x', padx=10, pady=10)
        
        tk.Button(button_frame, text="Calculate Change", command=self.calculate_change, 
                 bg='#2196F3', fg='white').pack(fill='x', pady=2)
        tk.Button(button_frame, text="Process Sale", command=self.process_sale, 
                 bg='#4CAF50', fg='white').pack(fill='x', pady=2)
        tk.Button(button_frame, text="Clear Cart", command=self.clear_cart, 
                 bg='#FF9800', fg='white').pack(fill='x', pady=2)
        
        # Initialize cart
        self.cart_items = []
        self.cart_total = 0.0
        
        # Load products for billing
        self.load_billing_products()
    
    def create_invoices_tab(self):
        """Create invoices management tab"""
        invoices_frame = ttk.Frame(self.notebook)
        self.notebook.add(invoices_frame, text="Invoices")
        
        # Invoice list
        tk.Label(invoices_frame, text="Invoice History", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Search frame
        search_frame = tk.Frame(invoices_frame)
        search_frame.pack(fill='x', padx=10, pady=5)
        
        tk.Label(search_frame, text="Search by Customer:").pack(side='left')
        self.invoice_search_entry = tk.Entry(search_frame, width=30)
        self.invoice_search_entry.pack(side='left', padx=5)
        tk.Button(search_frame, text="Search", command=self.search_invoices).pack(side='left', padx=5)
        tk.Button(search_frame, text="Show All", command=self.load_invoices).pack(side='left', padx=5)
        
        # Invoice tree
        invoice_columns = ('Invoice ID', 'Date', 'Customer', 'Total Amount', 'Payment Type')
        self.invoices_tree = ttk.Treeview(invoices_frame, columns=invoice_columns, show='headings', height=20)
        
        for col in invoice_columns:
            self.invoices_tree.heading(col, text=col)
            self.invoices_tree.column(col, width=120)
        
        self.invoices_tree.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Load invoices
        self.load_invoices()
    
    def create_barcodes_tab(self):
        """Create barcode generation tab"""
        barcode_frame = ttk.Frame(self.notebook)
        self.notebook.add(barcode_frame, text="Barcodes")
        
        # Barcode generation form
        form_frame = tk.Frame(barcode_frame, bg='white', relief='raised', bd=2)
        form_frame.pack(fill='x', padx=10, pady=10)
        
        tk.Label(form_frame, text="Barcode Generation", font=('Arial', 14, 'bold'), bg='white').pack(pady=10)
        
        # SKU selection
        tk.Label(form_frame, text="Select Product SKU:", bg='white').pack()
        self.barcode_sku_var = tk.StringVar()
        self.barcode_sku_combo = ttk.Combobox(form_frame, textvariable=self.barcode_sku_var, width=30)
        self.barcode_sku_combo.pack(pady=5)
        
        # Buttons
        button_frame = tk.Frame(form_frame, bg='white')
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="Generate Barcode", command=self.generate_barcode, 
                 bg='#4CAF50', fg='white').pack(side='left', padx=5)
        tk.Button(button_frame, text="View Barcode", command=self.view_barcode, 
                 bg='#2196F3', fg='white').pack(side='left', padx=5)
        
        # Barcode display area
        self.barcode_display_frame = tk.Frame(barcode_frame, bg='white', relief='raised', bd=2)
        self.barcode_display_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Update barcode combo
        self.update_barcode_combo()
    
    def clear_screen(self):
        """Clear all widgets from the screen"""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def add_product(self):
        """Add a new product"""
        try:
            # Get form data
            product_data = {}
            for field, entry in self.product_entries.items():
                value = entry.get().strip()
                if field in ['price', 'cost', 'quantity', 'min_stock']:
                    value = float(value) if value else 0
                product_data[field] = value
            
            # Validate required fields
            if not product_data['sku'] or not product_data['product_name']:
                messagebox.showerror("Error", "SKU and Product Name are required")
                return
            
            # Load existing products
            products_df = pd.read_excel('products.xlsx')
            
            # Check for duplicate SKU
            if product_data['sku'] in products_df['SKU'].values:
                messagebox.showerror("Error", "SKU already exists")
                return
            
            # Add new product
            new_product = pd.DataFrame([{
                'SKU': product_data['sku'],
                'Product_Name': product_data['product_name'],
                'Category': product_data['category'],
                'Price': product_data['price'],
                'Cost': product_data['cost'],
                'Quantity': product_data['quantity'],
                'Supplier': product_data['supplier'],
                'Min_Stock': product_data['min_stock']
            }])
            
            products_df = pd.concat([products_df, new_product], ignore_index=True)
            products_df.to_excel('products.xlsx', index=False)
            
            messagebox.showinfo("Success", "Product added successfully")
            self.clear_product_fields()
            self.load_products()
            self.update_stock_combo()
            self.update_barcode_combo()
            self.load_billing_products()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add product: {str(e)}")
    
    def update_product(self):
        """Update selected product"""
        selected = self.products_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a product to update")
            return
        
        try:
            # Get selected product SKU
            item = self.products_tree.item(selected[0])
            sku = item['values'][0]
            
            # Get form data
            product_data = {}
            for field, entry in self.product_entries.items():
                value = entry.get().strip()
                if field in ['price', 'cost', 'quantity', 'min_stock']:
                    value = float(value) if value else 0
                product_data[field] = value
            
            # Load and update products
            products_df = pd.read_excel('products.xlsx')
            idx = products_df[products_df['SKU'] == sku].index[0]
            
            products_df.loc[idx, 'Product_Name'] = product_data['product_name']
            products_df.loc[idx, 'Category'] = product_data['category']
            products_df.loc[idx, 'Price'] = product_data['price']
            products_df.loc[idx, 'Cost'] = product_data['cost']
            products_df.loc[idx, 'Quantity'] = product_data['quantity']
            products_df.loc[idx, 'Supplier'] = product_data['supplier']
            products_df.loc[idx, 'Min_Stock'] = product_data['min_stock']
            
            products_df.to_excel('products.xlsx', index=False)
            
            messagebox.showinfo("Success", "Product updated successfully")
            self.load_products()
            self.load_stock_data()
            self.load_billing_products()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update product: {str(e)}")
    
    def delete_product(self):
        """Delete selected product"""
        selected = self.products_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a product to delete")
            return
        
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this product?"):
            try:
                # Get selected product SKU
                item = self.products_tree.item(selected[0])
                sku = item['values'][0]
                
                # Load and update products
                products_df = pd.read_excel('products.xlsx')
                products_df = products_df[products_df['SKU'] != sku]
                products_df.to_excel('products.xlsx', index=False)
                
                messagebox.showinfo("Success", "Product deleted successfully")
                self.clear_product_fields()
                self.load_products()
                self.update_stock_combo()
                self.update_barcode_combo()
                self.load_billing_products()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete product: {str(e)}")
    
    def clear_product_fields(self):
        """Clear all product form fields"""
        for entry in self.product_entries.values():
            entry.delete(0, tk.END)
    
    def on_product_select(self, event):
        """Handle product selection in the tree"""
        selected = self.products_tree.selection()
        if selected:
            item = self.products_tree.item(selected[0])
            values = item['values']
            
            # Populate form fields
            fields = ['sku', 'product_name', 'category', 'price', 'cost', 'quantity', 'supplier', 'min_stock']
            for i, field in enumerate(fields):
                if i < len(values):
                    self.product_entries[field].delete(0, tk.END)
                    self.product_entries[field].insert(0, str(values[i]))
    
    def load_products(self):
        """Load products into the tree view"""
        try:
            products_df = pd.read_excel('products.xlsx')
            
            # Clear existing items
            for item in self.products_tree.get_children():
                self.products_tree.delete(item)
            
            # Add products to tree
            for _, row in products_df.iterrows():
                self.products_tree.insert('', 'end', values=(
                    row['SKU'], row['Product_Name'], row['Category'], 
                    f"${row['Price']:.2f}", f"${row['Cost']:.2f}", 
                    int(row['Quantity']), row['Supplier']
                ))
        except Exception as e:
            print(f"Error loading products: {e}")
    
    def adjust_stock(self, operation):
        """Adjust stock levels"""
        try:
            sku = self.stock_sku_var.get().split(' - ')[0].strip()
            qty_input = self.stock_qty_entry.get().strip()

            if not qty_input or qty_input in ['-', '+']:
                messagebox.showerror("Error", "Please enter a valid number (e.g., +3 or -2).")
                return

            try:
                qty_change = int(qty_input)
            except ValueError:
                messagebox.showerror("Error", "Quantity must be a valid integer.")
                return

            if operation == 'out':
                qty_change = -abs(qty_change)
            else:
                qty_change = abs(qty_change)
            
            # Load products
            products_df = pd.read_excel('products.xlsx')
            
            # Find product
            product_idx = products_df[products_df['SKU'] == sku].index
            if product_idx.empty:
                messagebox.showerror("Error", "Product not found")
                return
            
            # Update quantity
            current_qty = products_df.loc[product_idx[0], 'Quantity']
            new_qty = current_qty + qty_change
            
            if new_qty < 0:
                messagebox.showerror("Error", "Insufficient stock")
                return
            
            products_df.loc[product_idx[0], 'Quantity'] = new_qty
            products_df.to_excel('products.xlsx', index=False)
            
            messagebox.showinfo("Success", f"Stock updated. New quantity: {new_qty}")
            
            # Clear fields
            self.stock_qty_entry.delete(0, tk.END)
            self.stock_reason_entry.delete(0, tk.END)
            
            # Refresh displays
            self.load_stock_data()
            self.load_products()
            self.load_billing_products()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to adjust stock: {str(e)}")
    
    def load_stock_data(self):
        """Load stock data into the tree view"""
        try:
            products_df = pd.read_excel('products.xlsx')
            
            # Clear existing items
            for item in self.stock_tree.get_children():
                self.stock_tree.delete(item)
            
            # Add stock data to tree
            for _, row in products_df.iterrows():
                status = "Low Stock" if row['Quantity'] <= row['Min_Stock'] else "OK"
                self.stock_tree.insert('', 'end', values=(
                    row['SKU'], row['Product_Name'], int(row['Quantity']), 
                    int(row['Min_Stock']), status
                ))
                
                # Color code low stock items
                if status == "Low Stock":
                    item_id = self.stock_tree.get_children()[-1]
                    self.stock_tree.set(item_id, 'Status', status)
                    
        except Exception as e:
            print(f"Error loading stock data: {e}")
    
    def update_stock_combo(self):
        """Update stock SKU combo box"""
        try:
            products_df = pd.read_excel('products.xlsx')
            sku_list = [f"{row['SKU']} - {row['Product_Name']}" for _, row in products_df.iterrows()]
            self.stock_sku_combo['values'] = sku_list
        except Exception as e:
            print(f"Error updating stock combo: {e}")
    
    def search_products(self):
        """Search products for billing"""
        search_term = self.product_search_entry.get().lower()
        
        try:
            products_df = pd.read_excel('products.xlsx')
            
            # Clear existing items
            for item in self.billing_products_tree.get_children():
                self.billing_products_tree.delete(item)
            
            # Filter and add products
            for _, row in products_df.iterrows():
                if (search_term in row['Product_Name'].lower() or 
                    search_term in row['SKU'].lower() or
                    search_term in str(row['Category']).lower()):
                    
                    self.billing_products_tree.insert('', 'end', values=(
                        row['SKU'], row['Product_Name'], 
                        f"${row['Price']:.2f}", int(row['Quantity'])
                    ))
        except Exception as e:
            print(f"Error searching products: {e}")
    
    def load_billing_products(self):
        """Load all products for billing"""
        try:
            products_df = pd.read_excel('products.xlsx')
            
            # Clear existing items
            for item in self.billing_products_tree.get_children():
                self.billing_products_tree.delete(item)
            
            # Add all products
            for _, row in products_df.iterrows():
                self.billing_products_tree.insert('', 'end', values=(
                    row['SKU'], row['Product_Name'], 
                    f"${row['Price']:.2f}", int(row['Quantity'])
                ))
        except Exception as e:
            print(f"Error loading billing products: {e}")
    
    def add_to_cart(self, event):
        """Add selected product to cart"""
        selected = self.billing_products_tree.selection()
        if not selected:
            return
        
        item = self.billing_products_tree.item(selected[0])
        values = item['values']
        
        # Get product details
        sku = values[0]
        product_name = values[1]
        price = float(values[2].replace(',', ''))
        available_stock = int(values[3])
        
        if available_stock <= 0:
            messagebox.showerror("Error", "Product out of stock")
            return
        
        # Ask for quantity
        qty_window = tk.Toplevel(self.root)
        qty_window.title("Enter Quantity")
        qty_window.geometry("300x150")
        qty_window.transient(self.root)
        qty_window.grab_set()
        
        tk.Label(qty_window, text=f"Product: {product_name}").pack(pady=10)
        tk.Label(qty_window, text=f"Available: {available_stock}").pack()
        tk.Label(qty_window, text="Quantity:").pack()
        
        qty_entry = tk.Entry(qty_window, width=10)
        qty_entry.pack(pady=5)
        qty_entry.insert(0, "1")
        qty_entry.focus()
        
        def add_item():
            try:
                quantity = int(qty_entry.get())
                if quantity <= 0:
                    messagebox.showerror("Error", "Quantity must be positive")
                    return
                if quantity > available_stock:
                    messagebox.showerror("Error", "Not enough stock")
                    return
                
                # Add to cart
                total = price * quantity
                self.cart_items.append({
                    'sku': sku,
                    'name': product_name,
                    'price': price,
                    'quantity': quantity,
                    'total': total
                })
                
                # Update cart display
                self.update_cart_display()
                qty_window.destroy()
                
            except ValueError:
                messagebox.showerror("Error", "Invalid quantity")
        
        tk.Button(qty_window, text="Add to Cart", command=add_item).pack(pady=10)
        qty_entry.bind('<Return>', lambda e: add_item())
    
    def update_cart_display(self):
        """Update the cart tree view and total"""
        # Clear cart tree
        for item in self.cart_tree.get_children():
            self.cart_tree.delete(item)
        
        # Add cart items
        self.cart_total = 0
        for item in self.cart_items:
            self.cart_tree.insert('', 'end', values=(
                item['name'][:15] + '...' if len(item['name']) > 15 else item['name'],
                item['quantity'],
                f"${item['price']:.2f}",
                f"${item['total']:.2f}"
            ))
            self.cart_total += item['total']
        
        # Update total label
        self.total_label.config(text=f"Total: ${self.cart_total:.2f}")
    
    def calculate_change(self):
        """Calculate change amount"""
        try:
            received = float(self.received_entry.get())
            change = received - self.cart_total
            self.change_label.config(text=f"Change: ${change:.2f}")
        except ValueError:
            self.change_label.config(text="Change: Invalid amount")
    
    def process_sale(self):
        """Process the sale and create invoice"""
        if not self.cart_items:
            messagebox.showerror("Error", "Cart is empty")
            return
        
        customer_name = self.customer_entry.get() or "Walk-in Customer"
        
        try:
            # Generate invoice ID
            invoices_df = pd.read_excel('invoices.xlsx')
            invoice_id = f"INV{len(invoices_df) + 1:04d}"
            
            # Create invoice record
            invoice_data = {
                'Invoice_ID': invoice_id,
                'Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Customer_Name': customer_name,
                'Items': ', '.join([f"{item['name']} x{item['quantity']}" for item in self.cart_items]),
                'Total_Amount': self.cart_total,
                'Payment_Type': self.payment_var.get()
            }
            
            # Add to invoices
            new_invoice = pd.DataFrame([invoice_data])
            invoices_df = pd.concat([invoices_df, new_invoice], ignore_index=True)
            invoices_df.to_excel('invoices.xlsx', index=False)
            
            # Update stock quantities
            products_df = pd.read_excel('products.xlsx')
            for item in self.cart_items:
                product_idx = products_df[products_df['SKU'] == item['sku']].index[0]
                current_qty = products_df.loc[product_idx, 'Quantity']
                new_qty = current_qty - item['quantity']
                products_df.loc[product_idx, 'Quantity'] = new_qty
            
            products_df.to_excel('products.xlsx', index=False)
            
            messagebox.showinfo("Success", f"Sale processed successfully!\nInvoice ID: {invoice_id}")
            
            # Clear cart and refresh displays
            self.clear_cart()
            self.load_products()
            self.load_stock_data()
            self.load_billing_products()
            self.load_invoices()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process sale: {str(e)}")
    
    def clear_cart(self):
        """Clear the shopping cart"""
        self.cart_items = []
        self.cart_total = 0
        self.update_cart_display()
        self.customer_entry.delete(0, tk.END)
        self.received_entry.delete(0, tk.END)
        self.change_label.config(text="Change: $0.00")
    
    def load_invoices(self):
        """Load invoices into the tree view"""
        try:
            invoices_df = pd.read_excel('invoices.xlsx')
            
            # Clear existing items
            for item in self.invoices_tree.get_children():
                self.invoices_tree.delete(item)
            
            # Add invoices to tree (most recent first)
            for _, row in invoices_df.sort_values('Date', ascending=False).iterrows():
                self.invoices_tree.insert('', 'end', values=(
                    row['Invoice_ID'], row['Date'], row['Customer_Name'],
                    f"${row['Total_Amount']:.2f}", row['Payment_Type']
                ))
        except Exception as e:
            print(f"Error loading invoices: {e}")
    
    def search_invoices(self):
        """Search invoices by customer name"""
        search_term = self.invoice_search_entry.get().lower()
        
        try:
            invoices_df = pd.read_excel('invoices.xlsx')
            
            # Clear existing items
            for item in self.invoices_tree.get_children():
                self.invoices_tree.delete(item)
            
            # Filter and add invoices
            for _, row in invoices_df.iterrows():
                if search_term in row['Customer_Name'].lower():
                    self.invoices_tree.insert('', 'end', values=(
                        row['Invoice_ID'], row['Date'], row['Customer_Name'],
                        f"${row['Total_Amount']:.2f}", row['Payment_Type']
                    ))
        except Exception as e:
            print(f"Error searching invoices: {e}")
    
    def update_barcode_combo(self):
        """Update barcode SKU combo box"""
        try:
            products_df = pd.read_excel('products.xlsx')
            sku_list = [f"{row['SKU']} - {row['Product_Name']}" for _, row in products_df.iterrows()]
            self.barcode_sku_combo['values'] = sku_list
        except Exception as e:
            print(f"Error updating barcode combo: {e}")
    
    def generate_barcode(self):
        """Generate barcode for selected product"""
        sku_selection = self.barcode_sku_var.get()
        if not sku_selection:
            messagebox.showerror("Error", "Please select a product")
            return
        
        try:
            # Extract SKU from selection
            sku = sku_selection.split(' - ')[0]
            
            # Generate barcode
            from barcode import Code128
            from barcode.writer import ImageWriter
            
            # Create barcode
            barcode_obj = Code128(sku, writer=ImageWriter())
            filename = f"barcodes/{sku}_barcode"
            barcode_obj.save(filename)
            
            messagebox.showinfo("Success", f"Barcode generated and saved as {filename}.png")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate barcode: {str(e)}")
    
    def view_barcode(self):
        """View generated barcode"""
        sku_selection = self.barcode_sku_var.get()
        if not sku_selection:
            messagebox.showerror("Error", "Please select a product")
            return
        
        try:
            # Extract SKU from selection
            sku = sku_selection.split(' - ')[0]
            barcode_path = f"barcodes/{sku}_barcode.png"
            
            if not os.path.exists(barcode_path):
                messagebox.showerror("Error", "Barcode not found. Please generate it first.")
                return
            
            # Clear previous display
            for widget in self.barcode_display_frame.winfo_children():
                widget.destroy()
            
            # Load and display barcode image
            image = Image.open(barcode_path)
            image = image.resize((400, 200), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            
            label = tk.Label(self.barcode_display_frame, image=photo, bg='white')
            label.image = photo  # Keep a reference
            label.pack(pady=20)
            
            tk.Label(self.barcode_display_frame, text=f"Barcode for SKU: {sku}", 
                    font=('Arial', 12, 'bold'), bg='white').pack()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to view barcode: {str(e)}")

# Main application runner
if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryManagementApp(root)
    root.mainloop()