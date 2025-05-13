import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import os

# File paths
FAMILY_FILE = r"C:\Users\User\OneDrive - ensonmarket.com\shelf assignment\family information.xlsx"
OUTPUT_FILE = r"C:\Users\User\OneDrive - ensonmarket.com\shelf assignment\Shelf_Assignment_Reversed_Output.xlsx"

class ShelfAssignmentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shelf Assignment Editor")
        
        # Set the window size for a large screen
        self.root.geometry("1200x800")  # Adjusted for a 20+ inch screen
        print("Initializing ShelfAssignmentApp with window size 1200x800")
        
        # Initialize data
        self.df = None
        self.families = []
        self.categories = {}
        self.full_values = []  # To store the full list of values for filtering
        
        # Apply a modern theme and custom styles
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use the 'clam' theme for a modern look
        print("Applying styles with 'clam' theme")
        self.apply_styles()
        
        # Load data
        print("Loading data...")
        self.load_data()
        
        # Create tabbed interface
        print("Creating ttk.Notebook for tabbed interface")
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.table_tab = ttk.Frame(self.notebook, style="Custom.TFrame")
        self.shelf_tab = ttk.Frame(self.notebook, style="Custom.TFrame")
        self.notebook.add(self.table_tab, text="Table View")
        self.notebook.add(self.shelf_tab, text="Shelf View")
        print("Tabs created: Table View, Shelf View")
        
        # Create GUI elements for each tab
        print("Creating Table View tab...")
        self.create_table_tab()
        print("Creating Shelf View tab...")
        try:
            self.create_shelf_tab()
        except Exception as e:
            print(f"Error creating Shelf View tab: {str(e)}")
            messagebox.showerror("Error", f"Failed to create Shelf View tab: {str(e)}")

    def apply_styles(self):
        """Apply custom styles for a more artistic and readable GUI."""
        # Define fonts
        self.large_font = ('Helvetica', 14)  # Larger font for better readability
        self.dropdown_font = ('Helvetica', 16)  # Larger font for dropdown values
        self.button_font = ('Helvetica', 16, 'bold')
        self.shelf_text_font_base = 8  # Smaller base font size for shelf text labels
        self.label_font_base = 10  # Base font size for level and shelf labels
        
        # Create a custom style for frames with background color
        self.style.configure("Custom.TFrame", background="#e6ecf0")  # Light grayish-blue background
        
        # Configure Treeview style (table)
        self.style.configure("Treeview",
                             font=self.large_font,
                             rowheight=40,  # Increase row height to fit larger font
                             background="#f0f4f8",  # Light grayish-blue background
                             foreground="#333333",  # Dark gray text
                             fieldbackground="#f0f4f8")
        self.style.configure("Treeview.Heading",
                             font=('Helvetica', 16, 'bold'),
                             background="#4a90e2",  # Blue header background
                             foreground="#ffffff")  # White header text
        
        # Configure Button style
        self.style.configure("TButton",
                             font=self.button_font,
                             padding=10,
                             background="#4a90e2",  # Blue button background
                             foreground="#ffffff")  # White button text
        
        # Configure Combobox style with adjusted padding
        self.style.configure("TCombobox",
                             font=self.large_font,
                             background="#ffffff",  # White background
                             foreground="#333333",  # Dark gray text
                             padding=(10, 5, 15, 5))  # Right padding for wider arrow area
        self.style.map("TCombobox",
                       fieldbackground=[('readonly', '#ffffff')],
                       selectbackground=[('readonly', '#ffffff')],
                       selectforeground=[('readonly', '#333333')])
        # Customize the dropdown arrow area
        self.style.layout("TCombobox", [
            ('Combobox.field', {'sticky': 'nswe', 'children': [
                ('Combobox.downarrow', {'side': 'right', 'sticky': 'ns'}),
                ('Combobox.padding', {'sticky': 'nswe', 'children': [
                    ('Combobox.textarea', {'sticky': 'nswe'})
                ]}),
            ]}),
        ])
        
        # Configure the dropdown listbox style for larger font
        self.style.configure("TCombobox.Listbox",
                             font=self.dropdown_font)

    def load_data(self):
        """Load data from the Excel files."""
        try:
            # Read the output file
            if not os.path.exists(OUTPUT_FILE):
                messagebox.showerror("Error", f"Output file not found: {OUTPUT_FILE}")
                self.root.destroy()
                return
            self.df = pd.read_excel(OUTPUT_FILE)
            print(f"Read output file. Rows: {len(self.df)}")
            print(f"Columns in output file: {list(self.df.columns)}")
            
            # Read family information to get families and categories
            xls = pd.ExcelFile(FAMILY_FILE)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(FAMILY_FILE, sheet_name=sheet_name)
                print(f"\nReading sheet: {sheet_name}")
                
                # Log the raw data for the first few rows to understand the structure
                print(f"First few rows of the sheet:\n{df.head()}")
                
                # Read family name from cell A2 (row 2 in Excel, index 0 in pandas)
                family_row = 0  # Index 0 corresponds to row 2 in Excel
                family = str(df.iloc[family_row, 0]) if not pd.isna(df.iloc[family_row, 0]) else ""
                print(f"Family in cell A2 (row 2, index {family_row}): {family}")
                
                if family:
                    # Categories are in the same row as the family (row 2 in Excel, index 0 in pandas)
                    category_row = family_row  # Same row as the family
                    categories = df.iloc[category_row, 1:].dropna().tolist()
                    print(f"Raw categories in row 2 (B2 onward, index {category_row}): {categories}")
                    
                    # Ensure categories are strings
                    categories = [str(cat) for cat in categories]
                    print(f"Categories after converting to strings: {categories}")
                    
                    self.families.append(family)
                    self.categories[family] = categories
            print(f"\nFamilies loaded: {self.families}")
            print(f"Categories loaded: {self.categories}")
            
            # Ensure Family and Category columns exist
            if 'Family' not in self.df.columns:
                self.df['Family'] = ""
            if 'Category' not in self.df.columns:
                self.df['Category'] = ""
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            messagebox.showerror("Error", f"Error loading data: {str(e)}")
            self.root.destroy()

    def create_table_tab(self):
        """Create the table view tab (original GUI)."""
        # Create main frame
        frame = ttk.Frame(self.table_tab, style="Custom.TFrame")
        frame.pack(padx=20, pady=20, fill="both", expand=True)
        print("Created main frame for Table View tab")
        
        # Create Treeview to display data
        columns = list(self.df.columns)
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", style="Treeview")
        print("Created Treeview with columns:", columns)
        
        # Set column headings and widths
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)  # Increased width for larger font
        
        # Insert data into Treeview
        for idx, row in self.df.iterrows():
            self.tree.insert("", tk.END, values=list(row), iid=str(idx))
        print(f"Inserted {len(self.df)} rows into Treeview")
        
        # Add scrollbars
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        print("Added scrollbars to Treeview")
        
        # Layout Treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        print("Laid out Treeview and scrollbars")
        
        # Create Save button
        save_button = ttk.Button(frame, text="Save", command=self.save_data, style="TButton")
        save_button.grid(row=2, column=0, pady=20)
        print("Added Save button to Table View tab")
        
        # Variables for editing
        self.current_edit = None
        self.dropdown = None
        
        # Bind single-click to edit cells
        self.tree.bind("<Button-1>", self.on_single_click)
        print("Bound single-click event to Treeview")

    def create_shelf_tab(self):
        """Create the shelf view tab with 3D shelf visualization."""
        # Create main frame
        frame = ttk.Frame(self.shelf_tab, style="Custom.TFrame")
        frame.pack(padx=20, pady=20, fill="both", expand=True)
        print("Created main frame for Shelf View tab")
        
        # Create frame for dropdowns
        dropdown_frame = ttk.Frame(frame, style="Custom.TFrame")
        dropdown_frame.pack(fill="x", pady=10)
        print("Created dropdown frame for Shelf View tab")
        
        # Populate dropdowns for Section, Aisle, Side
        self.sections = sorted(self.df['Section'].unique().tolist())
        self.aisles = sorted(self.df['Aisle'].unique().tolist())
        self.sides = sorted(self.df['Side'].unique().tolist())
        print(f"Sections: {self.sections}")
        print(f"Aisles: {self.aisles}")
        print(f"Sides: {self.sides}")
        
        # Section dropdown
        ttk.Label(dropdown_frame, text="Section:", font=self.large_font).grid(row=0, column=0, padx=5)
        self.section_var = tk.StringVar()
        self.section_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.section_var, values=self.sections, state="readonly", style="TCombobox")
        self.section_dropdown.grid(row=0, column=1, padx=5)
        self.section_dropdown.bind("<<ComboboxSelected>>", self.update_shelf_view)
        print("Added Section dropdown")
        
        # Aisle dropdown
        ttk.Label(dropdown_frame, text="Aisle:", font=self.large_font).grid(row=0, column=2, padx=5)
        self.aisle_var = tk.StringVar()
        self.aisle_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.aisle_var, values=self.aisles, state="readonly", style="TCombobox")
        self.aisle_dropdown.grid(row=0, column=3, padx=5)
        self.aisle_dropdown.bind("<<ComboboxSelected>>", self.update_shelf_view)
        print("Added Aisle dropdown")
        
        # Side dropdown
        ttk.Label(dropdown_frame, text="Side:", font=self.large_font).grid(row=0, column=4, padx=5)
        self.side_var = tk.StringVar()
        self.side_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.side_var, values=self.sides, state="readonly", style="TCombobox")
        self.side_dropdown.grid(row=0, column=5, padx=5)
        self.side_dropdown.bind("<<ComboboxSelected>>", self.update_shelf_view)
        print("Added Side dropdown")
        
        # Family dropdown
        ttk.Label(dropdown_frame, text="Family:", font=self.large_font).grid(row=0, column=6, padx=5)
        self.family_var = tk.StringVar()
        self.family_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.family_var, values=self.families, state="readonly", style="TCombobox")
        self.family_dropdown.grid(row=0, column=7, padx=5)
        self.family_dropdown.bind("<<ComboboxSelected>>", self.update_category_dropdown)
        print("Added Family dropdown")
        
        # Category dropdown
        ttk.Label(dropdown_frame, text="Category:", font=self.large_font).grid(row=0, column=8, padx=5)
        self.category_var = tk.StringVar()
        self.category_dropdown = ttk.Combobox(dropdown_frame, textvariable=self.category_var, state="readonly", style="TCombobox")
        self.category_dropdown.grid(row=0, column=9, padx=5)
        print("Added Category dropdown")
        
        # Create Canvas for 3D shelf visualization
        self.canvas_frame = ttk.Frame(frame, style="Custom.TFrame")
        self.canvas_frame.pack(fill="both", expand=True)
        print("Created canvas frame for Shelf View tab")
        
        self.canvas = tk.Canvas(self.canvas_frame, bg="#ffffff")
        self.canvas.pack(fill="both", expand=True)
        print("Created canvas for 3D shelf visualization")
        
        # Bind mouse events for selection
        self.canvas.bind("<Button-1>", self.start_selection)
        self.canvas.bind("<B1-Motion>", self.update_selection)
        self.canvas.bind("<ButtonRelease-1>", self.end_selection)
        print("Bound mouse events for selection on canvas")
        
        # Bind resize event to redraw the shelf
        self.canvas.bind("<Configure>", self.on_resize)
        print("Bound resize event to canvas")
        
        # Variables for selection
        self.start_x = None
        self.start_y = None
        self.selection_rect = None
        self.selected_cells = set()  # Store (level, shelf) coordinates of selected cells
        
        # Variables for shelf sizing
        self.initial_cell_width = None
        self.initial_cell_height = None
        self.initial_aspect_ratio = None
        self.scale_factor = 1.0  # Scaling factor based on window size
        
        # Color mapping for categories (eye-friendly, high-contrast colors)
        self.category_colors = {}
        self.color_list = [
            "darkblue",    # Deep blue
            "darkgreen",   # Forest green
            "darkred",     # Deep red
            "purple4",     # Dark purple
            "darkorange",  # Muted orange
            "saddlebrown", # Earthy brown
            "deeppink4",   # Muted pink
            "teal",        # Teal
            "darkmagenta", # Dark magenta
            "olive",       # Olive green
            "navy",        # Navy blue
            "coral4",      # Muted coral
            "goldenrod",   # Muted gold
            "darkviolet",  # Dark violet
            "seagreen",    # Sea green
            "indigo"       # Indigo
        ]
        
        # Create Apply and Clear Selection buttons
        button_frame = ttk.Frame(frame, style="Custom.TFrame")
        button_frame.pack(pady=10)
        
        apply_button = ttk.Button(button_frame, text="Apply", command=self.apply_selection, style="TButton")
        apply_button.grid(row=0, column=0, padx=5)
        print("Added Apply button to Shelf View tab")
        
        clear_button = ttk.Button(button_frame, text="Clear Selection", command=self.clear_selection, style="TButton")
        clear_button.grid(row=0, column=1, padx=5)
        print("Added Clear Selection button to Shelf View tab")
        
        # Initialize the shelf view
        if self.sections:
            self.section_var.set(self.sections[0])
        if self.aisles:
            self.aisle_var.set(self.aisles[0])
        if self.sides:
            self.side_var.set(self.sides[0])
        print("Initialized shelf view with default dropdown values")
        self.update_shelf_view()

    def on_resize(self, event):
        """Handle window resize by redrawing the shelf with adjusted sizes."""
        # Recalculate the scale factor based on the new canvas size
        self.canvas.update_idletasks()
        new_width = self.canvas.winfo_width()
        new_height = self.canvas.winfo_height()
        
        # Initial canvas dimensions (approximated as before)
        initial_width = 1000
        initial_height = 600
        
        # Calculate scale factors for width and height
        scale_width = new_width / initial_width
        scale_height = new_height / initial_height
        
        # Use the smaller scale factor to maintain aspect ratio
        self.scale_factor = min(scale_width, scale_height)
        print(f"Window resized: new width={new_width}, new height={new_height}, scale_factor={self.scale_factor}")
        
        # Redraw the shelf with the new scale factor
        self.update_shelf_view()

    def update_shelf_view(self, event=None):
        """Update the 3D shelf visualization based on Section, Aisle, and Side selection."""
        section = self.section_var.get()
        aisle = self.aisle_var.get()
        side = self.side_var.get()
        
        if not section or not aisle or not side:
            print("No selection for Section, Aisle, or Side; skipping shelf view update")
            return
        
        print(f"Updating shelf view for Section: {section}, Aisle: {aisle}, Side: {side}")
        
        # Filter the DataFrame for the selected Section, Aisle, and Side
        filtered_df = self.df[
            (self.df['Section'] == section) &
            (self.df['Aisle'] == int(aisle)) &
            (self.df['Side'] == int(side))
        ]
        
        if filtered_df.empty:
            print("No data found for selected Section, Aisle, and Side; clearing canvas")
            self.canvas.delete("all")
            return
        
        # Determine the number of levels and shelves
        max_level = filtered_df['Level'].max()
        max_shelf = filtered_df['Shelf'].max()
        
        if not max_level or not max_shelf:
            print("Max level or max shelf not found; clearing canvas")
            self.canvas.delete("all")
            return
        
        self.max_level = int(max_level)
        self.max_shelf = int(max_shelf)
        print(f"Max Level: {self.max_level}, Max Shelf: {self.max_shelf}")
        
        # Clear the canvas
        self.canvas.delete("all")
        self.selected_cells.clear()
        print("Cleared canvas and selected cells")
        
        # Calculate base cell size (before scaling)
        canvas_width_base = 1000  # Base width for initial calculation
        canvas_height_base = 600  # Base height for initial calculation
        cell_width_base = canvas_width_base // self.max_shelf
        cell_height_base = canvas_height_base // self.max_level
        self.cell_width_base = min(cell_width_base, 60)  # Base width of each shelf
        self.cell_height_base = min(cell_height_base, 80)  # Base height of each shelf
        
        # Calculate the initial aspect ratio (only once)
        if self.initial_aspect_ratio is None:
            self.initial_cell_width = self.cell_width_base
            self.initial_cell_height = self.cell_height_base
            self.initial_aspect_ratio = self.initial_cell_width / self.initial_cell_height
            print(f"Initial aspect ratio: {self.initial_aspect_ratio}")
        
        # Apply the scale factor to maintain aspect ratio
        self.cell_width = self.cell_width_base * self.scale_factor
        self.cell_height = self.cell_height_base * self.scale_factor
        
        # Ensure the aspect ratio is maintained
        current_aspect_ratio = self.cell_width / self.cell_height
        if abs(current_aspect_ratio - self.initial_aspect_ratio) > 0.01:  # Small tolerance for floating-point errors
            self.cell_height = self.cell_width / self.initial_aspect_ratio
            print(f"Adjusted cell height to maintain aspect ratio: cell_width={self.cell_width}, cell_height={self.cell_height}")
        
        # Scale the depth and fonts
        self.depth = 10 * self.scale_factor  # Depth effect for 3D visualization
        shelf_font_size = int(self.shelf_text_font_base * self.scale_factor)
        label_font_size = int(self.label_font_base * self.scale_factor)
        self.shelf_text_font = ('Helvetica', max(shelf_font_size, 6), 'bold')  # Bold text for better visibility
        self.label_font = ('Helvetica', max(label_font_size, 6))
        print(f"Scaled sizes: cell_width={self.cell_width}, cell_height={self.cell_height}, depth={self.depth}, shelf_font_size={shelf_font_size}, label_font_size={label_font_size}")
        
        # Calculate the total size of the shelf grid (including space for labels)
        label_space_left = 50 * self.scale_factor  # Space for level labels on the left
        label_space_top = 30 * self.scale_factor   # Space for shelf labels on the top
        total_width = self.max_shelf * self.cell_width + self.depth + label_space_left
        total_height = self.max_level * self.cell_height + self.depth + label_space_top
        
        # Center the shelf grid in the canvas
        self.canvas.update_idletasks()
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        offset_x = (canvas_width - total_width) // 2 + label_space_left
        offset_y = (canvas_height - total_height) // 2 + label_space_top
        print(f"Centering shelf grid: offset_x={offset_x}, offset_y={offset_y}")
        
        # Draw shelf labels (Shelf 1, Shelf 2, etc.) above the grid
        for shelf in range(1, self.max_shelf + 1):
            label_x = (shelf - 1) * self.cell_width + offset_x + self.cell_width / 2
            label_y = offset_y - self.depth - 10 * self.scale_factor
            self.canvas.create_text(
                label_x, label_y,
                text=f"S{shelf}",
                font=self.label_font,
                fill="black",
                anchor="center"
            )
        
        # Draw level labels (Level 1, Level 2, etc.) to the left of the grid
        for level in range(1, self.max_level + 1):
            display_row = self.max_level - level
            label_y = display_row * self.cell_height + offset_y + self.cell_height / 2
            label_x = offset_x - self.depth - 30 * self.scale_factor
            self.canvas.create_text(
                label_x, label_y,
                text=f"L{level}",
                font=self.label_font,
                fill="black",
                anchor="center"
            )
        
        # Build category color mapping
        unique_categories = filtered_df['Category'].dropna().unique()
        self.category_colors.clear()
        for idx, category in enumerate(unique_categories):
            color = self.color_list[idx % len(self.color_list)]
            self.category_colors[str(category)] = color
        print(f"Category color mapping: {self.category_colors}")
        
        # Draw the 3D shelves with Level 1 at the top
        self.cell_coords = {}  # Store coordinates for each (level, shelf)
        for level in range(1, self.max_level + 1):
            # Reverse the level ordering: Level 1 at the top, max_level at the bottom
            display_row = self.max_level - level  # Level 1 -> row 0 (top), Level max_level -> row (max_level-1) (bottom)
            for shelf in range(1, self.max_shelf + 1):
                # Base coordinates for the shelf (top-left corner of the shelf face)
                x1 = (shelf - 1) * self.cell_width + offset_x
                y1 = display_row * self.cell_height + offset_y
                x2 = x1 + self.cell_width
                y2 = y1 + self.cell_height
                
                # Adjust for 3D effect (top-left corner shifted for perspective)
                x1_3d = x1 + self.depth
                y1_3d = y1
                x2_3d = x2 + self.depth
                y2_3d = y2
                
                # Draw the front face of the shelf (trapezoid for perspective)
                self.canvas.create_polygon(
                    x1_3d, y1_3d,  # Top-left
                    x2_3d, y1_3d,  # Top-right
                    x2, y2,        # Bottom-right
                    x1, y2,        # Bottom-left
                    fill="#d3d3d3", outline="black"  # Light gray for the front face
                )
                
                # Draw the top edge (for 3D effect)
                self.canvas.create_polygon(
                    x1_3d, y1_3d,  # Top-left of front face
                    x2_3d, y1_3d,  # Top-right of front face
                    x2_3d - self.depth, y1_3d - self.depth,  # Top-right shifted up
                    x1_3d - self.depth, y1_3d - self.depth,  # Top-left shifted up
                    fill="#f0f0f0", outline="black"  # Lighter gray for the top edge
                )
                
                # Draw the right edge (for 3D effect)
                self.canvas.create_polygon(
                    x2_3d, y1_3d,  # Top-right of front face
                    x2_3d - self.depth, y1_3d - self.depth,  # Top-right shifted up
                    x2 - self.depth, y2 - self.depth,  # Bottom-right shifted up
                    x2, y2,        # Bottom-right of front face
                    fill="#c0c0c0", outline="black"  # Darker gray for the right edge
                )
                
                # Store coordinates for selection (use the front face for selection purposes)
                self.cell_coords[(level, shelf)] = (x1, y1, x2, y2)
                
                # Add text label with Category value if available
                mask = (
                    (self.df['Section'] == section) &
                    (self.df['Aisle'] == int(aisle)) &
                    (self.df['Side'] == int(side)) &
                    (self.df['Level'] == level) &
                    (self.df['Shelf'] == shelf)
                )
                row = self.df[mask]
                if not row.empty:
                    category = str(row.iloc[0]['Category'])
                    if pd.isna(category) or category == "" or category == "nan":
                        continue
                    
                    # Determine the text color based on the category
                    text_color = self.category_colors.get(category, "black")
                    
                    # Split the category text into multiple lines if too long
                    max_width = self.cell_width - 10  # Approximate available width
                    font_size = max(shelf_font_size, 6)
                    avg_char_width = font_size * 0.6  # Rough estimate of character width
                    max_chars_per_line = int(max_width / avg_char_width)
                    
                    # Split the text into words
                    words = category.split()
                    lines = []
                    current_line = []
                    current_length = 0
                    
                    for word in words:
                        word_length = len(word)
                        if current_length + word_length + len(current_line) <= max_chars_per_line:
                            current_line.append(word)
                            current_length += word_length
                        else:
                            lines.append(" ".join(current_line))
                            current_line = [word]
                            current_length = word_length
                    if current_line:
                        lines.append(" ".join(current_line))
                    
                    # Draw each line of text
                    num_lines = len(lines)
                    line_spacing = font_size * 1.2  # Space between lines
                    total_text_height = num_lines * line_spacing
                    start_y = (y1 + y2) / 2 - total_text_height / 2 + line_spacing / 2
                    
                    for idx, line in enumerate(lines):
                        text_x = (x1 + x2) / 2 + self.depth / 2
                        text_y = start_y + idx * line_spacing
                        self.canvas.create_text(
                            text_x, text_y,
                            text=line,
                            font=self.shelf_text_font,
                            fill=text_color,
                            anchor="center"
                        )
        print(f"Drew 3D shelf grid with {self.max_level} levels and {self.max_shelf} shelves")

    def start_selection(self, event):
        """Start the selection process on mouse click."""
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.selection_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="blue", dash=(2, 2)
        )
        print(f"Started selection at ({self.start_x}, {self.start_y})")

    def update_selection(self, event):
        """Update the selection rectangle while dragging."""
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.selection_rect, self.start_x, self.start_y, current_x, current_y)
        
        # Highlight cells within the selection
        self.selected_cells.clear()
        for (level, shelf), (x1, y1, x2, y2) in self.cell_coords.items():
            # Check if the cell overlaps with the selection rectangle
            sel_x1, sel_y1, sel_x2, sel_y2 = self.canvas.coords(self.selection_rect)
            if (min(sel_x1, sel_x2) <= x2 and max(sel_x1, sel_x2) >= x1 and
                min(sel_y1, sel_y2) <= y2 and max(sel_y1, sel_y2) >= y1):
                self.selected_cells.add((level, shelf))
                # Highlight the front face of the shelf
                self.canvas.itemconfig(
                    self.canvas.find_enclosed(x1, y1, x2, y2),
                    fill="lightblue"
                )
            else:
                self.canvas.itemconfig(
                    self.canvas.find_enclosed(x1, y1, x2, y2),
                    fill="#d3d3d3"  # Reset to default front face color
                )
        print(f"Updated selection: {len(self.selected_cells)} cells selected")

    def end_selection(self, event):
        """End the selection process on mouse release and apply the selection."""
        self.canvas.delete(self.selection_rect)
        self.selection_rect = None
        self.start_x = None
        self.start_y = None
        print(f"Ended selection with {len(self.selected_cells)} cells selected")
        
        # Automatically apply the selection
        if self.selected_cells:
            self.apply_selection()

    def clear_selection(self):
        """Clear the current selection and reset highlights."""
        self.selected_cells.clear()
        for (level, shelf), (x1, y1, x2, y2) in self.cell_coords.items():
            self.canvas.itemconfig(
                self.canvas.find_enclosed(x1, y1, x2, y2),
                fill="#d3d3d3"  # Reset to default front face color
            )
        print("Cleared selection")

    def update_category_dropdown(self, event=None):
        """Update the Category dropdown based on the selected Family."""
        family = self.family_var.get()
        if family in self.categories:
            self.category_dropdown['values'] = self.categories[family]
            self.category_var.set(self.categories[family][0] if self.categories[family] else "")
        else:
            self.category_dropdown['values'] = ["No Categories Available"]
            self.category_var.set("No Categories Available")
        print(f"Updated Category dropdown for Family '{family}': {self.category_dropdown['values']}")

    def apply_selection(self):
        """Apply the selected Family and Category to the selected shelves in the Table View."""
        section = self.section_var.get()
        aisle = self.aisle_var.get()
        side = self.side_var.get()
        family = self.family_var.get()
        category = self.category_var.get()
        
        if not section or not aisle or not side or not family or not category:
            messagebox.showwarning("Warning", "Please select all dropdown values.")
            print("Apply failed: Missing dropdown values")
            return
        
        if not self.selected_cells:
            messagebox.showwarning("Warning", "Please select at least one shelf in the grid.")
            print("Apply failed: No cells selected")
            return
        
        # Update the DataFrame and Treeview
        updated_rows = 0
        for level, shelf in self.selected_cells:
            # Find the corresponding row in the DataFrame
            mask = (
                (self.df['Section'] == section) &
                (self.df['Aisle'] == int(aisle)) &
                (self.df['Side'] == int(side)) &
                (self.df['Level'] == level) &
                (self.df['Shelf'] == shelf)
            )
            row_idx = self.df.index[mask]
            if not row_idx.empty:
                row_idx = row_idx[0]
                # Update the DataFrame
                self.df.at[row_idx, 'Family'] = family
                self.df.at[row_idx, 'Category'] = category
                # Update the Treeview
                values = list(self.df.iloc[row_idx])
                self.tree.item(str(row_idx), values=values)
                updated_rows += 1
        messagebox.showinfo("Success", f"Family and Category values applied to {updated_rows} selected shelves.")
        print(f"Applied Family: {family}, Category: {category} to {updated_rows} shelves")
        
        # Refresh the shelf view to update the text labels
        self.update_shelf_view()

    def on_single_click(self, event):
        """Handle single-click to edit Family or Category cells in the Table View."""
        print("Single-click event triggered")
        
        # Remove any existing dropdown
        if self.dropdown is not None:
            self.dropdown.destroy()
            self.dropdown = None
        
        # Identify the cell clicked
        region = self.tree.identify("region", event.x, event.y)
        print(f"Region identified: {region}")
        if region != "cell":
            print("Not a cell region, exiting")
            return
        
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        print(f"Row ID: {row_id}, Column ID: {column_id}")
        
        column_idx = int(column_id.replace("#", "")) - 1
        column_name = self.df.columns[column_idx]
        print(f"Column name: {column_name}")
        
        # Only allow editing for Family and Category columns
        if column_name not in ["Family", "Category"]:
            print(f"Column {column_name} is not editable (Family or Category required)")
            return
        
        # Get the bounding box of the cell
        bbox = self.tree.bbox(row_id, column_id)
        print(f"Bounding box: {bbox}")
        if not bbox:
            print("Bounding box is empty, cannot place dropdown")
            return
        
        x, y, width, height = bbox
        
        # Calculate the maximum width to keep the down arrow visible
        window_width = self.root.winfo_width()
        max_width = window_width - x - 20  # Leave some margin
        adjusted_width = min(width + 20, max_width)  # Use the smaller of the two
        
        # Adjust x position if the combobox would extend beyond the window
        if x + adjusted_width > window_width:
            x = window_width - adjusted_width - 20  # Shift left to keep within window
        
        self.current_edit = (row_id, column_idx, column_name)
        
        # Create a dropdown (Combobox) at the cell's position
        self.dropdown = ttk.Combobox(self.tree, state="normal", style="TCombobox")  # Allow typing
        if column_name == "Family":
            self.full_values = self.families  # Store the full list for filtering
            self.dropdown["values"] = self.families
            current_value = str(self.df.at[int(row_id), "Family"])
            if pd.isna(current_value) or current_value == "nan":
                current_value = ""
            if current_value in self.families:
                self.dropdown.set(current_value)
            else:
                self.dropdown.set("")
            print(f"Family dropdown created with values: {self.families}, current: {current_value}")
        else:  # Category
            # Get the selected family in this row
            family = str(self.df.at[int(row_id), "Family"])
            if pd.isna(family) or family == "nan":
                family = ""
            if family in self.categories:
                self.full_values = self.categories[family]  # Store the full list for filtering
                self.dropdown["values"] = self.categories[family]
            else:
                self.full_values = ["No Categories Available"]
                self.dropdown["values"] = ["No Categories Available"]
            current_value = str(self.df.at[int(row_id), "Category"])
            if pd.isna(current_value) or current_value == "nan":
                current_value = ""
            if current_value in self.dropdown["values"]:
                self.dropdown.set(current_value)
            else:
                self.dropdown.set("")
            print(f"Category dropdown created for family '{family}' with values: {self.dropdown['values']}, current: {current_value}")
        
        # Position the dropdown with adjusted width to ensure down arrow is visible
        self.dropdown.place(x=x, y=y, width=adjusted_width, height=height)
        self.dropdown.lift()  # Ensure the dropdown is on top
        self.dropdown.focus_set()
        
        # Bind events for filtering, selection, and closing
        self.dropdown.bind("<KeyRelease>", self.on_key_release)
        self.dropdown.bind("<<ComboboxSelected>>", self.on_dropdown_select)
        self.dropdown.bind("<FocusOut>", self.on_dropdown_close)
        self.dropdown.bind("<Return>", self.on_dropdown_select)  # Allow Enter key to select

    def on_key_release(self, event):
        """Filter the combobox values based on the typed text and ensure dropdown pops up."""
        # Ignore arrow keys and Enter to allow navigation
        if event.keysym in ["Up", "Down", "Return"]:
            print(f"Arrow key or Enter pressed: {event.keysym}, skipping filter")
            return
        
        typed_text = self.dropdown.get().strip().lower()
        print(f"Key released, typed text: {typed_text}")
        
        if typed_text == "":
            # If the text is cleared, show the full list
            self.dropdown["values"] = self.full_values
            print(f"Restored full values: {self.full_values}")
        else:
            # Filter the values to those starting with the typed text (case-insensitive)
            filtered_values = [val for val in self.full_values if val.lower().startswith(typed_text)]
            self.dropdown["values"] = filtered_values
            print(f"Filtered values: {filtered_values}")
        
        # Ensure the dropdown pops up after filtering
        self.root.after(100, lambda: self.dropdown.event_generate('<Down>'))
        self.dropdown.focus_set()  # Ensure the combobox retains focus

    def on_dropdown_select(self, event=None):
        """Update the data when a dropdown selection is made or Enter is pressed."""
        print("Dropdown selection made")
        if self.current_edit is None:
            print("No current edit, exiting")
            return
        row_id, column_idx, column_name = self.current_edit
        selected_value = self.dropdown.get()
        print(f"Selected value: {selected_value} for {column_name} in row {row_id}")
        
        # Update the DataFrame
        self.df.at[int(row_id), column_name] = selected_value
        
        # If the Family value changed, reset the Category value in the same row
        if column_name == "Family":
            print(f"Family changed, resetting Category for row {row_id}")
            self.df.at[int(row_id), "Category"] = ""  # Reset Category to empty
            
        # Update the Treeview display
        values = list(self.df.iloc[int(row_id)])
        self.tree.item(row_id, values=values)
        
        # Update the Family and Category dropdowns in the Shelf View tab
        if column_name == "Family":
            self.family_var.set(selected_value)
            self.update_category_dropdown()
        elif column_name == "Category":
            self.category_var.set(selected_value)
        
        # Clean up
        self.dropdown.destroy()
        self.dropdown = None
        self.current_edit = None

    def on_dropdown_close(self, event):
        """Clean up the dropdown when it loses focus."""
        print("Dropdown lost focus, closing")
        if self.dropdown is not None:
            self.dropdown.destroy()
            self.dropdown = None
            self.current_edit = None

    def save_data(self):
        """Save the updated data back to the Excel file."""
        try:
            self.df.to_excel(OUTPUT_FILE, index=False)
            print(f"Updated data saved to: {OUTPUT_FILE}")
            messagebox.showinfo("Success", f"Data saved successfully to {OUTPUT_FILE}")
        except Exception as e:
            print(f"Error saving data: {str(e)}")
            messagebox.showerror("Error", f"Error saving data: {str(e)}")

def main():
    """Main function to launch the GUI."""
    if not os.path.exists(FAMILY_FILE):
        print(f"Family file not found: {FAMILY_FILE}")
        return
    if not os.path.exists(OUTPUT_FILE):
        print(f"Output file not found: {OUTPUT_FILE}")
        return
    
    root = tk.Tk()
    app = ShelfAssignmentApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()