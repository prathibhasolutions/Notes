from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.pdfgen import canvas
import os

# PDF file path
pdf_file = "DCA_Notes_Part1_Prathibha_Institute.pdf"

# Custom page with header/footer
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_elements(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_elements(self, page_count):
        self.saveState()
        # Header
        self.setFont('Helvetica-Bold', 10)
        self.drawString(1*cm, A4[1] - 1.5*cm, "Prathibha Institute")
        
        # Footer
        self.setFont('Helvetica', 9)
        self.drawCentredString(A4[0]/2, 1*cm, f"Page {self._pageNumber}")
        
        # Add line
        self.setStrokeColor(colors.HexColor('#4472C4'))
        self.setLineWidth(0.5)
        self.line(1*cm, A4[1] - 1.8*cm, A4[0] - 1*cm, A4[1] - 1.8*cm)
        self.line(1*cm, 1.5*cm, A4[0] - 1*cm, 1.5*cm)
        
        self.restoreState()

# Create PDF
doc = SimpleDocTemplate(pdf_file, pagesize=A4,
                        leftMargin=1*inch, rightMargin=1*inch,
                        topMargin=1.2*inch, bottomMargin=1*inch)

# Styles
styles = getSampleStyleSheet()
story = []

# Custom Styles
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Heading1'],
    fontSize=24,
    textColor=colors.HexColor('#1F4E78'),
    spaceAfter=30,
    alignment=TA_CENTER,
    fontName='Helvetica-Bold'
)

heading1_style = ParagraphStyle(
    'CustomHeading1',
    parent=styles['Heading1'],
    fontSize=18,
    textColor=colors.HexColor('#2E75B5'),
    spaceAfter=8,
    spaceBefore=8,
    fontName='Helvetica-Bold'
)

heading2_style = ParagraphStyle(
    'CustomHeading2',
    parent=styles['Heading2'],
    fontSize=14,
    textColor=colors.HexColor('#4472C4'),
    spaceAfter=6,
    spaceBefore=6,
    fontName='Helvetica-Bold'
)

heading3_style = ParagraphStyle(
    'CustomHeading3',
    parent=styles['Heading3'],
    fontSize=12,
    textColor=colors.HexColor('#5B9BD5'),
    spaceAfter=4,
    spaceBefore=4,
    fontName='Helvetica-Bold'
)

body_style = ParagraphStyle(
    'CustomBody',
    parent=styles['Normal'],
    fontSize=10,
    leading=13,
    alignment=TA_JUSTIFY,
    spaceAfter=4
)

# Title Page
story.append(Spacer(1, 2*inch))
story.append(Paragraph("Prathibha Institute", title_style))
story.append(Spacer(1, 0.3*inch))
story.append(Paragraph("Diploma in Computer Applications (DCA)", 
                      ParagraphStyle('subtitle', parent=styles['Normal'], fontSize=16, 
                                    alignment=TA_CENTER, textColor=colors.HexColor('#4472C4'))))
story.append(Spacer(1, 0.5*inch))
story.append(Paragraph("Comprehensive Study Notes", 
                      ParagraphStyle('subtitle2', parent=styles['Normal'], fontSize=14, 
                                    alignment=TA_CENTER, fontName='Helvetica-Bold')))
story.append(Spacer(1, 0.2*inch))
story.append(Paragraph("Part 1: Computer Fundamentals, MS Office & Internet", 
                      ParagraphStyle('subtitle3', parent=styles['Normal'], fontSize=13, 
                                    alignment=TA_CENTER)))
story.append(Spacer(1, 1.5*inch))
story.append(Paragraph("Professional Course Material for DCA Students", 
                      ParagraphStyle('desc', parent=styles['Normal'], fontSize=11, 
                                    alignment=TA_CENTER, textColor=colors.grey)))

story.append(PageBreak())

print("Generating PDF content...")

# ============ TOPIC 1: FUNDAMENTALS OF COMPUTERS ============
story.append(Paragraph("1. Fundamentals of Computers", heading1_style))
story.append(Spacer(1, 0.1*inch))

# 1.1 Introduction to Computers
story.append(Paragraph("1.1 Introduction to Computers", heading2_style))
story.append(Paragraph("What is a Computer?", heading3_style))
story.append(Paragraph("""A computer is an electronic device that accepts data as input, processes it according to 
a set of instructions (program), and produces information as output. The term "computer" is derived from the Latin 
word "computare" which means to calculate.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Key Characteristics of Computers:</b>", body_style))
characteristics = [
    "<b>Speed:</b> Computers can process millions of instructions per second",
    "<b>Accuracy:</b> Computers perform calculations with near-perfect accuracy",
    "<b>Storage:</b> Large amounts of data can be stored and retrieved quickly",
    "<b>Automation:</b> Can perform tasks automatically without human intervention",
    "<b>Versatility:</b> Can perform various types of tasks",
    "<b>Diligence:</b> Can work continuously without fatigue"
]
for char in characteristics:
    story.append(Paragraph(f"• {char}", body_style))
story.append(Spacer(1, 0.08*inch))

# History
story.append(Paragraph("History of Computers - Generations", heading3_style))
generations = [
    ("<b>First Generation (1940-1956): Vacuum Tubes</b>", 
     "Technology: Vacuum tubes for circuitry, magnetic drums for memory. Examples: ENIAC, UNIVAC. Characteristics: Very large in size, consumed lot of electricity, generated excessive heat, slow speed, less reliable."),
    
    ("<b>Second Generation (1956-1963): Transistors</b>", 
     "Technology: Transistors replaced vacuum tubes. Examples: IBM 1401, CDC 1604. Characteristics: Smaller, faster, more reliable, less heat generation. Programming: Assembly language, COBOL, FORTRAN."),
    
    ("<b>Third Generation (1964-1971): Integrated Circuits</b>", 
     "Technology: Integrated Circuits (ICs) - transistors on silicon chips. Examples: IBM 360 series. Characteristics: Smaller size, faster, more reliable, affordable. Development: Operating systems, multiprogramming."),
    
    ("<b>Fourth Generation (1971-Present): Microprocessors</b>", 
     "Technology: VLSI (Very Large Scale Integration) - thousands of ICs on single chip. Examples: IBM PC, Apple Macintosh. Characteristics: Portable, reliable, cheap. Development: GUI, internet, multimedia."),
    
    ("<b>Fifth Generation (Present-Future): Artificial Intelligence</b>", 
     "Technology: ULSI (Ultra Large Scale Integration), AI, parallel processing. Focus: Voice recognition, natural language processing, expert systems. Examples: AI systems, quantum computers.")
]

for gen_title, gen_desc in generations:
    story.append(Paragraph(gen_title, body_style))
    story.append(Paragraph(gen_desc, body_style))
    story.append(Spacer(1, 0.05*inch))

# 1.2 Hardware and Software
story.append(Paragraph("1.2 Computer Hardware and Software", heading2_style))
story.append(Paragraph("Hardware", heading3_style))
story.append(Paragraph("""Hardware refers to the physical components of a computer system that you can touch and see. 
It includes Input Devices, Output Devices, Processing Unit (CPU), and Storage Devices.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("Software", heading3_style))
story.append(Paragraph("""Software is a set of programs, procedures, and documentation that performs specific tasks 
on a computer system.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Types of Software:</b>", body_style))
software_types = [
    "<b>System Software:</b> Operating Systems (Windows, Linux, macOS), Device Drivers, Utility Programs (Antivirus, Disk Cleanup), Language Translators (Compilers, Interpreters)",
    "<b>Application Software:</b> Word Processors (MS Word), Spreadsheets (MS Excel), Database Management Systems (MS Access, Oracle), Graphics Software (Photoshop), Web Browsers (Chrome, Firefox)"
]
for sw in software_types:
    story.append(Paragraph(f"• {sw}", body_style))
story.append(Spacer(1, 0.08*inch))

# 1.3 Input and Output Devices
story.append(Paragraph("1.3 Input and Output Devices", heading2_style))
story.append(Paragraph("Input Devices", heading3_style))
story.append(Paragraph("Input devices are used to enter data and instructions into the computer.", body_style))
story.append(Spacer(1, 0.1*inch))

input_devices = [
    ("Keyboard", "Most common input device with alphanumeric keys, function keys, cursor control keys. Used for typing text and commands."),
    ("Mouse", "Pointing device with left, right, and scroll buttons. Used to point, click, drag, and select items on screen."),
    ("Scanner", "Converts printed documents and images into digital format. Types: Flatbed, Handheld, Drum scanners."),
    ("Microphone", "Input device for voice and audio. Used for voice recording, video conferencing, voice commands."),
    ("Webcam", "Captures video and images. Used for video conferencing, online meetings, photography."),
    ("Touch Screen", "Display screen sensitive to touch. Used in smartphones, tablets, ATMs, kiosks."),
    ("Barcode Reader", "Reads barcodes on products. Used in supermarkets, libraries, inventory management."),
    ("Biometric Devices", "Fingerprint scanner, iris scanner, face recognition. Used for authentication and security.")
]

for device, desc in input_devices:
    story.append(Paragraph(f"<b>{device}:</b> {desc}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Output Devices", heading3_style))
story.append(Paragraph("Output devices display or output the processed information.", body_style))
story.append(Spacer(1, 0.1*inch))

output_devices = [
    ("Monitor", "Visual display unit (VDU). Types: CRT, LCD, LED, OLED. Displays text, images, videos."),
    ("Printer", "Produces hard copy. Types: Impact (Dot Matrix), Non-Impact (Inkjet - sprays ink, Laser - uses laser beam and toner, Thermal - uses heat)."),
    ("Speakers", "Output audio and sound. Used for music, system sounds, multimedia."),
    ("Headphones", "Personal audio output device. Used for private listening."),
    ("Projector", "Projects computer output onto large screen. Used in presentations, classrooms, theaters."),
    ("Plotter", "Prints large format graphics. Used for architectural drawings, engineering designs.")
]

for device, desc in output_devices:
    story.append(Paragraph(f"<b>{device}:</b> {desc}", body_style))

# 1.4 Computer Memory
story.append(Paragraph("1.4 Computer Memory", heading2_style))
story.append(Paragraph("Primary Memory (Main Memory)", heading3_style))

memory_types = [
    ("<b>RAM (Random Access Memory):</b>", "Volatile memory - loses data when power is off. Faster access speed. Used for storing currently running programs and data. Types: DRAM, SRAM, DDR RAM. Capacity: 4GB to 32GB in modern computers."),
    ("<b>ROM (Read Only Memory):</b>", "Non-volatile memory - retains data when power is off. Stores permanent instructions (BIOS, firmware). Types: PROM, EPROM, EEPROM."),
    ("<b>Cache Memory:</b>", "Very fast memory between CPU and RAM. Stores frequently accessed data. Levels: L1 (fastest, smallest), L2, L3 (slowest, largest).")
]

for mem_title, mem_desc in memory_types:
    story.append(Paragraph(mem_title, body_style))
    story.append(Paragraph(mem_desc, body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Secondary Memory (Storage Devices)", heading3_style))
storage_devices = [
    ("Hard Disk Drive (HDD)", "Magnetic storage, 500GB to 10TB capacity, slower than SSD, cheaper per GB, used for mass storage."),
    ("Solid State Drive (SSD)", "Flash memory based, faster than HDD, more expensive, no moving parts, more durable."),
    ("Optical Discs", "CD (700MB), DVD (4.7GB), Blu-ray (25GB), uses laser technology, portable."),
    ("USB Flash Drive", "Portable, 4GB to 256GB+, plug and play, uses flash memory."),
    ("Memory Card", "SD Card, MicroSD, used in cameras, phones, tablets."),
    ("Cloud Storage", "Online storage (Google Drive, Dropbox, OneDrive), accessible from anywhere.")
]

for storage, desc in storage_devices:
    story.append(Paragraph(f"<b>{storage}:</b> {desc}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Memory Units", heading3_style))
memory_units = [
    "Bit: Smallest unit (0 or 1)",
    "Nibble: 4 bits",
    "Byte: 8 bits (can represent one character)",
    "Kilobyte (KB): 1024 bytes",
    "Megabyte (MB): 1024 KB",
    "Gigabyte (GB): 1024 MB",
    "Terabyte (TB): 1024 GB",
    "Petabyte (PB): 1024 TB"
]
for unit in memory_units:
    story.append(Paragraph(f"• <b>{unit}</b>", body_style))

# 1.5 Number Systems
story.append(Paragraph("1.5 Number Systems", heading2_style))
story.append(Paragraph("Types of Number Systems", heading3_style))

number_systems = [
    ("<b>Decimal (Base 10):</b>", "Uses digits 0-9. Most commonly used by humans. Example: 125, 456, 789."),
    ("<b>Binary (Base 2):</b>", "Uses digits 0, 1. Used by computers internally. Example: 1010, 1101, 11110."),
    ("<b>Octal (Base 8):</b>", "Uses digits 0-7. Used in computer programming. Example: 12, 45, 77."),
    ("<b>Hexadecimal (Base 16):</b>", "Uses digits 0-9 and A-F (A=10, B=11, C=12, D=13, E=14, F=15). Used in memory addresses, color codes. Example: 1A, FF, A5.")
]

for ns_title, ns_desc in number_systems:
    story.append(Paragraph(ns_title, body_style))
    story.append(Paragraph(ns_desc, body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Number System Conversions", heading3_style))
conversions = [
    ("<b>Decimal to Binary:</b>", "Divide decimal number by 2, note the remainder, repeat until quotient becomes 0, write remainders in reverse order. Example: (25) base-10 = (11001) base-2"),
    ("<b>Binary to Decimal:</b>", "Multiply each bit by 2^position and sum all values. Example: (1101) base-2 = 1×2³ + 1×2² + 0×2¹ + 1×2⁰ = 8 + 4 + 0 + 1 = (13) base-10"),
    ("<b>Decimal to Octal:</b>", "Divide by 8, note remainders. Example: (65) base-10 = (101) base-8"),
    ("<b>Decimal to Hexadecimal:</b>", "Divide by 16, note remainders. Example: (255) base-10 = (FF) base-16")
]

for conv_title, conv_desc in conversions:
    story.append(Paragraph(conv_title, body_style))
    story.append(Paragraph(conv_desc, body_style))

# 1.6 Operating System
story.append(Paragraph("1.6 Operating System Basics", heading2_style))
story.append(Paragraph("What is an Operating System?", heading3_style))
story.append(Paragraph("""An Operating System (OS) is system software that manages computer hardware and software 
resources and provides common services for computer programs. It acts as an interface between the user and computer hardware.""", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Functions of Operating System:</b>", body_style))
os_functions = [
    "Process Management: Managing running programs",
    "Memory Management: Allocating and deallocating memory",
    "File Management: Organizing files and directories",
    "Device Management: Managing input/output devices",
    "Security: Protecting system from unauthorized access",
    "User Interface: Providing interface for user interaction"
]
for func in os_functions:
    story.append(Paragraph(f"• {func}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Types of Operating Systems", heading3_style))
os_types = [
    "Single User, Single Tasking: MS-DOS",
    "Single User, Multi Tasking: Windows, macOS",
    "Multi User: Unix, Linux",
    "Real Time OS: Used in embedded systems, industrial control",
    "Mobile OS: Android, iOS",
    "Network OS: Windows Server, Linux Server"
]
for os_type in os_types:
    story.append(Paragraph(f"• {os_type}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Popular Operating Systems", heading3_style))
popular_os = [
    ("<b>Windows:</b>", "Developed by Microsoft. User-friendly GUI. Versions: Windows 7, 8, 10, 11. Widely used for desktop computers."),
    ("<b>Linux:</b>", "Open-source OS. Free to use and modify. Distributions: Ubuntu, Fedora, Debian. Popular in servers and programming."),
    ("<b>macOS:</b>", "Developed by Apple. Used in Apple Mac computers. Known for design and performance."),
    ("<b>Android:</b>", "Open-source mobile OS by Google. Most popular mobile OS."),
    ("<b>iOS:</b>", "Apple's mobile OS. Used in iPhone and iPad.")
]

for os_name, os_desc in popular_os:
    story.append(Paragraph(os_name, body_style))
    story.append(Paragraph(os_desc, body_style))

# ============ TOPIC 2: MS OFFICE ============
story.append(Paragraph("2. Microsoft Office Suite", heading1_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("Introduction to MS Office", heading2_style))
story.append(Paragraph("""Microsoft Office is a suite of productivity applications developed by Microsoft. 
It includes various programs for creating documents, spreadsheets, presentations, and managing emails.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Main Components:</b>", body_style))
ms_components = [
    "MS Word - Word Processing",
    "MS Excel - Spreadsheet",
    "MS PowerPoint - Presentations",
    "MS Access - Database Management",
    "MS Outlook - Email and Calendar"
]
for comp in ms_components:
    story.append(Paragraph(f"• {comp}", body_style))
story.append(Spacer(1, 0.1*inch))

# 2.1 Microsoft Word
story.append(Paragraph("2.1 Microsoft Word", heading2_style))
story.append(Paragraph("""MS Word is a word processing application used for creating, editing, formatting, and printing documents.""", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("MS Word Interface Components", heading3_style))
word_interface = [
    "Title Bar: Shows document name",
    "Menu Bar/Ribbon: Contains tabs (Home, Insert, Design, Layout, References, Mailings, Review, View)",
    "Quick Access Toolbar: Frequently used commands (Save, Undo, Redo)",
    "Ruler: Shows margins and tab stops",
    "Work Area: Where you type document content",
    "Status Bar: Shows page number, word count, zoom",
    "Scroll Bars: Navigate through document"
]
for interface in word_interface:
    story.append(Paragraph(f"• {interface}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Document Creation and Basic Operations", heading3_style))
story.append(Paragraph("<b>Creating, Saving, Opening:</b>", body_style))
basic_ops = [
    "New Document: File → New → Blank Document (or Ctrl + N)",
    "Save Document: File → Save (Ctrl + S). Formats: .docx, .doc, .pdf, .txt",
    "Open Document: File → Open (Ctrl + O)"
]
for op in basic_ops:
    story.append(Paragraph(f"• {op}", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Basic Editing Shortcuts:</b>", body_style))
shortcuts = [
    "Cut: Ctrl + X",
    "Copy: Ctrl + C",
    "Paste: Ctrl + V",
    "Undo: Ctrl + Z",
    "Redo: Ctrl + Y",
    "Find: Ctrl + F",
    "Replace: Ctrl + H"
]
for shortcut in shortcuts:
    story.append(Paragraph(f"• {shortcut}", body_style))

story.append(Paragraph("Text Formatting", heading3_style))
story.append(Paragraph("<b>Font Formatting:</b>", body_style))
font_formatting = [
    "Font Type: Times New Roman, Arial, Calibri, etc.",
    "Font Size: 8, 10, 12, 14, 16, 18, etc.",
    "Font Style: Bold (Ctrl + B), Italic (Ctrl + I), Underline (Ctrl + U)",
    "Font Color: Change text color",
    "Highlight: Background color for text",
    "Effects: Strikethrough, Subscript, Superscript, Small Caps"
]
for formatting in font_formatting:
    story.append(Paragraph(f"• {formatting}", body_style))

story.append(Paragraph("<b>Paragraph Formatting:</b>", body_style))
para_formatting = [
    "Alignment: Left (Ctrl + L), Center (Ctrl + E), Right (Ctrl + R), Justify (Ctrl + J)",
    "Line Spacing: Single, 1.5, Double",
    "Indentation: Left indent, Right indent, First line indent",
    "Bullets: Create bulleted lists",
    "Numbering: Create numbered lists",
    "Borders and Shading: Add borders and background colors"
]
for formatting in para_formatting:
    story.append(Paragraph(f"• {formatting}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Page Setup and Layout", heading3_style))
page_setup = [
    "Margins: Normal, Narrow, Wide, Custom",
    "Orientation: Portrait, Landscape",
    "Size: A4, Letter, Legal, Custom",
    "Columns: Single column, Two columns, Three columns",
    "Headers and Footers: Add page numbers, date, document title"
]
for setup in page_setup:
    story.append(Paragraph(f"• {setup}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Inserting Objects", heading3_style))
insert_objects = [
    "Pictures: Insert images from computer",
    "Online Pictures: Search and insert images from web",
    "Shapes: Rectangles, circles, arrows, callouts",
    "SmartArt: Visual diagrams and graphics",
    "Charts: Insert graphs and charts",
    "Table: Create tables",
    "WordArt: Decorative text",
    "Symbols: Special characters and symbols",
    "Hyperlink: Link to websites or documents"
]
for obj in insert_objects:
    story.append(Paragraph(f"• {obj}", body_style))

story.append(Paragraph("Working with Tables", heading3_style))
story.append(Paragraph("<b>Creating Tables:</b> Insert → Table → Select rows and columns", body_style))
story.append(Spacer(1, 0.1*inch))
story.append(Paragraph("<b>Table Operations:</b>", body_style))
table_ops = [
    "Insert/Delete Rows and Columns",
    "Merge Cells: Combine multiple cells into one",
    "Split Cells: Divide cell into multiple cells",
    "Table Styles: Apply predefined table designs",
    "Borders and Shading: Customize table appearance",
    "AutoFit: Automatically adjust column width"
]
for op in table_ops:
    story.append(Paragraph(f"• {op}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Mail Merge", heading3_style))
story.append(Paragraph("""Mail Merge is used to create personalized letters, labels, envelopes, or emails 
to multiple recipients.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Steps for Mail Merge:</b>", body_style))
mail_merge_steps = [
    "Create Main Document: Letter with placeholders",
    "Create Data Source: Excel file or Word table with recipient information",
    "Start Mail Merge: Mailings → Start Mail Merge → Letters",
    "Select Recipients: Mailings → Select Recipients → Use Existing List",
    "Insert Merge Fields: Insert fields like Name, Address, City",
    "Preview Results: Check how documents will look",
    "Finish & Merge: Create individual documents or print"
]
for step in mail_merge_steps:
    story.append(Paragraph(f"{mail_merge_steps.index(step) + 1}. {step}", body_style))

story.append(Paragraph("<b>Applications:</b> Invitation letters, Salary slips, Certificates, Mailing labels", body_style))

# 2.2 Microsoft Excel
story.append(Paragraph("2.2 Microsoft Excel", heading2_style))
story.append(Paragraph("""MS Excel is a spreadsheet application used for calculations, data analysis, creating charts, 
and managing data in tabular form.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Uses of Excel:</b>", body_style))
excel_uses = [
    "Financial calculations and budgeting",
    "Data analysis and reporting",
    "Statistical analysis",
    "Creating charts and graphs",
    "Database management",
    "Project management"
]
for use in excel_uses:
    story.append(Paragraph(f"• {use}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Excel Interface Components", heading3_style))
excel_interface = [
    "Title Bar: Shows workbook name",
    "Ribbon: Tabs (Home, Insert, Page Layout, Formulas, Data, Review, View)",
    "Name Box: Shows active cell reference",
    "Formula Bar: Displays/edits cell contents",
    "Worksheet Area: Grid of cells",
    "Column Headers: A, B, C... (up to XFD)",
    "Row Headers: 1, 2, 3... (up to 1,048,576)",
    "Sheet Tabs: Navigate between worksheets",
    "Status Bar: Shows calculations (Sum, Average, Count)"
]
for interface in excel_interface:
    story.append(Paragraph(f"• {interface}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Basic Excel Concepts", heading3_style))
excel_concepts = [
    "<b>Cell:</b> Intersection of row and column (e.g., A1, B5)",
    "<b>Cell Reference:</b> Address of a cell",
    "<b>Range:</b> Group of cells (e.g., A1:D10)",
    "<b>Worksheet:</b> Single spreadsheet page",
    "<b>Workbook:</b> Excel file containing multiple worksheets",
    "<b>Active Cell:</b> Currently selected cell"
]
for concept in excel_concepts:
    story.append(Paragraph(f"• {concept}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Data Types in Excel", heading3_style))
data_types = [
    "Text: Alphabetic and alphanumeric data",
    "Numbers: Numeric values",
    "Dates: Date and time values",
    "Formulas: Start with = sign"
]
for dtype in data_types:
    story.append(Paragraph(f"• {dtype}", body_style))

story.append(Paragraph("Formatting in Excel", heading3_style))
excel_formatting = [
    "Number Format: General, Number, Currency, Date, Time, Percentage",
    "Font: Type, size, color, bold, italic, underline",
    "Alignment: Horizontal and vertical alignment, merge cells",
    "Borders: Add borders to cells",
    "Fill Color: Background color",
    "Cell Styles: Predefined formatting styles"
]
for formatting in excel_formatting:
    story.append(Paragraph(f"• {formatting}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Formulas and Functions", heading3_style))
story.append(Paragraph("<b>Creating Formulas:</b> Start with = sign. Basic operators: + (add), - (subtract), * (multiply), / (divide), ^ (power)", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Cell References:</b>", body_style))
cell_refs = [
    "Relative Reference: A1 (changes when copied)",
    "Absolute Reference: $A$1 (doesn't change when copied)",
    "Mixed Reference: $A1 or A$1 (partially fixed)"
]
for ref in cell_refs:
    story.append(Paragraph(f"• {ref}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Important Functions", heading3_style))
functions_data = [
    ["Function", "Description"],
    ["=SUM(range)", "Adds all numbers in range"],
    ["=AVERAGE(range)", "Calculates average"],
    ["=MAX(range)", "Returns maximum value"],
    ["=MIN(range)", "Returns minimum value"],
    ["=COUNT(range)", "Counts numbers in range"],
    ["=IF(condition, true, false)", "Logical test"],
    ["=VLOOKUP(value, table, col)", "Vertical lookup"],
    ["=CONCATENATE(text1, text2)", "Joins text strings"],
    ["=TODAY()", "Current date"],
    ["=ROUND(number, decimals)", "Rounds number"]
]

table = Table(functions_data, colWidths=[2.5*inch, 3.5*inch])
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 11),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('FONTSIZE', (0, 1), (-1, -1), 10),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
]))
story.append(table)
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("Charts and Graphs", heading3_style))
story.append(Paragraph("<b>Creating Charts:</b> Select data → Insert → Charts → Choose chart type", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Types of Charts:</b>", body_style))
chart_types = [
    "Column Chart: Vertical bars for comparing values",
    "Bar Chart: Horizontal bars",
    "Line Chart: Shows trends over time",
    "Pie Chart: Shows proportions/percentages",
    "Area Chart: Similar to line chart with filled area",
    "Scatter Chart: Shows relationship between two variables"
]
for chart in chart_types:
    story.append(Paragraph(f"• {chart}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Data Management", heading3_style))
data_mgmt = [
    "<b>Sorting:</b> Select data → Data → Sort. Sort A to Z (ascending) or Z to A (descending)",
    "<b>Filtering:</b> Select data → Data → Filter. Click dropdown arrow to select values to display",
    "<b>Data Validation:</b> Data → Data Validation. Restrict input to specific values, create dropdown lists"
]
for mgmt in data_mgmt:
    story.append(Paragraph(f"• {mgmt}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Pivot Tables", heading3_style))
story.append(Paragraph("""Pivot Table is a powerful tool for summarizing, analyzing, and presenting large amounts of data.""", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Creating Pivot Table:</b>", body_style))
pivot_steps = [
    "Select data range",
    "Insert → PivotTable",
    "Choose where to place pivot table",
    "Drag fields to Rows, Columns, Values, and Filters"
]
for step in pivot_steps:
    story.append(Paragraph(f"{pivot_steps.index(step) + 1}. {step}", body_style))

# 2.3 Microsoft PowerPoint
story.append(Paragraph("2.3 Microsoft PowerPoint", heading2_style))
story.append(Paragraph("""MS PowerPoint is a presentation software used to create slide shows for business presentations, 
educational purposes, and training sessions.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Uses of PowerPoint:</b>", body_style))
ppt_uses = [
    "Business presentations",
    "Educational lectures",
    "Training materials",
    "Product demonstrations",
    "Conference presentations",
    "Marketing pitches"
]
for use in ppt_uses:
    story.append(Paragraph(f"• {use}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("PowerPoint Interface", heading3_style))
ppt_interface = [
    "Slide Pane: Shows current slide",
    "Thumbnail Pane: Shows all slides in presentation",
    "Notes Pane: Add speaker notes",
    "Ribbon: Tabs (Home, Insert, Design, Transitions, Animations, Slide Show, Review, View)",
    "Status Bar: Slide number, theme, view options"
]
for interface in ppt_interface:
    story.append(Paragraph(f"• {interface}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Creating Presentations", heading3_style))
story.append(Paragraph("<b>Adding Slides:</b> Home → New Slide (or Ctrl + M)", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Slide Layouts:</b>", body_style))
slide_layouts = [
    "Title Slide: For presentation title",
    "Title and Content: Text with bullet points",
    "Two Content: Two columns of content",
    "Comparison: Compare two items",
    "Blank: Empty slide for custom design"
]
for layout in slide_layouts:
    story.append(Paragraph(f"• {layout}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Working with Text", heading3_style))
story.append(Paragraph("<b>Text Formatting:</b> Font, size, color, bold, italic, underline, bullets, numbering, alignment", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Best Practices for Text:</b>", body_style))
text_practices = [
    "Use large, readable fonts (minimum 24pt)",
    "Limit text per slide (6 lines, 6 words per line rule)",
    "Use bullet points instead of paragraphs",
    "High contrast between text and background",
    "Consistent font throughout presentation"
]
for practice in text_practices:
    story.append(Paragraph(f"• {practice}", body_style))

story.append(Paragraph("Inserting Objects", heading3_style))
ppt_objects = [
    "Pictures: Insert → Pictures (from device or online)",
    "Shapes: Rectangles, circles, arrows, flowchart symbols",
    "SmartArt: Visual diagrams (Process, Hierarchy, Cycle, Relationship)",
    "Charts: Column, Bar, Line, Pie charts",
    "Tables: Insert → Table, specify rows and columns",
    "Audio/Video: Insert → Audio/Video (from computer or online)"
]
for obj in ppt_objects:
    story.append(Paragraph(f"• {obj}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Design and Themes", heading3_style))
design_elements = [
    "<b>Applying Themes:</b> Design → Themes. Predefined color schemes and fonts",
    "<b>Slide Background:</b> Design → Format Background. Solid fill, gradient, picture, texture",
    "<b>Master Slide:</b> View → Slide Master. Edit master layout, changes apply to all slides"
]
for element in design_elements:
    story.append(Paragraph(f"• {element}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Transitions", heading3_style))
story.append(Paragraph("""Transitions are effects that occur when moving from one slide to another.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Applying Transitions:</b> Transitions → Choose effect (Fade, Push, Wipe, etc.) → Set duration → Apply to one or all slides", 
body_style))
story.append(Spacer(1, 0.15*inch))

story.append(Paragraph("Animations", heading3_style))
story.append(Paragraph("""Animations are effects applied to individual objects on a slide.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Types of Animations:</b>", body_style))
animation_types = [
    "Entrance: How object appears (Fade In, Fly In, Zoom)",
    "Emphasis: Draw attention to object (Pulse, Spin, Grow)",
    "Exit: How object disappears (Fade Out, Fly Out)",
    "Motion Paths: Object moves along path"
]
for anim in animation_types:
    story.append(Paragraph(f"• {anim}", body_style))

story.append(Paragraph("<b>Animation Tips:</b> Use sparingly, don't distract from content, be consistent", body_style))

story.append(Paragraph("Slide Show", heading3_style))
story.append(Paragraph("<b>Running Slide Show:</b>", body_style))
slideshow_ops = [
    "From Beginning: Slide Show → From Beginning (F5)",
    "From Current Slide: Shift + F5",
    "Click or arrow keys to navigate",
    "Press Esc to exit",
    "Presenter View: See notes, timer, next slide"
]
for op in slideshow_ops:
    story.append(Paragraph(f"• {op}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("<b>Presentation Tips:</b>", body_style))
presentation_tips = [
    "Practice beforehand",
    "Know your audience",
    "Maintain eye contact",
    "Speak clearly and confidently",
    "Be prepared for questions"
]
for tip in presentation_tips:
    story.append(Paragraph(f"• {tip}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Saving and Sharing", heading3_style))
save_formats = [
    ".pptx - PowerPoint Presentation (editable)",
    ".ppsx - PowerPoint Show (opens directly in slide show)",
    ".pdf - PDF format (non-editable)",
    ".mp4 - Video format"
]
for format in save_formats:
    story.append(Paragraph(f"• {format}", body_style))

# ============ TOPIC 3: INTERNET & EMAIL ============
story.append(Paragraph("3. Internet and Email", heading1_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("3.1 Introduction to Internet", heading2_style))
story.append(Paragraph("What is Internet?", heading3_style))
story.append(Paragraph("""The Internet is a global network of billions of computers and electronic devices connected together. 
It allows users to access information, communicate, and share data across the world.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Key Features:</b>", body_style))
internet_features = [
    "Global connectivity",
    "Information sharing",
    "Communication platform",
    "Resource sharing",
    "24/7 availability"
]
for feature in internet_features:
    story.append(Paragraph(f"• {feature}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("<b>Brief History:</b>", body_style))
internet_history = [
    "1969: ARPANET created (first network)",
    "1983: TCP/IP protocol adopted",
    "1989: Tim Berners-Lee invented World Wide Web (WWW)",
    "1990s: Internet became public and commercial",
    "2000s: Broadband, social media, mobile internet",
    "Present: IoT, cloud computing, 5G"
]
for history in internet_history:
    story.append(Paragraph(f"• {history}", body_style))

story.append(Paragraph("Internet Terminology", heading3_style))
terminology_data = [
    ["Term", "Definition"],
    ["WWW", "World Wide Web - system of interlinked hypertext documents"],
    ["Website", "Collection of web pages under one domain name"],
    ["Web Page", "Single document on the web, written in HTML"],
    ["URL", "Uniform Resource Locator - web address"],
    ["HTTP/HTTPS", "Protocol for web communication (Secure)"],
    ["Browser", "Software to access web pages (Chrome, Firefox, Edge)"],
    ["Search Engine", "Website to search information (Google, Bing)"],
    ["IP Address", "Numerical address of device on network"],
    ["ISP", "Internet Service Provider"],
    ["Download", "Transferring data from internet to computer"],
    ["Upload", "Transferring data from computer to internet"]
]

table2 = Table(terminology_data, colWidths=[1.8*inch, 4.2*inch])
table2.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 11),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('FONTSIZE', (0, 1), (-1, -1), 9),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
]))
story.append(table2)
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("3.2 Internet Connectivity", heading2_style))
story.append(Paragraph("Types of Internet Connections", heading3_style))

connection_types = [
    ("<b>Dial-Up:</b>", "Uses telephone line and modem. Speed: 56 Kbps. Obsolete technology."),
    ("<b>Broadband - DSL:</b>", "Digital Subscriber Line, uses telephone line. Speed: 1-100 Mbps."),
    ("<b>Broadband - Cable:</b>", "Uses cable TV infrastructure. Speed: 10-500 Mbps."),
    ("<b>Broadband - Fiber Optic:</b>", "Uses fiber optic cables. Speed: 100-1000 Mbps (fastest)."),
    ("<b>Wireless - Wi-Fi:</b>", "Wireless local area network, range 30-50 meters."),
    ("<b>Wireless - Mobile Data:</b>", "3G, 4G, 5G cellular networks."),
    ("<b>Satellite:</b>", "Internet via satellite, for remote areas.")
]

for conn_type, conn_desc in connection_types:
    story.append(Paragraph(conn_type, body_style))
    story.append(Paragraph(conn_desc, body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Required Equipment", heading3_style))
equipment = [
    "Computer/Device: PC, laptop, smartphone, tablet",
    "Modem: Converts digital signals to analog",
    "Router: Distributes internet to multiple devices",
    "ISP Account: Subscription from Internet Service Provider",
    "Browser: Software to access websites"
]
for equip in equipment:
    story.append(Paragraph(f"• {equip}", body_style))

story.append(Paragraph("3.3 Web Browsers", heading2_style))
story.append(Paragraph("""A web browser is software that allows users to access, retrieve, and display content from 
the World Wide Web.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Popular Browsers:</b>", body_style))
browsers = [
    "Google Chrome: Fast, most popular",
    "Mozilla Firefox: Open-source, privacy-focused",
    "Microsoft Edge: Built into Windows",
    "Safari: Default browser for Apple devices",
    "Opera: Fast, built-in VPN",
    "Brave: Privacy-focused, blocks ads"
]
for browser in browsers:
    story.append(Paragraph(f"• {browser}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Browser Features", heading3_style))
browser_features = [
    "<b>Address Bar:</b> Enter website URL, search directly",
    "<b>Navigation:</b> Back, Forward, Refresh, Home, Stop buttons",
    "<b>Bookmarks/Favorites:</b> Save frequently visited websites (Ctrl + D)",
    "<b>Tabs:</b> Open multiple websites (Ctrl + T: new tab, Ctrl + W: close tab)",
    "<b>History:</b> View recently visited websites (Ctrl + H)",
    "<b>Downloads:</b> View downloaded files (Ctrl + J)",
    "<b>Incognito/Private Mode:</b> Browse without saving history (Ctrl + Shift + N)",
    "<b>Extensions:</b> Additional features (ad blockers, password managers)"
]
for feature in browser_features:
    story.append(Paragraph(f"• {feature}", body_style))

story.append(Paragraph("3.4 Search Engines", heading2_style))
story.append(Paragraph("""A search engine is a web-based tool that helps users find information on the internet 
by searching for keywords or phrases.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Popular Search Engines:</b> Google, Bing, Yahoo, DuckDuckGo", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("Search Tips", heading3_style))
search_tips = [
    "Use specific keywords",
    "Use multiple keywords for better results",
    "Use quotation marks for exact phrases: \"climate change\"",
    "Use minus sign to exclude words: jaguar -car",
    "Use site: to search specific website: site:wikipedia.org python",
    "Use filetype: to find specific files: filetype:pdf resume",
    "Use OR for alternatives: laptop OR notebook"
]
for tip in search_tips:
    story.append(Paragraph(f"• {tip}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Types of Searches", heading3_style))
search_types = [
    "Web Search: General web pages",
    "Image Search: Photos and images",
    "Video Search: Videos",
    "News Search: Latest news articles",
    "Maps: Location and directions",
    "Shopping: Products and prices"
]
for stype in search_types:
    story.append(Paragraph(f"• {stype}", body_style))

story.append(Paragraph("3.5 Email (Electronic Mail)", heading2_style))
story.append(Paragraph("Introduction to Email", heading3_style))
story.append(Paragraph("""Email (Electronic Mail) is a method of exchanging messages between people using 
electronic devices over the internet.""", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Advantages of Email:</b>", body_style))
email_advantages = [
    "Instant delivery",
    "Cost-effective (free)",
    "24/7 accessibility",
    "Can send to multiple recipients",
    "Attach files and documents",
    "Maintain record of communication",
    "Professional communication"
]
for advantage in email_advantages:
    story.append(Paragraph(f"• {advantage}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("<b>Popular Email Services:</b> Gmail, Outlook, Yahoo Mail, ProtonMail", body_style))
story.append(Spacer(1, 0.15*inch))

story.append(Paragraph("Email Address Structure", heading3_style))
story.append(Paragraph("<b>Format:</b> username@domain.com", body_style))
story.append(Spacer(1, 0.05*inch))
story.append(Paragraph("• <b>Username:</b> Unique identifier (e.g., john.doe, student123)", body_style))
story.append(Paragraph("• <b>@ Symbol:</b> Separates username and domain", body_style))
story.append(Paragraph("• <b>Domain:</b> Email service provider (gmail.com, outlook.com)", body_style))

story.append(Paragraph("Creating Email Account", heading3_style))
story.append(Paragraph("<b>Steps to Create Gmail Account:</b>", body_style))
gmail_steps = [
    "Go to www.gmail.com",
    "Click 'Create account'",
    "Enter personal information (name, username, password)",
    "Verify phone number",
    "Add recovery email (optional)",
    "Enter date of birth and gender",
    "Agree to terms and conditions",
    "Complete verification"
]
for step in gmail_steps:
    story.append(Paragraph(f"{gmail_steps.index(step) + 1}. {step}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("<b>Password Guidelines:</b>", body_style))
password_tips = [
    "Minimum 8 characters",
    "Mix of uppercase and lowercase letters",
    "Include numbers and special characters",
    "Don't use personal information",
    "Don't use common words"
]
for tip in password_tips:
    story.append(Paragraph(f"• {tip}", body_style))

story.append(Paragraph("Email Interface (Gmail)", heading3_style))
gmail_interface = [
    "Compose Button: Create new email",
    "Inbox: Received emails",
    "Starred: Important emails marked with star",
    "Sent: Emails you have sent",
    "Drafts: Emails saved but not sent",
    "Spam: Unwanted/junk emails",
    "Trash: Deleted emails",
    "Labels/Folders: Organize emails",
    "Search Box: Search emails",
    "Settings: Configure email preferences"
]
for interface in gmail_interface:
    story.append(Paragraph(f"• {interface}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Composing and Sending Email", heading3_style))
story.append(Paragraph("<b>Steps to Send Email:</b>", body_style))
send_steps = [
    "Click 'Compose' button",
    "<b>To:</b> Recipient's email address",
    "<b>Cc:</b> Carbon Copy - send copy to others (everyone sees)",
    "<b>Bcc:</b> Blind Carbon Copy - hidden from other recipients",
    "<b>Subject:</b> Brief description of email",
    "<b>Body:</b> Main message content",
    "Format text (bold, italic, underline, font, color)",
    "Attach files if needed",
    "Click 'Send'"
]
for step in send_steps:
    story.append(Paragraph(f"{send_steps.index(step) + 1}. {step}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("<b>Email Writing Tips:</b>", body_style))
writing_tips = [
    "Clear and descriptive subject line",
    "Professional greeting",
    "Keep message concise and clear",
    "Use proper grammar and spelling",
    "Professional tone for formal emails",
    "Appropriate closing (Regards, Best wishes)",
    "Proofread before sending"
]
for tip in writing_tips:
    story.append(Paragraph(f"• {tip}", body_style))

story.append(Paragraph("Email Attachments", heading3_style))
story.append(Paragraph("<b>Attaching Files:</b> Click attachment icon (paperclip) → Browse and select file → Wait for upload. Maximum size: 25MB (Gmail)", 
body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("<b>Supported File Types:</b>", body_style))
file_types = [
    "Documents: .doc, .docx, .pdf, .txt",
    "Spreadsheets: .xls, .xlsx",
    "Presentations: .ppt, .pptx",
    "Images: .jpg, .png, .gif",
    "Compressed: .zip, .rar"
]
for ftype in file_types:
    story.append(Paragraph(f"• {ftype}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Managing Emails", heading3_style))
email_mgmt = [
    "<b>Reply:</b> Send response to sender",
    "<b>Reply All:</b> Send response to all recipients",
    "<b>Forward:</b> Send email to someone else",
    "<b>Star:</b> Mark important emails",
    "<b>Labels:</b> Categorize emails (Work, Personal)",
    "<b>Archive:</b> Remove from inbox but keep in account",
    "<b>Delete:</b> Move to trash",
    "<b>Search:</b> Use search box, advanced search options"
]
for mgmt in email_mgmt:
    story.append(Paragraph(f"• {mgmt}", body_style))

story.append(Paragraph("3.6 Internet Security and Safety", heading2_style))
story.append(Paragraph("Online Threats", heading3_style))
threats = [
    "Virus: Malicious program that replicates and damages system",
    "Malware: Software designed to harm computer",
    "Phishing: Fake emails/websites to steal information",
    "Spam: Unwanted bulk emails",
    "Ransomware: Locks files and demands payment",
    "Spyware: Secretly monitors user activity",
    "Identity Theft: Stealing personal information",
    "Hacking: Unauthorized access to systems"
]
for threat in threats:
    story.append(Paragraph(f"• {threat}", body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("Security Best Practices", heading3_style))
story.append(Paragraph("<b>Password Security:</b>", body_style))
password_security = [
    "Use strong, unique passwords",
    "Don't share passwords",
    "Change passwords regularly",
    "Use password manager",
    "Enable two-factor authentication (2FA)"
]
for security in password_security:
    story.append(Paragraph(f"• {security}", body_style))

story.append(Paragraph("<b>Safe Browsing:</b>", body_style))
safe_browsing = [
    "Check website security (HTTPS, padlock icon)",
    "Don't click suspicious links",
    "Avoid downloading from untrusted sites",
    "Use secure Wi-Fi networks",
    "Clear browser cache regularly"
]
for practice in safe_browsing:
    story.append(Paragraph(f"• {practice}", body_style))

story.append(Paragraph("<b>Email Security:</b>", body_style))
email_security = [
    "Don't open emails from unknown senders",
    "Don't click links in suspicious emails",
    "Verify sender before sharing information",
    "Don't download unknown attachments",
    "Report spam and phishing"
]
for security in email_security:
    story.append(Paragraph(f"• {security}", body_style))

story.append(Paragraph("<b>Protection Tools:</b>", body_style))
protection = [
    "Install antivirus software (Windows Defender, Norton)",
    "Keep software updated",
    "Use firewall",
    "Regular system scans",
    "Backup important data"
]
for tool in protection:
    story.append(Paragraph(f"• {tool}", body_style))

story.append(Paragraph("3.7 Internet Applications and Services", heading2_style))
applications = [
    ("<b>E-Commerce:</b>", "Online buying and selling (Amazon, Flipkart). Online payment methods."),
    ("<b>Social Media:</b>", "Networking and content sharing (Facebook, Instagram, Twitter, LinkedIn)."),
    ("<b>Online Banking:</b>", "Managing bank accounts through internet. Check balance, transfer money, pay bills."),
    ("<b>Video Conferencing:</b>", "Real-time video communication (Zoom, Microsoft Teams, Google Meet)."),
    ("<b>Cloud Storage:</b>", "Store files online (Google Drive, Dropbox, OneDrive). Access from anywhere."),
    ("<b>Streaming Services:</b>", "Watch videos and listen to music online (YouTube, Netflix, Spotify)."),
    ("<b>Online Education:</b>", "Learning through internet platforms (Coursera, Udemy, Khan Academy).")
]

for app_title, app_desc in applications:
    story.append(Paragraph(app_title, body_style))
    story.append(Paragraph(app_desc, body_style))
story.append(Spacer(1, 0.08*inch))

story.append(Paragraph("3.8 Internet Etiquette (Netiquette)", heading2_style))
story.append(Paragraph("<b>Guidelines for Online Behavior:</b>", body_style))
netiquette = [
    "Be respectful and courteous",
    "Use appropriate language",
    "Don't type in ALL CAPS (considered shouting)",
    "Respect others' privacy",
    "Don't spam or send chain letters",
    "Give credit for others' work",
    "Think before posting",
    "Don't spread false information",
    "Respect copyright laws",
    "Be patient with beginners"
]
for rule in netiquette:
    story.append(Paragraph(f"• {rule}", body_style))

# Conclusion
story.append(Paragraph("Conclusion", heading1_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("""These comprehensive notes cover the fundamental topics of the DCA course Part 1:""", body_style))
story.append(Spacer(1, 0.08*inch))

conclusion_points = [
    "<b>Computer Fundamentals:</b> Understanding hardware, software, memory, number systems, and operating systems",
    "<b>MS Office:</b> Proficiency in Word (documents), Excel (spreadsheets, formulas, charts), and PowerPoint (presentations)",
    "<b>Internet & Email:</b> Using browsers, search engines, email communication, and online safety"
]
for point in conclusion_points:
    story.append(Paragraph(f"• {point}", body_style))
story.append(Spacer(1, 0.1*inch))

story.append(Paragraph("""These skills form the foundation of computer literacy and are essential for both academic 
and professional success. Regular practice with these applications will enhance your proficiency and productivity.""", 
body_style))
story.append(Spacer(1, 0.3*inch))

story.append(Paragraph("<b>For more information, contact:</b>", body_style))
story.append(Paragraph("<b>Prathibha Institute</b>", 
                      ParagraphStyle('contact', parent=styles['Normal'], fontSize=13, 
                                    fontName='Helvetica-Bold')))
story.append(Paragraph("Computer Training Center", body_style))

story.append(Spacer(1, 1*inch))
story.append(Paragraph("- End of Notes Part 1 -", 
                      ParagraphStyle('end', parent=styles['Normal'], fontSize=11, 
                                    alignment=TA_CENTER, textColor=colors.grey, fontName='Helvetica-Oblique')))

# Build PDF
print("Building PDF...")
doc.build(story, canvasmaker=NumberedCanvas)
print(f"✓ PDF generated successfully: {pdf_file}")
print(f"✓ Location: {os.path.abspath(pdf_file)}")
