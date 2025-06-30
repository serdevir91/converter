import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
import pythoncom

# Required libraries:
# pip install customtkinter pdf2docx docx2pdf pypiwin32 Pillow pillow-heif imgkit
# The 'pillow-heif' library is required for HEIC support.
# 'imgkit' and the 'wkhtmltoimage' tool are required for HTML to PNG conversion.
# https://wkhtmltopdf.org/downloads.html for wkhtmltoimage tool.
# Check for HEIC support
try:
    import pyheif
    pyheif.register_heif_opener()
    HEIC_SUPPORT = True
except ImportError:
    HEIC_SUPPORT = False

# Check for HTML support
try:
    import imgkit
    HTML_SUPPORT = True
except ImportError:
    HTML_SUPPORT = False

# Conversion libraries
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

# Main application class
class FileConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Window Settings ---
        self.title("CONVERTER")
        self.geometry("850x800") # Window size increased
        self.minsize(750, 600)
        
        # Theme and color settings
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.conversion_buttons = []
        self.setup_ui()

    def setup_ui(self):
        """Creates the application's user interface."""
        
        self.title_label = ctk.CTkLabel(self, text="CONVERTER", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=(20, 10))

        # Make the main frame scrollable
        self.scrollable_frame = ctk.CTkScrollableFrame(self)
        self.scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.scrollable_frame.grid_columnconfigure((0, 1), weight=1)

        # --- Add Conversion Cards ---
        self._create_conversion_card(
            parent=self.scrollable_frame, row=0, column=0, icon="📄 → 📝",
            title="PDF to Word", description="Convert multiple .pdf files to .docx format.",
            command=self.start_pdf_to_word_conversion
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=0, column=1, icon="📝 → 📄",
            title="Word to PDF", description="Convert multiple .docx files to .pdf format.",
            command=self.start_word_to_pdf_conversion
        )
        self._create_image_card(row=1, column=0, icon="🖼️ → JPG", title="PNG to JPG", 
                                in_format="PNG", out_format="JPG", save_kwargs={'format': 'JPEG', 'quality': 95}, convert_mode='RGB')
        self._create_image_card(row=1, column=1, icon="JPG → 🖼️", title="JPG to PNG",
                                in_format="JPG", out_format="PNG", save_kwargs={'format': 'PNG'})
        self._create_image_card(row=2, column=0, icon="🖼️ → WEBP", title="Image to WEBP",
                                in_format="PNG/JPG", out_format="WEBP", save_kwargs={'format': 'WEBP', 'quality': 85})
        self._create_image_card(row=2, column=1, icon="WEBP → 🖼️", title="WEBP to PNG",
                                in_format="WEBP", out_format="PNG", save_kwargs={'format': 'PNG'})
        self._create_image_card(row=3, column=0, icon="🖼️ → ICO", title="Image to ICO",
                                in_format="PNG/JPG", out_format="ICO", save_kwargs={'format': 'ICO', 'sizes': [(16,16), (32,32), (48,48), (64,64)]})
        self._create_image_card(row=3, column=1, icon="🎨 → 🔳", title="Image to Grayscale",
                                in_format="Image", out_format="PNG", save_kwargs={'format': 'PNG'}, convert_mode='L')
        self._create_image_card(row=4, column=0, icon="🍏 → 🖼️", title="HEIC to JPG",
                                in_format="HEIC", out_format="JPG", save_kwargs={'format': 'JPEG', 'quality': 95}, convert_mode='RGB', requires_heic=True)
        self._create_image_card(row=4, column=1, icon="🖼️ → BMP", title="Image to BMP",
                                in_format="PNG/JPG", out_format="BMP", save_kwargs={'format': 'BMP'})
        
        # HTML to PNG card
        html_command = self.start_html_to_png_conversion
        if not HTML_SUPPORT:
            html_command = self.show_html_error

        self._create_conversion_card(
            parent=self.scrollable_frame, row=5, column=0, icon="🌐 → 🖼️",
            title="HTML to PNG", description="Convert multiple .html files to .png images.",
            command=html_command
        )

        # Status label and progress bar
        self.status_label = ctk.CTkLabel(self, text="Please select an operation.", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=(5, 5))
        
        self.progress_bar = ctk.CTkProgressBar(self, mode='determinate')
        self.progress_bar.set(0)

    def _create_conversion_card(self, parent, row, column, icon, title, description, command):
        """Creates general-purpose conversion cards."""
        card_frame = ctk.CTkFrame(parent, corner_radius=15)
        card_frame.grid(row=row, column=column, padx=10, pady=10, sticky="nsew")
        inner_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        inner_frame.pack(padx=10, pady=10, fill="both", expand=True)
        ctk.CTkLabel(inner_frame, text=icon, font=ctk.CTkFont(size=40)).pack(pady=5)
        ctk.CTkLabel(inner_frame, text=title, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        ctk.CTkLabel(inner_frame, text=description, wraplength=280, justify="center").pack(pady=5, fill="x")
        button = ctk.CTkButton(inner_frame, text="Select Files", command=command, height=35)
        button.pack(pady=(15, 5), padx=20, fill="x")
        self.conversion_buttons.append(button)

    def _create_image_card(self, row, column, icon, title, in_format, out_format, save_kwargs, convert_mode=None, requires_heic=False):
        """Creates special cards for image conversion."""
        filetypes = {
            "PNG": [("PNG Images", "*.png")],
            "JPG": [("JPEG Images", "*.jpg *.jpeg")],
            "WEBP": [("WEBP Images", "*.webp")],
            "HEIC": [("HEIC Images", "*.heic *.heif")],
            "PNG/JPG": [("Image Files", "*.png *.jpg *.jpeg")],
            "Image": [("Image Files", "*.png *.jpg *.jpeg *.bmp")],
        }
        
        command = lambda: self.start_image_conversion(
            title_open=f"Select {in_format} Files", filetypes_open=filetypes.get(in_format),
            out_format=out_format, save_kwargs=save_kwargs, convert_mode=convert_mode
        )

        if requires_heic and not HEIC_SUPPORT:
            command = self.show_heic_error

        self._create_conversion_card(
            parent=self.scrollable_frame, row=row, column=column, icon=icon, title=title,
            description=f"Convert multiple .{in_format.lower()} files to .{out_format.lower()} format.",
            command=command
        )
            
    # --- UI Update Methods ---
    def update_status(self, message, color="white"):
        self.after(0, lambda: self.status_label.configure(text=message, text_color=color))

    def update_progress(self, value):
        self.after(0, lambda: self.progress_bar.set(value))

    def lock_buttons(self, locked=True):
        state = "disabled" if locked else "normal"
        self.after(0, lambda: [btn.configure(state=state) for btn in self.conversion_buttons])

    def show_progress_bar(self):
        self.after(0, lambda: self.progress_bar.pack(pady=(0, 10), padx=20, fill="x"))

    def hide_progress_bar(self):
        self.after(0, lambda: self.progress_bar.pack_forget())

    def show_heic_error(self):
        messagebox.showerror("Missing Library", "The 'pillow-heif' library is required for this feature.\n\nTo install, run: pip install pillow-heif")
    
    def show_html_error(self):
        messagebox.showerror(
            "Missing Dependency", 
            "Additional setup is required for HTML conversion:\n\n"
            "1. Install the Python library:\n"
            "   pip install imgkit\n\n"
            "2. Download and install the wkhtmltoimage tool and add it to your system's PATH.\n"
            "   (You can download it from its official website)"
        )

    # --- Batch Process Initiator ---
    def _initiate_batch_process(self, worker_function, file_types, title, **kwargs):
        """Manages file selection, folder selection, and thread initiation."""
        input_paths = filedialog.askopenfilenames(title=title, filetypes=file_types)
        if not input_paths: return

        output_dir = filedialog.askdirectory(title="Select Output Folder for Converted Files")
        if not output_dir: return

        all_args = {'input_paths': input_paths, 'output_dir': output_dir, **kwargs}
        threading.Thread(target=worker_function, kwargs=all_args, daemon=True).start()

    # --- Conversion Initiators ---
    def start_pdf_to_word_conversion(self):
        self._initiate_batch_process(self.convert_documents, [("PDF Files", "*.pdf")], "Select PDF Files to Convert", out_ext=".docx")
    
    def start_word_to_pdf_conversion(self):
        self._initiate_batch_process(self.convert_documents, [("Word Documents", "*.docx *.doc")], "Select Word Files to Convert", out_ext=".pdf")

    def start_image_conversion(self, **kwargs):
        title = kwargs.pop('title_open')
        filetypes = kwargs.pop('filetypes_open')
        self._initiate_batch_process(self.convert_images, filetypes, title, **kwargs)

    def start_html_to_png_conversion(self):
        self._initiate_batch_process(self.convert_html_to_png, [("HTML Files", "*.html *.htm")], "Select HTML Files to Convert", out_ext=".png")

    # --- Conversion Logic ---
    def convert_images(self, input_paths, output_dir, out_format, save_kwargs, convert_mode=None):
        """Converts multiple images."""
        self.lock_buttons(True)
        self.show_progress_bar()
        self.update_progress(0)
        success_count = 0
        total_files = len(input_paths)

        for i, path in enumerate(input_paths):
            self.update_status(f"Converting: {i+1}/{total_files} - {os.path.basename(path)}", "yellow")
            base_name = os.path.splitext(os.path.basename(path))[0]
            output_path = os.path.join(output_dir, f"{base_name}.{out_format.lower()}")
            try:
                with Image.open(path) as img:
                    if convert_mode and img.mode != convert_mode:
                        img = img.convert(convert_mode)
                    img.save(output_path, **save_kwargs)
                success_count += 1
            except Exception as e:
                print(f"Error ({path}): {e}")
            self.update_progress((i + 1) / total_files)

        self.hide_progress_bar()
        self.lock_buttons(False)
        self.update_status(f"Process complete: {success_count}/{total_files} files converted successfully.", "lightgreen")
        messagebox.showinfo("Process Complete", f"{success_count}/{total_files} files were converted successfully.")

    def convert_documents(self, input_paths, output_dir, out_ext):
        """Converts multiple PDF or Word documents."""
        self.lock_buttons(True)
        self.show_progress_bar()
        self.update_progress(0)
        success_count = 0
        total_files = len(input_paths)
        is_to_pdf = (out_ext == ".pdf")

        for i, path in enumerate(input_paths):
            self.update_status(f"Converting: {i+1}/{total_files} - {os.path.basename(path)}", "yellow")
            base_name = os.path.splitext(os.path.basename(path))[0]
            output_path = os.path.join(output_dir, f"{base_name}{out_ext}")
            
            try:
                if is_to_pdf:
                    pythoncom.CoInitialize()
                    convert(path, output_path)
                    pythoncom.CoUninitialize()
                else: # PDF to Word
                    cv = Converter(path)
                    cv.convert(output_path, start=0, end=None)
                    cv.close()
                
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    success_count += 1
            except Exception as e:
                print(f"Error ({path}): {e}")
                if is_to_pdf:
                    pythoncom.CoUninitialize()
            
            self.update_progress((i + 1) / total_files)

        self.hide_progress_bar()
        self.lock_buttons(False)
        self.update_status(f"Process complete: {success_count}/{total_files} files converted successfully.", "lightgreen")
        messagebox.showinfo("Process Complete", f"{success_count}/{total_files} files were converted successfully.")

    def convert_html_to_png(self, input_paths, output_dir, out_ext=".png"):
        """Converts multiple HTML files to PNG images."""
        self.lock_buttons(True)
        self.show_progress_bar()
        self.update_progress(0)
        success_count = 0
        total_files = len(input_paths)

        for i, path in enumerate(input_paths):
            self.update_status(f"Converting: {i+1}/{total_files} - {os.path.basename(path)}", "yellow")
            base_name = os.path.splitext(os.path.basename(path))[0]
            output_path = os.path.join(output_dir, f"{base_name}{out_ext}")
            
            try:
                imgkit.from_file(path, output_path, options={'enable-local-file-access': None})
                success_count += 1
            except OSError as e:
                if "No wkhtmltoimage executable found" in str(e):
                    self.after(0, lambda: messagebox.showerror("Error", "wkhtmltoimage tool not found. Please ensure it is installed and in your system's PATH."))
                    break # End the loop if there is an error
                else:
                    print(f"Error ({path}): {e}")
            except Exception as e:
                print(f"Error ({path}): {e}")

            self.update_progress((i + 1) / total_files)

        self.hide_progress_bar()
        self.lock_buttons(False)
        self.update_status(f"Process complete: {success_count}/{total_files} files converted successfully.", "lightgreen")
        if success_count > 0 or total_files == 0:
            messagebox.showinfo("Process Complete", f"{success_count}/{total_files} files were converted successfully.")

if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()
