import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
import pythoncom

# Gerekli kütüphaneler:
# pip install customtkinter pdf2docx docx2pdf pypiwin32 Pillow pillow-heif
# HEIC desteği için 'pillow-heif' kütüphanesi gereklidir.

# HEIC desteğini kontrol et ve Pillow'a kaydet
try:
    import pyheif
    pyheif.register_heif_opener()
    HEIC_SUPPORT = True
except ImportError:
    HEIC_SUPPORT = False

# Dönüştürme kütüphaneleri
from pdf2docx import Converter
from docx2pdf import convert
from PIL import Image

# Uygulamanın ana sınıfı
class FileConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Pencere Ayarları ---
        self.title("Converter")
        self.geometry("850x750") # Pencere boyutu artırıldı
        self.minsize(750, 550)
        
        # Tema ve renk ayarları
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.conversion_buttons = []

        # --- Arayüz Elemanları ---
        self.setup_ui()

    def setup_ui(self):
        """Uygulamanın arayüzünü oluşturur."""
        
        self.title_label = ctk.CTkLabel(self, text="CONVERTER", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=(20, 10))

        # Ana çerçeveyi kaydırılabilir yap (başlık yazısı kaldırıldı)
        self.scrollable_frame = ctk.CTkScrollableFrame(self)
        self.scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)
        self.scrollable_frame.grid_columnconfigure((0, 1), weight=1)

        # --- Dönüşüm Kartlarını Ekle ---
        # Mevcut kartlar
        self._create_conversion_card(
            parent=self.scrollable_frame, row=0, column=0, icon="📄 → 📝",
            title="PDF'den Word'e", description=".pdf uzantılı dosyayı .docx formatına dönüştürün.",
            command=self.start_pdf_to_word_conversion
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=0, column=1, icon="📝 → 📄",
            title="Word'den PDF'e", description=".docx uzantılı dosyayı .pdf formatına dönüştürün.",
            command=self.start_word_to_pdf_conversion
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=1, column=0, icon="🖼️ → JPG",
            title="PNG'den JPG'e", description=".png uzantılı resimleri .jpg formatına dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="PNG Seçin", filetypes_open=[("PNG Resimleri", "*.png")],
                title_save="JPG Kaydet", default_ext_save=".jpg", filetypes_save=[("JPEG Resimleri", "*.jpg")],
                save_kwargs={'format': 'JPEG', 'quality': 95}, convert_mode='RGB'
            )
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=1, column=1, icon="JPG → 🖼️",
            title="JPG'den PNG'ye", description=".jpg uzantılı resimleri .png formatına dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="JPG Seçin", filetypes_open=[("JPEG Resimleri", "*.jpg *.jpeg")],
                title_save="PNG Kaydet", default_ext_save=".png", filetypes_save=[("PNG Resimleri", "*.png")],
                save_kwargs={'format': 'PNG'}
            )
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=2, column=0, icon="🖼️ → WEBP",
            title="Resimden WEBP'ye", description=".png veya .jpg resimleri .webp formatına dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="Resim Seçin", filetypes_open=[("Resim Dosyaları", "*.png *.jpg *.jpeg")],
                title_save="WEBP Kaydet", default_ext_save=".webp", filetypes_save=[("WEBP Resimleri", "*.webp")],
                save_kwargs={'format': 'WEBP', 'quality': 85}
            )
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=2, column=1, icon="WEBP → 🖼️",
            title="WEBP'den PNG'ye", description=".webp uzantılı resimleri .png formatına dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="WEBP Seçin", filetypes_open=[("WEBP Resimleri", "*.webp")],
                title_save="PNG Kaydet", default_ext_save=".png", filetypes_save=[("PNG Resimleri", "*.png")],
                save_kwargs={'format': 'PNG'}
            )
        )
        self._create_conversion_card(
            parent=self.scrollable_frame, row=3, column=0, icon="🖼️ → ICO",
            title="Resimden ICO'ya", description="Resimleri .ico formatında ikona dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="Resim Seçin", filetypes_open=[("Resim Dosyaları", "*.png *.jpg *.jpeg")],
                title_save="ICO Kaydet", default_ext_save=".ico", filetypes_save=[("İkon Dosyaları", "*.ico")],
                save_kwargs={'format': 'ICO', 'sizes': [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]}
            )
        )

        # --- YENİ EKLENEN KARTLAR ---
        self._create_conversion_card(
            parent=self.scrollable_frame, row=3, column=1, icon="🎨 → 🔳",
            title="Resmi Siyah-Beyaz Yap", description="Bir resmi siyah-beyaz (gri tonlamalı) yapın.",
            command=lambda: self.start_image_conversion(
                title_open="Dönüştürülecek Resmi Seçin", filetypes_open=[("Resim Dosyaları", "*.png *.jpg *.jpeg *.bmp")],
                title_save="Siyah-Beyaz Resmi Kaydet", default_ext_save=".png", filetypes_save=[("PNG Resimleri", "*.png")],
                save_kwargs={'format': 'PNG'},
                convert_mode='L'
            )
        )

        # HEIC desteği kontrolü ve kartın oluşturulması
        heic_command = lambda: self.start_image_conversion(
            title_open="HEIC/HEIF Resim Seçin", filetypes_open=[("HEIC/HEIF Resimleri", "*.heic *.heif")],
            title_save="JPG Olarak Kaydet", default_ext_save=".jpg", filetypes_save=[("JPEG Resimleri", "*.jpg")],
            save_kwargs={'format': 'JPEG', 'quality': 95},
            convert_mode='RGB'
        )
        if not HEIC_SUPPORT:
            heic_command = self.show_heic_error
        
        self._create_conversion_card(
            parent=self.scrollable_frame, row=4, column=0, icon="🍏 → 🖼️",
            title="HEIC'den JPG'e", description="iPhone resimlerini (.heic) .jpg formatına dönüştürün.",
            command=heic_command
        )
        
        self._create_conversion_card(
            parent=self.scrollable_frame, row=4, column=1, icon="🖼️ → BMP",
            title="Resimden BMP'ye", description="PNG veya JPG resimleri .bmp formatına dönüştürün.",
            command=lambda: self.start_image_conversion(
                title_open="Resim Seçin", filetypes_open=[("Resim Dosyaları", "*.png *.jpg *.jpeg")],
                title_save="BMP Olarak Kaydet", default_ext_save=".bmp", filetypes_save=[("BMP Resimleri", "*.bmp")],
                save_kwargs={'format': 'BMP'}
            )
        )

        # Durum bilgisi ve ilerleme çubuğu
        self.status_label = ctk.CTkLabel(self, text="Lütfen bir işlem seçin.", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=(5, 5))
        
        self.progress_bar = ctk.CTkProgressBar(self, mode='indeterminate')

    def _create_conversion_card(self, parent, row, column, icon, title, description, command):
        """Dönüştürme seçenekleri için standart bir kart oluşturur."""
        card_frame = ctk.CTkFrame(parent, corner_radius=15)
        card_frame.grid(row=row, column=column, padx=10, pady=10, sticky="nsew")
        inner_frame = ctk.CTkFrame(card_frame, fg_color="transparent")
        inner_frame.pack(padx=10, pady=10, fill="both", expand=True)
        ctk.CTkLabel(inner_frame, text=icon, font=ctk.CTkFont(size=40)).pack(pady=5)
        ctk.CTkLabel(inner_frame, text=title, font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(10, 5))
        ctk.CTkLabel(inner_frame, text=description, wraplength=280, justify="center").pack(pady=5, fill="x")
        button = ctk.CTkButton(inner_frame, text="Başlat", command=command, height=35)
        button.pack(pady=(15, 5), padx=20, fill="x")
        self.conversion_buttons.append(button)
            
    # --- UI Güncelleme Metotları ---
    def update_status(self, message, color="white"):
        self.after(0, lambda: self.status_label.configure(text=message, text_color=color))

    def lock_buttons(self, locked=True):
        state = "disabled" if locked else "normal"
        self.after(0, lambda: [btn.configure(state=state) for btn in self.conversion_buttons])

    def show_progress_bar(self, indeterminate=True):
        def _show():
            self.progress_bar.pack(pady=(0, 10), padx=20, fill="x")
            if indeterminate:
                self.progress_bar.start()
        self.after(0, _show)

    def hide_progress_bar(self):
        self.after(0, lambda: (self.progress_bar.stop(), self.progress_bar.pack_forget()))

    # --- YENİ EKLENEN METOT ---
    def show_heic_error(self):
        """HEIC kütüphanesi bulunamadığında hata mesajı gösterir."""
        messagebox.showerror(
            "Eksik Kütüphane", 
            "HEIC desteği bulunamadı.\n\n"
            "Bu özellik için 'pillow-heif' kütüphanesi gereklidir.\n"
            "Lütfen terminal veya komut istemine aşağıdaki komutu yazarak kurun:\n\n"
            "pip install pillow-heif"
        )

    # --- Dönüştürme Başlatıcıları ---
    def start_pdf_to_word_conversion(self):
        threading.Thread(target=self.convert_pdf_to_word, daemon=True).start()

    def start_word_to_pdf_conversion(self):
        threading.Thread(target=self.convert_word_to_pdf, daemon=True).start()

    def start_image_conversion(self, **kwargs):
        threading.Thread(target=self.convert_image, kwargs=kwargs, daemon=True).start()

    # --- Dönüştürme Mantığı ---
    def convert_image(self, title_open, filetypes_open, title_save, default_ext_save, filetypes_save, save_kwargs, convert_mode=None):
        input_path = filedialog.askopenfilename(title=title_open, filetypes=filetypes_open)
        if not input_path: return
        
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = filedialog.asksaveasfilename(title=title_save, initialfile=f"{base_name}{default_ext_save}", defaultextension=default_ext_save, filetypes=filetypes_save)
        if not output_path: return

        self.lock_buttons(True)
        self.update_status(f"{save_kwargs.get('format', '')} formatına dönüştürülüyor...", "yellow")
        self.show_progress_bar()

        try:
            img = Image.open(input_path)
            if convert_mode and img.mode != convert_mode:
                img = img.convert(convert_mode)
            img.save(output_path, **save_kwargs)
            self.update_status(f"Başarıyla dönüştürüldü: {os.path.basename(output_path)}", "lightgreen")
        except Exception as e:
            self.update_status("Dönüştürme başarısız oldu!", "lightcoral")
            messagebox.showerror("Hata", f"Resim dönüştürme sırasında bir hata oluştu:\n{e}")
        finally:
            self.hide_progress_bar()
            self.lock_buttons(False)

    def convert_pdf_to_word(self):
        input_path = filedialog.askopenfilename(title="Dönüştürülecek PDF Dosyasını Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not input_path: return

        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = filedialog.asksaveasfilename(title="Word Dosyasını Kaydet", initialfile=f"{base_name}.docx", defaultextension=".docx", filetypes=[("Word Belgeleri", "*.docx")])
        if not output_path: return

        self.lock_buttons(True)
        self.update_status("PDF'den Word'e dönüştürülüyor...", "yellow")
        self.show_progress_bar()

        try:
            cv = Converter(input_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            self.update_status(f"Başarıyla dönüştürüldü: {os.path.basename(output_path)}", "lightgreen")
        except Exception as e:
            self.update_status("Bir hata oluştu!", "lightcoral")
            messagebox.showerror("Hata", f"Dönüştürme sırasında bir hata oluştu:\n{e}")
        finally:
            self.hide_progress_bar()
            self.lock_buttons(False)

    def convert_word_to_pdf(self):
        input_path = filedialog.askopenfilename(title="Dönüştürülecek Word Dosyasını Seçin", filetypes=[("Word Belgeleri", "*.docx *.doc")])
        if not input_path: return

        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = filedialog.asksaveasfilename(title="PDF Dosyasını Kaydet", initialfile=f"{base_name}.pdf", defaultextension=".pdf", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not output_path: return
            
        self.lock_buttons(True)
        self.update_status("Word'den PDF'e dönüştürülüyor...", "yellow")
        self.show_progress_bar()
        
        error_that_occurred = None
        try:
            pythoncom.CoInitialize()
            convert(input_path, output_path)
        except Exception as e:
            error_that_occurred = e
        finally:
            pythoncom.CoUninitialize()

        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            self.update_status(f"Başarıyla dönüştürüldü: {os.path.basename(output_path)}", "lightgreen")
        else:
            error_message = "Bu işlem için bilgisayarınızda Microsoft Word'ün kurulu olması veya yanıt veriyor olması gerekmektedir."
            self.update_status("Dönüştürme başarısız oldu!", "lightcoral")
            messagebox.showerror("Hata", error_message)
        
        self.hide_progress_bar()
        self.lock_buttons(False)

if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()
