# -*- coding: utf-8 -*-
"""
نظام طباعة جرعات الأدوية والروشتات لملصقات Xprinter (38x25 mm)
يدعم:
1. لوحة تحكم كاملة لإضافة وحذف الجرعات الشائعة وحفظها محلياً (JSON).
2. تجهيز قائمة لروشتة كاملة (Batch Printing) لطباعة عدة أصناف دفعة واحدة.
3. تحديد عدد النسخ لكل جرعة (تكرار الملصق للعلب المتعددة).
4. الإدخال اليدوي المباشر أو الاختيار السريع.
"""

import sys
import os
import json
import datetime
from PIL import Image, ImageDraw, ImageFont, ImageTk
import arabic_reshaper
from bidi.algorithm import get_display

# مكتبات الواجهة الرسومية
import tkinter as tk
from tkinter import ttk, messagebox

# مكتبات الطباعة على ويندوز
try:
    import win32print
    import win32ui
    from PIL import ImageWin
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

def resource_path(relative_path):
    """يحدد مسار الملف سواء كان البرنامج شغال كسكريبت أو كملف exe مجمّع (PyInstaller)"""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

APP_ICON_FILE = resource_path("app_icon.ico")
APP_LOGO_FILE = resource_path("app_logo.png")

DOSAGES_FILE = "dosages_config.json"
DEFAULT_DOSAGES = [
    "قرص بعد الأكل 3 مرات",
    "قرص كل 12 ساعة بعد الأكل",
    "قرص يومياً على الريق",
    "قرص يومياً قبل النوم",
    "ملعقة بعد الأكل 3 مرات",
    "كبسولة بعد الإفطار والعشاء",
    "نقط بالفم 3 مرات يومياً",
    "دهان مرتين يومياً",
    "قرص عند اللزوم"
]

def load_dosages():
    """تحميل الجرعات المحفوظة أو إنشاء الافتراضية"""
    if os.path.exists(DOSAGES_FILE):
        try:
            with open(DOSAGES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and data:
                    return data
        except Exception:
            pass
    save_dosages(DEFAULT_DOSAGES)
    return DEFAULT_DOSAGES[:]

def save_dosages(dosages_list):
    """حفظ الجرعات في ملف محلي"""
    try:
        with open(DOSAGES_FILE, "w", encoding="utf-8") as f:
            json.dump(dosages_list, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("خطأ في حفظ الجرعات:", e)

def get_available_printers():
    """الحصول على قائمة بجميع الطابعات في ويندوز"""
    printers = []
    if WIN32_AVAILABLE:
        try:
            for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
                printers.append(p[2])
        except Exception:
            pass
    return printers

def load_font(path, size):
    """
    يحمّل الخط بمحرك الرسم الأساسي (BASIC) صراحة. تشكيل النص العربي وترتيبه
    من اليمين لليسار بيتم يدويًا قبل الرسم (بدالة prepare_arabic) بمكتبات
    بايثون خالصة (arabic_reshaper + python-bidi) مالهاش أي اعتماد على مكتبات
    نظام خارجية زي raqm، اللي ممكن متتجمعش صح جوه ملف exe بعد التحزيم
    بـ PyInstaller وتسبب رجوع الحروف لشكلها المقلوب/المفصول.
    """
    if os.path.exists(path):
        try:
            return ImageFont.truetype(path, size, layout_engine=ImageFont.Layout.BASIC)
        except Exception:
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()

def prepare_arabic(text):
    """
    يجهز أي نص عربي للرسم بمكتبة PIL بشكل صحيح (استدعها مرة واحدة فقط لكل نص):
    1. reshape: يوصل حروف الكلمة ببعضها بشكلها الصحيح (بدل ما تطبع منفصلة).
    2. get_display: يرتب النص بصريًا من اليمين لليسار (RTL) بدل ما يطبع مقلوب.
    """
    if not text:
        return ""
    try:
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)
    except Exception:
        return text

def draw_prepared_text_right_aligned(draw, right_edge_x, y, prepared_text, font, fill="black"):
    """يرسم نصًا تم تجهيزه مسبقًا بـ prepare_arabic، محاذى لليمين عند إحداثي right_edge_x"""
    bbox = draw.textbbox((0, 0), prepared_text, font=font)
    text_w = bbox[2] - bbox[0]
    draw.text((right_edge_x - text_w, y), prepared_text, fill=fill, font=font)

def draw_prepared_text_centered(draw, center_x, center_y, prepared_text, font, fill="black"):
    """يرسم نصًا تم تجهيزه مسبقًا بـ prepare_arabic، في منتصف نقطة معينة أفقيا ورأسيا"""
    bbox = draw.textbbox((0, 0), prepared_text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = center_x - text_w / 2
    y = center_y - text_h / 2 - bbox[1]
    draw.text((x, y), prepared_text, fill=fill, font=font)

def render_label_image(pharmacy_name, drug_name, patient_name, dosage, date_str):
    """
    توليد صورة ملصق بدقة 203 DPI لمقاس 38mm x 25mm (~304 × 200 بكسل).
    المساحة صغيرة جداً فمش بتكفي غير الجرعة نفسها، فالملصق بيطبع الجرعة فقط
    (من غير اسم الصيدلية، اسم الدواء، التاريخ، أو أي إطار) مع تكبيرها لتملأ المساحة،
    وتصغير حجم الخط تلقائياً لو النص طويل عشان يفضل داخل حدود الملصق.
    """
    width = 304
    height = 200
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    font_bold_path = "C:/Windows/Fonts/arialbd.ttf"

    max_width = width - 16
    max_height = height - 16
    font_size = 34
    min_font_size = 14
    prepared_dosage = prepare_arabic(dosage)
    f_dosage = load_font(font_bold_path, font_size)
    bbox = draw.textbbox((0, 0), prepared_dosage, font=f_dosage)
    while (bbox[2] - bbox[0] > max_width or bbox[3] - bbox[1] > max_height) and font_size > min_font_size:
        font_size -= 2
        f_dosage = load_font(font_bold_path, font_size)
        bbox = draw.textbbox((0, 0), prepared_dosage, font=f_dosage)

    draw_prepared_text_centered(draw, width / 2, height / 2, prepared_dosage, f_dosage)

    return img

def print_image_to_printer(img, printer_name, copies=1):
    """إرسال الصورة للطباعة بعدد النسخ المحدد"""
    if not WIN32_AVAILABLE:
        raise RuntimeError("مكتبة pywin32 غير متوفرة. قم بتثبيتها عبر: pip install pywin32")

    hprinter = win32print.OpenPrinter(printer_name)
    try:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc(f"Dosage_Batch_{copies}")
        
        dib = ImageWin.Dib(img)
        handle = hdc.GetHandleOutput()

        for _ in range(copies):
            hdc.StartPage()
            dib.draw(handle, (0, 0, img.size[0], img.size[1]))
            hdc.EndPage()

        hdc.EndDoc()
        hdc.DeleteDC()
    finally:
        win32print.ClosePrinter(hprinter)


# نافذة إدارة الجرعات الشائعة (لوحة التحكم)
class DosageManagerDialog(tk.Toplevel):
    def __init__(self, parent, on_save_callback):
        super().__init__(parent)
        self.title("إدارة الجرعات الشائعة")
        self.geometry("450x420")
        self.resizable(False, False)
        self.on_save_callback = on_save_callback
        self.dosages = load_dosages()

        self.setup_ui()
        self.transient(parent)
        self.grab_set()

    def setup_ui(self):
        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="الجرعات المسجلة بالقائمة:", font=("Arial", 11, "bold")).pack(anchor="w")

        # قائمة الجرعات
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.listbox = tk.Listbox(list_frame, font=("Arial", 11), selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)
        
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for d in self.dosages:
            self.listbox.insert(tk.END, d)

        # إضافة جرعة جديدة
        add_frame = ttk.Frame(frame)
        add_frame.pack(fill=tk.X, pady=8)

        self.new_dosage_entry = ttk.Entry(add_frame, font=("Arial", 11))
        self.new_dosage_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        add_btn = ttk.Button(add_frame, text="إضافة جرعة", command=self.add_dosage)
        add_btn.pack(side=tk.RIGHT)

        # أزرار الحذف والإغلاق
        bottom_frame = ttk.Frame(frame)
        bottom_frame.pack(fill=tk.X, pady=(5, 0))

        del_btn = ttk.Button(bottom_frame, text="حذف الجرعة المحددة", command=self.delete_dosage)
        del_btn.pack(side=tk.LEFT)

        close_btn = ttk.Button(bottom_frame, text="حفظ وإغلاق", command=self.close_dialog)
        close_btn.pack(side=tk.RIGHT)

    def add_dosage(self):
        text = self.new_dosage_entry.get().strip()
        if text and text not in self.dosages:
            self.dosages.append(text)
            self.listbox.insert(tk.END, text)
            self.new_dosage_entry.delete(0, tk.END)
            save_dosages(self.dosages)
            self.on_save_callback()

    def delete_dosage(self):
        sel = self.listbox.curselection()
        if sel:
            idx = sel[0]
            val = self.dosages.pop(idx)
            self.listbox.delete(idx)
            save_dosages(self.dosages)
            self.on_save_callback()

    def close_dialog(self):
        self.destroy()


# التطبيق الرئيسي
class MainPharmacyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("منظومة طباعة جرعات الأدوية والروشتات (38 × 25 مم)")
        self.root.geometry("860x680")
        try:
            if os.path.exists(APP_ICON_FILE):
                self.root.iconbitmap(APP_ICON_FILE)
        except Exception:
            pass

        # المتغيرات الأساسية
        self.pharmacy_var = tk.StringVar(value="صيدليات دواء")
        self.patient_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y/%m/%d"))
        self.printer_var = tk.StringVar()
        self.status_var = tk.StringVar(value="جاهز للعمل")

        # متغيرات الإدخال الحالي
        self.drug_var = tk.StringVar()
        self.dosage_var = tk.StringVar()
        self.copies_var = tk.IntVar(value=1)

        # قائمة عناصر الروشتة (Batch Queue)
        self.batch_queue = []

        self.setup_ui()
        self.refresh_quick_dosages()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use("clam")

        # هيدر الصيدلية (اللوجو)
        header_frame = ttk.Frame(self.root, padding=(10, 8))
        header_frame.pack(fill=tk.X)
        try:
            if os.path.exists(APP_LOGO_FILE):
                logo_img = Image.open(APP_LOGO_FILE)
                logo_img.thumbnail((220, 80))
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                ttk.Label(header_frame, image=self.logo_photo).pack()
        except Exception:
            pass

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        # الشريط العلوي: إعدادات الصيدلية والطابعة
        top_bar = ttk.Frame(self.root, padding="10")
        top_bar.pack(fill=tk.X)

        # اختيار الطابعة
        printers = get_available_printers()
        default_p = ""
        for p in printers:
            if "370" in p or "xprinter" in p.lower() or "label" in p.lower():
                default_p = p
                break
        if not default_p and printers:
            default_p = printers[0]

        self.printer_var.set(default_p)

        ttk.Label(top_bar, text="الطابعة:").pack(side=tk.LEFT, padx=(0, 5))
        self.printer_combo = ttk.Combobox(
            top_bar, textvariable=self.printer_var, values=printers, state="readonly", width=25
        )
        self.printer_combo.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(top_bar, text="المريض:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(top_bar, textvariable=self.patient_var, width=18).pack(side=tk.LEFT, padx=(0, 15))

        manage_btn = ttk.Button(top_bar, text="⚙ لوحة تحكم الجرعات", command=self.open_dosage_manager)
        manage_btn.pack(side=tk.RIGHT)

        # فاصل
        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        # إطار العمل الرئيسي مقسم لعمودين (إدخال وتحديد الجرعات + قائمة الروشتة المجهزة)
        content_frame = ttk.Frame(self.root, padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # العمود الأيمن: الإدخال والجرعات السريعة
        left_input_frame = ttk.LabelFrame(content_frame, text=" إدخال بيانات الدواء والجرعة ", padding="10")
        left_input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        ttk.Label(left_input_frame, text="اسم الدواء / الصنف:", font=("Arial", 10, "bold")).pack(anchor="w")
        self.drug_entry = ttk.Entry(left_input_frame, textvariable=self.drug_var, font=("Arial", 11))
        self.drug_entry.pack(fill=tk.X, pady=(2, 10))
        self.drug_entry.focus()

        ttk.Label(left_input_frame, text="الجرعة (اختر سريعاً أو اكتب يدوياً):", font=("Arial", 10, "bold")).pack(anchor="w")
        
        # حاوية الأزرار السريعة للجرعات
        self.quick_container = ttk.Frame(left_input_frame)
        self.quick_container.pack(fill=tk.X, pady=(2, 8))

        # حقل الجرعة الفعلي (يمكن الكتابة فيه يدوياً أو التعديل على الاختيار)
        ttk.Label(left_input_frame, text="نص الجرعة النهائي:", font=("Arial", 9)).pack(anchor="w")
        self.dosage_entry = ttk.Entry(left_input_frame, textvariable=self.dosage_var, font=("Arial", 11))
        self.dosage_entry.pack(fill=tk.X, pady=(2, 10))

        # التحكم في عدد مرات التكرار (عدد النسخ)
        copies_row = ttk.Frame(left_input_frame)
        copies_row.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(copies_row, text="عدد النسخ (العلب):", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        spin = ttk.Spinbox(copies_row, from_=1, to=20, textvariable=self.copies_var, width=5)
        spin.pack(side=tk.LEFT)

        # أزرار الإضافة والطباعة الفردية
        action_row = ttk.Frame(left_input_frame)
        action_row.pack(fill=tk.X, pady=5)

        add_to_queue_btn = tk.Button(
            action_row, 
            text="➕ إضافة لقائمة الروشتة", 
            bg="#1976D2", 
            fg="white", 
            font=("Arial", 10, "bold"),
            command=self.add_item_to_batch
        )
        add_to_queue_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))

        quick_print_btn = tk.Button(
            action_row, 
            text="🖨 طباعة هذا الصنف فوراً", 
            bg="#388E3C", 
            fg="white", 
            font=("Arial", 10, "bold"),
            command=self.print_single_immediate
        )
        quick_print_btn.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(4, 0))

        # العمود الأيسر: جدول عناصر الروشتة (Batch List)
        right_batch_frame = ttk.LabelFrame(content_frame, text=" قائمة الروشتة المجهزة للطباعة ", padding="10")
        right_batch_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # جدول Treeview لعرض الأدوية المضافة
        columns = ("drug", "dosage", "copies")
        self.tree = ttk.Treeview(right_batch_frame, columns=columns, show="headings", height=12)
        self.tree.heading("drug", text="الدواء")
        self.tree.heading("dosage", text="الجرعة")
        self.tree.heading("copies", text="النسخ")

        self.tree.column("drug", width=110, anchor="center")
        self.tree.column("dosage", width=170, anchor="center")
        self.tree.column("copies", width=45, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # أزرار إدارة الجدول
        tree_btns = ttk.Frame(right_batch_frame)
        tree_btns.pack(fill=tk.X, pady=(0, 8))

        del_item_btn = ttk.Button(tree_btns, text="حذف الصنف", command=self.remove_selected_from_batch)
        del_item_btn.pack(side=tk.LEFT)

        clear_all_btn = ttk.Button(tree_btns, text="تفريغ القائمة", command=self.clear_batch)
        clear_all_btn.pack(side=tk.LEFT, padx=5)

        # زر الطباعة الكلية للروشتة
        self.batch_print_btn = tk.Button(
            right_batch_frame,
            text="🚀 طباعة جميع جرعات الروشتة دفعة واحدة",
            bg="#D32F2F",
            fg="white",
            font=("Arial", 11, "bold"),
            command=self.print_full_batch
        )
        self.batch_print_btn.pack(fill=tk.X)

        # شريط الحالة
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w", padding="3")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # ربط Enter بإضافة الصنف للقائمة
        self.root.bind("<Return>", lambda e: self.add_item_to_batch())

    def refresh_quick_dosages(self):
        """إعادة بناء شبكة الأزرار السريعة بناءً على ملف الجرعات المحفوظ"""
        for widget in self.quick_container.winfo_children():
            widget.destroy()

        dosages = load_dosages()
        # عرض الجرعات في شبكة عمودين لترتيب الشاشة
        row = 0
        col = 0
        for d in dosages:
            btn = ttk.Button(
                self.quick_container, 
                text=d, 
                command=lambda text=d: self.dosage_var.set(text)
            )
            btn.grid(row=row, column=col, sticky="ew", padx=2, pady=2)
            col += 1
            if col > 1:
                col = 0
                row += 1
        
        self.quick_container.columnconfigure(0, weight=1)
        self.quick_container.columnconfigure(1, weight=1)

    def open_dosage_manager(self):
        DosageManagerDialog(self.root, on_save_callback=self.refresh_quick_dosages)

    def add_item_to_batch(self):
        drug = self.drug_var.get().strip()
        dosage = self.dosage_var.get().strip()
        copies = self.copies_var.get()

        if not dosage:
            messagebox.showwarning("نقص بيانات", "يرجى اختيار أو كتابة الجرعة أولاً.")
            return

        item = {
            "drug": drug or "بدون اسم",
            "dosage": dosage,
            "copies": max(1, copies),
            "patient": self.patient_var.get().strip(),
            "date": self.date_var.get().strip(),
            "pharmacy": self.pharmacy_var.get().strip()
        }
        self.batch_queue.append(item)
        self.tree.insert("", tk.END, values=(item["drug"], item["dosage"], item["copies"]))

        # تفريغ الحقول وإعادة التركيز
        self.drug_var.set("")
        self.dosage_var.set("")
        self.copies_var.set(1)
        self.drug_entry.focus()
        self.status_var.set(f"تمت الإضافة للقائمة. إجمالي الروشتة: {len(self.batch_queue)} أصناف.")

    def remove_selected_from_batch(self):
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            self.tree.delete(sel[0])
            self.batch_queue.pop(idx)
            self.status_var.set("تم حذف الصنف من القائمة.")

    def clear_batch(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.batch_queue.clear()
        self.status_var.set("تم تفريغ قائمة الروشتة.")

    def print_single_immediate(self):
        drug = self.drug_var.get().strip()
        dosage = self.dosage_var.get().strip()
        copies = max(1, self.copies_var.get())

        if not dosage:
            messagebox.showwarning("نقص بيانات", "يرجى اختيار أو كتابة الجرعة للطباعة الفورية.")
            return

        printer = self.printer_var.get()
        if not printer:
            messagebox.showerror("خطأ", "لم يتم اختيار أي طابعة!")
            return

        img = render_label_image(
            pharmacy_name=self.pharmacy_var.get().strip(),
            drug_name=drug,
            patient_name=self.patient_var.get().strip(),
            dosage=dosage,
            date_str=self.date_var.get().strip()
        )

        try:
            print_image_to_printer(img, printer, copies=copies)
            self.status_var.set(f"تمت طباعة {copies} ملصق بنجاح.")
            self.drug_var.set("")
            self.dosage_var.set("")
            self.copies_var.set(1)
            self.drug_entry.focus()
        except Exception as e:
            messagebox.showerror("خطأ في الطباعة", str(e))

    def print_full_batch(self):
        if not self.batch_queue:
            messagebox.showinfo("تنبيه", "قائمة الروشتة فارغة! أضف الأصناف أولاً.")
            return

        printer = self.printer_var.get()
        if not printer:
            messagebox.showerror("خطأ", "لم يتم اختيار أي طابعة!")
            return

        total_labels = sum(item["copies"] for item in self.batch_queue)
        confirm = messagebox.askyesno(
            "تأكيد الطباعة", 
            f"سيتم طباعة الروشتة كاملة بعدد {len(self.batch_queue)} صنف بإجمالي {total_labels} ملصق.\nهل تريد المتابعة؟"
        )
        if not confirm:
            return

        try:
            for item in self.batch_queue:
                img = render_label_image(
                    pharmacy_name=item["pharmacy"],
                    drug_name=item["drug"],
                    patient_name=item["patient"],
                    dosage=item["dosage"],
                    date_str=item["date"]
                )
                print_image_to_printer(img, printer, copies=item["copies"])

            self.status_var.set(f"تمت طباعة الروشتة بالكامل بنجاح ({total_labels} ملصق).")
            self.clear_batch()
        except Exception as e:
            messagebox.showerror("خطأ أثناء الطباعة الجماعية", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = MainPharmacyApp(root)
    root.mainloop()
