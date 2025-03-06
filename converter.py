import threading
import os
import multiprocessing
from tkinter import filedialog, messagebox
import time
from docx2pdf import convert as conv_doc
from pptxtopdf import convert as conv_ppt



def convert_to_pdf_worker(file_path, output_dir,result_queue):
    """the function convert the file to pdf """
    try:
        # the file details
        print(f"קובץ מקורי: {file_path}")

        file_ext = os.path.splitext(file_path.lower())[1]
        file_name = os.path.basename(file_path)

        # creating the target file
        output_filename = os.path.splitext(file_name)[0] + ".pdf"
        pdf_path = os.path.join(output_dir, output_filename)

        # the progress update
        result_queue.put(("status", f"ממיר: {file_name}"))

        if file_ext == ".docx":
            try:
                # converting the file using docx2pdf library
                conv_doc(file_path, pdf_path)
                result_queue.put(("success", file_path))
            except Exception as word_error:
                result_queue.put(("error", f"שגיאה בהמרת Word: {str(word_error)}"))

        elif file_ext == ".pptx":
            try:
                # converting the file using pptTopdf library
                conv_ppt(file_path, pdf_path)
                result_queue.put(("success", file_path))
            except Exception as ppt_error:
                result_queue.put(("error", f"שגיאה בהמרת PowerPoint: {str(ppt_error)}"))

        else:
            result_queue.put(("error", f"פורמט לא נתמך: {file_ext}"))
            return

    except Exception as e:
        result_queue.put(("error", f"שגיאה כללית בהמרת {os.path.basename(file_path)}: {str(e)}"))


def conversion_manager(file_paths, output_dir, result_queue):
    """manage the conversion processes"""
    for file_path in file_paths:
        convert_to_pdf_worker(file_path, output_dir, result_queue)
        # time space between each conversion process
        time.sleep(0.5)


def update_ui(root, status_label, progress, result_queue, total_files, converted_files, on_complete):
    """updating the application interaction window during the conversion"""
    try:

        while not result_queue.empty():
            msg_type, msg = result_queue.get_nowait()
            if msg_type == "status":
                status_label.config(text=msg, fg="blue")
            elif msg_type == "success":
                converted_files[0] += 1
                progress_value = int((converted_files[0] / total_files) * 100)
                progress.set(progress_value)
            elif msg_type == "error":
                messagebox.showerror("שגיאה", msg)

        # check for end conversion
        if converted_files[0] >= total_files:
            progress.set(100)
            status_label.config(text="ההמרה הסתיימה בהצלחה! ✅", fg="green")
            on_complete()
            return

        # keep checking
        root.after(100, lambda: update_ui(root, status_label, progress, result_queue, total_files, converted_files,on_complete))

    except Exception as e:
        messagebox.showerror("שגיאה", f"שגיאה בעדכון הממשק: {str(e)}")

def start_process(file_paths, status_label, root, progress):
    """the function start the converting process with thread"""
    if not file_paths:
        return

    progress.set(0)

    output_dir = filedialog.askdirectory(title="בחר תיקייה לשמירת קבצי ה-PDF")
    if not output_dir:
        status_label.config(text="ההמרה בוטלה", fg="gray")
        return

    result_queue = multiprocessing.Queue()
    total_files = len(file_paths)

    # creating mission for the thread
    converter_thread = threading.Thread(
        target=conversion_manager,
        args=(file_paths, output_dir, result_queue)
    )
    converter_thread.daemon = True
    converter_thread.start()


    converted_files = [0]
    def on_complete():
        if converted_files[0] > 0:
            messagebox.showinfo("הודעה", f"כל הקבצים הומרו בהצלחה! 😀✅ \nהקבצים נשמרו ב: {output_dir}")

    status_label.config(text=f"מתחיל המרה של {total_files} קבצים...", fg="blue")
    root.update_idletasks()

    root.after(100, lambda: update_ui(root, status_label, progress, result_queue, total_files, converted_files, on_complete))

def browse_files(status_lable,root, progress=None):
    """choosing files to convert"""
    file_paths=filedialog.askopenfilenames(filetypes=[("Word & PowerPoint files", "*.docx;*.pptx")])
    if file_paths:
        start_process(file_paths, status_lable, root, progress)

