import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD


from converter import browse_files
def create_gui():
    root = TkinterDnD.Tk()
    root.title("Word & PowerPoint to PDF")
    root.geometry("500x350")
    root.resizable(False, False)

    # defining the interaction window's colors
    bg_color = "#f5f5f5"
    accent_color = "#3498db"
    text_color = "#2c3e50"

    root.configure(bg=bg_color)

    # creating our main frame
    main_frame = tk.Frame(root, bg=bg_color, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # title
    title_label = tk.Label(
        main_frame,
        text="   Word / PowerPoint  המרת קבצי \n PDF-ל    ",
        font=("Arial", 16, "bold"),
        bg=bg_color,
        fg=text_color
    )
    title_label.pack(pady=(0, 20))

    # icon for use
    icon_label = tk.Label(
        main_frame,
        text="📄",
        font=("Arial", 48),
        bg=bg_color
    )
    icon_label.pack(pady=10)


    # user guide
    status_lable = tk.Label(main_frame, text=" לחץ על הכפתור לבחירת קבצים",bg=bg_color, fg="gray", font=("Ariel", 12))
    status_lable.pack(pady=8)

    # value for the progress bar
    progress_var = tk.DoubleVar()

    # the choosing file button
    browse_button = tk.Button(main_frame, text="📂  Word / PowerPoint בחר קבצי",
                              command=lambda: browse_files(status_lable, root, progress_var),
                              font=("Ariel", 12), bg=accent_color, fg="white", activebackground="#2980b9",
                              activeforeground="white", bd=0, padx=15, pady=8, cursor="hand2")
    browse_button.pack(pady=15)


    # the progress bar
    progress = ttk.Progressbar(
        main_frame,
        orient="horizontal",
        length=460,
        mode="determinate",
        variable=progress_var
    )
    progress.pack(pady=10)


    def start_gui():
        """running the application interaction window"""
        # the interaction window's location on the screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        root.mainloop()

    return start_gui,status_lable,root

