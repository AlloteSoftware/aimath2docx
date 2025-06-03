import customtkinter as ctk
from tkinter import filedialog, messagebox
from aimath2docx import convert_file
import webbrowser
import os

def main():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title("AI Math to DOCX Converter")
    width = 500
    height = 240
    screen_width = app.winfo_screenwidth()
    screen_height = app.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    app.geometry(f"{width}x{height}+{x}+{y}")

    input_path = ctk.StringVar()

    def show_centered_message(title, message, parent=None):
        win = ctk.CTkToplevel(parent)
        win.title(title)
        win.resizable(False, False)

        # Получить системный фон окна customtkinter для Toplevel
        bg_color = ctk.ThemeManager.theme["CTkToplevel"]["fg_color"]
        # Определить индекс цвета по текущей теме
        if isinstance(bg_color, (list, tuple)):
            color_mode = ctk.ThemeManager.theme.get("color_mode", "Light")
            idx = 0 if color_mode == "Light" else 1
            bg_color = bg_color[idx]
        # Установить фон окна
        win.configure(fg_color=bg_color)

        # Фрейм с этим же цветом
        frame = ctk.CTkFrame(win, fg_color=bg_color)
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Лейбл с этим же цветом
        label = ctk.CTkLabel(
            frame,
            text=message,
            justify="center",
            font=("Segoe UI", 14),
            fg_color=bg_color
        )
        label.pack(pady=(10, 10))

        btn = ctk.CTkButton(frame, text="OK", command=win.destroy)
        btn.pack(pady=(0, 10))

        win.update_idletasks()  # Вычислить размеры

        req_width = frame.winfo_reqwidth() + 20   # немного запаса для рамки
        req_height = frame.winfo_reqheight() + 20

        # Центрирование
        if parent is not None:
            parent.update_idletasks()
            parent_x = parent.winfo_rootx()
            parent_y = parent.winfo_rooty()
            parent_w = parent.winfo_width()
            parent_h = parent.winfo_height()
            x = parent_x + (parent_w - req_width) // 2
            y = parent_y + (parent_h - req_height) // 2
        else:
            screen_w = win.winfo_screenwidth()
            screen_h = win.winfo_screenheight()
            x = (screen_w - req_width) // 2
            y = (screen_h - req_height) // 2

        win.geometry(f"{req_width}x{req_height}+{x}+{y}")

        win.grab_set()
        
    def browse_input():
        path = filedialog.askopenfilename(
            filetypes=[("Markdown files", "*.md")],
            title="Select Markdown file..."
        )
        if path:
            input_path.set(path)
            convert_button.configure(state="normal")

    def run_conversion():
        default_name = os.path.splitext(os.path.basename(input_path.get()))[0] + ".docx"
        save_path = filedialog.asksaveasfilename(
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            title="Save as..."
        )
        if not save_path:
            return  # если пользователь отменил выбор

        success = convert_file(input_path.get(), save_path)
        if success:
            show_centered_message("Success", f"Conversion complete.\nFile saved as:\n{save_path}", parent=app)
        else:
            show_centered_message("Error", "Error converting or saving file.", parent=app)

    ctk.CTkLabel(app, text="Input Markdown file:").pack(pady=(20, 0))
    ctk.CTkEntry(app, textvariable=input_path, width=400, state="readonly").pack()
    ctk.CTkButton(app, text="Browse", command=browse_input).pack(pady=(20, 0))

    convert_button = ctk.CTkButton(app, text="Convert", command=run_conversion, state="disabled")
    convert_button.pack(pady=(20, 5))

    footer = ctk.CTkLabel(
        app,
        text="Allote Software (c) 2025",
        text_color="black",
        font=("Segoe UI", 14, "bold"),
        cursor="hand2"
    )
    footer.place(relx=0.5, rely=1.0, anchor="s", y=-5)
    footer.bind("<Button-1>", lambda e: webbrowser.open_new("https://allote.software"))

    app.mainloop()

if __name__ == "__main__":
    main()
