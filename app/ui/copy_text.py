import tkinter as tk

def make_copyable_text(parent, text="", height=4):
    widget = tk.Text(
        parent,
        height=height,
        wrap="word",
        borderwidth=0,
        highlightthickness=1,
        relief="flat",
        padx=8,
        pady=8,
        font=("Segoe UI", 10),
        cursor="xterm",
    )
    widget.insert("1.0", text)
    widget.config(state="disabled")

    menu = tk.Menu(widget, tearoff=0)
    menu.add_command(label="Copy", command=lambda: _copy(widget))

    def popup(event):
        try:
            widget.focus_set()
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    widget.bind("<Button-3>", popup)
    return widget

def _copy(widget):
    try:
        selected = widget.get("sel.first", "sel.last")
    except tk.TclError:
        selected = widget.get("1.0", "end").strip()
    widget.clipboard_clear()
    widget.clipboard_append(selected)
