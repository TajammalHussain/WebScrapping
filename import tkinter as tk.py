import tkinter as tk

def button_click(number):
    current = entry.get()
    entry.delete(0, tk.END)
    entry.insert(tk.END, str(current) + str(number))

def button_clear():
    entry.delete(0, tk.END)
    result_label.config(text="")

def button_equal():
    try:
        result = eval(entry.get())
        entry.delete(0, tk.END)
        entry.insert(tk.END, result)
        result_label.config(text=result)
    except Exception as e:
        entry.delete(0, tk.END)
        entry.insert(tk.END, "Error")
        result_label.config(text="Error")

root = tk.Tk()
root.title("Calculator")

entry = tk.Entry(root, width=35, borderwidth=5)
entry.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

result_label = tk.Label(root, text="", width=35, borderwidth=5)
result_label.grid(row=1, column=0, columnspan=4, padx=10, pady=10)

buttons = [
    ("7", 2, 0), ("8", 2, 1), ("9", 2, 2), ("/", 2, 3),
    ("4", 3, 0), ("5", 3, 1), ("6", 3, 2), ("*", 3, 3),
    ("1", 4, 0), ("2", 4, 1), ("3", 4, 2), ("-", 4, 3),
    ("0", 5, 0), ("C", 5, 1), ("=", 5, 2), ("+", 5, 3)
]

for text, row, column in buttons:
    button = tk.Button(root, text=text, padx=30, pady=20, command=lambda text=text: button_click(text))
    button.grid(row=row, column=column)

root.mainloop()
