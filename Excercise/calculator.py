import tkinter as tk

# Create the main window
root = tk.Tk()
root.title("My Calculator")

# Global variable to store the expression
expression = ""

# Function to update expression
def press(num):
    global expression
    expression += str(num)
    equation.set(expression)

# Function to evaluate the final expression
def equal_press():
    try:
        global expression
        result = str(eval(expression))  # Evaluate the expression
        equation.set(result)
        expression = ""  # Clear the expression after evaluation
    except:
        equation.set("Error")
        expression = ""

# Function to clear the expression
def clear():
    global expression
    expression = ""
    equation.set("")

# StringVar to store the equation
equation = tk.StringVar()

# Create an Entry widget for the calculator display
display = tk.Entry(root, textvariable=equation, font=('Arial', 20), bd=8, insertwidth=2, width=14, borderwidth=4)
display.grid(row=0, column=0, columnspan=4)

# Create the buttons
buttons = [
    '7', '8', '9', '/',
    '4', '5', '6', '*',
    '1', '2', '3', '-',
    'C', '0', '=', '+'
]

# Adding buttons to the grid
row_value = 1
col_value = 0
for button in buttons:
    if button == 'C':
        tk.Button(root, text=button, command=clear, height=2, width=7).grid(row=row_value, column=col_value)  # Nút C
    elif button == '0':
        tk.Button(root, text=button, command=lambda button=button: press(button), height=2, width=7).grid(row=row_value, column=col_value)  # Nút 0
    elif button == '=':
        tk.Button(root, text=button, command=equal_press, height=2, width=7).grid(row=row_value, column=col_value)  # Nút =
    elif button == '+':
        tk.Button(root, text=button, command=lambda button=button: press(button), height=2, width=7).grid(row=row_value, column=col_value)  # Nút +
    else:
        tk.Button(root, text=button, command=lambda button=button: press(button), height=2, width=7).grid(row=row_value, column=col_value)

    col_value += 1
    if col_value > 3:
        col_value = 0
        row_value += 1

root.mainloop()
