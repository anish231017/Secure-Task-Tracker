import face_recognition
import os
from tkinter import *
from tkinter import simpledialog
from tkinter import messagebox

from PIL import Image, ImageTk
import cv2
import pandas as pd
from openpyxl import load_workbook

def set_background_image(window, image_path):
    background_image = Image.open(image_path)
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = Label(window, image=background_photo)
    background_label.image = background_photo
    background_label.place(relwidth=1, relheight=1)

def get_button(window, text, color, command, fg='white'):
    button = Button(
        window,
        text=text,
        activebackground=color,
        activeforeground="white",
        fg=fg,
        bg=color,
        command=command,
        height=2,
        width=20,
        font=('Helvetica', 16, 'bold')  # Adjust the font style and size
    )
    return button

def create_register_window():
    register_window = Toplevel(win)
    register_window.geometry("1200x520+350+120")
    register_window.title("User Registration")

    set_background_image(register_window,"mangerimage.jpg")


    def add_img():
        cv2image = cv2.cvtColor(cap.read()[1], cv2.COLOR_BGR2RGB)
        img = Image.fromarray(cv2image)
        imgtk = ImageTk.PhotoImage(image=img)
        capture_label.imgtk = imgtk
        capture_label.configure(image=imgtk)
        global recent
        recent = cap.read()[1]


    capture_label=Label(register_window)
    capture_label.grid(row=0, column=0)
    capture_label.place(x=30, y=25, width=640, height=440)

    add_img()

    entry_text = Entry(register_window,justify="left",font="Helvetica 32",width = 13,bg ="indigo",fg ="yellow")
    entry_text.place(x=850, y=150)
    entry_label = Label(register_window,text="Please Input Username: ",justify="left",font = "Helvetica 20 bold",bg ="blue",fg="yellow")
    entry_label.place(x=850, y=70)

    def accept():

        name = entry_text.get()
        cv2.imwrite(os.path.join(db_dir, "{}.jpg".format(name)),recent)
        messagebox.showinfo("Info", "User created successfully!")

        excel_file_path = "user_db.xlsx"
        registered_user = []
        registered_user.append(name)
        try:
            wb = load_workbook(excel_file_path)
            sheet = wb.active
            new_data = [registered_user]
            for data in new_data:
                sheet.append(data)

            wb.save(filename=excel_file_path)

        except FileNotFoundError:
            df = pd.DataFrame({"Name": registered_user})

            df.to_excel(excel_file_path, index=False)

        register_window.destroy()


    accept_button = get_button(register_window, 'Accept', 'green', accept)
    accept_button.place(x=850, y=300)

    def try_again():
        register_window.destroy()

    try_again_button = get_button(register_window, 'Try again', 'red', try_again)
    try_again_button.place(x=850, y=400)

def show_frames():
    cv2image = cv2.cvtColor(cap.read()[1], cv2.COLOR_BGR2RGB)
    img = Image.fromarray(cv2image)
    imgtk = ImageTk.PhotoImage(image=img)
    webcam_label.imgtk = imgtk
    webcam_label.configure(image=imgtk)
    webcam_label.after(20, show_frames)


db_dir = './db'
if not os.path.exists(db_dir):
    os.mkdir(db_dir)

def Manager_window():

    def assign_task():
        selected_employee = listbox_user.get(listbox_user.curselection())
        task_entry = simpledialog.askstring("Assign Task", f"Assign a task to {selected_employee}:",
                                            )

        if task_entry:
            # Update the task for the selected employee in the Excel file
            filepath = "user_db.xlsx"
            df = pd.read_excel(filepath)

            index = df[df['Name'] == selected_employee[2:]].index.item()
            df.at[index, 'Task'] = task_entry
            df.at[index, 'Status'] = "Not done"
            # Save the updated dataframe to the Excel file
            df.to_excel(filepath, index=False)

            # Update the task listbox
            update_task_listbox()

    def update_task_listbox():
        # Update the Task Listbox with the latest data
        df_task = pd.read_excel(filepath)
        values_task = df_task[column_name_task].tolist()
        values_status = df_status[column_name_status].tolist()
        listbox_status.delete(0,END)

        listbox_task.delete(0, END)
        for value in values_task:
            listbox_task.insert(END, f"\u2022 {value}")
        for value in values_status:
            listbox_status.insert(END, f"\u2022 {value}")

    manager_window = Tk()
    manager_window.geometry("1200x520+350+120")
    manager_window.title("Manager Window")
    manager_window.configure(bg = "lightblue")


    register_button_main_window = get_button(manager_window, 'Register User', 'grey', create_register_window, 'black')
    register_button_main_window.place(x=850, y=400)

    assign_task_button = get_button(manager_window, 'Assign Task', 'blue', assign_task, fg='white')
    assign_task_button.place(x=850, y=300)


    label_emp = Label(manager_window, text="EMPLOYEES ", justify="left", font="Helvetica 20 bold",bg="lightblue",fg="darkblue")
    label_emp.place(x=10, y=10)

    label_task = Label(manager_window, text="TASKS ", justify="left", font="Helvetica 20 bold",bg="lightblue",fg="darkblue")
    label_task.place(x=250, y=10)

    label_status = Label(manager_window, text="STATUS", justify="left", font="Helvetica 20 bold",bg="lightblue",fg="darkblue")
    label_status.place(x=470, y=10)

    filepath = "user_db.xlsx"
    #user listbox

    column_name = "Name"
    df = pd.read_excel(filepath)
    values = df[column_name].tolist()

    listbox_user = Listbox(manager_window, width=18, height=15,font="Helvetica 15",bg="lightgreen",fg="black")
    listbox_user.place(x=15,y=60)

    for value in values:
        listbox_user.insert(END, f"\u2022 {value}")

    #task listbox

    column_name_task = "Task"
    df_task = pd.read_excel(filepath)
    values_task = df_task[column_name_task].tolist()

    listbox_task = Listbox(manager_window, width=18, height=15, font="Helvetica 15",bg="lightgreen",fg="black")
    listbox_task.place(x=250, y=60)

    for value in values_task:
        listbox_task.insert(END, f"\u2022 {value}")

    #Status Listbox

    column_name_status = "Status"
    df_status = pd.read_excel(filepath)
    values_status = df_status[column_name_status].tolist()

    listbox_status = Listbox(manager_window, width=18, height=15, font="Helvetica 15 ",bg="lightgreen",fg="black")
    listbox_status.place(x=470, y=60)

    for value in values_status:
        listbox_status.insert(END, f"\u2022 {value}")

    manager_window.mainloop()

def show_task(name):
    filepath = "user_db.xlsx"

    column_name = "Name"
    df = pd.read_excel(filepath)
    name_dataset = df[column_name].tolist()

    index = name_dataset.index(name)
    task_column = "Task"
    task_dataset = df[task_column].tolist()

    if pd.isna(task_dataset[index]):

        messagebox.showinfo("Info", "No Task for you!")
    else:
        def submit_task():
            df.at[index, 'Status'] = "done"

            # Save the updated dataframe to the Excel file
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Info", "Task has been successfully submitted!")
            root.destroy()


        # Create the main window
        root = Tk()
        root.title("Task Manager")

        # Create a canvas
        canvas = Canvas(root, width=400, height=300, bg='#013358')
        canvas.pack()

        # Add text on top of the canvas
        task_label = Label(root, text=f"{task_dataset[index]}".upper(), font=("Helvetica", 30,"bold"),bg = "#013358",fg="lightblue",justify="center")
        task_label.place(x=85,y=40)
        text_label = Label(root, text="Your Today's Task", font=("Helvetica", 18))
        text_label.pack(pady=10)

        # Add a larger and green button to submit the task
        submit_button = Button(root, text="Submit Task", command=submit_task, bg='#4CAF50', fg='white',
                                     font=("Helvetica", 14))
        submit_button.pack(pady=15, ipadx=20, ipady=10)

        # Run the Tkinter event loop
        root.mainloop()


def login():
    name =""
    unknown_img_path = './.tmp.jpg'
    cv2.imwrite(unknown_img_path, cap.read()[1])
    unknown_image = face_recognition.load_image_file(unknown_img_path)
    resized_image = cv2.resize(unknown_image, (0, 0), fx=0.5, fy=0.5)
    face_locations = face_recognition.face_locations(resized_image)

    if not face_locations:
        print("No face detected.")
        messagebox.showinfo("Info", "No Face detected")
        os.remove(unknown_img_path)
        return

    top, right, bottom, left = face_locations[0]
    unknown_face_encoding = face_recognition.face_encodings(resized_image, [(top, right, bottom, left)])[0]
    for file_name in os.listdir(db_dir):
        file_path = os.path.join(db_dir, file_name)

        known_image = face_recognition.load_image_file(file_path)
        resized_known_image = cv2.resize(known_image, (0, 0), fx=0.5, fy=0.5)
        known_face_encoding = face_recognition.face_encodings(resized_known_image)[0]

        results = face_recognition.compare_faces([known_face_encoding], unknown_face_encoding,tolerance=0.4)

        if results[0]:

            print(f"Welcome, {os.path.splitext(file_name)[0]}!")
            name = os.path.splitext(file_name)[0]
            messagebox.showinfo("Info", "Welcome" +" "+ name + "!")
            break
    else:
        messagebox.showinfo("Info", "Your face is not recognised. "
                                    "Please try again or register yourself via manager!")
        print("Face not recognized.")

    filepath = "user_db.xlsx"

    column_name = "Name"
    df = pd.read_excel(filepath)
    values = df[column_name].tolist()

    if name == "Manager":

        Manager_window()

    elif name in values:
        show_task(name)

    os.remove(unknown_img_path)


# Main Window
win = Tk()
win.geometry("1200x520+350+100")
win.title("Login Window")
set_background_image(win, 'loginimage.jpg')

#Name label

secure_label_1 = Label(win, text="Secure", font=("Helvetica", 50, "bold"), fg="lightblue", bg="#013358")
secure_label_1.place(x=850, y=30)
secure_label_2 = Label(win, text="Task", font=("Helvetica", 50, "bold"), fg="lightblue", bg="#013358")
secure_label_2.place(x=880, y=127)
secure_label_3 = Label(win, text="Tracker", font=("Helvetica", 50, "bold"), fg="lightblue", bg="#013358")
secure_label_3.place(x=850, y=235)

# Webcam Label
webcam_label = Label(win)
webcam_label.grid(row=0, column=0)
webcam_label.place(x=30, y=25, width=640, height=440)
cap = cv2.VideoCapture(0)
show_frames()
# Login Button
login_button_main_window = get_button(win, 'Login', 'turquoise', login)
login_button_main_window.place(x=850, y=400)

win.mainloop()

