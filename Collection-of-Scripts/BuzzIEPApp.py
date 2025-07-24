import tkinter
import tkinter.messagebox
import customtkinter
import pandas as pd
import requests
import os

def start_buzz_auth():
    buzz_login_url = "https://api.agilixbuzz.com/cmd"
    username = Entry_username.get()
    password = Entry_password.get()

    headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json'
    }

    payload = {
        "request": {
            "cmd": "login3",
            "username": username,
            "password": password,
            "newsession": "true"
        }
    }

    response = requests.post(buzz_login_url, headers=headers, json=payload)
    response_json = response.json()

    if response_json['response']['code'] == "OK":
        global token
        token = response_json['response']['user']['token']
        print("Sucessful Login!")
        open_csv_window()
    else:
        print("Error: Login Issue")
        tkinter.messagebox.showerror(title="Login Failed",message="Invalid Login: Try again or exit")


def get_computer_files():
    global csv_path_accommodations
    global csv_path_user_data

    home_dir = os.path.expanduser("~")
    csv_file_accommodations = Entry_iep_csv.get()
    csv_file_user_data = Entry_xml_csv.get()
    
    csv_path_accommodations = (f"{home_dir}/Downloads/{csv_file_accommodations}")
    print(csv_path_accommodations)
    
    csv_path_user_data = (f"{home_dir}/Downloads/{csv_file_user_data}")
    print(csv_path_user_data)
    
    get_csv()


def get_csv():
    table = pd.read_csv(csv_path_accommodations, 
    usecols=["LMSID", "Accommodations"])

    for index, row in table.iterrows():
        LMSID = (row["LMSID"])
        Accommodations = (row["Accommodations"])
        start_buzz_api_accommodations(LMSID, Accommodations)
       
    start_buzz_api_user_data()


def start_buzz_api_accommodations(LMSID, Accommodations):
    token_string = (f"&_token={token}")
    url = (f"https://api.agilixbuzz.com/cmd?=putresource&entityid={LMSID}&path=Teacher/managed/iep.html&class=MISC{token_string}")

    payload = Accommodations
    headers = {
    'Content-Type': 'text/html',
    'Accept': 'application/json'
    }

    response = requests.post(url, headers=headers, data=payload)
    response_json = response.json()
    Response_textbox.insert(tkinter.END, f"{response_json}\n")
    Response_textbox.yview(tkinter.END)
    print(response_json)


def start_buzz_api_user_data():
    token_string = (f"&_token={token}")
    url = (f"https://api.agilixbuzz.com/cmd?cmd=updateusers{token_string}")

    table = pd.read_csv(csv_path_user_data, 
            usecols=["header"])
    if "header" in table.columns:
        convert_to_text = "\n".join(table["header"].astype(str))

    payload = convert_to_text
    headers = {
    'Content-Type': 'text/xml',
    'Accept': 'application/json'
    }

    response = requests.post(url, headers=headers, data=payload)
    response_json = response.json()
    Response_textbox.insert(tkinter.END, f"{response_json}\n")
    Response_textbox.yview(tkinter.END)
    print(response_json)


def start_exit():
    login_window.destroy()
    csv_window.destroy()


def open_csv_window():
    global Entry_iep_csv
    global Entry_xml_csv
    global csv_window
    global Response_textbox

    login_window.withdraw()

    csv_window = tkinter.Tk()
    csv_window.title("Buzz IEP Updater")
    csv_window.geometry("800x650")
    csv_window.configure(bg="#342e84")


    Label_xml_csv = customtkinter.CTkLabel(
        master=csv_window,
        text="Xml CSV",
        font=("Arial", 20),
        text_color="#ffffff",
        height=30,
        width=95,
        corner_radius=0,
        bg_color="#342e84",
        fg_color="#342e84",
    )

    Label_iep_csv = customtkinter.CTkLabel(
        master=csv_window,
        text="IEP CSV",
        font=("Arial", 20),
        text_color="#ffffff",
        height=30,
        width=95,
        corner_radius=0,
        bg_color="#342e84",
        fg_color="#342e84",
    )

    Entry_xml_csv = customtkinter.CTkEntry(
        master=csv_window,
        placeholder_text="Copy CSV name located in Downloads",
        placeholder_text_color="#454545",
        font=("Arial", 14),
        text_color="#000000",
        height=30,
        width=300,
        border_width=1,
        corner_radius=6,
        border_color="#000000",
        bg_color="#342e84",
        fg_color="#F0F0F0",
    )

    Entry_iep_csv = customtkinter.CTkEntry(
        master=csv_window,
        placeholder_text="Copy CSV name located in Downloads",
        placeholder_text_color="#454545",
        font=("Arial", 14),
        text_color="#000000",
        height=30,
        width=300,
        border_width=1,
        corner_radius=6,
        border_color="#000000",
        bg_color="#342e84",
        fg_color="#F0F0F0",
    )

    Button_start = customtkinter.CTkButton(
        master=csv_window,
        text="Start",
        command=get_computer_files,
        font=("undefined", 25),
        text_color="#ffffff",
        hover=True,
        hover_color="#ff0000",
        height=35,
        width=100,
        border_width=2,
        corner_radius=6,
        border_color="#000000",
        bg_color="#342e84",
        fg_color="#000000",
    )

    Button_exit = customtkinter.CTkButton(
        master=csv_window,
        text="EXIT",
        command=start_exit,
        font=("undefined", 25),
        text_color="#ffffff",
        hover=True,
        hover_color="#ff0000",
        height=35,
        width=100,
        border_width=2,
        corner_radius=6,
        border_color="#000000",
        bg_color="#342e84",
        fg_color="#000000",
    )

    Response_textbox = customtkinter.CTkTextbox(
        master=csv_window,
        width=720,
        height=300,
        font=("Arial", 12),
        bg_color="#f0f0f0",
        fg_color="#000000",
        corner_radius=6,
        border_color="#000000",
        border_width=1,
    )

    Label_xml_csv.place(x=360, y=60)
    Label_iep_csv.place(x=360, y=150)
    Entry_xml_csv.place(x=260, y=90)
    Entry_iep_csv.place(x=260, y=180)
    Button_start.place(x=250, y=270)
    Button_exit.place(x=470, y=270)
    Response_textbox.place(x=40, y=320)

    csv_window.mainloop()


login_window = tkinter.Tk()
login_window.title("Buzz Login")
login_window.geometry("800x350")
login_window.configure(bg="#342e84")

Label_username = customtkinter.CTkLabel(
    master=login_window,
    text="Username:",
    font=("Arial", 20),
    text_color="#FFFFFF",
    height=30,
    width=95,
    corner_radius=0,
    bg_color="#342e84",
    fg_color="#342e84",
    )

Label_password = customtkinter.CTkLabel(
    master=login_window,
    text="Password:",
    font=("Arial", 20),
    text_color="#FFFFFF",
    height=30,
    width=95,
    corner_radius=0,
    bg_color="#342e84",
    fg_color="#342e84",
    )

Entry_username = customtkinter.CTkEntry(
    master=login_window,
    placeholder_text="Username",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=225,
    border_width=1,
    corner_radius=6,
    border_color="#000000",
    bg_color="#342e84",
    fg_color="#F0F0F0",
    )

Entry_password = customtkinter.CTkEntry(
    master=login_window,
    placeholder_text="Password",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=225,
    border_width=1,
    corner_radius=6,
    border_color="#000000",
    bg_color="#342e84",
    fg_color="#F0F0F0",
    )

Button_login = customtkinter.CTkButton(
    master=login_window,
    text="Login",
    command=start_buzz_auth,
    font=("undefined", 14),
    text_color="#000000",
    hover=True,
    hover_color="#949494",
    height=30,
    width=95,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#342e84",
    fg_color="#F0F0F0",
    )

Label_username.place(x=230, y=90)
Entry_username.place(x=330, y=90)
Label_password.place(x=230, y=180)
Entry_password.place(x=330, y=180)
Button_login.place(x=360, y=270)

login_window.mainloop()