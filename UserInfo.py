import tkinter as tk
import customtkinter as ctk
import os
import json
import pyodbc
import pkg_resources
from CTkMessagebox import CTkMessagebox
from PIL import ImageTk, Image
from App import App
from cryptography.fernet import Fernet


class UserInfo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("600x440")
        self.resizable(False, False)
        self.title('Login')

        # Customize appearance
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("green")

        # Background image
        background_image_path = pkg_resources.resource_filename(__name__, 'Gainwell_background.png')
        self.img1 = ImageTk.PhotoImage(Image.open(background_image_path))
        self.l1 = ctk.CTkLabel(self, image=self.img1)
        self.l1.pack()

        # Main Frame
        self.frame = ctk.CTkFrame(self.l1, width=360, height=300, corner_radius=15)
        self.frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Label
        self.l2 = ctk.CTkLabel(self.frame, text="Log into your Account", font=ctk.CTkFont(size=25))
        self.l2.place(x=60, y=35)

        # Remember login details checkbox
        self.remember_var = tk.IntVar()
        self.remember_cb = ctk.CTkCheckBox(self.frame, text="Remember Me", variable=self.remember_var)
        self.remember_cb.place(x=60, y=195)

        # Username & Password Entry
        self.username_entry = ctk.CTkEntry(self.frame, width=220, placeholder_text="Username")
        self.username_entry.place(x=60, y=95)
        self.password_entry = ctk.CTkEntry(self.frame, width=220, placeholder_text="Password", show="*")
        self.password_entry.place(x=60, y=150)

        # Login Button
        self.button1 = ctk.CTkButton(self.frame, width=220, text='Login', font=ctk.CTkFont(size=15, weight="bold"),
                                     corner_radius=6, command=self.on_login)
        self.button1.place(x=60, y=235)

        # Check saved login
        self.check_saved_login()

    def encrypt_existing_credentials(self):
        # Load the existing plaintext login info
        with open("login_info.json", "r") as file:
            login_info = json.load(file)

        # Encrypt the credentials
        encrypted_text = self.encrypt_login(login_info['username'], login_info['password'])

        # Save the encrypted login info
        with open("login_info.json", "w") as file:
            json.dump({"encrypted_login": encrypted_text.decode()}, file)

    @staticmethod
    def generate_key():
        """
        Generates a key and saves it into a file
        """
        key = Fernet.generate_key()
        with open("secret.key", "wb") as key_file:
            key_file.write(key)

    @staticmethod
    def load_key():
        """
        Loads the key from the current directory named `secret.key`
        """
        if os.path.exists("secret.key"):
            return open("secret.key", "rb").read()
        else:
            UserInfo.generate_key()
            return open("secret.key", "rb").read()

    def encrypt_login(self, username, password):
        key = self.load_key()
        cipher_suite = Fernet(key)
        encrypted_text = cipher_suite.encrypt(f"{username}|{password}".encode())
        return encrypted_text

    def decrypt_login(self, encrypted_login):
        key = self.load_key()
        cipher_suite = Fernet(key)
        decrypted_text = cipher_suite.decrypt(encrypted_login).decode()
        username, password = decrypted_text.split("|")
        return username, password

    def on_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        # Only save login if remember is checked
        if self.remember_var.get() == 1:
            encrypted_login = self.encrypt_login(username, password)
            with open("login_info.json", "w") as file:
                json.dump({"encrypted_login": encrypted_login.decode()}, file)

        try:
            # Attempt to open a connection
            DSN = "EDWPROD"
            self.conn = pyodbc.connect(f'DSN={DSN};UID={username};PWD={password}')

            # Attempt connection to DB2
            DB2_DSN = "ECPROD"
            self.db2_conn = pyodbc.connect(f'DSN={DB2_DSN};UID={username};PWD={password}')

            # If successful, close the login screen and open the main app with connection
            self.destroy()
            main_app = App(self.conn, self.db2_conn)
            main_app.mainloop()

        except pyodbc.OperationalError as oe:
            # This will catch common database operational errors such as login failures
            CTkMessagebox(title="Connection Error", message="Failed to connect. Please check your credentials.",
                          icon="cancel")
        except pyodbc.Error as e:
            # This will catch all other pyodbc related errors
            CTkMessagebox(title="Error", message=str(e), icon="cancel")
        except Exception as e:
            # A general catch-all for other exceptions
            CTkMessagebox(title="Unexpected Error", message=str(e), icon="cancel")

    def check_saved_login(self):
        if os.path.exists("login_info.json"):
            with open("login_info.json", "r") as file:
                login_info = json.load(file)
                if "encrypted_login" in login_info:
                    encrypted_login = login_info["encrypted_login"].encode()
                    username, password = self.decrypt_login(encrypted_login)
                    self.username_entry.insert(0, username)
                    self.password_entry.insert(0, password)
                    self.remember_var.set(1)


if __name__ == "__main__":
    user_info = UserInfo()
    user_info.mainloop()
