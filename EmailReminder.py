import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
import pandas as pd
import smtplib
import tempfile
import csv
import os
import uuid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import logging
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master, bg="#202B3D")
        self.master = master
        self.master.geometry("1300x500")
        self.master.title("Email Sender")
        self.pack(fill="both", expand=True)
        self.create_widgets()
        self.update_counter()
        self.data_loaded = False



    def create_widgets(self): # GUI
        self.display_box_label = tk.Label(self, text="1.Students to Remind:", bg="#202B3D", fg="#FFFFFF")
        self.display_box_label.grid(row=1, column=0, sticky="w")
        self.display_box = tkinter.Text(self, height=10, width=90)
        self.display_box.grid(row=1, column=0, sticky="w")

        # Count Label
        self.count_label = tk.Label(self, text="Entries: 0", bg="#202B3D", fg="#FFFFFF")
        self.count_label.grid(row=2, column=0, sticky="w")

        # Buttons
        self.button1 = tk.Button(self, text="Load Excel File", command=self.button1_clicked, bg="#4CAF50", fg="#FFFFFF")
        self.button1.grid(row=5, column=0, padx=10, pady=10, sticky="w")

        self.button2 = tk.Button(self, text="Send Emails", command=self.send_emails, bg="#FF5722", fg="#FFFFFF")
        self.button2.grid(row=5, column=0, padx=10, pady=10, sticky="e")

        self.button3 = tk.Button(self, text="Load Log File", command=self.load_log_file, bg="#4CAF50", fg="#FFFFFF")
        self.button3.grid(row=6, column=0, padx=10, pady=10, sticky="w")

        self.clear_log_button = tk.Button(self, text="Clear Log File", command=self.clear_log_file, bg="#FF5722", fg="#FFFFFF")
        self.clear_log_button.grid(row=6, column=0, padx=10, pady=10, sticky="e")

        self.log_label = tk.Label(self, text="Email Log:", bg="#202B3D", fg="#FFFFFF")
        self.log_label.grid(row=0, column=1, sticky="w")
        self.log_textbox = tk.Text(self, height=10, width=90, bg="#FFFFFF", fg="#000000")
        self.log_textbox.grid(row=1, column=1, padx=10, pady=10)

        self.count_label_oatridge = tk.Label(self, text="Oatridge: 0", bg="#202B3D", fg="#FFFFFF")
        self.count_label_oatridge.grid(row=3, column=0, sticky="w")

        self.count_label_barony = tk.Label(self, text="Barony: 0", bg="#202B3D", fg="#FFFFFF")
        self.count_label_barony.grid(row=4, column=0, sticky="w")

        self.count_label_craibstone = tk.Label(self, text="Craibstone: 0", bg="#202B3D", fg="#FFFFFF")
        self.count_label_craibstone.grid(row=3, column=0, sticky="e")

        self.count_label_elmwood = tk.Label(self, text="Elmwood: 0", bg="#202B3D", fg="#FFFFFF")
        self.count_label_elmwood.grid(row=4, column=0, sticky="e")
        
        self.pie_chart_button = tk.Button(self, text="Show Pie Chart", command=self.display_pie_chart, bg="#2196F3", fg="#FFFFFF")
        self.pie_chart_button.grid(row=5, column=1, padx=10, pady=10, sticky="w")
        
        self.open_browser_button = tk.Button(self, text="Open Browser", command=self.open_browser)
        self.open_browser_button.grid(row=6, column=1, padx=10, pady=10, sticky="w")


        self.update_counter()
            
    def update_counter(self, event=None):
        num_entries = len(self.display_box.get("1.0", "end-1c").split("\n"))
        # Update the count label
        self.count_label.config(text=f"Entries: {num_entries}")

    def open_browser(self, search_query): # Function for Selenium Web browser
        try:
            # Create a non-headless browser instance (you can change the browser if needed)
            browser = webdriver.Chrome()

            # Construct the search URL based on the search query and the link you want to navigate to
           # search_url = f"https://www.example.com/search?q={search_query}"

            # Navigate to the search URL
            #browser.get(search_url)

            # Optionally, you can add a delay to wait for the page to load
            sleep(5)  # Adjust the sleep duration as needed

            # Close the browser when done
            browser.quit()
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            
    def button1_clicked(self):
        filename = tkinter.filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            try:
                df = pd.read_excel(filename)

                # Get the integer index for the "VQ Level (Derived Qualification) (Qualification)" column
                level_column_index = None
                for column_name in df.columns:
                    if column_name.strip().lower() == "vq level (derived qualification) (qualification)":  # Case-insensitive comparison with whitespace removed
                        level_column_index = df.columns.get_loc(column_name)
                        break

                # Check if the "VQ Level (Derived Qualification) (Qualification)" column was found
                if level_column_index is not None:
                    # Process the data based on the integer index
                    data = []
                    unique_individuals = set()  # Store unique individuals based on their email addresses
                    campus_counts = {"Barony": 0, "Oatridge": 0, "Elmwood": 0, "Craibstone": 0}  # Initialize campus counters

                    for _, row in df.iterrows():
                        # Access data using integer indices directly
                        NAME = row.iloc[5]
                        EMAIL = row.iloc[9]
                        PHONE = row.iloc[10]
                        LEVEL = row.iloc[level_column_index]  # Use the correct index for "VQ Level (Derived Qualification) (Qualification)"
                        FRAMEWORK = row.iloc[13]
                        TYPE = row.iloc[18]
                        # Use the exact column name from your Excel file for CAMPUS
                        CAMPUS = row.iloc[24]
                        
                        print(row.iloc[24])

                        # Check if this individual has been encountered before
                        if EMAIL not in unique_individuals:
                            # Update campus counts for the first occurrence of the individual
                            if CAMPUS in campus_counts:
                                campus_counts[CAMPUS] += 1

                            # Add the individual to the set of unique individuals
                            unique_individuals.add(EMAIL)

                            # Add the data to the list
                            data.append([NAME, EMAIL, PHONE, LEVEL, FRAMEWORK, TYPE, CAMPUS])  # Add additional columns as needed

                    # Save the processed data to a temporary CSV file
                    with open("temp_data.csv", mode="w", newline="") as file:
                        writer = csv.writer(file)
                        writer.writerows(data)

                    # Update the count labels
                    num_entries = len(data)
                    self.count_label.config(text=f"Entries: {num_entries}")

                    # Update the campus count labels
                    for campus, count in campus_counts.items():
                        if campus == "Barony":
                            self.count_label_barony.config(text=f"Barony: {count}")
                        elif campus == "Oatridge":
                            self.count_label_oatridge.config(text=f"Oatridge: {count}")
                        elif campus == "Elmwood":
                            self.count_label_elmwood.config(text=f"Elmwood: {count}")
                        elif campus == "Craibstone":
                            self.count_label_craibstone.config(text=f"Craibstone: {count}")

                    # Display the data in the display box
                    self.display_box.delete("1.0", tkinter.END)
                    for row in data:
                        self.display_box.insert(tkinter.END, f"{row}\n")
                    self.data_loaded = True

                else:
                    tkinter.messagebox.showerror("Error", "Column 'VQ Level (Derived Qualification) (Qualification)' not found.")

            except Exception as e:
                tkinter.messagebox.showerror("Error", f"Error loading Excel file: {str(e)}")

    def send_emails(self):
        # read data from temp_data.csv
        with open("temp_data.csv", mode="r") as file:
            reader = csv.reader(file)
            data = [row for row in reader]

        # group data by email
        df = pd.DataFrame(data, columns=["NAME", "EMAIL", "PHONE", "LEVEL", "FRAMEWORK", "TYPE", "CAMPUS"])
        grouped = df.groupby("EMAIL")

        try:
            # create an SMTP connection
            with smtplib.SMTP("relay.domain.ac.uk", 25) as server:  # Replace with your SMTP relay server
                # Log in to the SMTP server
                # server.login("your_username", "your_password")

                # Loop through groups of emails
                for email, group in grouped:
                    # Retrieve TYPE from the group DataFrame
                    TYPE = group["TYPE"].iloc[0]

                    # Connect to the SMTP server
                    server.connect("relay.domain.ac.uk", 25)  # Replace with your SMTP relay server

                    # Log in to the SMTP server if required
                    # server.login("your_username", "your_password")

                    # Create an email message
                    msg = MIMEMultipart()
                    msg["Subject"] = f'{TYPE} Confirmation Reminder'  # Use the correct variable for TYPE
                    msg["From"] = "luke.mcdonald@domain.ac.uk"  
                    name = group["NAME"].iloc[0]
                    framework = group["FRAMEWORK"].iloc[0]
                    campus = group["CAMPUS"].iloc[0]
                    level = group["LEVEL"].iloc[0]
                    email_body = f"""\
                <html>
                  <body>
                    <p>Dear {name} ,</p>
                    <p>This is a reminder to verify your claims/assignments for your {TYPE}. Please find the email/text from "SDS" by using your search and respond with a "y".</p>
                    <p>We have been notified that your claim is still unconfirmed and therefore still outstanding. This should be done as soon as possible to remain within contract bounds for your {level} course in {framework} at SRUC {campus}</p>
                    <p>If you cannot find the email/text it could be that it had been sent in the last couple weeks or it will be next friday. So please search for this using the search of "SDS" in your inboxes.</p>
                    <p>For those who have not received this yet please wait till next friday to check & contact me back.</p>
                    <p>Best Regards</p>
                    <br>Luke McDonald
                    <br>Work Based Learning Administrator (Campus Name)
                    <br>Business Address
                    <br>Campus:(Name)
                    <br>+44 00000000000000
                    <p><img src="https://d2kmbl0trprm6g.cloudfront.net/images/made/images/uploads/general/Uni-logo-SRUC_730_290_80.jpg" width="450" height="150"></p>
                  </body>
                </html>
               """
                    msg.attach(MIMEText(email_body, "html"))

                    # Send email to each recipient in the group
                    for _, row in group.iterrows():
                        recipient_email = row["EMAIL"]
                        try:
                            # Send the email
                            server.sendmail(msg["From"], recipient_email, msg.as_string())

                            # Log successful email send
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            log_msg = f"{timestamp} - Email sent to {recipient_email}"
                            self.log_textbox.insert("end", log_msg + "\n")
                            with open("log_file.txt", mode="a") as log_file:
                                log_file.write(log_msg + "\n")
                        except Exception as e:
                            # Log failed email send
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            log_msg = f"{timestamp} - Email to {recipient_email} failed: {str(e)}"
                            self.log_textbox.insert("end", log_msg + "\n")
                            with open("log_file.txt", mode="a") as log_file:
                                log_file.write(log_msg + "\n")

                    # Quit the SMTP server after sending emails to all recipients in the group
                    server.quit()

        except smtplib.SMTPException as e:
            # Handle SMTP exceptions
            print(f"SMTP Exception: {e}")
        except Exception as e:
            # Handle other exceptions
            print(f"Error: {e}")

    def load_log_file(self):
        try:
            with open("log_file.txt", "r") as f:
                contents = f.read()
                self.log_textbox.delete("1.0", "end")
                self.log_textbox.insert("end", contents)
        except FileNotFoundError:
            tk.messagebox.showinfo("File not found", "Log file not found.")

    def log_action(self, action):
        with open("log_file.txt", "a") as f:
            f.write(action + "\n")

    def clear_log_file(self):
        # Clear contents of log file
        with open("log_file.txt", "w") as log_file:
            log_file.write("")

        # Update log textbox
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.insert("end", "Log file cleared.\n")

    def create_pie_chart(self):
        if not self.data_loaded:
            return  # Exit if data is not loaded

        # Read data from temp_data.csv
        try:
            df = pd.read_csv("temp_data.csv")

            # Check if the 'CAMPUS' column exists in the DataFrame
            if 'CAMPUS' not in df.columns:
                tk.messagebox.showinfo("Column not found", "The 'CAMPUS' column does not exist in the data.")
                return

            # Calculate the percentage of students at each campus
            campus_counts = df['CAMPUS'].value_counts()
            labels = campus_counts.index
            sizes = campus_counts.values
            colors = ['gold', 'lightcoral', 'lightskyblue', 'lightgreen']  # You can define your own colors here

            # Create a pie chart
            fig, ax = plt.subplots()
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
            ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

            # Embed the pie chart in the Tkinter tab
            canvas = FigureCanvasTkAgg(fig, master=self.tab2)
            canvas.get_tk_widget().pack()
            canvas.draw()

        except FileNotFoundError:
            tk.messagebox.showinfo("File not found", "Data file (temp_data.csv) not found.")
            
    def display_pie_chart(self):
        # Gather data from the counters
        oatridge_count = int(self.count_label_oatridge.cget("text").split(":")[1].strip())
        barony_count = int(self.count_label_barony.cget("text").split(":")[1].strip())
        craibstone_count = int(self.count_label_craibstone.cget("text").split(":")[1].strip())
        elmwood_count = int(self.count_label_elmwood.cget("text").split(":")[1].strip())

        # Create a list of campus names and their corresponding counts
        campuses = ["Oatridge", "Barony", "Craibstone", "Elmwood"]
        counts = [oatridge_count, barony_count, craibstone_count, elmwood_count]

        # Create a pie chart
        plt.figure(figsize=(6, 6))
        plt.pie(counts, labels=campuses, autopct='%1.1f%%', startangle=140)
        plt.title("Campus Distribution")

        # Display the pie chart
        plt.axis('equal')  # Equal aspect ratio ensures that the pie chart is circular.
        plt.show()            

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
