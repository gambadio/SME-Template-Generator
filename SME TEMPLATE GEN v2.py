import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import docx
from PIL import ImageGrab, ImageTk, Image
from io import BytesIO
import clipboard

class ImageLabel(tk.Label):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.bind("<B1-Motion>", self.drag)
        self.bind("<Button-1>", self.click)
        self.bind("<ButtonRelease-1>", self.drop)

    def click(self, event):
        self.startX = event.x
        self.startY = event.y

    def drag(self, event):
        x = self.winfo_x() - self.startX + event.x
        y = self.winfo_y() - self.startY + event.y
        self.place(x=x, y=y)

    def drop(self, event):
        # This is where you would handle dropping the image label
        pass

class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

class IssueReportingApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # Initialize UI elements
        self.init_ui()

        # List to store images
        self.images = []

    def init_ui(self):
        self.title("Issue Reporting App")
        self.geometry("800x600")

        # Create a frame for the content on the left side
        self.content_frame = tk.Frame(self)
        self.content_frame.pack(side='left', fill='both', expand=True)

        # Create a scrollable frame on the right side for the images
        self.image_frame = ScrollableFrame(self)
        self.image_frame.pack(side='right', fill='both', expand=True)

        # Create input fields for the required information
        ttk.Label(self.content_frame, text="Req-Number").grid(row=0, column=0, sticky='w', pady=5)
        self.req_number = ttk.Entry(self.content_frame)
        self.req_number.grid(row=0, column=1, pady=5)

        # (Continue with rest of your code, replacing content_frame with self.content_frame)
        ttk.Label(self.content_frame, text="Company").grid(row=1, column=0, sticky='w', pady=5)
        self.company = ttk.Entry(self.content_frame)
        self.company.grid(row=1, column=1, pady=5)

        ttk.Label(self.content_frame, text="Country").grid(row=2, column=0, sticky='w', pady=5)
        self.country = ttk.Combobox(self.content_frame, values=['USA', 'UK', 'Germany', 'France'])  # Add more countries as needed
        self.country.grid(row=2, column=1, pady=5)

        ttk.Label(self.content_frame, text="Account name").grid(row=3, column=0, sticky='w', pady=5)
        self.account_name = ttk.Entry(self.content_frame)
        self.account_name.grid(row=3, column=1, pady=5)

        ttk.Label(self.content_frame, text="Impacted user's ID(s)").grid(row=4, column=0, sticky='w', pady=5)
        self.user_ids = ttk.Entry(self.content_frame)
        self.user_ids.grid(row=4, column=1, pady=5)

        ttk.Label(self.content_frame, text="Environment, Application, Sub Product, Dataset").grid(row=5, column=0, sticky='w', pady=5)
        self.dataset = ttk.Entry(self.content_frame)
        self.dataset.grid(row=5, column=1, pady=5)

        ttk.Label(self.content_frame, text="Report Details").grid(row=6, column=0, sticky='w', pady=5)
        self.report_details = scrolledtext.ScrolledText(self.content_frame, width=40, height=5)
        self.report_details.grid(row=6, column=1, pady=5)

        ttk.Label(self.content_frame, text="Is the issue replicable?").grid(row=7, column=0, sticky='w', pady=5)
        self.replicable = ttk.Combobox(self.content_frame, values=['YES', 'NO'])
        self.replicable.grid(row=7, column=1, pady=5)

        ttk.Label(self.content_frame, text="Steps/Troubleshooting").grid(row=8, column=0, sticky='w', pady=5)
        self.steps = scrolledtext.ScrolledText(self.content_frame, width=40, height=5)
        self.steps.grid(row=8, column=1, pady=5)

        ttk.Label(self.content_frame, text="Time and timezone of error").grid(row=9, column=0, sticky='w', pady=5)
        self.error_time = ttk.Entry(self.content_frame)
        self.error_time.grid(row=9, column=1, pady=5)

        ttk.Label(self.content_frame, text="Describe the issue").grid(row=10, column=0, sticky='w', pady=5)
        self.issue_description = scrolledtext.ScrolledText(self.content_frame, width=40, height=5)
        self.issue_description.grid(row=10, column=1, pady=5)

        # Image display area
        self.image_label = tk.Label(self.content_frame)  
        self.image_label.grid(row=11, column=0, columnspan=2)

        # Bind Ctrl+V for pasting images
        self.bind('<Control-v>', self.paste_screenshot)

        # Create buttons for generating output
        self.generate_button = ttk.Button(self.content_frame, text="Generate Word Document", command=self.generate_word_document)
        self.generate_button.grid(row=12, column=0, columnspan=2, pady=5)

        self.copy_output_button = ttk.Button(self.content_frame, text="Copy Output", command=self.generate_copy_ready_text)
        self.copy_output_button.grid(row=13, column=0, columnspan=2, pady=5)

        # Create a button for exiting the application
        self.exit_button = ttk.Button(self.content_frame, text="Exit", command=self.exit_application)
        self.exit_button.grid(row=14, column=0, columnspan=2, pady=5)

    def paste_screenshot(self, event):
        try:
            image = ImageGrab.grabclipboard()
            if isinstance(image, Image.Image):  # Make sure it's an image
                # Save the image and add it to the list
                filename = f"image_{len(self.images)}.png"
                image.save(filename)
                self.images.append(filename)

                # Display the image in the application
                image.thumbnail((100, 100))  # Reduce the size of the image
                photo = ImageTk.PhotoImage(image)
                img_label = ImageLabel(self.image_frame.scrollable_frame, image=photo)  # Add to ScrollableFrame
                img_label.image = photo
                img_label.place(x=100, y=100)  # Use place instead of grid to allow moving the label
            else:
                print("No image in the clipboard.")
        except Exception as e:
            print(e)  # This will print the actual error message
            messagebox.showerror('Error', 'Could not paste the image!')

    # (Continue with the rest of your code...)

    def exit_application(self):
        msg_box = messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application?', icon='warning')
        if msg_box == 'yes':
            self.destroy()
        else:
            messagebox.showinfo('Return', 'You will now return to the application screen')

if __name__ == '__main__':
    app = IssueReportingApp()
    app.mainloop()
