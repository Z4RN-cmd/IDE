import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from PIL import Image
import os, pygame
pygame.mixer.init()

# =========================
# APP
# =========================
def xp_start():
    pygame.mixer.Sound("xp.mp3").play()

app = ctk.CTk()

app.geometry("800x600")
app.title("IDE")

ctk.set_appearance_mode("light")


# =========================
# COLORS
# =========================
def on_startup_click():
    xp_start()
MAIN_COLOR = "#F00000"
DEFAULT_TEXT_COLOR = "#000000"
SECONDARY_COLOR = "#FFFFFF"





# =========================
# CREATE FILE POPUP
# =========================
def on_click():
    click_sound = pygame.mixer.Sound("click.mp3")
    click_sound.set_volume(1.0)
    click_sound.play()
def create_file_popup():

    popup = ctk.CTkToplevel(app)

    popup.geometry("400x250")
    popup.title("Create New File")
    popup.resizable(False, False)

    # Keep popup in front
    popup.transient(app)
    popup.grab_set()

    # =========================
    # POPUP TITLE
    # =========================

    popup_title = ctk.CTkLabel(
        popup,
        text="Enter file name:",
        font=("Inter", 20, "bold"),
        text_color=DEFAULT_TEXT_COLOR
    )

    popup_title.pack(
        pady=(30, 10)
    )

    # =========================
    # FILE NAME INPUT
    # =========================

    filename_entry = ctk.CTkEntry(
        popup,
        width=300,
        placeholder_text="Example: MyDocument"
    )

    filename_entry.pack(
        pady=10
    )

    filename_entry.focus()


    # =========================
    # CREATE FILE
    # =========================

    def create_file():

        filename = filename_entry.get().strip()

        # Check empty
        if not filename:

            CTkMessagebox(
                title="IDE",
                message="Please enter a file name!",
                icon="warning"
            )

            return

        # Add .txt automatically
        if not filename.lower().endswith(".txt"):
            filename += ".txt"

        # Check if file already exists
        if os.path.exists(filename):

            CTkMessagebox(
                title="IDE",
                message=(
                    f"The file already exists:\n\n"
                    f"{filename}"
                ),
                icon="warning"
            )

            return

        try:

            # Create empty file
            with open(filename, "w") as file:
                file.write("")

            # Get full path
            full_path = os.path.abspath(filename)

            # Success popup
            CTkMessagebox(
                title="IDE",
                message=(
                    "File created successfully!\n\n"
                    f"Created in:\n{full_path}"
                ),
                icon="check"
            )

            # Close popup
            popup.destroy()

        except Exception as error:

            CTkMessagebox(
                title="IDE",
                message=(
                    f"Failed to create file:\n\n"
                    f"{error}"
                ),
                icon="cancel"
            )


    # =========================
    # CREATE BUTTON
    # =========================

    create_button = ctk.CTkButton(
        popup,
        text="Create",
        width=150,
        fg_color=MAIN_COLOR,
        command=create_file
    )

    create_button.pack(
        pady=20
    )


# =========================
# SIDEBAR
# =========================
def notification_popup():
    CTkMessagebox(
        title="IDE",
        message="Support me for more features and updates!\n\n I don't have budget :(",
        icon="info"
    )
sidebar = ctk.CTkFrame(
    app,
    width=75,
    height=1000,
    fg_color=MAIN_COLOR,
    corner_radius=0
)

sidebar.place(
    x=0,
    y=0
)


# =========================
# TITLE
# =========================

title = ctk.CTkLabel(
    app,
    text="IDE Homepage",
    font=("Inter", 50, "bold", "italic"),
    text_color=DEFAULT_TEXT_COLOR,
    fg_color="transparent"
)

title.place(
    x=600,
    y=20
)


# =========================
# BLANK DOCUMENT IMAGE
# =========================

blank_image = ctk.CTkImage(
    light_image=Image.open("Blank_documents.png"),
    dark_image=Image.open("Blank_documents.png"),
    size=(200, 300)
)
coming_soon_image = ctk.CTkImage(
    light_image=Image.open("coming.png"),
    dark_image=Image.open("coming.png"),
    size=(200, 300)
)
def fah():
    pygame.mixer.Sound("fah.mp3").play()
def on_coming_soon_click():
    fah()
    notification_popup()
coming_soon = ctk.CTkButton(
    app,
    text="",
    image=coming_soon_image,
    width=130,
    height=130,
    fg_color="transparent",
    hover_color="#EEEEEE",
    corner_radius=15,
    command=on_coming_soon_click
)
coming_soon.place(
    x=500,
    y=180
)
# =========================
# BLANK DOCUMENT BUTTON
# =========================
show_image = ctk.CTkImage(
    light_image=Image.open("ishow.png"),
    dark_image=Image.open("ishow.png"),
    size=(130, 130)
)
ishow = ctk.CTkButton(
    app,
    text="",
    image=show_image,
    width=130,
    height=130,
    fg_color="transparent",
    hover_color="#EEEEEE",
    corner_radius=15,
    command=fah
)
ishow.place(
    x=1400,
    y=1
)
def on_blank_click():
    on_click()
    create_file_popup()

blank_button = ctk.CTkButton(
    app,
    text="",
    image=blank_image,
    width=130,
    height=130,
    fg_color="transparent",
    hover_color="#EEEEEE",
    corner_radius=15,
    command=on_blank_click
)

blank_button.place(
    x=200,
    y=180
)


# =========================
# LABEL
# =========================

document_label = ctk.CTkLabel(
    app,
    text="Blank Document",
    font=("Inter", 16, "bold"),
    text_color=DEFAULT_TEXT_COLOR,
    fg_color="transparent"
)

document_label.place(
    x=250,
    y=530
)


# =========================
# RUN APP
# =========================
xp_start()
app.mainloop()