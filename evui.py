from tkinter import *
from tkinter import messagebox


class Evui:
    """
    A wrapper for TKinter that makes it easy to create graphical user interfaces.
    """

    def __init__(self, **kwargs):
        pass


if __name__ == '__main__':

    root = Tk()
    root.geometry("400x200")
    root.resizable(False, False)
    root.title("Account Numbers")
    root.iconbitmap("favicon.ico")

    def button_click():
        text = account_number_input.get()
        if text != "":
            account_numbers = text.strip().replace(" ", "").split(",")
            sep = '\n    '
            user_confirmation = messagebox.askyesno(
                "Confirm Account Numbers",
                f"These are your account numbers:\n\n    {sep.join(account_numbers)}\n\nWould you like to continue?"
            )
            if user_confirmation:
                root.destroy()

    label = Label(
        root,
        text="Enter the account number(s) separated by commas",
        font=("Helvetica", 10)
    )
    account_number_input = Entry(
        root,
        width=40,
        font=("Helvetica", 10),
        relief=FLAT,
        borderwidth=10
    )
    run_bot_button = Button(
        root,
        text="Print Christmas Trees",
        padx=8,
        pady=5,
        font=("Helvetica", 10),
        command=button_click
    )

    label.place(relx=0.5, rely=0.25, anchor=CENTER)
    account_number_input.place(relx=0.5, rely=0.45, anchor=CENTER)
    run_bot_button.place(relx=0.5, rely=0.7, anchor=CENTER)

    root.mainloop()
