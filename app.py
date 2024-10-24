from tkinter import *
from chat import get_response, bot_name

BG_GRAY = "#EAEDED"
BG_COLOR = "#F5F5F5"
TEXT_COLOR = "#333333"
FONT = "Helvetica 14"
FONT_BOLD = "Helvetica 13 bold"

class ChatApplication:

    def __init__(self, parent):
        self.window = Toplevel(parent)
        self._setup_main_window()

    def run(self):
        self.window.mainloop()

    def _setup_main_window(self):
        self.window.title("Chatbot")
        self.window.resizable(width=False, height=False)
        self.window.configure(width=470, height=550, bg=BG_COLOR)


        head_label = Label(self.window, bg=BG_COLOR, fg=TEXT_COLOR,
                           text="Chào mừng bạn đến với Chatbot", font=FONT_BOLD, pady=10)
        head_label.place(relwidth=1)

        self.text_widget = Text(self.window, width=20, height=2, bg="#FFFFFF", fg=TEXT_COLOR,
                                font=FONT, padx=5, pady=5, wrap=WORD)
        self.text_widget.place(relheight=0.745, relwidth=1, rely=0.008)
        self.text_widget.configure(cursor="arrow", state="disabled")

        scrollbar = Scrollbar(self.text_widget)
        scrollbar.place(relheight=1, relx=0.974)
        scrollbar.configure(command=self.text_widget.yview)

        bottom_label = Label(self.window, bg=BG_GRAY, height=80)
        bottom_label.place(relwidth=1, rely=0.825)

        self.msg_entry = Entry(bottom_label, bg="#FFFFFF", fg=TEXT_COLOR, font=FONT)
        self.msg_entry.place(relwidth=0.74, relheight=0.06, rely=0.008, relx=0.011)
        self.msg_entry.focus()
        self.msg_entry.bind("<Return>", self._on_enter_pressed)

        # Nút gửi
        send_button = Button(bottom_label, text="➤", font=FONT_BOLD, width=5,
                             bg="#0084ff", fg="white", activebackground="#0056b3",
                             command=lambda: self._on_enter_pressed(None))
        send_button.place(relx=0.77, rely=0.008, relheight=0.06, relwidth=0.22)

    def _on_enter_pressed(self, event):
        msg = self.msg_entry.get()
        self._insert_message(msg, "You")

    def _insert_message(self, msg, sender):
        if not msg:
            return
        self.msg_entry.delete(0, END)
        msg_1 = f"{sender}: {msg}\n\n"
        self.text_widget.configure(state="normal")
        self.text_widget.insert(END, msg_1)
        self.text_widget.configure(state="disabled")

        msg_2 = f"{bot_name}: {get_response(msg)}\n\n"
        self.text_widget.configure(state="normal")
        self.text_widget.insert(END, msg_2)
        self.text_widget.configure(state="disabled")

        self.text_widget.see(END)


if __name__ == "__main__":
    app = ChatApplication()
    app.run()
