from gui.email_sender_gui import EmailSenderGUI
import ttkbootstrap as tb

if __name__ == "__main__":
    app = tb.Window(themename="superhero")
    gui = EmailSenderGUI(app)
    app.mainloop()