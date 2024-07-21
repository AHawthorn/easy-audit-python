import traceback

from app.ui import main_ui

if __name__ == "__main__":
    try:
        main_ui()
    except Exception as e:
        with open("error.log", "w") as f:
            f.write(str(e))
            f.write(traceback.format_exc())