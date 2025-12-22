import tkinter as tk
from tkinter import messagebox, filedialog
import sys

def select_folder(msg: str = "") -> str:
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    result = messagebox.askokcancel(
    title="디렉토리 선택",
    message=msg
    )

    # 2. 취소 클릭 시 즉시 종료
    if not result:
        root.destroy()
        sys.exit(0)

    # 3. 확인 클릭 시 디렉토리 선택
    folder_path = filedialog.askdirectory(
        title="메트릭 디렉토리 선택"
    )

    # 디렉토리 선택창에서 다시 취소한 경우도 종료
    if not folder_path:
        root.destroy()
        sys.exit(0)

    root.destroy()
    return folder_path