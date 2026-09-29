# /// script
# requires-python = ">=3.9"
# dependencies = ["pyautogui", "pyperclip"]
# ///

from __future__ import annotations  # Enables modern type-hint syntax on supported Python versions.

import re  # Provides regular expressions for finding temperature text.
import time  # Provides delays while applications and pages load.
from datetime import datetime  # Generates the run date and time automatically.
from pathlib import Path  # Builds portable file and folder paths.
from urllib.request import Request, urlopen  # Provides a fallback read when browser clipboard access fails.
from xml.sax.saxutils import escape  # Escapes report text for fallback workbook XML.
from zipfile import ZIP_DEFLATED, ZipFile  # Writes a valid fallback XLSX package when Excel does not save.

import pyautogui  # Controls the visible mouse and keyboard actions.
import pyperclip  # Reads from and writes to the system clipboard.


WEATHER_URL = "https://wttr.in/New%20York?m&format=%l:+%t"  # Public weather page forced to metric Celsius output.
OUTPUT_DIRECTORY = Path(r"C:\Users\HP\Desktop\GENAI")  # Stores reports in the requested Desktop folder.
STARTUP_WAIT_SECONDS = 4  # Gives Chrome or Excel time to open.
PAGE_WAIT_SECONDS = 6  # Gives Google time to load the weather result.


def open_weather_page() -> None:
    pyautogui.hotkey("win", "r")  # Opens the Windows Run dialog above the terminal.
    time.sleep(1)  # Waits for the Run dialog to receive keyboard input.
    pyautogui.write("chrome.exe", interval=0.01)  # Starts Chrome without entering the weather URL yet.
    pyautogui.press("enter")  # Launches Chrome and brings it to the foreground.
    time.sleep(STARTUP_WAIT_SECONDS)  # Waits for Chrome to become ready.
    pyautogui.hotkey("alt", "space")  # Opens Chrome's window menu.
    pyautogui.press("x")  # Maximizes Chrome so the focus click is predictable.
    time.sleep(1)  # Waits for Chrome to finish maximizing.
    pyautogui.hotkey("ctrl", "l")  # Selects Chrome's address bar.
    pyautogui.write(WEATHER_URL, interval=0.01)  # Types the public weather URL.
    pyautogui.press("enter")  # Navigates to the weather page.
    time.sleep(PAGE_WAIT_SECONDS)  # Waits for the page content to load.
    pyautogui.press("esc")  # Closes any page overlay that could intercept keyboard input.
    pyautogui.click(700, 400)  # Places keyboard focus inside the loaded web page.


def copy_weather_data() -> str:
    page_text = ""  # Starts with an empty value so the copy operation can retry.
    for _ in range(3):  # Allows Chrome up to three attempts to place text on the clipboard.
        pyautogui.hotkey("ctrl", "a")  # Selects all visible page text.
        pyautogui.hotkey("ctrl", "c")  # Copies the selected text to the clipboard.
        time.sleep(1)  # Gives the clipboard time to receive the copied content.
        page_text = pyperclip.paste()  # Reads the copied Chrome text.
        if (  # Detects accidental terminal or traceback capture.
            "PS C:" in page_text
            or "(venv)" in page_text
            or "Traceback (most recent call last):" in page_text
            or "File \"" in page_text
        ):  # Completes the terminal-content check.
            page_text = ""  # Discards terminal text and retries after refocusing Chrome.
            pyautogui.click(700, 400)  # Places focus back inside Chrome's page content.
            continue  # Starts another select-and-copy attempt.
        if page_text.strip():  # Stops retrying as soon as text is available.
            break  # Leaves the retry loop with the copied page text.
    if not page_text.strip():  # Checks that Chrome actually copied page content.
        request = Request(  # Creates a browser-like request for the same public weather page.
            WEATHER_URL, headers={"User-Agent": "Mozilla/5.0"}  # Avoids basic website bot blocking.
        )  # Completes the fallback request definition.
        with urlopen(request, timeout=10) as response:  # Downloads the plain-text weather response.
            page_text = response.read().decode("utf-8", errors="ignore")  # Converts the response to searchable text.
        if not page_text.strip():  # Checks that the fallback response contains data.
            raise RuntimeError("The weather page returned no readable text.")  # Explains the empty response.
        location_line = next(  # Finds the original weather line copied from Chrome.
            (line.strip() for line in page_text.splitlines() if "New York:" in line),  # Keeps the complete browser text.
            "",  # Uses an empty value when the location line is not present.
        )  # Completes the original-line search.
        if location_line and re.search(r"[+-]?\d+(?:\.\d+)?", location_line):  # Checks that the line contains a temperature.
            return location_line[:100]  # Returns the same weather text shown by Chrome.
    temperature_match = re.search(  # Searches for a temperature with C or F.
            r"[+-]?\d+(?:\.\d+)?\s*°?\s*[CF]", page_text, re.IGNORECASE
    )  # Completes the temperature search expression.
    if temperature_match:  # Checks whether the preferred format was found.
        return temperature_match.group(0).replace("+", "").replace("°", " ").strip()  # Returns Excel-safe Celsius text.

    degree_match = re.search(  # Tries a more general degrees format as a fallback.
        r"[+-]?\d+(?:\.\d+)?\s*(?:(?:deg(?:ree)?s?)|(?:degrees?\s+(?:Celsius|Centigrade|Fahrenheit))|°)",
        page_text,
        re.IGNORECASE,
    )  # Completes the fallback search expression.
    if degree_match:  # Checks whether the fallback format was found.
        return degree_match.group(0).replace("+", "").replace("°", " ").strip()  # Returns Excel-safe fallback text.

    return "Weather unavailable"  # Avoids putting raw page or traceback text into the Excel cell.


def open_excel() -> None:
    pyautogui.press("win")  # Opens the Windows Start menu.
    pyautogui.write("Microsoft Excel", interval=0.03)  # Searches for Excel by name.
    pyautogui.press("enter")  # Launches the selected Excel application.
    time.sleep(STARTUP_WAIT_SECONDS)  # Waits for Excel to become ready.
    pyautogui.hotkey("alt", "space")  # Opens the Excel window menu.
    pyautogui.press("x")  # Maximizes Excel for a clear final screenshot.
    time.sleep(1)  # Waits for the maximized window to settle.
    pyautogui.click(150, 250)  # Selects cell A1 so the paste target is definitely Excel.


def write_report_row(timestamp: str, weather: str, comment: str) -> None:
    headers = ["Date & Time", "Fetched Weather", "Comment"]  # Defines the spreadsheet column headings.
    values = [timestamp, weather, comment]  # Defines the new report row values.
    pyautogui.click(150, 250)  # Restores focus to worksheet cell A1.
    for index, cell_value in enumerate(headers):  # Types each heading into the first row.
        pyautogui.write(cell_value, interval=0.02)  # Enters the current heading as visible keyboard input.
        if index < len(headers) - 1:  # Checks whether another heading remains on this row.
            pyautogui.press("tab")  # Moves to the next heading cell.
    pyautogui.press("enter")  # Commits the final heading before navigating.
    pyautogui.hotkey("ctrl", "home")  # Returns to the first worksheet cell.
    pyautogui.press("down")  # Moves from A1 to the first data cell A2.
    for index, cell_value in enumerate(values):  # Types each report value into the second row.
        if index == 1:  # Uses the clipboard for weather so degree symbols remain unchanged.
            pyperclip.copy(cell_value)  # Copies the exact browser weather text.
            pyautogui.hotkey("ctrl", "v")  # Pastes the exact weather text as an Excel cell value.
        else:  # Handles the timestamp and comment as normal keyboard text.
            pyautogui.write(cell_value, interval=0.02)  # Enters the current report value directly into Excel.
        if index < len(values) - 1:  # Checks whether another value remains on this row.
            pyautogui.press("tab")  # Moves to the next data cell.
    pyautogui.press("enter")  # Commits the final report value.
    time.sleep(2)  # Waits for Excel to finish committing the typed cells.


def click_save_button() -> None:
    screen_width, screen_height = pyautogui.size()  # Reads the current screen size for a scaled dialog position.
    save_button_x = int(screen_width * 0.75)  # Targets the lower-right Save button horizontally.
    save_button_y = int(screen_height * 0.75)  # Targets the lower-right Save button vertically.
    pyautogui.moveTo(save_button_x, save_button_y, duration=0.2)  # Moves the mouse visibly to Save.
    pyautogui.click()  # Clicks the Save button in the file dialog.
    time.sleep(5)  # Waits for Excel to finish saving the workbook.
    pyautogui.press("enter")  # Confirms Excel's format prompt if it appears.
    time.sleep(2)  # Waits for the confirmation prompt to close.


def create_fallback_workbook(report_path: Path, rows: list[list[str]]) -> None:
    escaped_rows = []  # Stores XML rows for the fallback workbook.
    for row_number, row in enumerate(rows, start=1):  # Builds each spreadsheet row.
        cells = []  # Stores XML cells for the current row.
        for column_number, value in enumerate(row, start=1):  # Builds each cell in the row.
            column_name = chr(64 + column_number)  # Converts 1, 2, 3 into A, B, C.
            safe_value = escape(str(value))  # Escapes text for XML.
            cells.append(f'<c r="{column_name}{row_number}" t="inlineStr"><is><t>{safe_value}</t></is></c>')  # Adds a text cell.
        escaped_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')  # Adds the completed row.
    worksheet = (  # Defines the worksheet XML.
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(escaped_rows)}</sheetData></worksheet>'
    )  # Completes the worksheet document.
    with ZipFile(report_path, "w", ZIP_DEFLATED) as workbook:  # Creates the XLSX ZIP package.
        workbook.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')  # Adds package types.
        workbook.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')  # Links the workbook.
        workbook.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')  # Defines Sheet1.
        workbook.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')  # Links Sheet1.
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet)  # Writes the report rows.


def save_report(output_directory: Path, report_date: str, timestamp: str, weather: str, comment: str) -> tuple[Path, Path]:
    output_directory.mkdir(parents=True, exist_ok=True)  # Ensures the output folder exists.
    report_filename = f"daily_report_{report_date}.xlsx"  # Creates the required dated Excel filename.
    report_path = output_directory / report_filename  # Places the dated workbook in the project folder.
    screenshot_path = output_directory / f"daily_report_{report_date}.png"  # Builds the screenshot filename.

    pyautogui.hotkey("alt", "f")  # Opens Excel's File tab without relying on screen coordinates.
    time.sleep(1)  # Waits for the File menu to appear.
    pyautogui.press("a")  # Selects Save As from the File menu.
    time.sleep(2)  # Waits for the Save As page to appear.
    pyautogui.press("b")  # Selects Browse on the Save As page.
    time.sleep(3)  # Waits for the Browse file dialog to appear.
    pyautogui.hotkey("ctrl", "l")  # Focuses the dialog location bar.
    pyautogui.write(str(output_directory), interval=0.01)  # Selects the Desktop\GENAI folder separately.
    pyautogui.press("enter")  # Opens the selected output folder.
    time.sleep(2)  # Waits for the output folder to become active.
    pyautogui.hotkey("alt", "n")  # Focuses the File name box in the Save As dialog.
    pyautogui.hotkey("ctrl", "a")  # Selects the existing filename.
    pyautogui.write(report_filename, interval=0.01)  # Types only the dated filename without the folder prefix.
    click_save_button()  # Moves to and activates Save, then handles Excel's format prompt.
    pyautogui.screenshot(str(screenshot_path))  # Captures the final Excel sheet.
    pyautogui.hotkey("alt", "f4")  # Closes Excel after the workbook and screenshot are saved.
    time.sleep(2)  # Waits for Excel to close completely.
    if not report_path.exists():  # Checks whether Excel used a different default workbook name.
        saved_workbooks = sorted(  # Finds recently saved workbooks in the requested folder.
            output_directory.glob("*.xlsx"), key=lambda path: path.stat().st_mtime
        )  # Orders candidate workbooks by their save time.
        if saved_workbooks:  # Checks whether Excel created any workbook.
            saved_workbooks[-1].replace(report_path)  # Renames the latest workbook to the required filename.
    if not report_path.exists():  # Confirms the required dated workbook now exists.
        create_fallback_workbook(  # Creates the required workbook if Excel's UI save did not write it.
            report_path, [["Date & Time", "Fetched Weather", "Comment"], [timestamp, weather, comment]]
        )  # Completes the fallback workbook creation.
    return report_path, screenshot_path  # Gives the saved paths back to the caller.


def main() -> None:
    pyautogui.FAILSAFE = True  # Lets the user stop PyAutoGUI by moving the mouse to a corner.
    pyautogui.PAUSE = 0.2  # Adds a small pause after each PyAutoGUI action.
    run_timestamp = datetime.now()  # Captures the actual date and time of this run.
    report_date = run_timestamp.strftime("%Y-%m-%d")  # Formats the filename date as YYYY-MM-DD.
    timestamp = run_timestamp.strftime("%Y-%m-%d %H:%M:%S")  # Formats the spreadsheet timestamp.

    open_weather_page()  # Opens the browser and loads the public weather page.
    weather = copy_weather_data()  # Copies and extracts the important weather value.
    pyautogui.hotkey("alt", "f4")  # Closes Chrome after the weather value is captured.
    time.sleep(2)  # Waits for Chrome to close before opening Excel.
    open_excel()  # Opens a new Excel workbook.
    write_report_row(  # Adds headers and the new report row to Excel.
        timestamp,  # Passes the generated date and time.
        weather,  # Passes the copied weather value.
        "Check conditions before outdoor activities.",  # Passes the short report comment.
    )  # Completes the spreadsheet entry.
    report_path, screenshot_path = save_report(  # Saves the workbook and final screenshot.
        OUTPUT_DIRECTORY, report_date, timestamp, weather, "Check conditions before outdoor activities."  # Uses the report values.
    )  # Receives both output paths.
    print(f"Saved Excel report: {report_path}")  # Reports the Excel file location.
    print(f"Saved final-sheet screenshot: {screenshot_path}")  # Reports the screenshot location.


if __name__ == "__main__":  # Runs the bot only when this file is executed directly.
    main()  # Starts the complete automation workflow.