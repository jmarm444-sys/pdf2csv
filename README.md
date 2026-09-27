#pdf2csv

extract values from pdf format of usb holdings to .csv for excel, only pulls out symbol, quantity and per-share cost basis.
The pdf statement table has eight data columns. The CSV keeps three:
    sym, qty, per-share
where per-share is cost basis divided by quantity (half-up to the cent).
Example: AAPL, quantity 3, cost basis 945.39 -> aapl,3,315.13

Requirements (Windows 10/11):
    pip install pymupdf winocr pillow

Run with no arguments for the Browse/Convert window.
Run with two arguments for the command line:
    python pdf2csv_usbholdings.py input.pdf output.csv

