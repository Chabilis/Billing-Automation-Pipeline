# Simple Design GUI

This small Tkinter app collects shipment fields (Waybill, Date, Plate, Origin, Reference, Seal, Total Blocks) and provides a simple visual preview based on a selected design number (1–7).

Usage

- Run: `python GUI_userinput.py`
- Enter the fields on the left.
- Choose a Design number (1-7) and click `Preview Design`.
- To save the preview as a PostScript file, click `Save Preview (PS)`. The file `preview.ps` will be created in the current folder.

Notes

- Saving uses the Tkinter `postscript` method; no external packages required.
- To convert `preview.ps` to PNG, use external tools like `imagemagick` (`magick preview.ps preview.png`) or open in a viewer that supports PostScript.
