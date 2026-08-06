def go_to_a1(sheet):
    """Leave the tab with A1 selected and scrolled to the top-left."""
    try:
        sheet.api.Activate()
    except Exception:
        pass

    try:
        sheet.book.app.api.Goto(sheet.api.Range("A1"), True)
        return
    except Exception:
        pass

    try:
        sheet.api.Range("A1").Select()
        return
    except Exception:
        pass

    try:
        win = sheet.book.app.api.ActiveWindow
        win.ScrollRow = 1
        win.ScrollColumn = 1
    except Exception:
        print("    NOTE: could not reset the view to A1 (harmless).")
