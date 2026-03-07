Build instructions for peru_compras_bot.exe (GUI)

1) (Optional) Create and activate a virtual environment:
   python -m venv .venv
   .venv\Scripts\activate

2) Build executable (Windows):
   build_exe.bat

3) Result:
   dist\peru_compras_bot.exe

Important notes:
- The generated executable opens a graphical interface (no terminal window).
- The user can select any Excel file from the interface (no need to place files manually).
- Filtros (Acuerdo, Catálogo, Categoría) are editable directly in the GUI.
- Chrome/Chromium must be installed on the machine.
- webdriver-manager downloads the compatible chromedriver automatically.
