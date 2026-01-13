
pyinstaller --noconfirm --onefile --windowed --name "pdf-assembler" --clean ^
    --hidden-import=PySide6.QtXml ^
    --collect-all=fitz ^
    main.py
