import fitz
import os

def test_chinese_output():
    doc = fitz.open()
    page = doc.new_page()
    
    text = "測試頁面 name"
    print(f"Testing text: {text}")
    
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc", # Microsoft JhengHei
        "C:/Windows/Fonts/msyh.ttc", # Microsoft YaHei
        "C:/Windows/Fonts/simsun.ttc", # SimSun
        "C:/Windows/Fonts/arial.ttf" # Fallback
    ]
    
    font_file_used = None
    for fp in font_paths:
        if os.path.exists(fp):
            print(f"Found font: {fp}")
            try:
                # Try to load it to see if it works
                font = fitz.Font(fontfile=fp)
                font_file_used = fp
                print(f"Successfully loaded font: {fp}")
                break
            except Exception as e:
                print(f"Failed to load font {fp}: {e}")
                
    p = fitz.Point(100, 100)
    
    try:
        if font_file_used:
            print(f"Inserting text with fontfile: {font_file_used}")
            # Note: insert_text with fontfile requires fontname usually if not providing a font object, 
            # or simply passing the file path might work but fitz usually needs 'fontname' to reference it 
            # OR use 'fontfile' argument effectively.
            # In PyMuPDF, insert_text(point, text, ..., fontfile=path, ...)
            # Note: insert_text with fontfile requires fontname usually if not providing a font object, 
            # or simply passing the file path might work but fitz usually needs 'fontname' to reference it 
            # OR use 'fontfile' argument effectively.
            # In PyMuPDF, insert_text(point, text, ..., fontfile=path, ...)
            page.insert_text(p, text, fontsize=20, fontfile=font_file_used, fontname="custom_font")
        else:
            print("Using fallback 'china-ts'")
            page.insert_text(p, text, fontsize=20, fontname="china-ts")
            
        doc.save("test_chinese_output.pdf")
        print("Saved test_chinese_output.pdf")
        
    except Exception as e:
        print(f"Error inserting text: {e}")

if __name__ == "__main__":
    test_chinese_output()
