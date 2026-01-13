import fitz

def test_font():
    doc = fitz.open()
    page = doc.new_page()
    
    text = "測試 Testing Page 1/1"
    
    # Test 1: china-ts
    try:
        page.insert_text((50, 50), f"china-ts: {text}", fontsize=20, fontname="china-ts", color=(0,0,0))
        print("Success: china-ts")
    except Exception as e:
        print(f"Failed: china-ts - {e}")

    # Test 2: Standard helv (no Chinese)
    try:
        page.insert_text((50, 100), "Helv: Testing", fontsize=20, fontname="helv", color=(0,0,0))
        print("Success: helv")
    except Exception as e:
        print(f"Failed: helv - {e}")
        
    doc.save("test_font.pdf")
    print("Saved test_font.pdf")

if __name__ == "__main__":
    test_font()
