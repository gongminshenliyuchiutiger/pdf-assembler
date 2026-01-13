import fitz

def test_rotation():
    doc = fitz.open()
    page = doc.new_page(width=200, height=400) # Portrait
    
    print(f"Original Rect: {page.rect}")
    
    # 2. Rotate 90
    page.set_rotation(90)
    print(f"Rotated 90 Rect: {page.rect}")
    
    # 3. Add Text "Visual Top Left"
    # Try inserting at (20, 20)
    page.insert_text((20, 20), "TL", fontsize=20, fontname="helv", rotate=0)

    # 4. Add Text "Visual Bottom Right"
    # Visual W=400, H=200. BR=(380, 180)
    # insert at (380, 180)
    page.insert_text((380, 180), "BR", fontsize=20, fontname="helv", rotate=0)

    # Analyze placements
    blocks = page.get_text("dict")["blocks"]
    for b in blocks:
        for l in b["lines"]:
            for s in l["spans"]:
                print(f"Text '{s['text']}' at bbox: {s['bbox']} origin: {s['origin']}")

    doc.save("test_rotation.pdf")
    print("Saved test_rotation.pdf")

if __name__ == "__main__":
    test_rotation()
