import fitz
import os

def test_overlay_logic():
    print("Testing Overlay Logic...")
    
    # Create dummy page
    doc = fitz.open()
    page = doc.new_page(width=100, height=200) # Portrait
    
    # Mock settings
    pos = "Bottom-Right"
    margin = 20
    size = 12
    text = "TEST"
    
    rotations = [0, 90, 180, 270]
    
    for rot in rotations:
        page.set_rotation(rot)
        rect = page.rect
        w = rect.width
        h = rect.height
        
        print(f"\n--- Rotation: {rot} ---")
        print(f"Visual Rect: {w}x{h}")
        
        # LOGIC FROM Main.py (Simplified)
        vx, vy = 0, 0
        if 'Bottom' in pos:
            vy = h - margin
        if 'Right' in pos:
            vx = w - margin
            
        # Mock Text Width
        width = 20
        vx -= width
        
        print(f"Visual Target: ({vx}, {vy}) (Should be Bottom-Right corner - margin)")
        
        # Transform
        p_vis = fitz.Point(vx, vy)
        mat = page.derotation_matrix
        p_phys = p_vis * mat
        
        print(f"Physical Point: {p_phys}")
        
        # Insert Text
        try:
             page.insert_text(p_phys, f"Rot{rot}", fontsize=size, rotate=-rot)
             print(f"Inserted text at {p_phys} with rot {-rot}")
        except Exception as e:
             print(f"Error: {e}")

    doc.save("verify_overlay.pdf")
    print("\nSaved verify_overlay.pdf. Please inspect visually.")

if __name__ == "__main__":
    test_overlay_logic()
