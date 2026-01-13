
import fitz
import os

def debug_overlay():
    # Create a new PDF
    doc = fitz.open()
    page = doc.new_page(width=595, height=842) # A4
    
    # Rotation test cases
    rotations = [0, 90, 180, 270]
    
    font_path = "C:/Windows/Fonts/msjh.ttc"
    font = fitz.Font(fontfile=font_path) if os.path.exists(font_path) else fitz.Font("heiti")
    
    for rot in rotations:
        p = doc.new_page(width=595, height=842)
        p.set_rotation(rot)
        
        # Overlay Logic
        # Position: Bottom-Right
        margin = 20
        size = 12
        text = f"Page Rot {rot}"
        
        rect = p.rect
        w = rect.width
        h = rect.height
        
        # Calc Vis Coords
        # Bottom-Right
        vx = w - margin
        vy = h - margin
        
        # Alignment
        width = font.text_length(text, fontsize=size)
        vx -= width
        
        p_vis = fitz.Point(vx, vy)
        mat = p.derotation_matrix
        p_phys = p_vis * mat
        
        text_rot = -rot
        
        print(f"Rot: {rot}")
        print(f"  Rect: {rect}")
        print(f"  Vis: ({vx}, {vy})")
        print(f"  Phys: {p_phys}")
        print(f"  Text Rot: {text_rot}")
        
        try:
            p.insert_text(p_phys, text, fontsize=size, rotate=text_rot, fontfile=font_path if os.path.exists(font_path) else None, fontname="heiti" if not os.path.exists(font_path) else None)
        except Exception as e:
            print(f"Error: {e}")

    doc.save("debug_overlay.pdf")
    print("Saved debug_overlay.pdf")

if __name__ == "__main__":
    debug_overlay()
