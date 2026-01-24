import pyautogui as pag
import comtypes.client


def make_icon_finder(icons_dir):
    def icon(name):
        for candidate in (name, name.swapcase()):
            path = icons_dir / candidate
            if path.exists():
                return str(path)
        raise FileNotFoundError(f"Icon not found: {name}")

    return icon


# def connect_to_sap2000():
#     """Connect to the active SAP2000 instance"""
#     global sap_object, sap_model
#     try:
#         sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
#         sap_model = sap_object.SapModel
#         print("✓ Connected to SAP2000")
#         return True
#     except Exception as e:
#         print(f"✗ Could not connect to SAP2000: {e}")
#         return False


def get_node_distance(sap_model, node1_id, node2_id):
    """Calculate the distance between two nodes"""
    node1 = sap_model.PointObj.GetCoordCartesian(str(node1_id))
    node2 = sap_model.PointObj.GetCoordCartesian(str(node2_id))
    distance = (
        (node2[0] - node1[0]) ** 2
        + (node2[1] - node1[1]) ** 2
        + (node2[2] - node1[2]) ** 2
    ) ** 0.5
    return distance


def get_lp(L, db, fye):
    """Calculate the plastic hinge length"""
    lp = max(0.08 * L + 0.15 * fye * db, 0.3 * fye * db)
    return lp


def click_image(image_path, n=1, confidence=0.8):

    try:
        # Find the image on screen
        location = pag.locateOnScreen(image_path, confidence=confidence)
        # location = pag.locateOnScreen(image_path)
        print(location)
        print(1)
        if location:
            # Get center coordinates
            center = pag.center(location)

            # Click the center
            if n == 1:
                pag.click(center.x, center.y)
            else:
                pag.doubleClick(center.x, center.y)
            print(f"✓ Clicked {image_path} at ({center.x}, {center.y})")
            return True
        else:
            print(f"✗ Could not find {image_path}")
            return False

    except Exception as e:
        print(f"✗ Error: {e}")
        return False


def click_image_offset(
    image_path,
    x_offset_ratio=0.5,
    y_offset_ratio=0.5,
    n=1,
    confidence=0.8,
):

    try:
        # Find the image on screen
        location = pag.locateOnScreen(image_path, confidence=confidence)

        if location:
            # Calculate click position based on ratios
            left, top, width, height = location
            target_x = left + (width * x_offset_ratio)
            target_y = top + (height * y_offset_ratio)

            # Perform the click(s)
            if n == 1:
                pag.click(target_x, target_y)
            elif n == 2:
                pag.doubleClick(target_x, target_y)
            else:
                # For more than 2 clicks
                pag.click(target_x, target_y, clicks=n, interval=0.1)

            print(
                f"✓ Clicked {image_path} at ({target_x:.0f}, {target_y:.0f}) "
                f"[x_ratio={x_offset_ratio}, y_ratio={y_offset_ratio}]"
            )
            return True
        else:
            print(f"✗ Could not find {image_path}")
            return False

    except Exception as e:
        print(f"✗ Error: {e}")
        return False


def maximize_window_temporarily():
    """Maximize window with faster keystrokes"""
    # Save current PAUSE setting
    original_pause = pag.PAUSE

    # Temporarily reduce PAUSE for this operation
    pag.PAUSE = 0.1  # Much faster

    # Perform the maximize sequence
    pag.hotkey("alt")
    pag.hotkey("space")
    pag.hotkey("x")

    # Restore original PAUSE
    pag.PAUSE = original_pause
