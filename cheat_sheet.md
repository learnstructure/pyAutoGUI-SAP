# 🚀 pyAutoGUI Complete Cheat Sheet
## Import & Setup
```
import pyautogui as pag
import time

## Essential safety settings

pag.FAILSAFE = True        ## Move mouse to (0,0) to abort
pag.PAUSE = 0.5            ## Delay between pyautogui actions (seconds)
```

# 📏 Screen Information 
## Basic screen info
```
width, height = pag.size()                ## Screen resolution
current_x, current_y = pag.position()     ## Current mouse position
```

## Multi-monitor support (Windows)
```
screen_count = pag.screen()._numMonitors  ## Number of monitors
primary_monitor = pag.screen()._primary   ## Primary monitor info
```

## Get active window info (requires PyGetWindow)
` pip install PyGetWindow `
```
import pygetwindow as gw
active = gw.getActiveWindow()
active_title = active.title
```

# 🖱️ MOUSE FUNCTIONS
## Position & Movement
```
## Get position
x, y = pag.position()                     ## Current position
pos = pag.position()                      ## Returns Point(x, y)

## Absolute movement
pag.moveTo(x, y)                          ## Move to coordinates
pag.moveTo(x, y, duration=1.5)            ## Smooth movement (seconds)
pag.moveTo(x, y, tween=pag.easeInOutQuad) ## Animation curve

## Relative movement
pag.move(offset_x, offset_y)              ## Move relative to current
pag.move(50, 0, duration=1)               ## Move right 50px over 1 sec

## Tweening functions (for smooth animations)
pag.easeInQuad    ## Start slow, end fast
pag.easeOutQuad   ## Start fast, end slow  
pag.easeInOutQuad ## Slow start/end, fast middle
pag.easeInBounce  ## Bouncing effect at end
pag.easeInElastic ## Elastic effect
```

## Click Operations
### Basic clicks
```
pag.click()                               ## Click at current position
pag.click(x=100, y=200)                   ## Click at coordinates
pag.click(button='left')                  ## Specify button
pag.click(button='right')                 ## Right click
pag.click(button='middle')                ## Middle click
```

### Multiple clicks
```
pag.doubleClick()                         ## Double left click
pag.tripleClick()                         ## Triple left click
pag.rightClick()                          ## Right click (convenience)
pag.middleClick()                         ## Middle click (convenience)
```

### Click with modifiers
```
pag.click(clicks=2, interval=0.25)        ## Double click with 0.25s between
pag.click(x, y, clicks=3, interval=0.1)   ## Triple click fast
```

### Press and release separately
```
pag.mouseDown()                           ## Press mouse button
pag.mouseUp()                             ## Release mouse button
pag.mouseDown(x=100, y=200, button='right')
pag.mouseUp(x=100, y=200, button='right')
```

### Drag & Drop
```
## Drag (click, move, release)
pag.dragTo(x, y, duration=1)              ## Drag to absolute position
pag.drag(offset_x, offset_y, duration=1)  ## Drag relative distance
```

### Drag with button specification
```
pag.dragTo(x, y, duration=1, button='left')
pag.drag(50, 0, duration=0.5, button='right')
```

### Press-drag-release separately
```
pag.mouseDown(x=100, y=100)
pag.moveTo(200, 200, duration=1)
pag.mouseUp()
```

## Scrolling
### Vertical scrolling
```
pag.scroll(10)           ## Scroll up 10 "clicks"
pag.scroll(-10)          ## Scroll down 10 "clicks"
pag.scroll(500)          ## Large scroll up
pag.scroll(-500)         ## Large scroll down
```

### Horizontal scrolling (platform dependent)
```
pag.hscroll(10)          ## Scroll right (if supported)
pag.hscroll(-10)         ## Scroll left (if supported)
pag.scroll(10, x=100, y=100)  ## Scroll at specific position
```

# ⌨️ KEYBOARD FUNCTIONS
## Typing Text
```
## Basic typing
pag.write('Hello World!')                 ## Type string
pag.write('Hello', interval=0.1)          ## Type with delay between keys
pag.write('Hello\n')                      ## \n = Enter key
pag.write(['h', 'e', 'l', 'l', 'o'])     ## List of characters
```

## Special characters
```
pag.write('Email: test@example.com\nPhone: (123) 456-7890\n')
pag.write('C:\\Users\\Name\\Documents')   ## Backslashes need escaping
```

## Key Presses
```
## Single key presses
pag.press('enter')       ## Press and release Enter
pag.press('tab')         ## Tab key
pag.press('f5')          ## Function key
pag.press('left')        ## Arrow key
pag.press('capslock')    ## Toggle caps lock
pag.press('numlock')     ## Toggle num lock

## Multiple keys
pag.press(['enter', 'tab', 'space'])  ## Press multiple keys
```

## Press and hold
```
pag.keyDown('shift')     ## Hold Shift key down
pag.keyUp('shift')       ## Release Shift key
```

## Hotkeys (Keyboard Shortcuts)
```
## Common shortcuts
pag.hotkey('ctrl', 'c')           ## Copy
pag.hotkey('ctrl', 'v')           ## Paste
pag.hotkey('ctrl', 'x')           ## Cut
pag.hotkey('ctrl', 'a')           ## Select all
pag.hotkey('ctrl', 's')           ## Save
pag.hotkey('ctrl', 'z')           ## Undo
pag.hotkey('ctrl', 'y')           ## Redo
pag.hotkey('ctrl', 'f')           ## Find
pag.hotkey('ctrl', 'n')           ## New
pag.hotkey('ctrl', 'o')           ## Open
pag.hotkey('ctrl', 'p')           ## Print
pag.hotkey('alt', 'tab')          ## Switch window
pag.hotkey('win', 'd')            ## Show desktop (Windows)
pag.hotkey('command', 'space')    ## Spotlight (Mac)

## Complex shortcuts
pag.hotkey('ctrl', 'shift', 'esc')       ## Task Manager
pag.hotkey('win', 'shift', 's')          ## Snipping Tool
pag.hotkey('alt', 'f4')                  ## Close window

## With pause between keys
pag.hotkey('ctrl', 'c', interval=0.1)    ## With delay
Special Keys Reference
text
## Alphabet and Numbers: 'a' to 'z', '0' to '9'
## Function Keys: 'f1' to 'f24'
## Arrow Keys: 'left', 'right', 'up', 'down'
## Modifier Keys: 'shift', 'ctrl', 'alt', 'win', 'command'
## Navigation: 'enter', 'esc', 'backspace', 'tab', 'space'
## Lock Keys: 'capslock', 'numlock', 'scrolllock'
## Multimedia: 'volumemute', 'volumeup', 'volumedown'
## Special: 'insert', 'delete', 'home', 'end', 'pageup', 'pagedown'
## Punctuation: '!', '@', '#', '$', etc. (use shift + number)

```

# 🖼️ SCREEN & IMAGE FUNCTIONS
## Screenshots
```
## Take screenshot
screenshot = pag.screenshot()                     ## Full screen
screenshot.save('screenshot.png')                 ## Save to file

## Screenshot of region
region = pag.screenshot(region=(x, y, width, height))
region.save('region.png')

## Screenshot with custom name
pag.screenshot('my_screenshot.png')               ## Direct save
pag.screenshot('region.png', region=(0,0,300,400))

## In-memory screenshot
screenshot = pag.screenshot()                     ## PIL Image object
pixels = screenshot.load()                        ## Pixel access
color = pixels[100, 200]                          ## Get pixel color

```

## Pixel Color Operations
```
## Get pixel color
color = pag.pixel(x, y)                           ## Returns RGB tuple
print(f"Color at ({x},{y}): {color}")

## Check pixel color
matches = pag.pixelMatchesColor(x, y, (255, 0, 0))  ## Exact match
matches = pag.pixelMatchesColor(x, y, (255, 0, 0), tolerance=10)  ## Within range

## Sample multiple pixels
colors = []
for i in range(10):
    colors.append(pag.pixel(x + i, y))

```

## Image Recognition (locateOnScreen)
```
## Basic image search
try:
    location = pag.locateOnScreen('image.png')     ## Find image
    print(f"Found at: {location}")                 ## Box(left, top, width, height)
except pag.ImageNotFoundException:
    print("Image not found")

## Search with confidence (0.0 to 1.0)
location = pag.locateOnScreen('button.png', confidence=0.9)

## Search in region (faster)
location = pag.locateOnScreen('icon.png', region=(0, 0, 500, 500))

## Grayscale search (faster, less accurate)
location = pag.locateOnScreen('image.png', grayscale=True)

## Get center of found image
if location:
    center = pag.center(location)                  ## Point(x, y)
    pag.click(center)

## Find all matching images
all_locations = list(pag.locateAllOnScreen('item.png'))
for loc in all_locations:
    pag.click(pag.center(loc))

## Find with minimum matches
locations = pag.locateAllOnScreen('dot.png', minSearchTime=2)
```

## Advanced Image Search
```
## Pre-load image for multiple searches
import pyscreeze
target_image = pyscreeze.load('target.png')

## Search with loaded image
location = pag.locate(target_image, pag.screenshot())

## Custom confidence per channel
location = pag.locateOnScreen('image.png', 
                             confidence=0.9,
                             rgb=True)  ## Compare RGB instead of grayscale
```

# ⚙️ UTILITY FUNCTIONS
## Timing & Delays
```
## Using pyautogui's built-in pause
pag.PAUSE = 1.0                    ## 1 second between ALL pyautogui actions

## Manual delays
import time
time.sleep(2.5)                    ## Delay everything
pag.sleep(2.5)                     ## Alias for time.sleep

## Between specific actions
pag.moveTo(100, 100)
time.sleep(0.5)                    ## Wait for UI
pag.click()
```

## Alert & Message Boxes
```
## Display messages
pag.alert('Operation completed!')           ## OK button
pag.alert('Warning!', 'Alert Title')        ## With custom title

## Confirmations
response = pag.confirm('Continue?')         ## Returns 'OK' or 'Cancel'
response = pag.confirm('Save changes?', buttons=['Save', 'Don\'t Save', 'Cancel'])

## Prompts for input
name = pag.prompt('Enter your name:')       ## Returns text or None
name = pag.prompt('Name:', 'Enter your name', 'Default Name')

## Password input
password = pag.password('Enter password:')  ## Shows asterisks
password = pag.password('Password:', 'Login', mask='*')
```

## Coordinate Helpers
```
## Get coordinates interactively
print("Move mouse to desired position...")
time.sleep(3)                                ## Time to move mouse
x, y = pag.position()
print(f"Position: ({x}, {y})")

## Real-time position display
print("Press Ctrl+C to stop position display")
try:
    while True:
        x, y = pag.position()
        position_str = f"X: {x:4} Y: {y:4}"
        print(position_str, end='')
        print('\b' * len(position_str), end='', flush=True)
        time.sleep(0.1)
except KeyboardInterrupt:
    print('\nDone.')
```

# 🔄 COMBINED OPERATIONS
## Common Patterns
```
## Click and type
pag.click(100, 200)
pag.write('Hello World', interval=0.05)

## Select all and delete
pag.hotkey('ctrl', 'a')
pag.press('delete')

## Copy and paste
pag.hotkey('ctrl', 'c')
pag.click(destination_x, destination_y)
pag.hotkey('ctrl', 'v')

## Drag selection
pag.moveTo(start_x, start_y)
pag.mouseDown()
pag.moveTo(end_x, end_y, duration=1)
pag.mouseUp()
```
## Error Handling
```
try:
    location = pag.locateOnScreen('button.png', confidence=0.9, minSearchTime=2)
    if location:
        pag.click(pag.center(location))
    else:
        print("Button not found, trying alternative...")
        pag.click(backup_x, backup_y)
except pag.ImageNotFoundException:
    print("Image file not found")
except Exception as e:
    print(f"Unexpected error: {e}")
    ## Fallback action
    pag.press('esc')
```

# 📊 PERFORMANCE TIPS
```
## 1. Use regions for faster image search
pag.locateOnScreen('icon.png', region=(0, 0, 500, 500))

## 2. Lower confidence for faster matching (if possible)
pag.locateOnScreen('button.png', confidence=0.7)

## 3. Use grayscale for faster image recognition
pag.locateOnScreen('image.png', grayscale=True)

## 4. Cache images if searching multiple times
target_image = pyscreeze.load('target.png')

## 5. Reduce screenshot size
pag.screenshot(region=(x, y, width, height))
```

# 🎮 GAME & APPLICATION AUTOMATION
```
## Rapid clicking
for _ in range(10):
    pag.click()
    time.sleep(0.1)

## Pattern movement (for games)
points = [(100, 100), (200, 100), (200, 200), (100, 200)]
for point in points:
    pag.moveTo(point, duration=0.25)
    pag.click()

## Hold key for duration
pag.keyDown('w')
time.sleep(2)  ## Move forward for 2 seconds
pag.keyUp('w')

## Combination for gaming
pag.keyDown('ctrl')
pag.click()    ## Ctrl + Click
pag.keyUp('ctrl')
```

# 📁 FILE OPERATIONS WITH AUTOMATION
```
## Save file dialog
pag.hotkey('ctrl', 's')
time.sleep(1)                    ## Wait for dialog
pag.write('document.txt')
pag.press('enter')

## Open file dialog  
pag.hotkey('ctrl', 'o')
time.sleep(1)
pag.write('C:\\Users\\Name\\file.txt')
pag.press('enter')
```

# 🎯 PRO TIPS & TRICKS
```
## 1. Record mouse position with hotkey
import keyboard  ## pip install keyboard

def record_position():
    x, y = pag.position()
    print(f"Position recorded: ({x}, {y})")
    return (x, y)

## 2. Create reusable click function with retry
def click_with_retry(image_path, max_attempts=3, confidence=0.9):
    for attempt in range(max_attempts):
        try:
            location = pag.locateOnScreen(image_path, confidence=confidence)
            if location:
                pag.click(pag.center(location))
                return True
        except:
            time.sleep(1)
    return False

## 3. Smooth mouse movements for human-like behavior
def human_like_move(x, y):
    import random
    duration = random.uniform(0.5, 1.5)
    pag.moveTo(x, y, duration=duration)

## 4. Type with human-like variability
def human_type(text):
    import random
    for char in text:
        pag.write(char, interval=random.uniform(0.05, 0.2))
```

# ⚡ QUICK REFERENCE - Most Used
```
## TOP 10 MOST USED FUNCTIONS
pag.click(x, y)                 ## 1. Click at position
pag.write('text')               ## 2. Type text
pag.hotkey('ctrl', 'c')         ## 3. Keyboard shortcut
pag.moveTo(x, y)                ## 4. Move mouse
pag.press('enter')              ## 5. Press key
pag.screenshot('file.png')      ## 6. Take screenshot
pag.locateOnScreen('img.png')   ## 7. Find image
pag.dragTo(x, y)                ## 8. Drag mouse
pag.scroll(10)                  ## 9. Scroll
pag.alert('Message')            ## 10. Show alert
```

# 🚨 TROUBLESHOOTING
```
## Debug mode
import logging
logging.basicConfig(level=logging.DEBUG)

## Check if point is on screen
def is_on_screen(x, y):
    width, height = pag.size()
    return 0 <= x <= width and 0 <= y <= height

## Safe click with bounds checking
def safe_click(x, y):
    if is_on_screen(x, y):
        pag.click(x, y)
    else:
        print(f"Point ({x}, {y}) is off-screen")

## Monitor execution time
import datetime
start_time = datetime.datetime.now()
## ... automation code ...
print(f"Execution time: {datetime.datetime.now() - start_time}")
```


# 📋 CHEAT SHEET PRINTABLE VERSION
```
MOUSE:
  moveTo(x,y)       - Move to coordinates
  click()           - Click at position
  dragTo(x,y)       - Drag to position
  scroll(n)         - Scroll n clicks
  position()        - Get current position

KEYBOARD:
  write(text)       - Type text
  press(key)        - Press key
  hotkey(k1,k2)     - Key combination
  keyDown(key)      - Hold key
  keyUp(key)        - Release key

SCREEN:
  screenshot()      - Capture screen
  locateOnScreen()  - Find image
  pixel(x,y)        - Get color
  size()            - Screen resolution

UTILITY:
  PAUSE = n         - Delay between actions
  FAILSAFE = True   - Safety feature
  alert(text)       - Show message box
```