from PIL import Image, ImageDraw, ImageFont

# Create a logo similar to Prathibha Institute
width, height = 500, 500
img = Image.new('RGB', (width, height), 'white')
draw = ImageDraw.Draw(img)

# Draw circle border
circle_margin = 30
draw.ellipse([circle_margin, circle_margin, width-circle_margin, height-circle_margin], 
             outline='#1F4E78', width=8)

# Draw large P
try:
    font_large = ImageFont.truetype("arial.ttf", 120)
    font_med = ImageFont.truetype("arialbd.ttf", 50)
    font_small = ImageFont.truetype("arialbd.ttf", 35)
except:
    font_large = ImageFont.load_default()
    font_med = ImageFont.load_default()
    font_small = ImageFont.load_default()

# Draw P letter
draw.text((190, 100), "P", fill='#1F4E78', font=font_large)

# Draw institute name
draw.text((80, 320), "PRATHIBHA", fill='#1F4E78', font=font_med)
draw.text((100, 380), "INSTITUTE", fill='#1F4E78', font=font_small)

# Save
img.save('logo.png', 'PNG')
print("Logo created successfully!")
