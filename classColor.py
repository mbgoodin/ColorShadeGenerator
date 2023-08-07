from main import lighten
from main import darken
from main import rgb_to_hex

class Color:
    def __init__(self):
        pass

    def convert(self, type, shade, hex):
        if type == "darken":
            return rgb_to_hex(darken(hex_code=hex,shade=shade))
        if type == "lighten":
            return rgb_to_hex(lighten(hex_code=hex,shade=shade))
        else:
            print("Please enter valide arguments")



color = Color()
print(color.convert(type="darken",shade=9, hex="00376d"))