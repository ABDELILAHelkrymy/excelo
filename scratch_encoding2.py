text = "ε΋΍έόϟ΍ϙϭηϟ΍ϥϳϋ"
bytes_arr = bytes([ord(c) & 0xFF for c in text])

decoded = bytes_arr.decode('cp1256', errors='replace')
print(ascii(decoded))

decoded_utf8 = bytes_arr.decode('utf-8', errors='replace')
print(ascii(decoded_utf8))

