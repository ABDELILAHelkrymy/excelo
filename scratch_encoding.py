text = "ε΋΍έόϟ΍ϙϭηϟ΍ϥϳϋ"
try:
    print(text.encode('cp1253').decode('cp1256'))
except Exception as e:
    print("cp1253->cp1256 failed:", e)

try:
    # Maybe utf-8 bytes interpreted as something else?
    pass
except Exception as e:
    pass

# Try to just see the codepoints
print([hex(ord(c)) for c in text])
