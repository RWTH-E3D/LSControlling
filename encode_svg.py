import base64

with open("fakultaet3.svg", "rb") as svg_file:
    encoded_string = base64.b64encode(svg_file.read()).decode('utf-8')
    with open("lscontrolling_logo.py", "w") as py_file:
        py_file.write(f"lscontrolling_logo = b'{encoded_string}'\n")

print(f"Logo written to lscontrolling_logo.py")
