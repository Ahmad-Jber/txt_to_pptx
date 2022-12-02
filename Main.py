from Presentation_Management import create_powerpoint, delete_slide, self

text_file = open("PPT.txt", "r")
data = text_file.read()
text_file.close()

# You can name PCT.pptx as you like, but must end with .pptx

create_powerpoint(data, "PCT.pptx")
for slide in self.slides:
    if slide.placeholders[1].text == "":
        delete_slide(self, slide, "PCT.pptx")
self.save("PCT.pptx")
