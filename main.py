import comtypes.client
from transformers import pipeline
import os

os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "Coffee Shop Business Pitch Deck.pptx")
def write_in_file(file_path, data_to_append):
    with open(file_path, 'a') as file:
        file.write(data_to_append + "\n")

def extract_slide_content(ppt_file):
    if os.path.exists("output.txt"):
        os.remove("output.txt")

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(ppt_file)

    slides_data = {}

    for slide_index, slide in enumerate(presentation.Slides, start=1):
        slide_data = {"headings": [], "paragraphs": []}

        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_frame = shape.TextFrame
                for paragraph in text_frame.TextRange.Paragraphs():
                    text = paragraph.Text.strip()
                    write_in_file("output.txt", text) 

        slides_data[f"{slide_index}"] = slide_data

    presentation.Close()
    powerpoint.Quit()

    return slides_data

ppt_file_path = file_path
slides_content = extract_slide_content(ppt_file_path)

with open('output.txt', 'r') as file:
    text_data = file.read()

print(f"Length of text_data: {len(text_data)} characters")

summarizer = pipeline("summarization", model="facebook/bart-large-cnn")

max_input_length = 1024  # Set the maximum input length for the model
chunks = [text_data[i:i + max_input_length] for i in range(0, len(text_data), max_input_length)]

summaries = []
for chunk in chunks:
    summary = summarizer(
        chunk,
        max_length=150,
        min_length=30,
        do_sample=False,
        truncation=True
    )
    summaries.append(summary[0]['summary_text'])

final_summary = " ".join(summaries)

print("Final Summary:")
print(final_summary)

write_in_file('summery.txt', final_summary)
