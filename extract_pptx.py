import zipfile
import xml.etree.ElementTree as ET

def extract_text_from_pptx(pptx_path, txt_path):
    namespaces = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    text_content = []
    
    with zipfile.ZipFile(pptx_path, 'r') as z:
        slide_files = [f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
        # Sort by slide number
        slide_files.sort(key=lambda x: int(x.replace('ppt/slides/slide', '').replace('.xml', '')))
        
        for slide_file in slide_files:
            slide_xml = z.read(slide_file)
            root = ET.fromstring(slide_xml)
            slide_texts = []
            for node in root.iterfind('.//a:t', namespaces):
                if node.text:
                    slide_texts.append(node.text)
            if slide_texts:
                text_content.append(" ".join(slide_texts))
                
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write('\n\n--- Slide ---\n'.join(text_content))

files = [
    ("c:/Users/Chu/Downloads/正規語言/9781284077247_PPTx_CH05.pptx", "c:/Users/Chu/Downloads/正規語言/ch5.txt"),
    ("c:/Users/Chu/Downloads/正規語言/9781284077247_PPTx_CH06 (1).pptx", "c:/Users/Chu/Downloads/正規語言/ch6.txt"),
    ("c:/Users/Chu/Downloads/正規語言/9781284077247_PPTx_CH07 (1).pptx", "c:/Users/Chu/Downloads/正規語言/ch7.txt")
]

for pptx, txt in files:
    try:
        extract_text_from_pptx(pptx, txt)
        print(f"Extracted {pptx} to {txt}")
    except Exception as e:
        print(f"Failed to extract {pptx}: {e}")
