# AI Slide Generator

An **automatic slide deck generator** that transforms a single topic into a full PowerPoint presentation—complete with subtopics, detailed descriptions, and AI‑generated images. Built with Google Gemini (GenAI) and [python‑pptx](https://python-pptx.readthedocs.io/).

---

## 🚀 Features

- **One‑step subtopic + description** generation via a single prompt  
- **Hyper‑realistic 3D images** illustrating each subtopic  
- **Automated PPTX creation** with custom layouts, text wrapping & styling  
- **In‑memory image handling** (no temporary files)  
- **Modular, easy‑to‑customize code**

---

## 📋 Prerequisites

- **Python 3.8+**  
- A valid **Google Gemini API key**

---

## 🔧 Installation

1. Clone this repo
   ```bash
   git clone https://github.com/yourusername/ai-slide-generator.git
   cd ai-slide-generator
   ```

2. Install dependencies
   ```bash
   pip install python-pptx pillow google-genai
   ```

3. Set your API key as an environment variable
   ```bash
   export GOOGLE_GEMINI_API_KEY="your_api_key_here"
   # on Windows (PowerShell):
   setx GOOGLE_GEMINI_API_KEY "your_api_key_here"
   ```

---

## 🎬 Usage

Run the main script and enter your topic when prompted:

```bash
python slide_generator.py
```

```
Enter the main topic for the presentation: Pollution in New Delhi
```

- The script will generate **6 subtopics**, fetch descriptions and images, then build and open `Pollution in New Delhi.pptx`.

---

## 🔍 How It Works

1. **`get_subtopics(topic)`**  
   Queries Gemini to return exactly six numbered subtopics.  

2. **`generate_description(subtopic, topic)`**  
   Fetches a concise, 1100‑character description for each subtopic.  

3. **`generate_image(subtopic, topic)`**  
   Creates a hyper‑realistic 3D image reflecting the concept—retries up to 5 times if the model is busy.  

4. **`build_presentation(slides_info, topic)`**  
   Uses `python-pptx` to assemble slides with:  
   - Bold, centered titles  
   - Justified, wrapped descriptions  
   - Centered, resized images  
   - Light‑blue backgrounds for a clean look

---

## ✏️ Customization

- **Change number of subtopics** by editing the prompt or slice count (`[:6]`).  
- **Adjust slide layout** by tweaking `Inches(...)`, font sizes, or background color (`RGBColor`).  
- **Swap AI model** by modifying the `model=` argument in the `generate_content` calls.

---
