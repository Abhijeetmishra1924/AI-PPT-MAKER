
!pip install --quiet groq python-pptx pillow

from groq import Groq
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageDraw, ImageFilter
from google.colab import output, files
from IPython.display import HTML, display
import os, re, time, random, shutil, zipfile

# -------------------------

# -------------------------
GROQ_API_KEY = ""
client = Groq(api_key=GROQ_API_KEY)

# -------------------------
# Setup folders
# -------------------------
os.makedirs("/mnt/data", exist_ok=True)
os.makedirs("/content/bg_cache", exist_ok=True)
os.makedirs("/content/output_multi", exist_ok=True)

# -------------------------
# Style palettes (4 styles)
# Each style: top, bottom, title_rgb, text_rgb, accent_rgb, bg_mode
# bg_mode: "gradient" or "white" or "dark"
# -------------------------
STYLES = {
    "corporate_clean": {
        "name":"Corporate Clean",
        "colors":((240,245,250),(255,255,255),(20,30,40),(40,50,60),(80,130,200)),
        "bg_mode":"white"
    },
    "modern_gradient": {
        "name":"Modern Gradient",
        "colors":((18,24,56),(80,150,240),(255,255,255),(230,240,255),(255,200,220)),
        "bg_mode":"gradient"
    },
    "dark_executive": {
        "name":"Dark Executive",
        "colors":((8,10,16),(28,40,56),(240,245,250),(200,220,240),(220,180,80)),
        "bg_mode":"dark"
    },
    "minimal_whitepaper": {
        "name":"Minimal Whitepaper",
        "colors":((255,255,255),(245,245,245),(10,10,10),(40,40,40),(80,80,120)),
        "bg_mode":"white"
    }
}

# -------------------------
# Background generator (gradient or subtle)
# -------------------------
def make_background(path, top, bottom, width=1920, height=1080, seed=None, mode="gradient"):
    if mode == "white":
        # simple light texture
        img = Image.new("RGB", (width, height), top)
        draw = ImageDraw.Draw(img)
        # faint diagonal lines for texture
        for i in range(0, width, 80):
            draw.line([(i,0),(i - width, height)], fill=(top[0]-3 if top[0]>10 else top[0],)*3, width=1)
        img = img.filter(ImageFilter.GaussianBlur(radius=0.4))
        img.save(path)
        return path
    elif mode == "dark":
        # subtle dark gradient
        img = Image.new("RGB", (width, height), top)
        draw = ImageDraw.Draw(img)
        for y in range(height):
            t = y/(height-1)
            r = int(top[0] + (bottom[0]-top[0])*t)
            g = int(top[1] + (bottom[1]-top[1])*t)
            b = int(top[2] + (bottom[2]-top[2])*t)
            draw.line([(0,y),(width,y)], fill=(r,g,b))
        # subtle overlay
        overlay = Image.new("RGBA",(width,height),(255,255,255,0))
        od = ImageDraw.Draw(overlay)
        random.seed(seed or 0)
        for i in range(5):
            alpha = int(18 + i*6)
            color = (255,255,255,alpha)
            offset = int((i-2)*width*0.15)
            pts = [(-width+offset, int(height*0.1)), (int(width*0.4)+offset,int(height*0.05)), (int(width*0.9)+offset,int(height*0.9)), (int(width*0.3)+offset,int(height*0.95))]
            od.polygon(pts, fill=color)
        combined = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
        combined = combined.filter(ImageFilter.GaussianBlur(radius=1.1))
        combined.save(path)
        return path
    else:
        # default gradient
        img = Image.new("RGB", (width, height), top)
        draw = ImageDraw.Draw(img)
        for y in range(height):
            t = y/(height-1)
            r = int(top[0] + (bottom[0]-top[0])*t)
            g = int(top[1] + (bottom[1]-top[1])*t)
            b = int(top[2] + (bottom[2]-top[2])*t)
            draw.line([(0,y),(width,y)], fill=(r,g,b))
        img.save(path)
        return path

# -------------------------
# Prompt for deep corporate mixed-length slides (10-12 slides)
# -------------------------
def generate_deep_deck(topic, min_slides=10, max_slides=12):
    # We'll instruct model to produce between min and max slides and varied lengths
    prompt = f"""
You are a professional corporate research writer. Produce a detailed presentation about: {topic}

Create between {min_slides} and {max_slides} slides. Produce rich, mixed-length content: some slides should have short paragraphs (3–4 lines), others long detailed paragraphs (7–10 lines), and some slides can combine paragraph + bullets.

Use this exact structure outline (but vary phrasing and content naturally):

1) Title Slide
- Title line
- One-line subtitle (8–14 words)
- One short tagline (optional)

2) Executive Summary
- 2 short paragraphs (2 lines each), describing the deck overview.

3) Background & Context
- 3–6 lines (historical and context)

4) Key Applications
- 1 detailed paragraphs (each 4–8 lines) describing real-world applications and examples.

5) Benefits & Impact
- 2 paragraphs (short + medium) and then 3 short bullets (use "- text")

6) Market Landscape
- 3–5 lines, mention leading companies, adoption trends, and a brief metric

7) Technical Approach / How it Works
- 3–6 lines describing architecture or approach

8) Case Study (1 example)
- 3–6 lines of detailed case study

9) Challenges & Ethics
- 3–6 lines with regulatory and risk aspects

10) Future Outlook & Roadmap
- 3–6 lines with predicted trends

11) Conclusion & Recommendations
- 2–3 lines concluding and 2 short bullets as calls-to-action

RULES:
- Separate slides with a blank line.
- For each slide, first line must be the slide title.
- When bullets are requested, format as "- text" each on its own line.
- Use natural corporate tone, include one or two example company names where relevant.
- Do can include charts if need; text-only content.
- Return ONLY the slides in plain text with blank line separators.
"""
    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=2000
    )
    # Access content
    return resp.choices[0].message.content

# -------------------------
# Helpers: parse raw blocks into slides
# -------------------------
def split_blocks(raw):
    blocks = [b.strip() for b in re.split(r"\n\s*\n", raw) if b.strip()]
    return blocks

def parse_block(block):
    lines = [l.rstrip() for l in block.split("\n") if l.strip()]
    title = lines[0]
    body = lines[1:]
    return title, body

# -------------------------
# PPT builder for deep slides (handles long paragraphs, bullets, mixed layouts)
# -------------------------
def build_deep_presentation(raw, style_key):
    if style_key not in STYLES:
        raise ValueError("Unknown style.")
    style = STYLES[style_key]
    top, bottom, title_rgb, text_rgb, accent_rgb = style["colors"][0], style["colors"][1], style["colors"][2], style["colors"][3], style["colors"][4]
    bg_mode = style["bg_mode"]
    bg_path = f"/content/bg_cache/{style_key}_deep.jpg"
    if not os.path.exists(bg_path):
        make_background(bg_path, top, bottom, seed=sum(top)+len(style_key), mode=bg_mode)

    prs = Presentation()
    # choose slide count based on model output
    blocks = split_blocks(raw)
    # ensure between 10 and 12 slides
    if len(blocks) < 10:
        # pad with summary slides
        while len(blocks) < 10:
            blocks.append("Additional\nThis slide was auto-added to reach the minimum slide count.")
    blocks = blocks[:12]  # cap at 12

    for idx, block in enumerate(blocks[:12]):
        title, body_lines = parse_block(block)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # background
        slide.shapes.add_picture(bg_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # accent bar
        try:
            bar = slide.shapes.add_textbox(0, 0, int(prs.slide_width * 0.18), int(prs.slide_height * 0.06))
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(*accent_rgb)
            bar.line.fill.background()
        except Exception:
            pass

        # Title placement
        if idx == 0:
            # Title slide: big title + short subtitle + tagline
            tbox = slide.shapes.add_textbox(Inches(1), Inches(0.9), prs.slide_width - Inches(2), Inches(1.6))
            tf = tbox.text_frame
            tf.text = title
            p = tf.paragraphs[0]
            p.font.size = Pt(56)
            p.font.bold = True
            # choose color based on style (dark text for light background)
            p.font.color.rgb = RGBColor(*title_rgb)

            # subtitle: first body line truncated
            subtitle = ""
            tagline = ""
            if body_lines:
                # combine all body lines then split into sentences
                joined = " ".join(body_lines)
                sents = re.split(r'(?<=[.!?]) +', joined)
                subtitle = " ".join(sents[:1]) if sents else ""
                tagline = " ".join(sents[1:2]) if len(sents) > 1 else ""
            sub_box = slide.shapes.add_textbox(Inches(1), Inches(2.6), prs.slide_width - Inches(2), Inches(1.1))
            st = sub_box.text_frame
            st.word_wrap = True
            st.text = subtitle
            st.paragraphs[0].font.size = Pt(20)
            st.paragraphs[0].font.color.rgb = RGBColor(*text_rgb)
            if tagline:
                tag_box = slide.shapes.add_textbox(Inches(1), Inches(3.6), prs.slide_width - Inches(2), Inches(0.8))
                tt = tag_box.text_frame
                tt.word_wrap = True
                tt.text = tagline
                tt.paragraphs[0].font.size = Pt(14)
                tt.paragraphs[0].font.color.rgb = RGBColor(*text_rgb)
        else:
            # For other slides, inspect body_lines: if bullets present, render bullets; else paragraph(s)
            bullets = [ln[2:].strip() for ln in body_lines if ln.startswith("- ")]
            paragraphs = [ln for ln in body_lines if not ln.startswith("- ")]
            # Title
            tbox = slide.shapes.add_textbox(Inches(0.9), Inches(0.8), prs.slide_width - Inches(2), Inches(1))
            ttf = tbox.text_frame
            ttf.text = title
            ttf.paragraphs[0].font.size = Pt(28)
            ttf.paragraphs[0].font.bold = True
            ttf.paragraphs[0].font.color.rgb = RGBColor(*title_rgb)

            if bullets:
                # Render bullets in a single textbox with wrapping
                bbox = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), prs.slide_width - Inches(2), Inches(5))
                btf = bbox.text_frame
                btf.text = ""
                for b in bullets[:6]:
                    p = btf.add_paragraph()
                    p.text = "• " + b
                    p.level = 0
                    p.font.size = Pt(20)
                    p.font.color.rgb = RGBColor(*text_rgb)
            else:
                # Mixed-length paragraphs: join but split into 1-3 paragraph blocks depending on desired density
                joined = " ".join(paragraphs).strip()
                if not joined:
                    joined = "Key point: further details on this topic."
                # heuristics: for earlier slides make longer paragraphs, for later slides shorter
                if idx in (1,3,4):   # executive summary, key apps -> longer
                    # split into up to 3 paragraphs
                    sents = re.split(r'(?<=[.!?]) +', joined)
                    paras = []
                    cur = []
                    for s in sents:
                        cur.append(s)
                        if len(cur) >= 3:
                            paras.append(" ".join(cur))
                            cur = []
                    if cur:
                        paras.append(" ".join(cur))
                    paras = paras[:3]
                elif idx in (2,5,6):  # background, market, tech -> medium
                    sents = re.split(r'(?<=[.!?]) +', joined)
                    paras = [" ".join(sents[:max(1,int(len(sents)/3))])]
                    # ensure 1-2 paras
                else:
                    # later slides shorter: 1 short paragraph
                    sents = re.split(r'(?<=[.!?]) +', joined)
                    paras = [" ".join(sents[:3])]
                # render paragraphs
                top_y = Inches(1.9)
                for i, para in enumerate(paras):
                    pbox = slide.shapes.add_textbox(Inches(0.9), top_y + Inches(i*1.2), prs.slide_width - Inches(2), Inches(1.6))
                    ptf = pbox.text_frame
                    ptf.word_wrap = True
                    ptf.text = para
                    ptf.paragraphs[0].font.size = Pt(20)
                    ptf.paragraphs[0].font.color.rgb = RGBColor(*text_rgb)

    # filename depending on style
    fname = f"/mnt/data/Deep_PPT_{style_key}.pptx"
    prs.save(fname)
    return fname

# -------------------------
# Orchestrator: generate for one style or all styles and zip
# -------------------------
def generate_and_package(topic, style_choice):
    # style_choice can be one of keys or "all"
    results = []
    if style_choice == "all":
        # generate for every style
        for k in STYLES.keys():
            raw = generate_deep_deck(topic, min_slides=10, max_slides=12)
            path = build_deep_presentation(raw, k)
            results.append(path)
        # create zip
        zip_path = "/mnt/data/Deep_PPT_AllStyles.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            for p in results:
                zf.write(p, os.path.basename(p))
        return zip_path
    else:
        raw = generate_deep_deck(topic, min_slides=10, max_slides=12)
        path = build_deep_presentation(raw, style_choice)
        return path

# -------------------------
# Backend called from JS UI
# -------------------------
def backend_final(topic, style_key):
    if not topic:
        return {"status":"error","msg":"Please enter a topic."}
    try:
        out = generate_and_package(topic, style_key)
        # trigger download
        time.sleep(0.2)
        files.download(out)
        return {"status":"ok","path": out}
    except Exception as e:
        return {"status":"error","msg": str(e)}

output.register_callback("backend_final", backend_final)

# -------------------------
# HTML UI (Modern Dark) with Style selector (includes "All")
# -------------------------
html = """
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap" rel="stylesheet">
<style>
body {background:#07101a;color:#eaf3ff;font-family:Inter,Arial;margin:12px;}
.card{max-width:980px;margin:14px auto;padding:18px;border-radius:12px;background:linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));box-shadow:0 10px 40px rgba(0,0,0,0.6);}
h1{margin:0;font-size:20px;color:#fff}
label{display:block;margin-top:10px;color:#9fb3d8;font-size:13px}
input,select{width:100%;padding:10px;border-radius:8px;background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.04);color:#eaf3ff}
button{margin-top:12px;padding:11px 16px;border-radius:10px;border:none;background:linear-gradient(90deg,#2d6cdf,#8456ff);color:#fff;font-weight:700;cursor:pointer}
.status{margin-top:10px;padding:10px;border-radius:8px;background:rgba(255,255,255,0.02);color:#bcd5ff}
.small{font-size:12px;color:#9fb3d8;margin-top:8px}
</style>

<div class="card">
  <h1>Deep Corporate PPT Generator — Mixed-Length (10–12 slides)</h1>
  <p class="small">Generates deep research-style decks (text-only). Choose a design style or select "All" to generate one deck per style and download a ZIP.</p>

  <label>Topic</label>
  <input id="topic" placeholder="e.g. AI in Healthcare: Opportunities, Challenges & Roadmap" />

  <label>Design Style</label>
  <select id="style">
    <option value="corporate_clean">Corporate Clean</option>
    <option value="modern_gradient">Modern Gradient</option>
    <option value="dark_executive">Dark Executive</option>
    <option value="minimal_whitepaper">Minimal Whitepaper</option>
    <option value="all">All (generate for all styles & zip)</option>
  </select>

  <div style="display:flex;gap:12px;align-items:center">
    <button id="gen">Generate Deck</button>
    <div id="status" class="status">Status: Idle</div>
  </div>
  <div class="small">Note: Generation may take 20–90s depending on Colab resources and Groq latency. Large decks may be sizeable.</div>
</div>

<script>
document.getElementById('gen').onclick = async function() {
  const topic = document.getElementById('topic').value.trim();
  const style = document.getElementById('style').value;
  const status = document.getElementById('status');
  if(!topic){ status.innerText = 'Status: Enter a topic'; return; }
  status.innerText = 'Status: Generating (this may take a while)...';
  try {
    const res = await google.colab.kernel.invokeFunction('backend_final', [topic, style], {});
    if(res && res.status === 'ok') {
      status.innerText = 'Status: Download started. Check your browser downloads.';
    } else {
      status.innerText = 'Status: Error: ' + (res.msg || 'unknown');
    }
  } catch (err) {
    status.innerText = 'Status: Execution error';
  }
}
</script>
"""

display(HTML(html))
