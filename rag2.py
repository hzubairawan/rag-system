"""
DocMind - chat with your documents (single-file version).

Upload PDF, Word, PowerPoint, Excel, CSV, HTML, JSON and text files, then ask questions.
Answers come only from your documents and always cite their sources.

Run:    streamlit run docmind_app.py
Install: pip install streamlit pandas langchain-core langchain-community langchain-text-splitters \
             langchain-openai langchain-google-genai faiss-cpu pypdf python-docx python-pptx \
             openpyxl beautifulsoup4
"""

from __future__ import annotations

import os

import streamlit as st
import streamlit.components.v1 as components

APP_NAME = "DocMind"
APP_VERSION = "v1.0"
SAMPLE_QUESTIONS = [
    "Give me a summary of the key points",
    "What are the main topics covered?",
    "List important dates, numbers and names",
]

st.set_page_config(
    page_title=f"{APP_NAME} · Chat with your documents",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)


def ui_html(markup: str) -> None:
    """Render raw HTML. Newlines/indentation are stripped so Markdown never turns it into a code block."""
    st.markdown(" ".join(line.strip() for line in markup.splitlines()), unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Design system (works in both light and dark themes)
# ----------------------------------------------------------------------------
ui_html(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root{--glass:rgba(128,128,128,.10);--glass-border:rgba(128,128,128,.28);}
footer, [data-testid="stAppDeployButton"], .stDeployButton {display:none !important;}
header[data-testid="stHeader"] {background: transparent;}
[data-testid="stAppViewContainer"], [data-testid="stMain"], .main {background: transparent !important;}
.stApp {
  background-image:
    radial-gradient(900px 520px at 8% -8%, rgba(99,102,241,.30), transparent 62%),
    radial-gradient(800px 520px at 95% 2%, rgba(236,72,153,.22), transparent 62%),
    radial-gradient(900px 620px at 50% 112%, rgba(34,211,238,.18), transparent 62%);
  background-size: 135% 135%;
  animation: aurora 24s ease-in-out infinite alternate;
}
@keyframes aurora {0% {background-position: 0% 0%;} 100% {background-position: 100% 100%;}}
.block-container {padding-top: 2.2rem; max-width: 1180px;}

/* glass + 3D tilt */
.glass {
  background: var(--glass); border: 1px solid var(--glass-border); border-radius: 18px;
  backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
  box-shadow: 0 1px 0 rgba(255,255,255,.10) inset, 0 18px 40px -22px rgba(0,0,0,.55);
}
.tilt {transition: transform .35s cubic-bezier(.2,.8,.2,1), box-shadow .35s; transform-style: preserve-3d; will-change: transform;}
.tilt:hover {
  transform: perspective(900px) rotateX(5deg) rotateY(-7deg) translateY(-8px) scale(1.02);
  box-shadow: 0 34px 60px -24px rgba(139,92,246,.65);
}

/* hero */
.badge {
  display: inline-flex; align-items: center; gap: .45rem; padding: .28rem .8rem; border-radius: 999px;
  font-size: .78rem; font-weight: 600; border: 1px solid var(--glass-border); background: var(--glass);
}
.badge .dot {width: 8px; height: 8px; border-radius: 50%; background: #22C55E; box-shadow: 0 0 10px #22C55E;}
.hero-title {
  font-size: 3.3rem; line-height: 1.05; font-weight: 800; letter-spacing: -.02em; margin: .9rem 0 .8rem;
  background: linear-gradient(90deg, #818CF8, #C084FC, #F472B6, #22D3EE, #818CF8); background-size: 200% auto;
  -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
  animation: shine 7s linear infinite;
}
@keyframes shine {to {background-position: 200% center;}}
.hero-sub {font-size: 1.1rem; opacity: .78; max-width: 34rem; line-height: 1.6;}
.pills {display: flex; flex-wrap: wrap; gap: .5rem; margin-top: 1.2rem;}
.pill {padding: .35rem .85rem; border-radius: 999px; font-size: .82rem; font-weight: 600;
  border: 1px solid rgba(168,85,247,.45); background: rgba(168,85,247,.10);}

/* sections, cards, steps */
.section-title {font-size: 1.35rem; font-weight: 800; margin: 2rem 0 .8rem; letter-spacing: -.01em;}
.grid3 {display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 1rem; margin: .4rem 0 1rem;}
.card {padding: 1.25rem;}
.card .ico {
  width: 48px; height: 48px; border-radius: 14px; display: grid; place-items: center; font-size: 1.45rem; margin-bottom: .8rem;
  background: linear-gradient(135deg, #6366F1, #A855F7, #EC4899);
  box-shadow: 0 12px 26px -8px rgba(168,85,247,.75), inset 0 2px 0 rgba(255,255,255,.35); transform: translateZ(30px);
}
.card h4 {margin: .1rem 0 .35rem; font-size: 1.04rem;}
.card p {margin: 0; opacity: .72; font-size: .9rem; line-height: 1.5;}
.steps {display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem;}
.step {padding: 1.2rem 1.3rem;}
.step .num {font-size: .72rem; font-weight: 800; letter-spacing: .14em; color: #A855F7;}
.step h4 {margin: .35rem 0; font-size: 1.05rem;}
.step p {margin: 0; opacity: .72; font-size: .9rem; line-height: 1.5;}
.fmt-row {display: flex; flex-wrap: wrap; gap: .6rem; margin: .4rem 0;}
.fmt {
  display: inline-flex; align-items: center; gap: .4rem; padding: .45rem .9rem; border-radius: 12px; font-weight: 700; font-size: .84rem;
  border: 1px solid var(--glass-border); background: var(--glass); transition: transform .25s, box-shadow .25s;
}
.fmt:hover {transform: translateY(-5px) rotateX(10deg); box-shadow: 0 16px 28px -16px rgba(139,92,246,.8);}
.notice {padding: 1rem 1.2rem; margin: 1rem 0; border-left: 4px solid #A855F7;}

/* brand + stats */
.brand {display: flex; align-items: center; gap: .75rem; margin-bottom: .3rem;}
.logo {
  width: 44px; height: 44px; border-radius: 13px; display: grid; place-items: center; font-size: 1.35rem;
  background: linear-gradient(135deg, #6366F1, #A855F7, #EC4899);
  box-shadow: 0 12px 26px -8px rgba(168,85,247,.85), inset 0 2px 0 rgba(255,255,255,.4);
  transform: perspective(300px) rotateY(-14deg) rotateX(8deg);
}
.brand-name {font-weight: 800; font-size: 1.3rem; line-height: 1.1;}
.brand-sub {font-size: .72rem; opacity: .6;}
.compact {display: flex; align-items: center; gap: .9rem; padding: .7rem 1.1rem; margin-bottom: .8rem;}
.compact .logo {width: 38px; height: 38px; font-size: 1.1rem;}
.compact .t {
  font-weight: 800; font-size: 1.4rem; margin-right: auto;
  background: linear-gradient(90deg, #818CF8, #C084FC, #F472B6); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
}
.stats {display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: .9rem; margin: .4rem 0 1rem;}
.stat {padding: .9rem 1rem; display: flex; align-items: center; gap: .85rem;}
.stat .si {
  font-size: 1.45rem; width: 46px; height: 46px; border-radius: 13px; display: grid; place-items: center;
  background: linear-gradient(135deg, rgba(99,102,241,.40), rgba(236,72,153,.32));
}
.stat .sv {font-size: 1.45rem; font-weight: 800; line-height: 1.1;}
.stat .sv.sm {font-size: 1.02rem;}
.stat .sl {font-size: .7rem; opacity: .65; text-transform: uppercase; letter-spacing: .09em;}
.footer {text-align: center; opacity: .6; font-size: .82rem; margin: 2.8rem 0 1rem; padding-top: 1.1rem; border-top: 1px solid var(--glass-border);}

/* Streamlit widgets */
div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {border-radius: 12px; transition: transform .2s ease, box-shadow .2s ease;}
div[data-testid="stButton"] button:hover, div[data-testid="stDownloadButton"] button:hover {transform: translateY(-2px);}
button[kind="primary"], button[data-testid="stBaseButton-primary"] {
  background: linear-gradient(135deg, #6366F1, #A855F7, #EC4899) !important; border: none !important; color: #fff !important;
  font-weight: 700 !important; box-shadow: 0 12px 26px -10px rgba(168,85,247,.9);
}
button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {transform: translateY(-3px) scale(1.01); box-shadow: 0 20px 36px -12px rgba(236,72,153,.85);}
[data-testid="stFileUploaderDropzone"] {
  border: 2px dashed rgba(168,85,247,.6) !important; border-radius: 18px !important; padding: 1.5rem !important;
  background: linear-gradient(135deg, rgba(99,102,241,.10), rgba(236,72,153,.08)) !important; transition: all .3s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {border-color: #EC4899 !important; transform: translateY(-3px); box-shadow: 0 0 0 5px rgba(236,72,153,.10), 0 24px 44px -22px rgba(168,85,247,.8);}
[data-testid="stChatMessage"] {
  background: var(--glass); border: 1px solid var(--glass-border); border-radius: 18px; padding: 1rem 1.15rem; margin-bottom: .75rem;
  backdrop-filter: blur(10px); box-shadow: 0 16px 32px -24px rgba(0,0,0,.6); animation: rise .4s ease both;
}
@keyframes rise {from {opacity: 0; transform: translateY(12px);} to {opacity: 1; transform: none;}}
[data-testid="stChatInput"] {border-radius: 16px; border: 1px solid rgba(168,85,247,.5); box-shadow: 0 14px 36px -18px rgba(168,85,247,.85);}
[data-testid="stExpander"] {border-radius: 14px; border: 1px solid var(--glass-border) !important; background: var(--glass);}
section[data-testid="stSidebar"] {border-right: 1px solid var(--glass-border);}
section[data-testid="stSidebar"] h3 {font-size: .72rem; letter-spacing: .1em; opacity: .65;}
::-webkit-scrollbar {width: 10px; height: 10px;}
::-webkit-scrollbar-thumb {background: linear-gradient(#6366F1, #A855F7); border-radius: 10px;}

/* typography (icons keep their own font) */
.stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp button, .stApp input, .stApp textarea {
  font-family: 'Plus Jakarta Sans', 'Segoe UI', system-ui, -apple-system, sans-serif;
}
/* top bar */
.topbar {position: relative; z-index: 2;}
.topbar .tag {font-size: .78rem; opacity: .6; font-weight: 600; letter-spacing: .04em;}
/* sign-in card */
[data-testid="stForm"] {
  position: relative; overflow: hidden; background: var(--glass); border: 1px solid var(--glass-border) !important; border-radius: 24px;
  padding: 1.7rem 1.6rem 1.4rem; backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px);
  box-shadow: 0 1px 0 rgba(255,255,255,.12) inset, 0 40px 80px -36px rgba(139,92,246,.7);
}
[data-testid="stForm"]::before {content: ""; position: absolute; inset: 0 0 auto 0; height: 4px; background: linear-gradient(90deg, #6366F1, #A855F7, #EC4899, #22D3EE);}
div[data-baseweb="input"] {border-radius: 12px !important; border: 1px solid rgba(168,85,247,.45) !important; transition: box-shadow .2s, border-color .2s;}
div[data-baseweb="input"]:focus-within {border-color: #A855F7 !important; box-shadow: 0 0 0 4px rgba(168,85,247,.22);}
.gate-note {text-align: center; font-size: .82rem; opacity: .75; margin: .9rem 0 .4rem;}
.gate-note a {color: #A855F7; font-weight: 700; text-decoration: none;}
.trust {display: flex; flex-wrap: wrap; justify-content: center; gap: .5rem; margin-top: .6rem;}
.trust span {font-size: .76rem; font-weight: 600; padding: .28rem .7rem; border-radius: 999px; border: 1px solid var(--glass-border); background: var(--glass);}
/* account card */
.account {display: flex; align-items: center; gap: .75rem; padding: .75rem .9rem; margin: .2rem 0 .6rem;}
.account .av {
  width: 38px; height: 38px; border-radius: 50%; display: grid; place-items: center; font-weight: 800; color: #fff;
  background: linear-gradient(135deg, #22C55E, #22D3EE); box-shadow: 0 8px 20px -8px rgba(34,197,94,.8);
}
.account .an {font-weight: 800; font-size: .95rem; line-height: 1.1;}
.account .as {font-size: .74rem; opacity: .65;}
hr {border: none; height: 1px; background: linear-gradient(90deg, transparent, rgba(168,85,247,.5), transparent);}
@media (max-width: 700px) {.hero-title {font-size: 2.3rem;} .topbar .tag {display: none;}}
@media (prefers-reduced-motion: reduce) {* {animation: none !important; transition: none !important;}}
</style>
"""
)

# ----------------------------------------------------------------------------
# 3D hero (Three.js). Loaded in an iframe from cdnjs; shows an emoji if WebGL/CDN is unavailable.
# ----------------------------------------------------------------------------
HERO_3D = """
<!doctype html><html><head><meta charset="utf-8">
<style>
html,body{margin:0;height:100%;background:transparent;overflow:hidden}
#c{width:100%;height:100%;display:block}
#fb{display:none;position:absolute;inset:0;align-items:center;justify-content:center;font:700 96px system-ui}
</style></head><body>
<canvas id="c"></canvas><div id="fb">&#128218;</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
(function(){
  var canvas=document.getElementById('c'), fb=document.getElementById('fb');
  function fallback(){fb.style.display='flex';canvas.style.display='none';}
  if(typeof THREE==='undefined'){fallback();return;}
  var renderer;
  try{renderer=new THREE.WebGLRenderer({canvas:canvas,alpha:true,antialias:true});}catch(e){fallback();return;}
  function W(){return canvas.clientWidth||420;} function H(){return canvas.clientHeight||330;}
  renderer.setPixelRatio(Math.min(window.devicePixelRatio||1,2));
  renderer.setSize(W(),H(),false);
  var scene=new THREE.Scene();
  var camera=new THREE.PerspectiveCamera(42,W()/H(),0.1,100); camera.position.set(0,0,8.4);
  scene.add(new THREE.AmbientLight(0xffffff,0.55));
  var l1=new THREE.PointLight(0x818cf8,2.2,40); l1.position.set(6,5,7); scene.add(l1);
  var l2=new THREE.PointLight(0xf472b6,2.0,40); l2.position.set(-7,-4,6); scene.add(l2);
  var l3=new THREE.PointLight(0x22d3ee,1.4,40); l3.position.set(0,6,-4); scene.add(l3);
  var world=new THREE.Group(); scene.add(world);

  var core=new THREE.Mesh(new THREE.IcosahedronGeometry(1.25,1),
    new THREE.MeshStandardMaterial({color:0x6366f1,metalness:0.55,roughness:0.28,flatShading:true,emissive:0x312e81,emissiveIntensity:0.6}));
  world.add(core);
  var shell=new THREE.Mesh(new THREE.IcosahedronGeometry(1.75,1),
    new THREE.MeshBasicMaterial({color:0xc084fc,wireframe:true,transparent:true,opacity:0.35}));
  world.add(shell);
  var ring=new THREE.Mesh(new THREE.TorusGeometry(2.5,0.03,12,160),new THREE.MeshBasicMaterial({color:0x22d3ee,transparent:true,opacity:0.75}));
  ring.rotation.x=1.2; world.add(ring);
  var ring2=new THREE.Mesh(new THREE.TorusGeometry(2.95,0.025,12,160),new THREE.MeshBasicMaterial({color:0xf472b6,transparent:true,opacity:0.6}));
  ring2.rotation.x=-0.9; ring2.rotation.y=0.5; world.add(ring2);

  var palette=[0x6366f1,0xa855f7,0xec4899,0x22d3ee,0xf59e0b,0x34d399], cards=[], i, j;
  for(i=0;i<6;i++){
    var g=new THREE.Group();
    g.add(new THREE.Mesh(new THREE.BoxGeometry(0.95,1.25,0.07),
      new THREE.MeshStandardMaterial({color:0xffffff,metalness:0.1,roughness:0.35,emissive:palette[i],emissiveIntensity:0.35})));
    var band=new THREE.Mesh(new THREE.BoxGeometry(0.95,0.22,0.09),new THREE.MeshStandardMaterial({color:palette[i],metalness:0.3,roughness:0.3}));
    band.position.y=0.5; g.add(band);
    for(j=0;j<4;j++){
      var line=new THREE.Mesh(new THREE.BoxGeometry(0.62-(j%2)*0.18,0.05,0.09),new THREE.MeshBasicMaterial({color:0x94a3b8}));
      line.position.set(-0.05-((j%2)?0.09:0),0.2-j*0.17,0); g.add(line);
    }
    cards.push({g:g,a:i/6*Math.PI*2,r:3.3+(i%2)*0.35,s:0.25+(i%3)*0.05,y:(i%3-1)*0.9,p:i});
    world.add(g);
  }

  var N=420,pos=new Float32Array(N*3);
  for(i=0;i<N;i++){
    var rr=4+Math.random()*5,th=Math.random()*6.283,ph=Math.acos(2*Math.random()-1);
    pos[i*3]=rr*Math.sin(ph)*Math.cos(th); pos[i*3+1]=rr*Math.sin(ph)*Math.sin(th)*0.7; pos[i*3+2]=rr*Math.cos(ph)-2;
  }
  var pg=new THREE.BufferGeometry(); pg.setAttribute('position',new THREE.BufferAttribute(pos,3));
  var stars=new THREE.Points(pg,new THREE.PointsMaterial({color:0xc7d2fe,size:0.05,transparent:true,opacity:0.8}));
  scene.add(stars);

  var tx=0,ty=0,cx=0,cy=0;
  window.addEventListener('pointermove',function(e){
    var r=canvas.getBoundingClientRect();
    tx=((e.clientX-r.left)/r.width-0.5)*2; ty=((e.clientY-r.top)/r.height-0.5)*2;
  });
  window.addEventListener('resize',function(){
    renderer.setSize(W(),H(),false); camera.aspect=W()/H(); camera.updateProjectionMatrix();
  });

  var clock=new THREE.Clock();
  function frame(){
    var t=clock.getElapsedTime();
    cx+=(tx-cx)*0.06; cy+=(ty-cy)*0.06;
    world.rotation.y=cx*0.55+Math.sin(t*0.3)*0.15; world.rotation.x=cy*0.35;
    core.rotation.x=t*0.35; core.rotation.y=t*0.5; core.scale.setScalar(1+Math.sin(t*1.6)*0.03);
    shell.rotation.x=-t*0.2; shell.rotation.y=t*0.25;
    ring.rotation.z=t*0.3; ring2.rotation.z=-t*0.22;
    for(var k=0;k<cards.length;k++){
      var c=cards[k],a=c.a+t*c.s;
      c.g.position.set(Math.cos(a)*c.r,c.y+Math.sin(t*0.9+c.p)*0.18,Math.sin(a)*c.r*0.55);
      c.g.rotation.y=-a*0.6+0.3; c.g.rotation.z=Math.sin(t+c.p)*0.12;
    }
    stars.rotation.y=t*0.02;
    renderer.render(scene,camera);
    requestAnimationFrame(frame);
  }
  frame();
})();
</script></body></html>
"""


def render_topbar() -> None:
    if st.session_state.get("auth_ok"):
        status = "Knowledge base ready" if st.session_state.get("kb") is not None else f"Signed in · {st.session_state.get('provider', '')}"
    else:
        status = "Secure sign-in"
    ui_html(
        f"""
        <div class="glass compact topbar"><div class="logo">📚</div><div class="t">{APP_NAME}</div>
        <span class="tag">AI DOCUMENT ASSISTANT</span>
        <div class="badge"><span class="dot"></span> {status}</div></div>
        """
    )


render_topbar()


INSTALL_CMD = (
    "pip install streamlit pandas langchain-core langchain-community langchain-text-splitters "
    "langchain-openai langchain-google-genai faiss-cpu pypdf python-docx python-pptx openpyxl beautifulsoup4"
)

# Heavy libraries are imported AFTER the page header is drawn, so the page is never blank,
# and a missing package shows a clear message instead of a white screen.
with st.spinner("Loading AI libraries… (the first launch can take up to a minute)"):
    try:
        import hashlib
        import io
        import json
        from dataclasses import dataclass, field
        from datetime import datetime
        from pathlib import Path
        from html import escape as esc
        from typing import Callable, Iterator

        import faiss  # noqa: F401  (fail early if faiss-cpu is missing)
        import pandas as pd
        from langchain_community.vectorstores import FAISS
        from langchain_core.documents import Document
        from langchain_core.messages import AIMessage, BaseMessage, HumanMessage
        from langchain_core.output_parsers import StrOutputParser
        from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
        from langchain_text_splitters import RecursiveCharacterTextSplitter
    except Exception as _import_error:  # noqa: BLE001
        st.error("❌ A required package could not be loaded. Install everything, then restart the app.")
        st.code(INSTALL_CMD, language="bash")
        st.exception(_import_error)
        st.stop()



# ============================================================================
# 1) FILE LOADERS
# ============================================================================



MAX_FILE_MB = 50


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _decode(data: bytes) -> str:
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _doc(text: str, location: str, **extra) -> Document:
    return Document(page_content=text, metadata={"location": location, **extra})


def _df_to_docs(df, label: str = "", rows_per_block: int = 40) -> list[Document]:
    """Turn a DataFrame into blocks of rows, repeating the header in every block."""
    df = df.dropna(how="all").dropna(axis=1, how="all").fillna("")
    if df.empty:
        return []
    header = " | ".join(str(c) for c in df.columns)
    docs = []
    for start in range(0, len(df), rows_per_block):
        block = df.iloc[start : start + rows_per_block]
        lines = [header]
        lines += [" | ".join(str(v) for v in row) for row in block.itertuples(index=False)]
        location = f"{label}rows {start + 1}-{start + len(block)}"
        docs.append(_doc("\n".join(lines), location))
    return docs


# ----------------------------------------------------------------------------
# Individual loaders
# ----------------------------------------------------------------------------
def _load_pdf(data: bytes) -> list[Document]:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted and not reader.decrypt(""):
        raise ValueError("PDF is password-protected.")
    docs = []
    for number, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        if text:
            docs.append(_doc(text, f"Page {number}", page=number))
    if not docs:
        raise ValueError("No extractable text found (scanned PDF? OCR is not supported).")
    return docs


def _load_docx(data: bytes) -> list[Document]:
    from docx import Document as DocxDocument
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = DocxDocument(io.BytesIO(data))
    sections: list[tuple[str, list[str]]] = [("Document start", [])]

    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph = Paragraph(child, document)
            text = paragraph.text.strip()
            if not text:
                continue
            style = getattr(getattr(paragraph, "style", None), "name", "") or ""
            if style.startswith("Heading") or style == "Title":
                sections.append((text[:80], []))
            sections[-1][1].append(text)
        elif child.tag.endswith("}tbl"):
            table = Table(child, document)
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                if any(cells):
                    sections[-1][1].append(" | ".join(cells))

    docs = [_doc("\n".join(lines), f"Section: {title}") for title, lines in sections if lines]
    if not docs:
        raise ValueError("The document appears to be empty.")
    return docs


def _load_pptx(data: bytes) -> list[Document]:
    from pptx import Presentation

    presentation = Presentation(io.BytesIO(data))
    docs = []
    for number, slide in enumerate(presentation.slides, start=1):
        parts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                parts.append(shape.text_frame.text.strip())
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    parts.append(" | ".join(cell.text.strip() for cell in row.cells))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                parts.append(f"Speaker notes: {notes}")
        if parts:
            docs.append(_doc("\n".join(parts), f"Slide {number}", page=number))
    if not docs:
        raise ValueError("No text found in the presentation.")
    return docs


def _load_csv(data: bytes) -> list[Document]:
    import pandas as pd

    df = pd.read_csv(io.StringIO(_decode(data)), sep=None, engine="python", dtype=str)
    docs = _df_to_docs(df)
    if not docs:
        raise ValueError("The CSV file is empty.")
    return docs


def _load_xlsx(data: bytes) -> list[Document]:
    import pandas as pd

    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=str)
    docs = []
    for sheet_name, df in sheets.items():
        docs += _df_to_docs(df, label=f"Sheet '{sheet_name}', ")
    if not docs:
        raise ValueError("The workbook has no data.")
    return docs


def _load_html(data: bytes) -> list[Document]:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(_decode(data), "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = "\n".join(line.strip() for line in soup.get_text("\n").splitlines() if line.strip())
    if not text:
        raise ValueError("No readable text found in the HTML file.")
    return [_doc(text, "Web page")]


def _load_json(data: bytes) -> list[Document]:
    raw = _decode(data)
    try:
        raw = json.dumps(json.loads(raw), indent=2, ensure_ascii=False)
    except json.JSONDecodeError:
        pass  # fall back to the raw text
    if not raw.strip():
        raise ValueError("The JSON file is empty.")
    return [_doc(raw, "JSON data")]


def _load_text(data: bytes) -> list[Document]:
    text = _decode(data).strip()
    if not text:
        raise ValueError("The file is empty.")
    return [_doc(text, "Text")]


# ----------------------------------------------------------------------------
# Registry
# ----------------------------------------------------------------------------
LOADERS: dict[str, Callable[[bytes], list[Document]]] = {
    "pdf": _load_pdf,
    "docx": _load_docx,
    "pptx": _load_pptx,
    "csv": _load_csv,
    "xlsx": _load_xlsx,
    "html": _load_html,
    "htm": _load_html,
    "json": _load_json,
    "txt": _load_text,
    "md": _load_text,
    "log": _load_text,
    "py": _load_text,
}

SUPPORTED_TYPES = list(LOADERS)


def load_file(name: str, data: bytes) -> list[Document]:
    """Load a file and return its documents tagged with `source` and `location`."""
    extension = Path(name).suffix.lower().lstrip(".")
    loader = LOADERS.get(extension)
    if loader is None:
        raise ValueError(f"Unsupported file type: .{extension or '?'}")
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        raise ValueError(f"File is larger than {MAX_FILE_MB} MB.")
    docs = loader(data)
    for doc in docs:
        doc.metadata["source"] = name
        doc.metadata.setdefault("location", "Document")
    return docs


# ============================================================================
# 2) RAG ENGINE
# ============================================================================




# ----------------------------------------------------------------------------
# Providers
# ----------------------------------------------------------------------------
PROVIDERS = {
    "OpenAI": {
        "chat_models": ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"],
        "embedding_model": "text-embedding-3-small",
        "key_env": "OPENAI_API_KEY",
        "placeholder": "sk-...",
        "key_url": "https://platform.openai.com/api-keys",
    },
    "Google Gemini": {
        "chat_models": ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro"],
        "embedding_model": "models/gemini-embedding-001",
        "key_env": "GOOGLE_API_KEY",
        "placeholder": "AIzaSy...",
        "key_url": "https://aistudio.google.com/apikey",
    },
}

HISTORY_WINDOW = 6  # number of previous messages sent to the model


def get_llm(provider: str, model: str, api_key: str, temperature: float = 0.0):
    if provider == "OpenAI":
        from langchain_openai import ChatOpenAI

        return ChatOpenAI(model=model, temperature=temperature, api_key=api_key)
    from langchain_google_genai import ChatGoogleGenerativeAI

    return ChatGoogleGenerativeAI(
        model=model, temperature=temperature, google_api_key=api_key, max_retries=2
    )


def get_embeddings(provider: str, api_key: str):
    model = PROVIDERS[provider]["embedding_model"]
    if provider == "OpenAI":
        from langchain_openai import OpenAIEmbeddings

        return OpenAIEmbeddings(model=model, api_key=api_key)
    from langchain_google_genai import GoogleGenerativeAIEmbeddings

    return GoogleGenerativeAIEmbeddings(model=model, google_api_key=api_key)


# ----------------------------------------------------------------------------
# Prompts
# ----------------------------------------------------------------------------
ANSWER_PROMPT = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            "You are DocMind, an assistant that answers questions strictly from the "
            "document excerpts provided below.\n\n"
            "Rules:\n"
            "1. Use ONLY the context below. Never rely on outside knowledge.\n"
            "2. If the answer is not in the context, say: "
            "\"I couldn't find that in the uploaded documents.\" and, if useful, "
            "suggest what the user could ask instead.\n"
            "3. Cite the sources you used inline, like (report.pdf, Page 3).\n"
            "4. Be clear and well structured. Use Markdown (bullets, bold, tables) when it helps.\n"
            "5. Reply in the same language as the user's question.\n\n"
            "Context:\n{context}",
        ),
        MessagesPlaceholder("history"),
        ("human", "{question}"),
    ]
)

CONDENSE_PROMPT = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            "Given the chat history and a follow-up message, rewrite the follow-up as a "
            "standalone search question that makes sense without the history. "
            "Do NOT answer it. If it is already standalone, return it unchanged. "
            "Return only the question.",
        ),
        MessagesPlaceholder("history"),
        ("human", "{question}"),
    ]
)

SUMMARY_PROMPT = ChatPromptTemplate.from_template(
    "Below are excerpts sampled from the user's documents.\n\n{context}\n\n"
    "Write a concise, well-structured summary in Markdown with:\n"
    "1. **Overview** (2-3 sentences)\n"
    "2. **Key points** (bullets, mention the file name when relevant)\n"
    "3. **Notable facts** (important numbers, dates, names)\n"
    "Use only the information in the excerpts."
)


# ----------------------------------------------------------------------------
# Knowledge base
# ----------------------------------------------------------------------------
@dataclass
class KnowledgeBase:
    store: FAISS
    chunks: list[Document]
    file_table: list[dict]
    skipped: list[tuple[str, str]]
    signature: str
    words: int = 0
    extra: dict = field(default_factory=dict)


def make_signature(files: list[tuple[str, bytes]], provider: str, chunk_size: int, overlap: int) -> str:
    digest = hashlib.sha256()
    for name, data in files:
        digest.update(name.encode())
        digest.update(hashlib.sha256(data).digest())
    digest.update(f"{provider}|{chunk_size}|{overlap}".encode())
    return digest.hexdigest()


def build_knowledge_base(
    files: list[tuple[str, bytes]],
    *,
    provider: str,
    api_key: str,
    chunk_size: int,
    chunk_overlap: int,
    signature: str,
    on_step: Callable[[str], None] | None = None,
    on_progress: Callable[[float], None] | None = None,
    batch_size: int = 64,
) -> KnowledgeBase:
    step = on_step or (lambda _msg: None)
    progress = on_progress or (lambda _v: None)

    # 1) Load every file, skipping the ones that fail
    raw_docs: list[Document] = []
    skipped: list[tuple[str, str]] = []
    for name, data in files:
        step(f"📖 Reading **{name}**")
        try:
            raw_docs += load_file(name, data)
        except Exception as err:  # noqa: BLE001 - report per-file problems to the user
            skipped.append((name, str(err)))
    if not raw_docs:
        reasons = "; ".join(f"{n}: {r}" for n, r in skipped) or "no readable content"
        raise ValueError(f"None of the files could be read ({reasons}).")

    # 2) Split into chunks
    step("✂️ Splitting into chunks")
    splitter = RecursiveCharacterTextSplitter(chunk_size=chunk_size, chunk_overlap=chunk_overlap)
    chunks = splitter.split_documents(raw_docs)
    for i, chunk in enumerate(chunks):
        chunk.metadata["chunk_id"] = i

    # 3) Embed + index in batches (progress bar + friendlier on rate limits)
    step(f"🧠 Creating embeddings for {len(chunks)} chunks")
    embeddings = get_embeddings(provider, api_key)
    store: FAISS | None = None
    for start in range(0, len(chunks), batch_size):
        batch = chunks[start : start + batch_size]
        if store is None:
            store = FAISS.from_documents(batch, embeddings)
        else:
            store.add_documents(batch)
        progress(min(start + batch_size, len(chunks)) / len(chunks))

    # 4) Per-file statistics
    table = []
    for name, data in files:
        if any(n == name for n, _ in skipped):
            continue
        file_docs = [d for d in raw_docs if d.metadata["source"] == name]
        words = sum(len(d.page_content.split()) for d in file_docs)
        table.append(
            {
                "File": name,
                "Type": Path(name).suffix.lower().lstrip(".").upper(),
                "Size (KB)": round(len(data) / 1024, 1),
                "Sections": len(file_docs),
                "Words": words,
                "Chunks": sum(1 for c in chunks if c.metadata["source"] == name),
            }
        )

    return KnowledgeBase(
        store=store,
        chunks=chunks,
        file_table=table,
        skipped=skipped,
        signature=signature,
        words=sum(row["Words"] for row in table),
    )


# ----------------------------------------------------------------------------
# Retrieval + generation
# ----------------------------------------------------------------------------
def _to_messages(history: list[dict]) -> list[BaseMessage]:
    messages: list[BaseMessage] = []
    for item in history[-HISTORY_WINDOW:]:
        cls = HumanMessage if item["role"] == "user" else AIMessage
        messages.append(cls(content=item["content"]))
    return messages


def condense_question(llm, history: list[dict], question: str) -> str:
    """Turn a follow-up like 'and what about 2023?' into a standalone search query."""
    if not history:
        return question
    chain = CONDENSE_PROMPT | llm | StrOutputParser()
    rewritten = chain.invoke({"history": _to_messages(history), "question": question}).strip()
    return rewritten or question


def retrieve(kb: KnowledgeBase, query: str, k: int, use_mmr: bool) -> list[Document]:
    if use_mmr:  # Maximal Marginal Relevance: relevant AND diverse passages
        return kb.store.max_marginal_relevance_search(query, k=k, fetch_k=max(k * 4, 20))
    return kb.store.similarity_search(query, k=k)


def format_docs(docs: list[Document]) -> str:
    return "\n\n---\n\n".join(
        f"[Source: {d.metadata.get('source', '?')} · {d.metadata.get('location', '?')}]\n{d.page_content}"
        for d in docs
    )


def answer_stream(llm, question: str, docs: list[Document], history: list[dict]) -> Iterator[str]:
    chain = ANSWER_PROMPT | llm | StrOutputParser()
    yield from chain.stream(
        {"context": format_docs(docs), "question": question, "history": _to_messages(history)}
    )


def summarize(llm, kb: KnowledgeBase, max_chunks: int = 14) -> str:
    step = max(1, len(kb.chunks) // max_chunks)
    sample = kb.chunks[::step][:max_chunks]
    context = "\n\n".join(
        f"[{c.metadata['source']} · {c.metadata['location']}]\n{c.page_content[:1500]}" for c in sample
    )
    return (SUMMARY_PROMPT | llm | StrOutputParser()).invoke({"context": context})


# ----------------------------------------------------------------------------
# Errors
# ----------------------------------------------------------------------------
def friendly_error(err: Exception) -> str:
    msg = str(err).lower()
    if any(s in msg for s in ("incorrect api key", "api key not valid", "api_key_invalid", "401", "403", "permission")):
        return "🔑 This API key looks invalid or lacks permission. Please check it and try again."
    if any(s in msg for s in ("quota", "429", "resource_exhausted", "rate limit", "insufficient")):
        return "⏳ Rate limit or quota reached. Wait a moment, or check your billing/quota, then try again."
    if "404" in msg or "not found" in msg or "does not exist" in msg:
        return "🧭 This model isn't available for your key. Pick another model in the sidebar."
    if "safety" in msg or "blocked" in msg:
        return "🛡️ The response was blocked by safety filters. Try rephrasing."
    return f"⚠️ Something went wrong: {err}"


# ============================================================================
# 3) STREAMLIT APP
# ============================================================================




st.session_state.setdefault("messages", [])
st.session_state.setdefault("kb", None)
st.session_state.setdefault("summary", None)
st.session_state.setdefault("auth_ok", False)
st.session_state.setdefault("provider", None)
st.session_state.setdefault("api_key", "")


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def pack_sources(docs) -> list[dict]:
    seen, sources = set(), []
    for doc in docs:
        key = (doc.metadata.get("source"), doc.metadata.get("location"), doc.page_content[:80])
        if key in seen:
            continue
        seen.add(key)
        sources.append(
            {
                "source": doc.metadata.get("source", "?"),
                "location": doc.metadata.get("location", "?"),
                "text": doc.page_content[:600],
            }
        )
    return sources


def render_sources(sources: list[dict]) -> None:
    if not sources:
        return
    with st.expander(f"🔎 Sources · {len(sources)} passages"):
        for item in sources:
            st.markdown(f"**📄 {item['source']}** · `{item['location']}`")
            st.text(item["text"])
            st.divider()


def export_markdown() -> str:
    lines = [f"# {APP_NAME} chat export", f"_Exported {datetime.now():%Y-%m-%d %H:%M}_", ""]
    for msg in st.session_state.messages:
        who = "You" if msg["role"] == "user" else APP_NAME
        lines += [f"**{who}:**", "", msg["content"], ""]
        if msg.get("sources"):
            cited = sorted({f"{s['source']} ({s['location']})" for s in msg["sources"]})
            lines += ["_Sources: " + "; ".join(cited) + "_", ""]
    return "\n".join(lines)


def detect_provider(key: str) -> str | None:
    if key.startswith("AIza"):
        return "Google Gemini"
    if key.startswith("sk-"):
        return "OpenAI"
    return None


def verify_api_key(provider: str, api_key: str) -> None:
    """Cheap live check: embeds one word. Raises if the key is invalid or lacks access."""
    get_embeddings(provider, api_key).embed_query("connection test")


def mask_key(key: str) -> str:
    return f"{key[:4]}…{key[-3:]}" if len(key) > 10 else "••••"


def sign_out() -> None:
    st.session_state.auth_ok = False
    st.session_state.provider = None
    st.session_state.api_key = ""
    st.session_state.kb = None
    st.session_state.summary = None
    st.session_state.messages = []
    st.session_state.pop("pending_prompt", None)


def render_gate() -> None:
    """Full-screen sign-in: nothing in the app is reachable without a verified API key."""
    left, right = st.columns([1.1, 0.9], gap="large")
    with left:
        ui_html(
            f"""
            <div class="badge"><span class="dot"></span> AI-powered · Private · Cited answers</div>
            <div class="hero-title">Chat with your<br>documents.</div>
            <div class="hero-sub">{APP_NAME} reads your PDFs, Word files, spreadsheets, slides and more,
            then answers in seconds, with the exact page as proof.</div>
            <div class="pills">
              <span class="pill">⚡ Streaming answers</span><span class="pill">📎 Page-level citations</span>
              <span class="pill">🌍 Any language</span><span class="pill">🔒 Session-only memory</span>
            </div>
            """
        )
        try:
            components.html(HERO_3D, height=300)
        except Exception:  # noqa: BLE001
            pass
    with right:
        ui_html('<div style="height:1.6rem"></div>')
        with st.form("signin_form", clear_on_submit=False):
            ui_html(
                """
                <div class="brand" style="margin-bottom:1rem"><div class="logo">🔑</div>
                <div><div class="brand-name">Sign in to DocMind</div>
                <div class="brand-sub">Enter your API key to unlock the app</div></div></div>
                """
            )
            choice = st.radio("AI provider", ["Auto-detect", *PROVIDERS], horizontal=True)
            key_input = st.text_input("API key", type="password", placeholder="Paste your OpenAI (sk-…) or Gemini (AIza…) key")
            submitted = st.form_submit_button("🔓 Unlock DocMind", type="primary", use_container_width=True)
        ui_html(
            """
            <div class="gate-note">Don't have a key? Get one from
            <a href="https://platform.openai.com/api-keys" target="_blank">OpenAI ↗</a> or
            <a href="https://aistudio.google.com/apikey" target="_blank">Google AI Studio ↗</a></div>
            <div class="trust"><span>🔒 Key stays in this session</span><span>🚫 No documents stored</span><span>⚡ No account needed</span></div>
            """
        )
        if submitted:
            key = (key_input or "").strip()
            provider_choice = detect_provider(key) if choice == "Auto-detect" else choice
            if not key:
                st.error("Please paste your API key to continue.")
            elif provider_choice is None:
                st.error("We couldn't recognise this key. Please choose the provider manually.")
            else:
                ok = False
                with st.spinner("Verifying your key…"):
                    try:
                        verify_api_key(provider_choice, key)
                        ok = True
                    except Exception as err:  # noqa: BLE001
                        st.error(friendly_error(err))
                if ok:
                    st.session_state.auth_ok = True
                    st.session_state.provider = provider_choice
                    st.session_state.api_key = key
                    st.rerun()
    render_landing()


FORMAT_ICONS = {
    "pdf": "📕", "docx": "📘", "pptx": "📙", "xlsx": "📗", "csv": "📊",
    "html": "🌐", "json": "🧾", "txt": "📝", "md": "📝", "log": "📋", "py": "🐍",
}


def render_footer() -> None:
    ui_html(
        f"""
        <div class="footer">© {datetime.now().year} {APP_NAME} · Powered by LangChain, FAISS, OpenAI &amp; Gemini<br>
        AI can make mistakes. Always verify important information using the cited sources.</div>
        """
    )


def render_landing() -> None:
    ui_html('<div class="section-title">How it works</div>')
    ui_html(
        """
        <div class="steps">
          <div class="glass step tilt"><div class="num">STEP 1</div><h4>🔑 Unlock with your key</h4>
            <p>Sign in with your OpenAI or Gemini API key. It is verified once and stays in your session only.</p></div>
          <div class="glass step tilt"><div class="num">STEP 2</div><h4>📤 Upload documents</h4>
            <p>Drop one or many files, in any mix of formats, then press <b>Process documents</b>.</p></div>
          <div class="glass step tilt"><div class="num">STEP 3</div><h4>💬 Ask and verify</h4>
            <p>Get answers with citations and see the exact passages they came from.</p></div>
        </div>
        """
    )
    ui_html('<div class="section-title">Built for real work</div>')
    features = [
        ("📚", "Every file type", "PDF, Word, PowerPoint, Excel, CSV, HTML, JSON and text, all in one conversation."),
        ("📎", "Verifiable answers", "Each answer cites the file and page, and the source passages are one click away."),
        ("💬", "Follow-up aware", "Ask “and what about last year?” and DocMind remembers what you were discussing."),
        ("⚡", "Real-time streaming", "Answers appear word by word, so there is no waiting for a full reply."),
        ("🧠", "OpenAI or Gemini", "Switch AI providers and models in one click, with your own key."),
        ("🔒", "Privacy-minded", "Documents are processed in memory for your session and never saved by DocMind."),
    ]
    cards = "".join(
        f'<div class="glass card tilt"><div class="ico">{icon}</div><h4>{title}</h4><p>{text}</p></div>'
        for icon, title, text in features
    )
    ui_html(f'<div class="grid3">{cards}</div>')
    ui_html('<div class="section-title">Supported formats</div>')
    tiles = "".join(
        f'<span class="fmt">{FORMAT_ICONS.get(ext, "📄")} .{ext}</span>' for ext in SUPPORTED_TYPES if ext != "htm"
    )
    ui_html(f'<div class="fmt-row">{tiles}</div>')
    render_footer()


# ----------------------------------------------------------------------------
# Sign-in gate: nothing below runs until an API key has been verified
# ----------------------------------------------------------------------------
if not st.session_state.auth_ok:
    render_gate()
    st.stop()

provider = st.session_state.provider
api_key = st.session_state.api_key
cfg = PROVIDERS[provider]

# ----------------------------------------------------------------------------
# Sidebar
# ----------------------------------------------------------------------------
with st.sidebar:
    ui_html(
        f"""
        <div class="brand"><div class="logo">📚</div>
        <div><div class="brand-name">{APP_NAME}</div><div class="brand-sub">Chat with your documents · {APP_VERSION}</div></div></div>
        """
    )

    ui_html(
        f"""
        <div class="glass account"><div class="av">✓</div>
        <div><div class="an">Signed in</div><div class="as">{esc(provider)} · {esc(mask_key(api_key))}</div></div></div>
        """
    )
    if st.button("🚪 Sign out", use_container_width=True):
        sign_out()
        st.rerun()

    st.markdown("### MODEL")
    model_name = st.selectbox("Chat model", cfg["chat_models"])

    st.markdown("### SETTINGS")
    with st.expander("✂️ Indexing"):
        chunk_size = st.slider("Chunk size", 500, 2000, 1000, step=100)
        chunk_overlap = st.slider("Chunk overlap", 0, 400, 150, step=50)
        st.caption("Changing these requires re-processing your documents.")
    with st.expander("🎯 Retrieval & answers"):
        top_k = st.slider("Passages to retrieve (k)", 2, 10, 4)
        use_mmr = st.toggle("Diverse results (MMR)", value=True, help="Avoids near-duplicate passages.")
        temperature = st.slider("Creativity", 0.0, 1.0, 0.0, 0.1, help="0 = strictly factual.")

    st.markdown("### CHAT")
    if st.session_state.messages:
        if st.button("🧹 Clear chat", use_container_width=True):
            st.session_state.messages = []
            st.rerun()
        st.download_button(
            "⬇️ Export chat (.md)",
            data=export_markdown(),
            file_name="docmind_chat.md",
            mime="text/markdown",
            use_container_width=True,
        )
    else:
        st.caption("Your conversation will appear here.")

    st.markdown("### ABOUT")
    with st.expander("ℹ️ Privacy & info"):
        st.caption(
            "• Documents are processed in memory for this session and are not saved by DocMind.\n\n"
            "• Relevant excerpts are sent to the AI provider you choose in order to generate answers.\n\n"
            "• AI can make mistakes. Verify important details using the cited sources."
        )
    st.caption(f"{APP_NAME} {APP_VERSION} · © {datetime.now().year}")

# ----------------------------------------------------------------------------
# Upload
# ----------------------------------------------------------------------------
uploaded = st.file_uploader(
    "Drop your documents here",
    type=SUPPORTED_TYPES,
    accept_multiple_files=True,
    help=f"Up to {MAX_FILE_MB} MB per file. You can upload several files of different types at once.",
)

if not uploaded:
    st.session_state.kb = None
    st.session_state.messages = []
    st.session_state.summary = None
    render_landing()
    st.stop()

files = [(f.name, f.getvalue()) for f in uploaded]
signature = make_signature(files, provider, chunk_size, chunk_overlap)
kb = st.session_state.kb

# ----------------------------------------------------------------------------
# Index documents
# ----------------------------------------------------------------------------
if kb is None or kb.signature != signature:
    label = "⚡ Process documents" if kb is None else "🔄 Re-process (files or settings changed)"
    if not st.button(label, type="primary", use_container_width=True):
        st.info("Click the button to read, split and index your documents. This only takes a moment.")
        st.stop()

    error = None
    new_kb = None
    with st.status("Processing documents…", expanded=True) as status:
        bar = st.progress(0.0)
        try:
            new_kb = build_knowledge_base(
                files,
                provider=provider,
                api_key=api_key,
                chunk_size=chunk_size,
                chunk_overlap=chunk_overlap,
                signature=signature,
                on_step=st.write,
                on_progress=bar.progress,
            )
            status.update(label="Ready! Ask away 🎉", state="complete", expanded=False)
        except Exception as err:  # noqa: BLE001
            error = err
            status.update(label="Processing failed", state="error")
    if error:
        st.error(friendly_error(error))
        st.stop()

    st.session_state.kb = new_kb
    st.session_state.messages = []
    st.session_state.summary = None
    st.rerun()

# ----------------------------------------------------------------------------
# Knowledge base overview
# ----------------------------------------------------------------------------
ui_html(
    f"""
    <div class="stats">
      <div class="glass stat tilt"><div class="si">📄</div><div><div class="sv">{len(kb.file_table)}</div><div class="sl">Documents</div></div></div>
      <div class="glass stat tilt"><div class="si">🧩</div><div><div class="sv">{len(kb.chunks):,}</div><div class="sl">Chunks</div></div></div>
      <div class="glass stat tilt"><div class="si">✍️</div><div><div class="sv">{kb.words:,}</div><div class="sl">Words</div></div></div>
      <div class="glass stat tilt"><div class="si">🤖</div><div><div class="sv sm">{provider}</div><div class="sl">Provider</div></div></div>
    </div>
    """
)

for name, reason in kb.skipped:
    st.warning(f"Skipped **{name}**: {reason}")
if len(kb.chunks) > 2500:
    st.warning("Large knowledge base: indexing and answers may be slower and cost more.")

with st.expander("📚 Knowledge base details", expanded=False):
    st.dataframe(pd.DataFrame(kb.file_table), use_container_width=True, hide_index=True)
    if st.button("📝 Summarize all documents"):
        with st.spinner("Writing summary…"):
            try:
                llm = get_llm(provider, model_name, api_key, temperature)
                st.session_state.summary = summarize(llm, kb)
            except Exception as err:  # noqa: BLE001
                st.error(friendly_error(err))
    if st.session_state.summary:
        st.markdown(st.session_state.summary)

# ----------------------------------------------------------------------------
# Chat
# ----------------------------------------------------------------------------
if not st.session_state.messages:
    ui_html('<div class="section-title" style="margin-top:.6rem">💡 Try asking</div>')
    cols = st.columns(len(SAMPLE_QUESTIONS))
    for i, question in enumerate(SAMPLE_QUESTIONS):
        if cols[i].button(question, key=f"sample_{i}", use_container_width=True):
            st.session_state.pending_prompt = question
            st.rerun()

AVATARS = {"user": "👤", "assistant": "📚"}

for msg in st.session_state.messages:
    with st.chat_message(msg["role"], avatar=AVATARS[msg["role"]]):
        st.markdown(msg["content"])
        if msg["role"] == "assistant":
            render_sources(msg.get("sources", []))

prompt = st.chat_input("Ask anything about your documents…")
if not prompt:
    prompt = st.session_state.pop("pending_prompt", None)

if prompt:
    history = list(st.session_state.messages)
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user", avatar=AVATARS["user"]):
        st.markdown(prompt)

    answer, sources = None, []
    with st.chat_message("assistant", avatar=AVATARS["assistant"]):
        try:
            llm = get_llm(provider, model_name, api_key, temperature)
            with st.spinner("Searching your documents…"):
                standalone = condense_question(llm, history, prompt)
                docs = retrieve(kb, standalone, top_k, use_mmr)
            answer = st.write_stream(answer_stream(llm, prompt, docs, history))
            sources = pack_sources(docs)
            render_sources(sources)
        except Exception as err:  # noqa: BLE001
            st.error(friendly_error(err))

    if answer and answer.strip():
        st.session_state.messages.append({"role": "assistant", "content": answer, "sources": sources})
        st.rerun()
    else:
        st.session_state.messages.pop()  # drop the unanswered question