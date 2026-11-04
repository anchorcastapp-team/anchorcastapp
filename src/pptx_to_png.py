#!/usr/bin/env python3
"""AnchorCast fallback converter. JSON on stdout; no downloads or pip installs.
Native PowerPoint / LibreOffice are attempted by presentation-import.js first.
The PPTX fallback is basic static text/image rendering, NOT a full Office renderer.
"""
import json, os, sys, io

def error(message, code):
    return {"success": False, "error": message, "code": code}

def convert(source, output):
    os.makedirs(output, exist_ok=True)
    ext = os.path.splitext(source)[1].lower()
    if ext == '.pdf':
        try:
            import fitz
        except ImportError:
            return error('Python is installed, but PDF rendering requires PyMuPDF in this interpreter, or the Poppler pdftoppm program.', 'missing-pdf-renderer')
        doc = fitz.open(source)
        try:
            if doc.needs_pass:
                return error('The PDF is password protected. Export an unlocked copy.', 'encrypted-file')
            if not 0 < len(doc) <= 1000:
                return error('The PDF has no pages or exceeds the 1,000-page import limit.', 'page-limit')
            slides = []
            for i, page in enumerate(doc):
                scale = 1920 / max(page.rect.width, page.rect.height)
                target = os.path.join(output, 'slide-%d.png' % (i + 1))
                page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).save(target)
                slides.append({'index': i, 'imagePath': target})
            return {'success': True, 'slides': slides, 'count': len(slides), 'method': 'python-pymupdf'}
        finally:
            doc.close()
    if ext != '.pptx':
        return error('Python is available, but this fallback supports .pptx only. Import .ppt/.odp using LibreOffice, or export to PPTX/PDF first.', 'needs-office-converter')
    try:
        from pptx import Presentation
        from PIL import Image, ImageDraw, ImageFont
    except ImportError as exc:
        return error('Python is installed, but this interpreter is missing presentation packages. Install python-pptx and Pillow for this interpreter. Details: ' + str(exc), 'missing-python-packages')
    prs = Presentation(source)
    if not 0 < len(prs.slides) <= 1000:
        return error('The PPTX has no slides or exceeds the 1,000-slide import limit.', 'slide-limit')
    scale = 1920 / max(int(prs.slide_width), int(prs.slide_height))
    width, height = max(1,round(prs.slide_width*scale)), max(1,round(prs.slide_height*scale))
    font_paths = [os.path.join(os.environ.get('WINDIR','C:/Windows'),'Fonts','arial.ttf'),
                  '/System/Library/Fonts/Supplemental/Arial.ttf', '/Library/Fonts/Arial.ttf',
                  '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']
    def font(size):
        for p in font_paths:
            try: return ImageFont.truetype(p,size)
            except OSError: pass
        return ImageFont.load_default()
    slides=[]
    for index, sl in enumerate(prs.slides):
        bg=(255,255,255)
        try:
            rgb=sl.background.fill.fore_color.rgb
            if rgb is not None: bg=tuple(rgb)
        except (AttributeError,TypeError): pass
        im=Image.new('RGB',(width,height),bg);draw=ImageDraw.Draw(im)
        for shape in sl.shapes:
            x,y,w,h=[round(v*scale) for v in [shape.left,shape.top,shape.width,shape.height]]
            if getattr(shape,'shape_type',None)==13:
                try:
                    pic=Image.open(io.BytesIO(shape.image.blob)).convert('RGBA')
                    pic.thumbnail((max(w,1),max(h,1)))
                    im.paste(pic,(x,y),pic)
                except Exception: pass
            if not getattr(shape,'has_text_frame',False):continue
            for para in shape.text_frame.paragraphs:
                if not para.text.strip():continue
                pts=24;colour=(0,0,0)
                for run in para.runs:
                    if run.font.size:pts=run.font.size.pt
                    try:
                        rgb=run.font.color.rgb
                        if rgb is not None:colour=tuple(rgb)
                    except (AttributeError,TypeError):pass
                # Office coordinates use EMUs; one point is 12,700 EMUs.
                px=max(10,round(pts*12700*scale));f=font(px)
                line=''
                for word in para.text.split():
                    candidate=(line+' '+word).strip()
                    if line and draw.textbbox((0,0),candidate,font=f)[2]>max(w,1):
                        draw.text((x,y),line,fill=colour,font=f);y+=round(px*1.3);line=word
                    else:line=candidate
                if line:draw.text((x,y),line,fill=colour,font=f);y+=round(px*1.3)
        target=os.path.join(output,'slide-%d.png'%(index+1));im.save(target,'PNG');slides.append({'index':index,'imagePath':target})
    return {'success':True,'slides':slides,'count':len(slides),'method':'python-basic',
            'warning':'Basic PPTX rendering was used. Complex layouts, charts, animations and some styling may differ. For faithful static slides, export to PDF from PowerPoint or use LibreOffice.'}

if __name__ == '__main__':
    try:
        result=convert(sys.argv[1],sys.argv[2]) if len(sys.argv)==3 else error('Usage: pptx_to_png.py input output', 'arguments')
    except Exception as exc:
        result=error('The file could not be converted: '+str(exc), 'conversion-failed')
    print(json.dumps(result))
