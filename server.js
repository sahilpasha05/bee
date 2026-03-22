const express = require('express');
const cors = require('cors');
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

app.get('/', (req, res) => res.json({ status: 'ok', service: 'n8n-doc-microservice' }));
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ── /extract-text ─────────────────────────────────────────────────────────────
app.post('/extract-text', (req, res) => {
    const { fileExt, base64Data, executionId } = req.body;
    if (!base64Data) return res.json({ stdout: '', stderr: 'No data provided', returncode: 1 });

    const inputPath = `/tmp/brief_${executionId}.${fileExt}`;

    try {
        // Write file from base64
        fs.writeFileSync(inputPath, Buffer.from(base64Data, 'base64'));

        let stdout = '';

        if (fileExt === 'pdf') {
            stdout = execSync(`pdftotext -layout "${inputPath}" -`, { 
                encoding: 'utf8', 
                maxBuffer: 50 * 1024 * 1024 
            });
        } else if (fileExt === 'docx') {
            // Write python script to file to avoid shell escaping issues
            const pyScript = `/tmp/extract_${executionId}.py`;
            fs.writeFileSync(pyScript, `
import sys
from docx import Document
doc = Document('${inputPath}')
lines = [p.text for p in doc.paragraphs if p.text.strip()]
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if cell.text.strip():
                lines.append(cell.text.strip())
print('\\n'.join(lines))
`);
            stdout = execSync(`python3 ${pyScript}`, { 
                encoding: 'utf8', 
                maxBuffer: 50 * 1024 * 1024 
            });
            try { fs.unlinkSync(pyScript); } catch(e) {}
        } else if (fileExt === 'pptx') {
            const pyScript = `/tmp/extract_${executionId}.py`;
            fs.writeFileSync(pyScript, `
from pptx import Presentation
prs = Presentation('${inputPath}')
lines = []
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t: lines.append(t)
print('\\n'.join(lines))
`);
            stdout = execSync(`python3 ${pyScript}`, { 
                encoding: 'utf8', 
                maxBuffer: 50 * 1024 * 1024 
            });
            try { fs.unlinkSync(pyScript); } catch(e) {}
        } else if (fileExt === 'txt') {
            stdout = Buffer.from(base64Data, 'base64').toString('utf8');
        }

        // Cleanup input file
        try { fs.unlinkSync(inputPath); } catch(e) {}

        res.json({ stdout: (stdout || '').trim(), stderr: '', returncode: 0 });
    } catch (e) {
        try { fs.unlinkSync(inputPath); } catch(e2) {}
        res.json({ stdout: '', stderr: e.message || String(e), returncode: 1 });
    }
});

// ── /deps ─────────────────────────────────────────────────────────────────────
app.post('/deps', (req, res) => {
    res.json({ stdout: 'DEPS_OK', stderr: '', returncode: 0 });
});

// ── /generate-charts ──────────────────────────────────────────────────────────
app.post('/generate-charts', (req, res) => {
    const { executionId, docSections, workableTask, fullDocumentText } = req.body;
    const workDir = `/tmp/charts_${executionId}`;
    fs.mkdirSync(workDir, { recursive: true });

    fs.writeFileSync(`${workDir}/sections.json`, JSON.stringify(docSections || []));
    fs.writeFileSync(`${workDir}/task.json`, JSON.stringify(workableTask || {}));
    fs.writeFileSync(`${workDir}/text.json`, JSON.stringify({ text: (fullDocumentText || '').substring(0, 5000) }));

    const pyScript = `${workDir}/script.py`;
    fs.writeFileSync(pyScript, `
import sys, os, json, base64, re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

exec_id = "${executionId}"
work_dir = "${workDir}"
os.makedirs(work_dir, exist_ok=True)

try:
    with open(f"{work_dir}/sections.json") as f: sections = json.load(f)
    with open(f"{work_dir}/task.json") as f: task = json.load(f)
    with open(f"{work_dir}/text.json") as f: full_text = json.load(f)["text"]
except:
    sections = []; task = {}; full_text = ""

charts_generated = []
task_str = str(task).lower()
sections_str = str(sections).lower()

has_resource = any(k in task_str + sections_str for k in ['resource','gantt','early start','late start','labourer','activity schedule'])
has_bsc = any(k in task_str for k in ['balanced scorecard','bsc'])
has_swot = any(k in task_str for k in ['swot','strengths','weaknesses'])
has_pestel = any(k in task_str for k in ['pestel','pestle'])
has_stakeholder = 'stakeholder' in task_str
has_cost = any(k in task_str for k in ['cost overrun','escalat'])
has_bmc = any(k in task_str for k in ['business model','value proposition','bmc'])

if has_swot:
    fig,axes = plt.subplots(2,2,figsize=(14,9))
    fig.patch.set_facecolor('#F8F9FA')
    swot_data = {
        ('Strengths','#27AE60',axes[0,0]):['Core project management capability','Strong stakeholder engagement','Established methodologies','Clear governance structure'],
        ('Weaknesses','#E74C3C',axes[0,1]):['Resource constraints','Communication gaps','Scope management issues','Risk underestimation'],
        ('Opportunities','#2980B9',axes[1,0]):['Technology adoption','Lessons learned application','Industry best practice','Regulatory alignment'],
        ('Threats','#E67E22',axes[1,1]):['External risk factors','Stakeholder resistance','Budget overruns','Schedule delays']
    }
    for (title,color,ax),items in swot_data.items():
        ax.set_facecolor(color+'18'); ax.set_xlim(0,1); ax.set_ylim(0,1)
        ax.set_xticks([]); ax.set_yticks([])
        for spine in ax.spines.values(): spine.set_edgecolor(color); spine.set_linewidth(2)
        ax.text(0.5,0.93,title,ha='center',va='top',fontsize=13,fontweight='bold',color=color,transform=ax.transAxes)
        ax.axhline(y=0.87,color=color,linewidth=1.5,alpha=0.4)
        for i,item in enumerate(items):
            ax.text(0.05,0.76-i*0.145,f'• {item}',ha='left',va='top',fontsize=9,color='#2C3E50',transform=ax.transAxes)
    plt.suptitle('Figure: SWOT Analysis',fontsize=13,fontweight='bold',y=1.01)
    plt.tight_layout()
    path = f"{work_dir}/swot.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#F8F9FA'); plt.close()
    charts_generated.append(('swot',path))

if has_stakeholder:
    fig,ax = plt.subplots(figsize=(13,8))
    ax.set_xlim(0,10); ax.set_ylim(0,8); ax.axis('off')
    fig.patch.set_facecolor('#FDFEFE')
    for r,color,alpha in [(3.5,'#1F3864',0.08),(2.5,'#2E86AB',0.08),(1.5,'#44BBA4',0.08)]:
        circle = plt.Circle((5,4),r,color=color,alpha=alpha,zorder=1)
        ax.add_patch(circle)
    ax.text(5,4,'PROJECT\\nCORE',ha='center',va='center',fontsize=10,fontweight='bold',color='white',zorder=5,
            bbox=dict(boxstyle='round,pad=0.4',facecolor='#1F3864',edgecolor='white'))
    stakeholders = [(5,7.2,'Client / Sponsor','#C73E1D'),(5,0.8,'End Users','#C73E1D'),
        (1.8,6.2,'Regulatory Bodies','#E67E22'),(8.2,6.2,'Investors / Funders','#E67E22'),
        (1.0,4.0,'Contractors','#2E86AB'),(9.0,4.0,'Government','#2E86AB'),
        (1.8,1.8,'Community','#44BBA4'),(8.2,1.8,'Media / Press','#44BBA4')]
    for x,y,name,color in stakeholders:
        ax.annotate('',xy=(5,4),xytext=(x,y),arrowprops=dict(arrowstyle='-',color=color,alpha=0.3,lw=1))
        ax.text(x,y,name,ha='center',va='center',fontsize=8,
                bbox=dict(boxstyle='round,pad=0.3',facecolor=color+'22',edgecolor=color,linewidth=1.2))
    ax.set_title("Figure: Stakeholder Map\\nSource: Author's own analysis",fontsize=10,fontweight='bold',pad=10)
    plt.tight_layout()
    path = f"{work_dir}/stakeholder.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('stakeholder',path))

if has_pestel:
    fig,ax = plt.subplots(figsize=(13,8))
    ax.set_xlim(0,13); ax.set_ylim(0,8); ax.axis('off')
    pestel = [
        ('P\\nPolitical','#2E86AB',1.0,['Policy changes','Trade regulations','Political stability','Government incentives']),
        ('E\\nEconomic','#27AE60',3.2,['Market growth','Inflation rates','Consumer spending','Currency risk']),
        ('S\\nSocial','#F18F01',5.4,['Demographic shifts','Cultural trends','Consumer behaviour','Social responsibility']),
        ('T\\nTechnological','#9B59B6',7.6,['Digital disruption','Innovation cycles','Automation','Data analytics']),
        ('E\\nEnvironmental','#16A085',9.8,['Climate change','Sustainability','Carbon targets','Resource scarcity']),
        ('L\\nLegal','#E74C3C',12.0,['Compliance','IP protection','Consumer law','Regulatory standards']),
    ]
    for title,color,x,items in pestel:
        rect = mpatches.FancyBboxPatch((x-0.9,0.3),1.9,7.4,boxstyle='round,pad=0.15',facecolor=color+'20',edgecolor=color,linewidth=2)
        ax.add_patch(rect)
        ax.text(x,7.3,title,ha='center',va='top',fontsize=9,fontweight='bold',color=color)
        ax.axhline(y=6.6,xmin=(x-0.9)/13,xmax=(x+1.0)/13,color=color,alpha=0.4,lw=1)
        for i,item in enumerate(items):
            ax.text(x,6.1-i*1.2,item,ha='center',va='top',fontsize=7.5,color='#2C3E50')
    ax.set_title('Figure: PESTEL Analysis',fontsize=11,fontweight='bold',pad=10,color='#1F3864')
    plt.tight_layout()
    path = f"{work_dir}/pestel.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('pestel',path))

# Always generate a risk matrix for project management briefs
if 'risk' in task_str or 'risk' in sections_str:
    fig,ax = plt.subplots(figsize=(10,8))
    ax.set_xlim(0,5); ax.set_ylim(0,5)
    colors = [['#92D050','#92D050','#FFFF00','#FF0000','#FF0000'],
              ['#92D050','#FFFF00','#FFFF00','#FF0000','#FF0000'],
              ['#92D050','#FFFF00','#FFFF00','#FFA500','#FF0000'],
              ['#92D050','#92D050','#FFFF00','#FFA500','#FFA500'],
              ['#92D050','#92D050','#92D050','#FFFF00','#FFFF00']]
    for i in range(5):
        for j in range(5):
            rect = mpatches.Rectangle((j,4-i),1,1,facecolor=colors[i][j],edgecolor='white',linewidth=2,alpha=0.8)
            ax.add_patch(rect)
            ax.text(j+0.5,4-i+0.5,str((i+1)*(j+1)),ha='center',va='center',fontsize=12,fontweight='bold',color='#1F3864')
    risks = [(3.5,4.5,'Schedule delay','#1F3864'),(2.5,3.5,'Budget overrun','#1F3864'),
             (1.5,2.5,'Stakeholder conflict','#1F3864'),(4.5,1.5,'Technical failure','#1F3864'),
             (0.5,0.5,'Force majeure','#1F3864')]
    for x,y,label,color in risks:
        ax.plot(x,y,'o',markersize=12,color=color,zorder=5)
        ax.text(x+0.1,y+0.1,label,fontsize=7.5,color=color)
    ax.set_xlabel('Likelihood →',fontsize=11,fontweight='bold')
    ax.set_ylabel('Impact →',fontsize=11,fontweight='bold')
    ax.set_xticks([0.5,1.5,2.5,3.5,4.5])
    ax.set_xticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
    ax.set_yticks([0.5,1.5,2.5,3.5,4.5])
    ax.set_yticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
    ax.set_title("Figure: Risk Matrix\\nSource: Author's own analysis",fontsize=11,fontweight='bold',pad=12)
    plt.tight_layout()
    path = f"{work_dir}/risk_matrix.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('risk_matrix',path))

manifest = {}
for name,path in charts_generated:
    if os.path.exists(path):
        with open(path,'rb') as f:
            manifest[name] = {'path':path,'b64':base64.b64encode(f.read()).decode()}

with open(f"{work_dir}/manifest.json",'w') as f:
    json.dump({'charts':manifest,'count':len(manifest),'exec_id':exec_id},f)

print(f"CHARTS_DONE:{len(manifest)}")
for name in manifest:
    print(f"CHART:{name}")
`);

    try {
        const stdout = execSync(`python3 ${pyScript}`, {
            encoding: 'utf8',
            maxBuffer: 100 * 1024 * 1024,
            timeout: 120000
        });
        res.json({ stdout: stdout.trim(), stderr: '', returncode: 0 });
    } catch (e) {
        res.json({ stdout: e.stdout ? e.stdout.trim() : '', stderr: e.message || '', returncode: 1 });
    }
});

// ── /export-docx ──────────────────────────────────────────────────────────────
app.post('/export-docx', (req, res) => {
    const {
        executionId, studentName, studentId, programme, university,
        submissionDate, workableTask, totalWordCount, targetWordCount, docSections
    } = req.body;

    const workDir = `/tmp/charts_${executionId}`;
    const {
        Document, Packer, Paragraph, TextRun, ImageRun,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
        WidthType, VerticalAlign, PageNumber, PageBreak, TabStopType
    } = require('docx');

    let charts = {};
    try {
        if (fs.existsSync(`${workDir}/manifest.json`)) {
            const manifest = JSON.parse(fs.readFileSync(`${workDir}/manifest.json`, 'utf8'));
            charts = manifest.charts || {};
        }
    } catch(e) {}

    const FONT = 'Arial';
    const SZ = 24, SZS = 20, SZH1 = 28, SZH2 = 26, LINE = 360;

    const tr = (t, o={}) => new TextRun({
        text: String(t||''), font: FONT, size: o.size||SZ,
        bold: o.bold||false, italics: o.italic||false, color: o.color||'000000'
    });
    const blk = () => new Paragraph({ spacing:{line:LINE,before:0,after:0}, children:[tr('')] });

    function mkP(runs, o={}) {
        return new Paragraph({
            alignment: o.center?AlignmentType.CENTER:o.right?AlignmentType.RIGHT:
                       o.left?AlignmentType.LEFT:AlignmentType.JUSTIFIED,
            spacing:{line:LINE,before:o.before||0,after:o.after||160},
            children: Array.isArray(runs)?runs:[tr(runs,o)]
        });
    }

    function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{line:LINE,before:280,after:140},children:[tr(t,{bold:true,size:SZH1,color:'1F3864'})]})}
    function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{line:LINE,before:200,after:100},children:[tr(t,{bold:true,size:SZH2,color:'2C5282'})]})}

    function embedChart(name, wCm, hCm, caption) {
        if (!charts[name]) return [];
        try {
            const buf = Buffer.from(charts[name].b64, 'base64');
            return [
                new Paragraph({alignment:AlignmentType.CENTER,spacing:{line:LINE,before:60,after:0},
                    children:[new ImageRun({data:Uint8Array.from(buf),transformation:{
                        width:Math.round(wCm*360000/9144),
                        height:Math.round(hCm*360000/9144)
                    },type:'png'})]}),
                mkP([tr(caption,{size:SZS,italic:true,color:'444444'})],{center:true,before:20,after:60}),
                mkP([tr("Source: Author's own analysis.",{size:SZS,italic:true,color:'666666'})],{left:true,before:0,after:160})
            ];
        } catch(e) { return [mkP(`[Chart: ${name}]`,{italic:true})]; }
    }

    function getChartsForSection(title, content) {
        const text = (title + ' ' + content).toLowerCase();
        const injected = [];
        if (text.includes('risk') && charts['risk_matrix'])
            injected.push(['risk_matrix',14,10,'Figure: Risk Assessment Matrix']);
        if (text.includes('stakeholder') && charts['stakeholder'])
            injected.push(['stakeholder',14,8,'Figure: Stakeholder Map']);
        if (text.includes('swot') && charts['swot'])
            injected.push(['swot',15,9,'Figure: SWOT Analysis']);
        if (text.includes('pestel') && charts['pestel'])
            injected.push(['pestel',15,8,'Figure: PESTEL Analysis']);
        return injected;
    }

    try {
        const cover = [
            blk(),blk(),blk(),blk(),
            mkP([tr((university||'University').toUpperCase(),{bold:true,size:28,color:'1F3864'})],{center:true,after:40}),
            mkP([tr(programme||'Programme',{size:22})],{center:true,after:200}),
            blk(),blk(),
            mkP([tr((workableTask?.document_type||'Academic Report').toUpperCase(),{bold:true,size:32,color:'1F3864'})],{center:true,after:300}),
            blk(),blk(),blk(),
            mkP([tr('Student Name: '+(studentName||''),{size:22})],{center:true,after:20}),
            mkP([tr('Student ID: '+(studentId||''),{size:22})],{center:true,after:20}),
            mkP([tr('Programme: '+(programme||''),{size:22})],{center:true,after:20}),
            mkP([tr('Submission Date: '+(submissionDate||''),{size:22})],{center:true,after:20}),
            mkP([tr('Referencing: '+(workableTask?.referencing_style||'Harvard'),{size:22})],{center:true,after:20}),
            blk(),
            mkP([tr('Word Count: '+(totalWordCount||0)+' words (Target: '+(targetWordCount||0)+')',{size:22,bold:true})],{center:true}),
            new Paragraph({children:[new PageBreak()]})
        ];

        const tocItems = (docSections||[]).map((s,i) => {
            const title = (s.sectionNumber?s.sectionNumber+'  ':'')+s.title;
            return new Paragraph({
                spacing:{line:LINE,before:40,after:40},
                children:[tr(title,{size:SZS})]
            });
        });
        const toc = [
            h1('Table of Contents'),
            ...tocItems,
            new Paragraph({children:[new PageBreak()]})
        ];

        const mainContent = [];
        for (const sec of (docSections||[])) {
            const titleLower = (sec.title||'').toLowerCase();
            if (titleLower.includes('references') || titleLower.includes('bibliography')) {
                mainContent.push(new Paragraph({children:[new PageBreak()]}));
            }
            mainContent.push(h1((sec.sectionNumber?sec.sectionNumber+'  ':'')+sec.title));
            const paras = (sec.content||'').split(/\n\n+/).filter(p=>p.trim());
            for (const para of paras) {
                const t = para.trim();
                if (!t) continue;
                if (t.match(/^\d+\.\d+\s+\w/) && t.length < 120) {
                    mainContent.push(h2(t));
                } else {
                    mainContent.push(mkP(t));
                }
            }
            const chartMatches = getChartsForSection(sec.title||'', sec.content||'');
            for (const [name, wCm, hCm, caption] of chartMatches) {
                mainContent.push(...embedChart(name, wCm, hCm, caption));
            }
            mainContent.push(blk());
        }

        mainContent.push(
            new Paragraph({
                border:{top:{style:BorderStyle.SINGLE,size:6,color:'CCCCCC',space:4}},
                spacing:{line:LINE,before:80,after:60},
                children:[tr(`Word Count: ${totalWordCount||0} words (Target: ${targetWordCount||0}). Excludes cover page, TOC, tables, figures, and references.`,{size:SZS,italic:true,color:'666666'})]
            })
        );

        const doc = new Document({
            styles:{
                default:{document:{run:{font:FONT,size:SZ}}}
            },
            sections:[{
                properties:{page:{size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
                headers:{default:new Header({children:[new Paragraph({
                    alignment:AlignmentType.RIGHT,
                    border:{bottom:{style:BorderStyle.SINGLE,size:6,color:'1F3864',space:4}},
                    spacing:{before:0,after:80},
                    children:[tr(`${university||'University'}  |  ${workableTask?.document_type||'Report'}  |  ${studentName||''}`,{size:18,italic:true,color:'555555'})]
                })]})},
                footers:{default:new Footer({children:[new Paragraph({
                    alignment:AlignmentType.CENTER,
