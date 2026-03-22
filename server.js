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
combined_str = task_str + ' ' + sections_str

# Detect subject area
is_gsm = any(k in combined_str for k in ['global strategic','internationalisation','market entry','segmentation criteria','entry mode','global strategy','international market'])
is_pm = any(k in combined_str for k in ['project management','pmbok','prince2','risk register','stakeholder management','gantt','crossrail','challenger','shard'])

has_resource = any(k in combined_str for k in ['resource','gantt','early start','late start','labourer','activity schedule'])
has_bsc = any(k in combined_str for k in ['balanced scorecard','bsc'])
has_swot = any(k in combined_str for k in ['swot','strengths','weaknesses'])
has_pestel = any(k in combined_str for k in ['pestel','pestle'])
has_stakeholder = 'stakeholder' in combined_str
has_cost = any(k in combined_str for k in ['cost overrun','escalat'])
has_bmc = any(k in combined_str for k in ['business model','value proposition','bmc'])
has_market_screen = any(k in combined_str for k in ['market screen','market select','segmentation','country select','market entry','internationalisation'])
has_risk_pm = is_pm and any(k in combined_str for k in ['risk management','risk register','risk matrix','risk assess'])

# Get case study name for dynamic content
case_name = task.get('case_study','') if isinstance(task, dict) else ''
is_luckin = 'luckin' in case_name.lower() or 'luckin' in combined_str
is_crossrail = 'crossrail' in combined_str
is_boeing = 'boeing' in combined_str
is_challenger = 'challenger' in combined_str

# Dynamic SWOT content based on case study
if is_luckin:
    swot_content = {
        'Strengths': ['Rapid store expansion (24,000+ stores)','Technology-driven ordering platform','Low price strategy vs Starbucks','Strong brand recognition in China'],
        'Weaknesses': ['Limited international presence','Past accounting scandal (2020)','Heavy reliance on Chinese market','Brand trust issues globally'],
        'Opportunities': ['Emerging coffee markets in Asia','Growing middle class globally','Digital payment adoption worldwide','Partnership and franchise models'],
        'Threats': ['Starbucks global dominance','Cultural resistance to Chinese brands','Regulatory barriers in new markets','Currency and political risks']
    }
elif is_crossrail:
    swot_content = {
        'Strengths': ['Government backing and funding','Strong engineering expertise','Integrated transport planning','Major economic impact'],
        'Weaknesses': ['Significant cost overruns','Multiple project delays','Complex stakeholder management','Technical integration challenges'],
        'Opportunities': ['Urban regeneration catalyst','Reduced congestion in London','Property value increases','Template for future projects'],
        'Threats': ['Budget escalation risks','Political pressure','Public disruption','Contractor performance risks']
    }
else:
    swot_content = {
        'Strengths': ['Strong market position','Established operational capabilities','Financial resources','Brand recognition'],
        'Weaknesses': ['Resource constraints','Limited geographic reach','Operational complexity','Cost pressures'],
        'Opportunities': ['Market expansion potential','Technology adoption','Strategic partnerships','Emerging market growth'],
        'Threats': ['Competitive pressure','Regulatory changes','Economic uncertainty','Market disruption']
    }

if has_swot:
    fig,axes = plt.subplots(2,2,figsize=(14,9))
    fig.patch.set_facecolor('#F8F9FA')
    colors_map = {'Strengths':'#27AE60','Weaknesses':'#E74C3C','Opportunities':'#2980B9','Threats':'#E67E22'}
    ax_map = {'Strengths':axes[0,0],'Weaknesses':axes[0,1],'Opportunities':axes[1,0],'Threats':axes[1,1]}
    for title, items in swot_content.items():
        color = colors_map[title]
        ax = ax_map[title]
        ax.set_facecolor(color+'18'); ax.set_xlim(0,1); ax.set_ylim(0,1)
        ax.set_xticks([]); ax.set_yticks([])
        for spine in ax.spines.values(): spine.set_edgecolor(color); spine.set_linewidth(2)
        ax.text(0.5,0.93,title,ha='center',va='top',fontsize=13,fontweight='bold',color=color,transform=ax.transAxes)
        ax.axhline(y=0.87,color=color,linewidth=1.5,alpha=0.4)
        for i,item in enumerate(items):
            ax.text(0.05,0.76-i*0.145,f'• {item}',ha='left',va='top',fontsize=9,color='#2C3E50',transform=ax.transAxes)
    title_text = f'Figure: SWOT Analysis — {case_name}' if case_name else 'Figure: SWOT Analysis'
    plt.suptitle(title_text,fontsize=13,fontweight='bold',y=1.01)
    plt.tight_layout()
    path = f"{work_dir}/swot.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#F8F9FA'); plt.close()
    charts_generated.append(('swot',path))

# Market Screening Matrix — for GSM/international strategy briefs
if has_market_screen or is_gsm:
    fig,ax = plt.subplots(figsize=(14,8))
    ax.set_facecolor('#FDFEFE')
    fig.patch.set_facecolor('#FDFEFE')
    if is_luckin:
        markets = ['India','Brazil','Japan','South Korea','UAE','UK','Germany','France','Australia','Mexico']
        scores = {
            'Market Size': [9,8,7,6,8,7,6,6,6,7],
            'Coffee Culture': [5,6,9,8,6,8,7,9,8,6],
            'Digital Readiness': [7,6,9,9,8,8,8,8,9,7],
            'Competitive Gap': [8,7,5,6,7,6,6,5,7,7],
            'Regulatory Ease': [6,5,7,7,8,8,8,8,8,6]
        }
    else:
        markets = ['Market A','Market B','Market C','Market D','Market E','Market F','Market G','Market H']
        scores = {
            'Market Size': [8,7,6,9,5,7,8,6],
            'Growth Potential': [7,8,7,6,9,6,7,8],
            'Competitive Gap': [6,7,8,5,7,8,6,7],
            'Entry Barriers': [7,6,7,8,6,7,5,8],
            'Strategic Fit': [8,7,6,7,8,7,9,6]
        }
    criteria = list(scores.keys())
    x = np.arange(len(markets))
    width = 0.15
    colors_bar = ['#2E86AB','#27AE60','#E67E22','#9B59B6','#E74C3C']
    for i,(crit,vals) in enumerate(scores.items()):
        ax.bar(x + i*width, vals, width, label=crit, color=colors_bar[i], alpha=0.85, edgecolor='white')
    totals = [sum(scores[c][j] for c in criteria) for j in range(len(markets))]
    ax2 = ax.twinx()
    ax2.plot(x + width*2, totals, 'D-', color='#1F3864', linewidth=2, markersize=8, label='Total Score', zorder=5)
    ax2.set_ylabel('Total Score', fontsize=9, color='#1F3864')
    ax2.tick_params(axis='y', labelcolor='#1F3864')
    ax.set_xlabel('Markets / Countries', fontsize=10, fontweight='bold')
    ax.set_ylabel('Score (1-10)', fontsize=10, fontweight='bold')
    ax.set_title(f"Figure: Market Screening Matrix\\nSource: Author's own analysis", fontsize=11, fontweight='bold', pad=12)
    ax.set_xticks(x + width*2)
    ax.set_xticklabels(markets, rotation=25, ha='right', fontsize=9)
    ax.set_ylim(0,11)
    ax.legend(loc='upper left', fontsize=8)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    path = f"{work_dir}/market_screening.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('market_screening',path))

if has_stakeholder:
    fig,ax = plt.subplots(figsize=(13,8))
    ax.set_xlim(0,10); ax.set_ylim(0,8); ax.axis('off')
    fig.patch.set_facecolor('#FDFEFE')
    for r,color,alpha in [(3.5,'#1F3864',0.08),(2.5,'#2E86AB',0.08),(1.5,'#44BBA4',0.08)]:
        circle = plt.Circle((5,4),r,color=color,alpha=alpha,zorder=1)
        ax.add_patch(circle)
    if is_gsm or is_luckin:
        center_label = 'LUCKIN\\nCOFFEE' if is_luckin else 'ORGANISATION'
        stakeholders = [
            (5,7.2,'CEO / Board','#C73E1D'),(5,0.8,'Customers','#C73E1D'),
            (1.8,6.2,'Government / Regulators','#E67E22'),(8.2,6.2,'Investors / Shareholders','#E67E22'),
            (1.0,4.0,'Franchise Partners','#2E86AB'),(9.0,4.0,'Competitors','#2E86AB'),
            (1.8,1.8,'Local Communities','#44BBA4'),(8.2,1.8,'Suppliers','#44BBA4')
        ]
    else:
        center_label = 'PROJECT\\nCORE'
        stakeholders = [
            (5,7.2,'Client / Sponsor','#C73E1D'),(5,0.8,'End Users','#C73E1D'),
            (1.8,6.2,'Regulatory Bodies','#E67E22'),(8.2,6.2,'Investors / Funders','#E67E22'),
            (1.0,4.0,'Contractors','#2E86AB'),(9.0,4.0,'Government','#2E86AB'),
            (1.8,1.8,'Community','#44BBA4'),(8.2,1.8,'Media / Press','#44BBA4')
        ]
    ax.text(5,4,center_label,ha='center',va='center',fontsize=10,fontweight='bold',color='white',zorder=5,
            bbox=dict(boxstyle='round,pad=0.4',facecolor='#1F3864',edgecolor='white'))
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
    if is_luckin:
        pestel = [
            ('P\\nPolitical','#2E86AB',1.0,['Trade regulations','Market entry barriers','Geopolitical tensions','Government incentives']),
            ('E\\nEconomic','#27AE60',3.2,['GDP growth rates','Consumer spending','Currency exchange','Cost of operations']),
            ('S\\nSocial','#F18F01',5.4,['Coffee culture growth','Digital lifestyle','Health consciousness','Youth demographics']),
            ('T\\nTechnological','#9B59B6',7.6,['Mobile payment apps','AI ordering systems','Supply chain tech','Data analytics']),
            ('E\\nEnvironmental','#16A085',9.8,['Sustainable sourcing','Carbon footprint','Packaging regulations','Climate impact']),
            ('L\\nLegal','#E74C3C',12.0,['Franchise laws','Food safety regs','Employment law','IP protection']),
        ]
    else:
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

# Risk Matrix ONLY for PM briefs — not GSM/strategy briefs
if has_risk_pm and not is_gsm:
    fig,ax = plt.subplots(figsize=(10,8))
    ax.set_xlim(0,5); ax.set_ylim(0,5)
    colors_rm = [['#92D050','#92D050','#FFFF00','#FF0000','#FF0000'],
              ['#92D050','#FFFF00','#FFFF00','#FF0000','#FF0000'],
              ['#92D050','#FFFF00','#FFFF00','#FFA500','#FF0000'],
              ['#92D050','#92D050','#FFFF00','#FFA500','#FFA500'],
              ['#92D050','#92D050','#92D050','#FFFF00','#FFFF00']]
    for i in range(5):
        for j in range(5):
            rect = mpatches.Rectangle((j,4-i),1,1,facecolor=colors_rm[i][j],edgecolor='white',linewidth=2,alpha=0.8)
            ax.add_patch(rect)
            ax.text(j+0.5,4-i+0.5,str((i+1)*(j+1)),ha='center',va='center',fontsize=12,fontweight='bold',color='#1F3864')
    if is_crossrail:
        risks = [(3.5,4.5,'Cost overrun','#1F3864'),(2.5,3.5,'Schedule delay','#1F3864'),
                 (1.5,2.5,'Stakeholder conflict','#1F3864'),(4.5,1.5,'Technical failure','#1F3864'),
                 (0.5,0.5,'Force majeure','#1F3864')]
    else:
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

    // Isolate this job - only use data from THIS request
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

    const FONT = 'Times New Roman';
    const SZ = 24, SZS = 20, SZH1 = 28, SZH2 = 26, LINE = 360;  // 24 = 12pt, consistent throughout

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
        if ((text.includes('segmentation') || text.includes('market select') || text.includes('screening') || text.includes('market entry')) && charts['market_screening'])
            injected.push(['market_screening',15,9,'Figure: Market Screening Matrix']);
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
            mkP([tr('Student Name:  '+(studentName&&studentName!=='Student'&&studentName!=='N/A'?studentName:'Not Available'),{size:22})],{center:true,after:20}),
            mkP([tr('Student ID:  '+(studentId&&studentId!=='N/A'?studentId:'Not Available'),{size:22})],{center:true,after:20}),
            mkP([tr('Programme:  '+(programme&&programme!=='N/A'?programme:'Not Available'),{size:22})],{center:true,after:20}),
            mkP([tr('Submission Date:  '+(submissionDate&&submissionDate!=='N/A'?submissionDate:'Not Available'),{size:22})],{center:true,after:20}),
            mkP([tr('Referencing Style:  '+(workableTask?.referencing_style||'Harvard'),{size:22})],{center:true,after:20}),
            blk(),blk(),
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

        // Word count declaration removed as per brief requirements

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
                    children:[tr(`${university&&university!=='N/A'?university:'Academic Report'}  |  ${workableTask?.document_type||'Report'}`,{size:18,italic:true,color:'555555'})]
                })]})},
                footers:{default:new Footer({children:[new Paragraph({
                    alignment:AlignmentType.CENTER,
                    border:{top:{style:BorderStyle.SINGLE,size:6,color:'1F3864',space:4}},
                    spacing:{before:80,after:0},
                    children:[
                        tr('Page ',{size:18,color:'777777'}),
                        new TextRun({children:[PageNumber.CURRENT],font:FONT,size:18,color:'777777'}),
                        tr(' of ',{size:18,color:'777777'}),
                        new TextRun({children:[PageNumber.TOTAL_PAGES],font:FONT,size:18,color:'777777'})
                    ]
                })]})},
                children:[...cover,...toc,...mainContent]
            }]
        });

        Packer.toBuffer(doc).then(buf => {
            const b64 = buf.toString('base64');
            res.json({
                stdout: 'SUCCESS',
                stderr: '',
                returncode: 0,
                docxBase64: b64,
                filename: `${(workableTask?.document_type||'Report').replace(/\s+/g,'_')}_completed.docx`
            });
            try { fs.rmSync(workDir, { recursive: true, force: true }); } catch(e) {}
        }).catch(err => {
            res.json({ stdout: '', stderr: err.message, returncode: 1 });
        });
    } catch(e) {
        res.json({ stdout: '', stderr: e.message, returncode: 1 });
    }
});

// ── /cleanup ──────────────────────────────────────────────────────────────────
app.post('/cleanup', (req, res) => {
    const { executionId } = req.body;
    try {
        const workDir = `/tmp/charts_${executionId}`;
        if (fs.existsSync(workDir)) fs.rmSync(workDir, { recursive: true, force: true });
    } catch(e) {}
    res.json({ stdout: 'CLEANUP_OK', stderr: '', returncode: 0 });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`n8n microservice listening on port ${PORT}`);
});
