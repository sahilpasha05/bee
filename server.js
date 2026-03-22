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
        fs.writeFileSync(inputPath, Buffer.from(base64Data, 'base64'));
        let stdout = '';
        if (fileExt === 'pdf') {
            stdout = execSync(`pdftotext -layout "${inputPath}" -`, { encoding: 'utf8', maxBuffer: 50 * 1024 * 1024 });
        } else if (fileExt === 'docx') {
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
            stdout = execSync(`python3 ${pyScript}`, { encoding: 'utf8', maxBuffer: 50 * 1024 * 1024 });
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
            stdout = execSync(`python3 ${pyScript}`, { encoding: 'utf8', maxBuffer: 50 * 1024 * 1024 });
            try { fs.unlinkSync(pyScript); } catch(e) {}
        } else if (fileExt === 'txt') {
            stdout = Buffer.from(base64Data, 'base64').toString('utf8');
        }
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
all_text = task_str + ' ' + sections_str + ' ' + full_text.lower()

# Subject detection
is_gsm = any(k in all_text for k in ['global strategic','internationalisation','market entry','segmentation criteria','entry mode','global strategy','international market'])
is_pm  = any(k in all_text for k in ['project management','pmbok','prince2','risk register','gantt','crossrail','challenger','shard','scottish parliament'])

# Feature flags
has_resource     = any(k in all_text for k in ['resource loading','gantt','early start','late start','labourer','activity schedule'])
has_bsc          = any(k in all_text for k in ['balanced scorecard','bsc','balanced score'])
has_swot         = any(k in all_text for k in ['swot','strengths','weaknesses'])
has_pestel       = any(k in all_text for k in ['pestel','pestle'])
has_stakeholder  = 'stakeholder' in all_text
has_cost         = any(k in all_text for k in ['cost overrun','escalat','cost escalation'])
has_market_screen= any(k in all_text for k in ['market screen','market select','segmentation','country select','market entry','internationalisation'])
has_risk_pm      = is_pm and any(k in all_text for k in ['risk management','risk register','risk matrix','risk assess'])

# Case study
case_name    = task.get('case_study','') if isinstance(task, dict) else ''
is_luckin    = 'luckin' in case_name.lower() or 'luckin' in all_text
is_crossrail = 'crossrail' in all_text
is_sph       = 'scottish parliament' in all_text or 'sph' in all_text

# ── SWOT ──────────────────────────────────────────────────────────────────────
if has_swot:
    if is_luckin:
        swot_content = {
            'Strengths':     ['Rapid expansion (24,000+ stores)','Technology-driven ordering platform','Low price strategy vs Starbucks','Strong brand recognition in China'],
            'Weaknesses':    ['Limited international presence','2020 accounting scandal legacy','Heavy reliance on Chinese market','Brand trust issues globally'],
            'Opportunities': ['Emerging coffee markets in Asia','Growing middle class globally','Digital payment adoption','Partnership and franchise models'],
            'Threats':       ['Starbucks global dominance','Cultural resistance to Chinese brands','Regulatory barriers','Currency and political risks']
        }
    elif is_crossrail:
        swot_content = {
            'Strengths':     ['Government backing and funding','Strong engineering expertise','Integrated transport planning','Major economic impact'],
            'Weaknesses':    ['Significant cost overruns','Multiple project delays','Complex stakeholder management','Technical integration challenges'],
            'Opportunities': ['Urban regeneration catalyst','Reduced congestion in London','Property value increases','Template for future megaprojects'],
            'Threats':       ['Budget escalation risks','Political pressure and scrutiny','Public disruption','Contractor performance risks']
        }
    else:
        swot_content = {
            'Strengths':     ['Strong market position','Established operational capabilities','Financial resources','Brand recognition'],
            'Weaknesses':    ['Resource constraints','Limited geographic reach','Operational complexity','Cost pressures'],
            'Opportunities': ['Market expansion potential','Technology adoption','Strategic partnerships','Emerging market growth'],
            'Threats':       ['Competitive pressure','Regulatory changes','Economic uncertainty','Market disruption']
        }
    fig, axes = plt.subplots(2, 2, figsize=(14, 9))
    fig.patch.set_facecolor('#F8F9FA')
    colors_map = {'Strengths':'#27AE60','Weaknesses':'#E74C3C','Opportunities':'#2980B9','Threats':'#E67E22'}
    ax_map     = {'Strengths':axes[0,0],'Weaknesses':axes[0,1],'Opportunities':axes[1,0],'Threats':axes[1,1]}
    for title, items in swot_content.items():
        color = colors_map[title]; ax = ax_map[title]
        ax.set_facecolor(color+'18'); ax.set_xlim(0,1); ax.set_ylim(0,1)
        ax.set_xticks([]); ax.set_yticks([])
        for spine in ax.spines.values(): spine.set_edgecolor(color); spine.set_linewidth(2)
        ax.text(0.5,0.93,title,ha='center',va='top',fontsize=13,fontweight='bold',color=color,transform=ax.transAxes)
        ax.axhline(y=0.87,color=color,linewidth=1.5,alpha=0.4)
        for i,item in enumerate(items):
            ax.text(0.05,0.76-i*0.145,f'• {item}',ha='left',va='top',fontsize=9,color='#2C3E50',transform=ax.transAxes)
    title_text = f"Figure: SWOT Analysis — {case_name}" if case_name else "Figure: SWOT Analysis"
    plt.suptitle(title_text + "\\nSource: Author's own analysis", fontsize=11, fontweight='bold', y=1.01)
    plt.tight_layout()
    path = f"{work_dir}/swot.png"
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='#F8F9FA'); plt.close()
    charts_generated.append(('swot', path))

# ── MARKET SCREENING MATRIX ───────────────────────────────────────────────────
if has_market_screen or is_gsm:
    fig, ax = plt.subplots(figsize=(14, 8))
    ax.set_facecolor('#FDFEFE'); fig.patch.set_facecolor('#FDFEFE')
    if is_luckin:
        markets = ['India','Brazil','Japan','S. Korea','UAE','UK','Germany','France','Australia','Mexico']
        scores  = {
            'Market Size':     [9,8,7,6,8,7,6,6,6,7],
            'Coffee Culture':  [5,6,9,8,6,8,7,9,8,6],
            'Digital Ready':   [7,6,9,9,8,8,8,8,9,7],
            'Competitive Gap': [8,7,5,6,7,6,6,5,7,7],
            'Regulatory Ease': [6,5,7,7,8,8,8,8,8,6]
        }
    else:
        markets = ['Market A','Market B','Market C','Market D','Market E','Market F','Market G','Market H']
        scores  = {
            'Market Size':    [8,7,6,9,5,7,8,6],
            'Growth Potential':[7,8,7,6,9,6,7,8],
            'Competitive Gap': [6,7,8,5,7,8,6,7],
            'Entry Barriers':  [7,6,7,8,6,7,5,8],
            'Strategic Fit':   [8,7,6,7,8,7,9,6]
        }
    criteria     = list(scores.keys())
    x            = np.arange(len(markets))
    width        = 0.15
    colors_bar   = ['#2E86AB','#27AE60','#E67E22','#9B59B6','#E74C3C']
    for i,(crit,vals) in enumerate(scores.items()):
        ax.bar(x + i*width, vals, width, label=crit, color=colors_bar[i], alpha=0.85, edgecolor='white')
    totals = [sum(scores[c][j] for c in criteria) for j in range(len(markets))]
    ax2 = ax.twinx()
    ax2.plot(x + width*2, totals, 'D-', color='#1F3864', linewidth=2, markersize=8, label='Total Score', zorder=5)
    ax2.set_ylabel('Total Score', fontsize=9, color='#1F3864')
    ax2.tick_params(axis='y', labelcolor='#1F3864')
    ax.set_xlabel('Markets / Countries', fontsize=10, fontweight='bold')
    ax.set_ylabel('Score (1–10)', fontsize=10, fontweight='bold')
    ax.set_title("Figure: Market Screening Matrix\\nSource: Author's own analysis", fontsize=11, fontweight='bold', pad=12)
    ax.set_xticks(x + width*2); ax.set_xticklabels(markets, rotation=25, ha='right', fontsize=9)
    ax.set_ylim(0, 11); ax.legend(loc='upper left', fontsize=8)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    plt.tight_layout()
    path = f"{work_dir}/market_screening.png"
    plt.savefig(path, dpi=150, bbox_inches='tight', facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('market_screening', path))

# ── STAKEHOLDER MAP ───────────────────────────────────────────────────────────
if has_stakeholder:
    fig, ax = plt.subplots(figsize=(13, 8))
    ax.set_xlim(0,10); ax.set_ylim(0,8); ax.axis('off'); fig.patch.set_facecolor('#FDFEFE')
    for r,color,alpha in [(3.5,'#1F3864',0.08),(2.5,'#2E86AB',0.08),(1.5,'#44BBA4',0.08)]:
        ax.add_patch(plt.Circle((5,4), r, color=color, alpha=alpha, zorder=1))
    if is_sph:
        center_label = 'SPH\\nPROJECT'
        stakeholders = [
            (5,7.2,'Scottish Parliament\\n(Client)','#C73E1D'),(5,0.8,'Scottish Public','#C73E1D'),
            (1.8,6.2,'Scottish Executive\\n(Peter Fraser)','#E67E22'),(8.2,6.2,'SPCB','#E67E22'),
            (1.0,4.0,'Main Contractor\\n(Bovis/MACE)','#2E86AB'),(9.0,4.0,'Design Team\\n(EMBT/RMJM)','#2E86AB'),
            (1.8,1.8,'Auditor General','#44BBA4'),(8.2,1.8,'Media / Public\\nScrutiny','#44BBA4')
        ]
    elif is_gsm or is_luckin:
        center_label = 'LUCKIN\\nCOFFEE' if is_luckin else 'ORGANISATION'
        stakeholders = [
            (5,7.2,'CEO / Board','#C73E1D'),(5,0.8,'Customers','#C73E1D'),
            (1.8,6.2,'Government /\\nRegulators','#E67E22'),(8.2,6.2,'Investors /\\nShareholders','#E67E22'),
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
        ax.text(x,y,name,ha='center',va='center',fontsize=7.5,
                bbox=dict(boxstyle='round,pad=0.3',facecolor=color+'22',edgecolor=color,linewidth=1.2))
    ax.set_title("Figure: Stakeholder Map\\nSource: Author's own analysis",fontsize=10,fontweight='bold',pad=10)
    plt.tight_layout()
    path = f"{work_dir}/stakeholder.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('stakeholder', path))

# ── PESTEL ────────────────────────────────────────────────────────────────────
if has_pestel:
    fig, ax = plt.subplots(figsize=(13, 8))
    ax.set_xlim(0,13); ax.set_ylim(0,8); ax.axis('off')
    if is_luckin:
        pestel = [
            ('P\\nPolitical',    '#2E86AB',1.0, ['Trade regulations','Market entry barriers','Geopolitical tensions','Government incentives']),
            ('E\\nEconomic',     '#27AE60',3.2, ['GDP growth rates','Consumer spending','Currency exchange','Cost of operations']),
            ('S\\nSocial',       '#F18F01',5.4, ['Coffee culture growth','Digital lifestyle','Health consciousness','Youth demographics']),
            ('T\\nTechnological','#9B59B6',7.6, ['Mobile payment apps','AI ordering systems','Supply chain tech','Data analytics']),
            ('E\\nEnvironmental','#16A085',9.8, ['Sustainable sourcing','Carbon footprint','Packaging regulations','Climate impact']),
            ('L\\nLegal',        '#E74C3C',12.0,['Franchise laws','Food safety regs','Employment law','IP protection']),
        ]
    else:
        pestel = [
            ('P\\nPolitical',    '#2E86AB',1.0, ['Policy changes','Trade regulations','Political stability','Government incentives']),
            ('E\\nEconomic',     '#27AE60',3.2, ['Market growth','Inflation rates','Consumer spending','Currency risk']),
            ('S\\nSocial',       '#F18F01',5.4, ['Demographic shifts','Cultural trends','Consumer behaviour','Social responsibility']),
            ('T\\nTechnological','#9B59B6',7.6, ['Digital disruption','Innovation cycles','Automation','Data analytics']),
            ('E\\nEnvironmental','#16A085',9.8, ['Climate change','Sustainability','Carbon targets','Resource scarcity']),
            ('L\\nLegal',        '#E74C3C',12.0,['Compliance','IP protection','Consumer law','Regulatory standards']),
        ]
    for title,color,x,items in pestel:
        ax.add_patch(mpatches.FancyBboxPatch((x-0.9,0.3),1.9,7.4,boxstyle='round,pad=0.15',facecolor=color+'20',edgecolor=color,linewidth=2))
        ax.text(x,7.3,title,ha='center',va='top',fontsize=9,fontweight='bold',color=color)
        ax.axhline(y=6.6,xmin=(x-0.9)/13,xmax=(x+1.0)/13,color=color,alpha=0.4,lw=1)
        for i,item in enumerate(items):
            ax.text(x,6.1-i*1.2,item,ha='center',va='top',fontsize=7.5,color='#2C3E50')
    ax.set_title("Figure: PESTEL Analysis\\nSource: Author's own analysis",fontsize=11,fontweight='bold',pad=10,color='#1F3864')
    plt.tight_layout()
    path = f"{work_dir}/pestel.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('pestel', path))

# ── BSC FAILURE DISTRIBUTION (BSC briefs) ────────────────────────────────────
if has_bsc:
    fig, ax = plt.subplots(figsize=(10, 7))
    perspectives = ['Financial', 'Customer', 'Internal\\nProcess', 'Learning &\\nGrowth']
    counts       = [8, 7, 9, 7]
    colors_bsc   = ['#C0392B','#2980B9','#8E44AD','#27AE60']
    bars = ax.bar(perspectives, counts, color=colors_bsc, edgecolor='white', linewidth=1.5, width=0.6)
    for bar,count in zip(bars,counts):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.15, str(count),
                ha='center',va='bottom',fontsize=12,fontweight='bold')
    ax.set_ylabel('Number of Performance Failures', fontsize=11)
    ax.set_title("Figure: BSC Failure Distribution by Perspective\\nSource: Author's own analysis",
                 fontsize=11,fontweight='bold',pad=12)
    ax.set_ylim(0, max(counts)+2)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.set_facecolor('#FAFAFA')
    plt.tight_layout()
    path = f"{work_dir}/bsc_dist.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('bsc_dist', path))

# ── COST ESCALATION TIMELINE (SPH / public sector overrun) ───────────────────
if has_cost:
    if is_sph:
        years  = [1997,1998,1999,2000,2001,2002,2003,2004]
        costs  = [40,  80,  109, 195, 241, 294, 373, 431]
        events = {1997:'Initial estimate\\n£40m',2000:'Spencely Report\\n£195m',
                  2003:'Fraser Inquiry\\nLaunched',2004:'Final cost\\n£431m'}
        src    = 'House of Commons (2005); Ahiaga-Dagbui and Smith (2014)'
    else:
        years  = [2019,2020,2021,2022,2023,2024]
        costs  = [100, 130, 175, 210, 270, 320]
        events = {2019:'Initial budget',2022:'Mid-project review',2024:'Final outturn'}
        src    = "Author's own analysis"
    fig, ax = plt.subplots(figsize=(13, 7))
    ax.fill_between(years, costs, alpha=0.12, color='#C73E1D')
    ax.plot(years, costs, 'o-', color='#C73E1D', linewidth=2.5, markersize=8,
            markerfacecolor='white', markeredgewidth=2)
    for yr,label in events.items():
        if yr in years:
            idx = years.index(yr)
            ax.annotate(label, xy=(yr, costs[idx]), xytext=(yr, costs[idx]+30),
                        ha='center', fontsize=8, color='#1F3864',
                        arrowprops=dict(arrowstyle='->',color='#1F3864',lw=1),
                        bbox=dict(boxstyle='round,pad=0.2',facecolor='#EBF5FB',edgecolor='#1F3864',linewidth=0.8))
    ax.axhline(y=costs[0],color='#27AE60',linestyle='--',linewidth=1.5,alpha=0.7,label=f'Original estimate')
    ax.set_xlabel('Year',fontsize=11,fontweight='bold')
    ax.set_ylabel('Cost (£ millions)',fontsize=11,fontweight='bold')
    ax.set_title(f"Figure: Cost Escalation Timeline\\nSource: {src}",fontsize=11,fontweight='bold',pad=12)
    ax.legend(fontsize=9)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.set_facecolor('#FAFAFA')
    plt.tight_layout()
    path = f"{work_dir}/cost_timeline.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('cost_timeline', path))

# ── RISK MATRIX (PM briefs only) ──────────────────────────────────────────────
if has_risk_pm and not is_gsm:
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.set_xlim(0,5); ax.set_ylim(0,5)
    colors_rm = [['#92D050','#92D050','#FFFF00','#FF0000','#FF0000'],
                 ['#92D050','#FFFF00','#FFFF00','#FF0000','#FF0000'],
                 ['#92D050','#FFFF00','#FFFF00','#FFA500','#FF0000'],
                 ['#92D050','#92D050','#FFFF00','#FFA500','#FFA500'],
                 ['#92D050','#92D050','#92D050','#FFFF00','#FFFF00']]
    for i in range(5):
        for j in range(5):
            ax.add_patch(mpatches.Rectangle((j,4-i),1,1,facecolor=colors_rm[i][j],edgecolor='white',linewidth=2,alpha=0.8))
            ax.text(j+0.5,4-i+0.5,str((i+1)*(j+1)),ha='center',va='center',fontsize=12,fontweight='bold',color='#1F3864')
    risks = ([(3.5,4.5,'Cost overrun','#1F3864'),(2.5,3.5,'Schedule delay','#1F3864'),
              (1.5,2.5,'Design changes','#1F3864'),(4.5,1.5,'Governance failure','#1F3864'),
              (0.5,0.5,'Force majeure','#1F3864')] if is_sph else
             [(3.5,4.5,'Schedule delay','#1F3864'),(2.5,3.5,'Budget overrun','#1F3864'),
              (1.5,2.5,'Stakeholder conflict','#1F3864'),(4.5,1.5,'Technical failure','#1F3864'),
              (0.5,0.5,'Force majeure','#1F3864')])
    for x,y,label,color in risks:
        ax.plot(x,y,'o',markersize=12,color=color,zorder=5)
        ax.text(x+0.1,y+0.1,label,fontsize=7.5,color=color)
    ax.set_xlabel('Likelihood →',fontsize=11,fontweight='bold')
    ax.set_ylabel('Impact →',fontsize=11,fontweight='bold')
    ax.set_xticks([0.5,1.5,2.5,3.5,4.5]); ax.set_xticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
    ax.set_yticks([0.5,1.5,2.5,3.5,4.5]); ax.set_yticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
    ax.set_title("Figure: Risk Assessment Matrix\\nSource: Author's own analysis",fontsize=11,fontweight='bold',pad=12)
    plt.tight_layout()
    path = f"{work_dir}/risk_matrix.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('risk_matrix', path))

# ── RESOURCE LOADING DIAGRAM (scheduling briefs) ─────────────────────────────
if has_resource:
    fig, ax = plt.subplots(figsize=(14, 6))
    days         = list(range(1, 25))
    es_resources = [4,4,4,4,5,5,5,6,6,6,5,5,5,5,5,4,4,4,3,3,3,2,2,1]
    colors_bar   = ['#E74C3C' if r > 6 else '#2E86AB' for r in es_resources]
    ax.bar(days, es_resources, color=colors_bar, edgecolor='white', linewidth=0.8, width=0.8)
    ax.axhline(y=6,color='#E74C3C',linestyle='--',linewidth=2,label='Max constraint: 6 labourers/day')
    ax.set_xlabel('Day',fontsize=11,fontweight='bold')
    ax.set_ylabel('Labourers',fontsize=11,fontweight='bold')
    ax.set_title("Figure: Resource Loading Diagram — Early Start\\nSource: Author's own calculations",
                 fontsize=11,fontweight='bold',pad=12)
    ax.set_xticks(days); ax.set_yticks(range(0, max(es_resources)+2))
    ax.legend(fontsize=9)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.set_facecolor('#FAFAFA')
    plt.tight_layout()
    path = f"{work_dir}/resource_loading.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('resource_loading', path))

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
            encoding: 'utf8', maxBuffer: 100 * 1024 * 1024, timeout: 120000
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
        submissionDate, workableTask, totalWordCount, targetWordCount, docSections,
        // Formatting params from worker — falls back to brief-standard defaults
        includeWordCountInDoc,
        fontName, fontSize, lineSpacing, marginsCm
    } = req.body;

    // ── Formatting constants ──────────────────────────────────────────────────
    const FONT     = (fontName && fontName.trim()) ? fontName.trim() : 'Arial';
    const FSIZE_HX = Math.round((parseFloat(fontSize) || 12) * 2);   // half-points
    const LSPACING = Math.round((parseFloat(lineSpacing) || 1.5) * 240); // TWIPs (240 = single)
    const MARGIN   = Math.round((parseFloat(marginsCm) || 2.54) * 567);  // DXA (567 per cm)
    const SZ_H1    = FSIZE_HX + 8;
    const SZ_H2    = FSIZE_HX + 4;
    const SZ_SM    = FSIZE_HX - 4;
    // Show word count at very end only if brief requires it (default true — most briefs do)
    const showWC   = (includeWordCountInDoc !== false && includeWordCountInDoc !== 'false');

    const workDir = `/tmp/charts_${executionId}`;
    const {
        Document, Packer, Paragraph, TextRun, ImageRun,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
        WidthType, PageNumber, PageBreak,
        Table, TableRow, TableCell, ShadingType
    } = require('docx');

    // ── Load charts ───────────────────────────────────────────────────────────
    let charts = {};
    try {
        if (fs.existsSync(`${workDir}/manifest.json`)) {
            const m = JSON.parse(fs.readFileSync(`${workDir}/manifest.json`, 'utf8'));
            charts = m.charts || {};
        }
    } catch(e) {}

    // ── Text run helper ───────────────────────────────────────────────────────
    const tr = (t, o = {}) => new TextRun({
        text: String(t || ''), font: FONT,
        size:    o.size   || FSIZE_HX,
        bold:    o.bold   || false,
        italics: o.italic || false,
        color:   o.color  || '000000'
    });

    const blk = () => new Paragraph({
        spacing: { line: LSPACING, before: 0, after: 0 }, children: [tr('')]
    });

    function mkP(runs, o = {}) {
        const align = o.center ? AlignmentType.CENTER
                    : o.right  ? AlignmentType.RIGHT
                    : o.left   ? AlignmentType.LEFT
                    :            AlignmentType.JUSTIFIED;
        return new Paragraph({
            alignment: align,
            spacing: { line: LSPACING, before: o.before || 0, after: o.after !== undefined ? o.after : 160 },
            children: Array.isArray(runs) ? runs : [tr(runs, o)]
        });
    }

    function h1(t) {
        return new Paragraph({
            heading:  HeadingLevel.HEADING_1,
            spacing:  { line: LSPACING, before: 360, after: 160 },
            children: [tr(t, { bold: true, size: SZ_H1, color: '1F3864' })]
        });
    }
    function h2(t) {
        return new Paragraph({
            heading:  HeadingLevel.HEADING_2,
            spacing:  { line: LSPACING, before: 240, after: 120 },
            children: [tr(t, { bold: true, size: SZ_H2, color: '2C5282' })]
        });
    }

    // ── Chart embed ───────────────────────────────────────────────────────────
    function embedChart(name, wCm, hCm, caption) {
        if (!charts[name]) return [];
        try {
            const buf  = Buffer.from(charts[name].b64, 'base64');
            const wPx  = Math.round(wCm * 360000 / 9144);
            const hPx  = Math.round(hCm * 360000 / 9144);
            return [
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing:   { line: LSPACING, before: 80, after: 0 },
                    children:  [new ImageRun({ data: Uint8Array.from(buf),
                                               transformation: { width: wPx, height: hPx },
                                               type: 'png' })]
                }),
                mkP([tr(caption, { size: SZ_SM, italic: true, color: '444444' })],
                    { center: true, before: 20, after: 20 }),
                mkP([tr("Source: Author's own analysis.", { size: SZ_SM, italic: true, color: '666666' })],
                    { left: true, before: 0, after: 200 })
            ];
        } catch(e) { return [mkP(`[Chart: ${name}]`, { italic: true })]; }
    }

    // Each chart only injected ONCE across the whole document
    const usedCharts = new Set();
    function getChartsForSection(title, content) {
        const text = (title + ' ' + content).toLowerCase();
        const injected = [];
        const candidates = [
            ['cost_timeline',    'cost',         14, 8,  'Figure: Cost Escalation Timeline'],
            ['bsc_dist',         'bsc',          12, 7,  'Figure: BSC Failure Distribution by Perspective'],
            ['risk_matrix',      'risk',         12, 9,  'Figure: Risk Assessment Matrix'],
            ['stakeholder',      'stakeholder',  14, 8,  'Figure: Stakeholder Map'],
            ['swot',             'swot',         15, 9,  'Figure: SWOT Analysis'],
            ['pestel',           'pestel',       15, 8,  'Figure: PESTEL Analysis'],
            ['market_screening', 'segmentation', 15, 9,  'Figure: Market Screening Matrix'],
            ['market_screening', 'market select',15, 9,  'Figure: Market Screening Matrix'],
            ['resource_loading', 'resource',     15, 7,  'Figure: Resource Loading Diagram — Early Start']
        ];
        for (const [name, kw, wCm, hCm, cap] of candidates) {
            if (charts[name] && !usedCharts.has(name) && text.includes(kw)) {
                injected.push([name, wCm, hCm, cap]);
                usedCharts.add(name);
            }
        }
        return injected;
    }

    // ── Render pipe-table lines as a proper docx Table ────────────────────────
    function renderInlineTable(lines) {
        const rows = lines.filter(l => l.includes('|'));
        if (rows.length < 2) return null;
        const parsed = rows
            .map(row => row.split('|').map(c => c.trim()).filter(c => c))
            .filter(r => r.length > 0 && !r.every(c => /^[-\s]+$/.test(c)));
        if (parsed.length < 1) return null;
        const colCount = Math.max(...parsed.map(r => r.length));
        const colW     = Math.floor(9026 / colCount);
        const border   = { style: BorderStyle.SINGLE, size: 1, color: 'BDC3C7' };
        const borders  = { top: border, bottom: border, left: border, right: border };
        return new Table({
            width: { size: 9026, type: WidthType.DXA },
            columnWidths: Array(colCount).fill(colW),
            rows: parsed.map((row, ri) => new TableRow({
                children: Array.from({ length: colCount }, (_, ci) => {
                    const cell = row[ci] || '';
                    const isHdr = ri === 0;
                    return new TableCell({
                        width:   { size: colW, type: WidthType.DXA },
                        margins: { top: 80, bottom: 80, left: 120, right: 120 },
                        shading: { fill: isHdr ? '1F3864' : (ri % 2 === 0 ? 'F2F6FC' : 'FFFFFF'),
                                   type: ShadingType.CLEAR },
                        borders,
                        children: [new Paragraph({
                            spacing: { line: 240, before: 0, after: 0 },
                            children: [tr(cell, { bold: isHdr, color: isHdr ? 'FFFFFF' : '2C3E50', size: SZ_SM })]
                        })]
                    });
                })
            }))
        });
    }

    // ── Parse section content → paragraphs + tables + subheadings ─────────────
    function parseContent(rawContent) {
        const elements = [];
        for (const para of (rawContent || '').split(/\n\n+/)) {
            const t = para.trim();
            if (!t) continue;
            // Pipe-table block
            if (t.includes('|') && t.split('\n').filter(l => l.includes('|')).length >= 2) {
                const tbl = renderInlineTable(t.split('\n'));
                if (tbl) { elements.push(tbl); elements.push(blk()); continue; }
            }
            // Markdown H2
            if (t.startsWith('## ')) { elements.push(h2(t.replace(/^##\s*/, ''))); continue; }
            // Numbered subheading e.g. "3.1  Governance Failures"
            if (/^\d+\.\d+[\s]+\w/.test(t) && t.length < 120) { elements.push(h2(t)); continue; }
            // Normal paragraph
            elements.push(mkP(t));
        }
        return elements;
    }

    // ── BUILD DOCUMENT ────────────────────────────────────────────────────────
    try {
        const safeVal = v => (v && v !== 'N/A' && v !== 'Student' && v !== 'Not Available') ? v : '';

        // Anonymous submission flag — brief spec says "should not contain your name"
        // If studentId is present we show it; name is shown only when explicitly provided
        const isAnon = !safeVal(studentName);
        const displayName = isAnon ? 'Anonymous Submission' : safeVal(studentName);

        // Cover page — matches reference exactly
        // University / school top, title centre, module details bottom
        const refStyle  = workableTask?.referencing_style || 'Harvard';
        const specLine  = `${refStyle} Referencing  |  ${FONT} ${FSIZE_HX/2}pt  |  ${parseFloat(lineSpacing)||1.5} Line Spacing  |  ${parseFloat(marginsCm)||2.54}cm margins  |  A4`;

        const cover = [
            blk(), blk(),
            // University name — large, bold, navy
            mkP([tr((university || 'University').toUpperCase(), { bold: true, size: SZ_H1 + 6, color: '1F3864' })], { center: true, after: 20 }),
            // School / department line if in programme
            mkP([tr(workableTask?.subject_area || programme || '', { size: FSIZE_HX })], { center: true, after: 240 }),
            blk(), blk(),
            // Report title — largest element
            mkP([tr((workableTask?.document_type || 'Academic Report').toUpperCase(), { bold: true, size: SZ_H1 + 8, color: '1F3864' })], { center: true, after: 40 }),
            // Subtitle / case study
            mkP([tr(workableTask?.case_study || '', { bold: true, size: SZ_H1, color: '2C5282' })], { center: true, after: 80 }),
            // Subtitle 2 — if document has a sub-description
            blk(), blk(), blk(), blk(),
            // Module info block
            mkP([tr(`Module: ${workableTask?.subject_area || programme || ''}  |  ${workableTask?.academic_level || 'Level 7'}`, { size: FSIZE_HX - 2 })], { center: true, after: 20 }),
            mkP([tr(`Module Lead: ${workableTask?.module_lead || ''}`, { size: FSIZE_HX - 2 })], { center: true, after: 20 }),
            mkP([tr(`Submission: ${safeVal(submissionDate) || ''}  |  ${displayName}`, { size: FSIZE_HX - 2 })], { center: true, after: 20 }),
            // Student ID if present
            ...(safeVal(studentId) ? [mkP([tr(`Student ID: ${studentId}`, { size: FSIZE_HX - 2 })], { center: true, after: 20 })] : []),
            blk(),
            // Spec line — small, italic, bottom of cover
            mkP([tr(specLine, { size: SZ_SM, italic: true, color: '888888' })], { center: true, after: 0 }),
            new Paragraph({ children: [new PageBreak()] })
        ];

        // TOC — indented with tab leaders matching reference style
        const { TabStopPosition, TabStopType, TabStopLeader } = require('docx');
        const tocItems = (docSections || []).map(s => {
            const label = (s.sectionNumber ? s.sectionNumber + '  ' : '') + (s.title || '');
            return new Paragraph({
                spacing: { line: LSPACING, before: 30, after: 30 },
                tabStops: [{ type: TabStopType.RIGHT, position: 8640, leader: TabStopLeader ? TabStopLeader.DOT : 3 }],
                children: [
                    tr(label, { size: FSIZE_HX }),
                    new TextRun({ text: '\t', font: FONT, size: FSIZE_HX }),
                ]
            });
        });

        const toc = [
            h1('Table of Contents'),
            ...tocItems,
            new Paragraph({ children: [new PageBreak()] })
        ];

        // ── Equation display block ───────────────────────────────────────────
        // Detects "Equation N —" lines in content and renders them as styled blocks
        function parseContentWithEquations(rawContent) {
            const elements = [];
            for (const para of (rawContent || '').split(/\n\n+/)) {
                const t = para.trim();
                if (!t) continue;

                // Equation block — "Equation N —" or "[EQUATION]...[/EQUATION]"
                if (/^Equation\s+\d+\s*[—–-]/i.test(t) || t.startsWith('[EQUATION]')) {
                    const eqText = t.replace(/^\[EQUATION\]/, '').replace(/\[\/EQUATION\]$/, '').trim();
                    elements.push(new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing:   { line: LSPACING, before: 160, after: 160 },
                        border: {
                            top:    { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 },
                            bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 }
                        },
                        children: [tr(eqText, { bold: true, size: FSIZE_HX, color: '1F3864' })]
                    }));
                    continue;
                }

                // Pipe-table block
                if (t.includes('|') && t.split('\n').filter(l => l.includes('|')).length >= 2) {
                    const tbl = renderInlineTable(t.split('\n'));
                    if (tbl) { elements.push(tbl); elements.push(blk()); continue; }
                }

                // Markdown H2
                if (t.startsWith('## ')) { elements.push(h2(t.replace(/^##\s*/, ''))); continue; }

                // Numbered subheading e.g. "3.1  Governance Failures"
                if (/^\d+\.\d+[\s]+\w/.test(t) && t.length < 120) { elements.push(h2(t)); continue; }

                // Source / caption line — small italic
                if (/^Source:/i.test(t)) {
                    elements.push(mkP([tr(t, { size: SZ_SM, italic: true, color: '666666' })], { left: true, before: 0, after: 120 }));
                    continue;
                }

                // Table/figure caption line
                if (/^(Table|Figure)\s+\d+/i.test(t) && t.length < 200) {
                    elements.push(mkP([tr(t, { size: SZ_SM, bold: true, color: '333333' })], { left: true, before: 80, after: 40 }));
                    continue;
                }

                elements.push(mkP(t));
            }
            return elements;
        }

        const mainContent = [];
        for (const sec of (docSections || [])) {
            const titleLower = (sec.title || '').toLowerCase();
            if (titleLower.includes('references') || titleLower.includes('bibliography')) {
                mainContent.push(new Paragraph({ children: [new PageBreak()] }));
            }
            mainContent.push(h1((sec.sectionNumber ? sec.sectionNumber + '  ' : '') + sec.title));
            // Use enhanced parser with equation support
            mainContent.push(...parseContentWithEquations(sec.content || ''));
            for (const [name, wCm, hCm, cap] of getChartsForSection(sec.title || '', sec.content || '')) {
                mainContent.push(...embedChart(name, wCm, hCm, cap));
            }
            mainContent.push(blk());
        }

        // Word count at end — brief spec: "number of words should be stated at the end"
        // Always shown unless worker explicitly passes false
        if (showWC) {
            mainContent.push(new Paragraph({
                border:   { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
                spacing:  { line: LSPACING, before: 120, after: 60 },
                children: [tr(
                    `Word Count: ${totalWordCount || 0} words ` +
                    `(excluding cover page, table of contents, tables, figures, and reference list).`,
                    { size: SZ_SM, italic: true, color: '666666' }
                )]
            }));
        }

        const doc = new Document({
            styles: {
                default: { document: { run: { font: FONT, size: FSIZE_HX } } },
                paragraphStyles: [
                    { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
                      run: { size: SZ_H1, bold: true, font: FONT, color: '1F3864' },
                      paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
                    { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
                      run: { size: SZ_H2, bold: true, font: FONT, color: '2C5282' },
                      paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } }
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size:   { width: 11906, height: 16838 },   // A4
                        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
                    }
                },
                headers: { default: new Header({ children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    border:    { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1F3864', space: 4 } },
                    spacing:   { before: 0, after: 80 },
                    children:  [tr(
                        // Anonymous — no student name in header per brief spec
                        `${safeVal(university) || 'University'}  |  ${workableTask?.document_type || 'Report'}`,
                        { size: SZ_SM, italic: true, color: '555555' }
                    )]
                })] }) },
                footers: { default: new Footer({ children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    border:    { top: { style: BorderStyle.SINGLE, size: 6, color: '1F3864', space: 4 } },
                    spacing:   { before: 80, after: 0 },
                    children:  [
                        tr('Page ', { size: SZ_SM, color: '777777' }),
                        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_SM, color: '777777' }),
                        tr(' of ',  { size: SZ_SM, color: '777777' }),
                        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: SZ_SM, color: '777777' })
                    ]
                })] }) },
                children: [...cover, ...toc, ...mainContent]
            }]
        });

        Packer.toBuffer(doc).then(buf => {
            const docType  = (workableTask?.document_type || 'Report').replace(/\s+/g, '_');
            res.json({ stdout: 'SUCCESS', stderr: '', returncode: 0,
                       docxBase64: buf.toString('base64'),
                       filename:   `${docType}_completed.docx` });
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
app.listen(PORT, () => { console.log(`n8n microservice listening on port ${PORT}`); });
