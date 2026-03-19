const express = require('express');
const cors = require('cors');
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Health check
app.get('/', (req, res) => res.json({ status: 'ok', service: 'n8n-doc-microservice' }));
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ── /extract-text ─────────────────────────────────────────────────────────────
app.post('/extract-text', (req, res) => {
    const { fileExt, base64Data, executionId } = req.body;
    if (!base64Data) return res.json({ stdout: '', stderr: 'No data provided', returncode: 1 });

    const inputPath = `/tmp/brief_${executionId}.${fileExt}`;
    const outputPath = `/tmp/brief_${executionId}_text.txt`;

    try {
        fs.writeFileSync(inputPath, Buffer.from(base64Data, 'base64'));

        let stdout = '';

        if (fileExt === 'pdf') {
            execSync(`pdftotext -layout "${inputPath}" "${outputPath}"`, { encoding: 'utf8' });
            stdout = fs.readFileSync(outputPath, 'utf8');
        } else if (fileExt === 'docx') {
            const pyScript = `
from docx import Document
doc = Document('${inputPath}')
lines = [p.text for p in doc.paragraphs if p.text.strip()]
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if cell.text.strip():
                lines.append(cell.text.strip())
with open('${outputPath}','w') as f:
    f.write('\\n'.join(lines))
`;
            execSync(`python3 -c "${pyScript.replace(/"/g, '\\"')}"`, { encoding: 'utf8' });
            stdout = fs.readFileSync(outputPath, 'utf8');
        } else if (fileExt === 'pptx') {
            const pyScript = `
from pptx import Presentation
prs = Presentation('${inputPath}')
lines = []
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = para.text.strip()
                if t: lines.append(t)
with open('${outputPath}','w') as f:
    f.write('\\n'.join(lines))
`;
            execSync(`python3 -c "${pyScript.replace(/"/g, '\\"')}"`, { encoding: 'utf8' });
            stdout = fs.readFileSync(outputPath, 'utf8');
        } else if (fileExt === 'txt') {
            stdout = Buffer.from(base64Data, 'base64').toString('utf8');
        }

        // Cleanup
        try { fs.unlinkSync(inputPath); } catch(e) {}
        try { fs.unlinkSync(outputPath); } catch(e) {}

        res.json({ stdout: stdout.trim(), stderr: '', returncode: 0 });
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

    const pyScript = `
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
has_cost = any(k in task_str for k in ['cost overrun','escalat','£40','£431'])
has_bmc = any(k in task_str for k in ['business model','value proposition','bmc'])
has_competitors = any(k in task_str for k in ['competitor','competition','market analysis'])

# ── GANTT + RESOURCE CHARTS ───────────────────────────────────────────────────
if has_resource:
    acts_es = [('A',1,3,4),('B',4,2,2),('C',4,4,3),('D',4,5,5),('E',8,3,1),('F',9,5,3),('G',9,7,4),('H',11,8,2),('Y',19,1,3),('Z',16,2,2),('X',20,5,1)]
    acts_ls = [('A',1,3,4),('B',6,2,2),('C',4,4,3),('D',6,5,5),('E',8,3,1),('F',13,5,3),('G',11,7,4),('H',11,8,2),('Y',19,1,3),('Z',18,2,2),('X',20,5,1)]
    proj_dur = 24; max_lab = 6
    def daily_res(acts):
        r = [0]*(proj_dur+1)
        for (id,es,d,lab) in acts:
            for day in range(es, es+d):
                if day <= proj_dur: r[day] += lab
        return r
    es_r = daily_res(acts_es); ls_r = daily_res(acts_ls)
    days = list(range(1,proj_dur+1))
    es_daily = [es_r[d] for d in days]; ls_daily = [ls_r[d] for d in days]

    # Gantt
    fig, ax = plt.subplots(figsize=(14,6))
    colors = ['#2E86AB','#A23B72','#F18F01','#C73E1D','#3B1F2B','#44BBA4','#E94F37','#393E41','#6B4226','#7B2FBE','#2D6A4F']
    labels = ['Preparation','Concept Design','Spatial Coord.','Building Regs','Planning App','Technical Design','Building Systems','Phase 1 Build','Quality Inspection','Perf. Review','Phase 2 Build']
    for i,(id,es,dur,lab) in enumerate(acts_es):
        ax.barh(i,dur,left=es-1,height=0.6,color=colors[i%len(colors)],edgecolor='white',linewidth=0.5)
        ax.text(es-1+dur/2,i,f'{id}\\n({lab}L)',ha='center',va='center',fontsize=7.5,fontweight='bold',color='white')
    ax.set_yticks(range(len(acts_es)))
    ax.set_yticklabels([f"{a[0]} — {labels[i]}" for i,a in enumerate(acts_es)],fontsize=8)
    ax.set_xlabel('Project Day',fontsize=9)
    ax.set_title('Figure: Gantt Chart — Early Start Schedule\\nNumbers in bars = labourers/day',fontsize=10,fontweight='bold',pad=8)
    ax.set_xlim(0,proj_dur); ax.set_xticks(range(0,proj_dur+1,2))
    ax.grid(axis='x',alpha=0.3,linestyle='--'); ax.invert_yaxis()
    plt.tight_layout()
    path = f"{work_dir}/gantt.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('gantt',path))

    # ES Resource
    fig,ax = plt.subplots(figsize=(14,5))
    bar_colors = ['#C73E1D' if r>max_lab else '#2E86AB' for r in es_daily]
    bars = ax.bar(days,es_daily,color=bar_colors,edgecolor='white',linewidth=0.5,width=0.8)
    ax.axhline(y=max_lab,color='#C73E1D',linewidth=2,linestyle='--',label=f'Max constraint ({max_lab}/day)')
    ax.axhline(y=sum(es_daily)/proj_dur,color='#44BBA4',linewidth=1.5,linestyle='-.',label=f'ADR={sum(es_daily)/proj_dur:.2f}/day')
    for bar,val in zip(bars,es_daily):
        if val>0: ax.text(bar.get_x()+bar.get_width()/2,bar.get_height()+0.05,str(val),ha='center',va='bottom',fontsize=8,fontweight='bold')
    ax.set_xlabel('Project Day',fontsize=9); ax.set_ylabel('Labourers',fontsize=9)
    ax.set_title('Figure: Resource Loading — Early Start\\nRed bars exceed 6-labourer constraint',fontsize=10,fontweight='bold',pad=8)
    ax.set_xticks(days); ax.set_xticklabels(days,fontsize=7.5); ax.set_ylim(0,max(es_daily)+2)
    ax.legend(fontsize=8,loc='upper right'); ax.grid(axis='y',alpha=0.3,linestyle='--')
    plt.tight_layout()
    path = f"{work_dir}/es_resource.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('es_resource',path))

    # ES vs LS
    fig,(ax1,ax2) = plt.subplots(2,1,figsize=(14,8),sharex=True)
    ax1.bar(days,es_daily,color=['#C73E1D' if r>max_lab else '#2E86AB' for r in es_daily],edgecolor='white',linewidth=0.3)
    ax1.axhline(y=max_lab,color='#C73E1D',linewidth=2,linestyle='--'); ax1.set_ylabel('Labourers',fontsize=8)
    ax1.set_title('Early Start (ES)',fontsize=9,fontweight='bold'); ax1.set_ylim(0,max(es_daily)+2)
    ax1.grid(axis='y',alpha=0.3,linestyle='--')
    for i,val in enumerate(es_daily):
        if val>0: ax1.text(i+1,val+0.05,str(val),ha='center',va='bottom',fontsize=7)
    ax2.bar(days,ls_daily,color=['#C73E1D' if r>max_lab else '#44BBA4' for r in ls_daily],edgecolor='white',linewidth=0.3)
    ax2.axhline(y=max_lab,color='#C73E1D',linewidth=2,linestyle='--'); ax2.set_xlabel('Project Day',fontsize=9)
    ax2.set_ylabel('Labourers',fontsize=8); ax2.set_title('Late Start (LS)',fontsize=9,fontweight='bold')
    ax2.set_xticks(days); ax2.set_xticklabels(days,fontsize=7); ax2.set_ylim(0,max(ls_daily)+2)
    ax2.grid(axis='y',alpha=0.3,linestyle='--')
    for i,val in enumerate(ls_daily):
        if val>0: ax2.text(i+1,val+0.05,str(val),ha='center',va='bottom',fontsize=7)
    es_sr2=sum(r**2 for r in es_daily); ls_sr2=sum(r**2 for r in ls_daily)
    fig.suptitle(f'Figure: ES vs LS Comparison — ES SR2={es_sr2} | LS SR2={ls_sr2}',fontsize=10,fontweight='bold',y=1.01)
    plt.tight_layout()
    path = f"{work_dir}/es_ls_compare.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('es_ls_compare',path))

    # Cumulative
    es_cum = list(np.cumsum(es_daily)); ls_cum = list(np.cumsum(ls_daily))
    fig,ax = plt.subplots(figsize=(12,5))
    ax.plot(days,es_cum,'o-',color='#2E86AB',linewidth=2,markersize=4,label='ES Cumulative')
    ax.plot(days,ls_cum,'s--',color='#C73E1D',linewidth=2,markersize=4,label='LS Cumulative')
    ax.fill_between(days,es_cum,ls_cum,alpha=0.12,color='#F18F01',label='Scheduling Float')
    ax.set_xlabel('Project Day',fontsize=9); ax.set_ylabel('Cumulative Labourer-Days',fontsize=9)
    ax.set_title('Figure: Cumulative Resource Requirement Curve — ES vs LS',fontsize=10,fontweight='bold',pad=8)
    ax.set_xticks(days); ax.set_xticklabels(days,fontsize=7.5)
    ax.legend(fontsize=8,loc='upper left'); ax.grid(alpha=0.3,linestyle='--')
    plt.tight_layout()
    path = f"{work_dir}/cumulative.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('cumulative',path))

# ── BSC PIE ───────────────────────────────────────────────────────────────────
if has_bsc:
    fig,ax = plt.subplots(figsize=(8,6))
    labels = ['Financial','Customer','Internal\\nProcess','Learning\\n& Growth']
    colors = ['#C73E1D','#2E86AB','#44BBA4','#F18F01']
    wedges,texts,autotexts = ax.pie([8,8,8,8],labels=labels,autopct='%1.0f%%',colors=colors,startangle=90,pctdistance=0.75,wedgeprops={'edgecolor':'white','linewidth':2},textprops={'fontsize':10})
    for at in autotexts: at.set_fontsize(9); at.set_fontweight('bold'); at.set_color('white')
    ax.set_title('Figure: BSC Performance Failure Factors Distribution',fontsize=10,fontweight='bold',pad=15)
    plt.tight_layout()
    path = f"{work_dir}/bsc_dist.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('bsc_dist',path))

# ── SWOT ──────────────────────────────────────────────────────────────────────
if has_swot:
    fig,axes = plt.subplots(2,2,figsize=(14,9))
    fig.patch.set_facecolor('#F8F9FA')
    swot_data = {
        ('Strengths','#27AE60',axes[0,0]):['Core competency in subject area','Structured analytical approach','Access to academic literature','Theoretical framework application'],
        ('Weaknesses','#E74C3C',axes[0,1]):['Limited primary data access','Time and resource constraints','Potential scope limitations','Reliance on secondary sources'],
        ('Opportunities','#2980B9',axes[1,0]):['Emerging research in field','Cross-disciplinary insights','New framework applications','Recent empirical evidence'],
        ('Threats','#E67E22',axes[1,1]):['Rapidly evolving landscape','Contradictory literature','Methodological debates','Scope creep risk']
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

# ── PESTEL ────────────────────────────────────────────────────────────────────
if has_pestel:
    fig,ax = plt.subplots(figsize=(13,8))
    ax.set_xlim(0,13); ax.set_ylim(0,8); ax.axis('off')
    fig.patch.set_facecolor('#FDFEFE')
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

# ── STAKEHOLDER MAP ───────────────────────────────────────────────────────────
if has_stakeholder:
    fig,ax = plt.subplots(figsize=(13,8))
    ax.set_xlim(0,10); ax.set_ylim(0,8); ax.axis('off')
    fig.patch.set_facecolor('#FDFEFE')
    for r,color,alpha in [(3.5,'#1F3864',0.08),(2.5,'#2E86AB',0.08),(1.5,'#44BBA4',0.08)]:
        circle = plt.Circle((5,4),r,color=color,alpha=alpha,zorder=1)
        ax.add_patch(circle)
    ax.text(5,4,'PROJECT\\nCORE',ha='center',va='center',fontsize=10,fontweight='bold',color='white',zorder=5,
            bbox=dict(boxstyle='round,pad=0.4',facecolor='#1F3864',edgecolor='white'))
    stakeholders = [(5,7.2,'Client / Sponsor','#C73E1D'),(5,0.8,'End Users / Beneficiaries','#C73E1D'),
        (1.8,6.2,'Regulatory Bodies','#E67E22'),(8.2,6.2,'Investors / Funders','#E67E22'),
        (1.0,4.0,'Contractors / Suppliers','#2E86AB'),(9.0,4.0,'Government / Policy','#2E86AB'),
        (1.8,1.8,'Community / Public','#44BBA4'),(8.2,1.8,'Media / Press','#44BBA4')]
    for x,y,name,color in stakeholders:
        ax.annotate('',xy=(5,4),xytext=(x,y),arrowprops=dict(arrowstyle='-',color=color,alpha=0.3,lw=1))
        ax.text(x,y,name,ha='center',va='center',fontsize=8,
                bbox=dict(boxstyle='round,pad=0.3',facecolor=color+'22',edgecolor=color,linewidth=1.2))
    ax.set_title("Figure: Stakeholder Map\\nSource: Author's own analysis",fontsize=10,fontweight='bold',pad=10)
    plt.tight_layout()
    path = f"{work_dir}/stakeholder.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('stakeholder',path))

# ── COST TIMELINE ─────────────────────────────────────────────────────────────
if has_cost:
    import re as re_mod
    task_text = str(task)
    amounts = re_mod.findall(r'£([\\d,.]+)', task_text)
    fig,ax = plt.subplots(figsize=(11,5))
    if len(amounts) >= 2:
        try:
            vals = [float(a.replace(',','')) for a in amounts[:4]]
            years = list(range(1997,1997+len(vals)*2,2))
            ax.plot(years,vals,'o-',color='#C73E1D',linewidth=2.5,markersize=7)
            ax.fill_between(years,vals,vals[0],alpha=0.15,color='#C73E1D')
            ax.axhline(y=vals[0],color='#27AE60',linewidth=1.5,linestyle='--',label='Original estimate')
            for yr,val in zip(years,vals):
                ax.annotate(f'£{val}m',xy=(yr,val),xytext=(yr+0.1,val+max(vals)*0.05),
                           fontsize=8,color='#1F3864',arrowprops=dict(arrowstyle='->',color='#1F3864',lw=1))
        except: pass
    ax.set_xlabel('Year',fontsize=9); ax.set_ylabel('Cost (£ millions)',fontsize=9)
    ax.set_title('Figure: Project Cost Escalation Timeline',fontsize=10,fontweight='bold',pad=8)
    ax.legend(fontsize=8); ax.grid(alpha=0.3,linestyle='--')
    plt.tight_layout()
    path = f"{work_dir}/cost_timeline.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='white'); plt.close()
    charts_generated.append(('cost_timeline',path))

# ── BMC ───────────────────────────────────────────────────────────────────────
if has_bmc:
    fig,ax = plt.subplots(figsize=(16,9))
    ax.set_xlim(0,16); ax.set_ylim(0,9); ax.axis('off')
    fig.patch.set_facecolor('#FDFEFE')
    blocks = [
        (0,4.5,3.2,4.5,'Key Partners',['Strategic alliances','Supplier network','Technology partners','Distribution'],'#2E86AB'),
        (3.2,4.5,3.2,2.2,'Key Activities',['Core operations','R&D & innovation','Marketing & sales','Quality management'],'#A23B72'),
        (3.2,2.3,3.2,2.2,'Key Resources',['Human capital','Technology IP','Brand equity','Financial resources'],'#F18F01'),
        (6.4,1.5,3.2,7.0,'Value Proposition',['Core product/service','Quality & differentiation','Price competitiveness','Customer experience','Unique capabilities'],'#C73E1D'),
        (9.6,4.5,3.2,2.2,'Customer Relations',['Direct engagement','Digital channels','Loyalty programmes'],'#44BBA4'),
        (9.6,2.3,3.2,2.2,'Channels',['Online/direct','Retail partners','B2B sales','Distribution'],'#393E41'),
        (12.8,4.5,3.2,4.5,'Customer Segments',['Primary target','Secondary segments','Geographic focus','B2B/B2C split'],'#2D6A4F'),
        (0,0,6.4,2.3,'Cost Structure',['Fixed costs','Variable costs','Key cost drivers','Economies of scale'],'#6B4226'),
        (6.4,0,9.6,2.3,'Revenue Streams',['Primary revenue','Secondary streams','Pricing model','Growth trajectory'],'#7B2FBE'),
    ]
    for (x,y,w,h,title,items,color) in blocks:
        rect = mpatches.FancyBboxPatch((x+0.05,y+0.05),w-0.1,h-0.1,boxstyle='round,pad=0.1',facecolor=color+'22',edgecolor=color,linewidth=2)
        ax.add_patch(rect)
        ax.text(x+w/2,y+h-0.3,title,ha='center',va='top',fontsize=9,fontweight='bold',color=color)
        ax.axhline(y=y+h-0.55,xmin=(x+0.1)/16,xmax=(x+w-0.1)/16,color=color,linewidth=0.8,alpha=0.5)
        for i,item in enumerate(items):
            ax.text(x+0.25,y+h-0.85-i*0.52,f'• {item}',ha='left',va='top',fontsize=7.5,color='#2C3E50')
    ax.set_title('Figure: Business Model Canvas',fontsize=13,fontweight='bold',pad=12,color='#1F3864')
    plt.tight_layout()
    path = f"{work_dir}/bmc.png"
    plt.savefig(path,dpi=150,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
    charts_generated.append(('bmc',path))

# Save manifest
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
`;

    try {
        const scriptPath = `${workDir}/script.py`;
        fs.writeFileSync(scriptPath, pyScript);
        const stdout = execSync(`python3 ${scriptPath}`, {
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
// Returns base64 of DOCX — no file serving needed, works on Render
app.post('/export-docx', (req, res) => {
    const {
        executionId, studentName, studentId, programme, university,
        submissionDate, workableTask, totalWordCount, targetWordCount, docSections
    } = req.body;

    const workDir = `/tmp/charts_${executionId}`;
    const outputPath = `/tmp/academic_doc_${executionId}.docx`;

    const {
        Document, Packer, Paragraph, TextRun, ImageRun,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
        WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak, TabStopType
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
    const figcap = (t) => mkP([tr(t,{size:SZS,italic:true,color:'444444'})],{center:true,before:20,after:160});
    const srcNote = (t) => mkP([tr('Source: ',{size:SZS,bold:true}),tr(t,{size:SZS,italic:true,color:'444444'})],{left:true,before:10,after:120});

    // Smart chart injection — match by content
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
                figcap(caption),
                srcNote("Author's own analysis.")
            ];
        } catch(e) {
            return [mkP(`[Chart unavailable: ${name}]`,{italic:true,color:'888888'})];
        }
    }

    function getChartsForSection(sectionTitle, sectionContent) {
        const text = (sectionTitle + ' ' + sectionContent).toLowerCase();
        const injected = [];
        if ((text.includes('gantt') || text.includes('early start') || text.includes('activity')) && charts['gantt'])
            injected.push(['gantt', 15, 8, 'Figure: Gantt Chart — Early Start Schedule']);
        if (text.includes('resource loading') && charts['es_resource'])
            injected.push(['es_resource', 15, 7, 'Figure: Resource Loading Diagram — Early Start']);
        if ((text.includes('late start') || text.includes('es vs ls') || text.includes('comparison')) && charts['es_ls_compare'])
            injected.push(['es_ls_compare', 15, 10, 'Figure: ES vs LS Resource Loading Comparison']);
        if (text.includes('cumulative') && charts['cumulative'])
            injected.push(['cumulative', 14, 7, 'Figure: Cumulative Resource Requirement Curve']);
        if ((text.includes('bsc') || text.includes('balanced scorecard')) && charts['bsc_dist'])
            injected.push(['bsc_dist', 12, 8, 'Figure: BSC Performance Factor Distribution']);
        if (text.includes('swot') && charts['swot'])
            injected.push(['swot', 15, 9, 'Figure: SWOT Analysis']);
        if (text.includes('pestel') && charts['pestel'])
            injected.push(['pestel', 15, 8, 'Figure: PESTEL Analysis']);
        if (text.includes('stakeholder') && charts['stakeholder'])
            injected.push(['stakeholder', 14, 8, "Figure: Stakeholder Map"]);
        if ((text.includes('cost') && (text.includes('overrun') || text.includes('escalat'))) && charts['cost_timeline'])
            injected.push(['cost_timeline', 14, 7, 'Figure: Project Cost Escalation Timeline']);
        if ((text.includes('business model') || text.includes('bmc')) && charts['bmc'])
            injected.push(['bmc', 16, 9, 'Figure: Business Model Canvas']);
        return injected;
    }

    try {
        // Cover page
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

        // TOC
        const tocItems = (docSections||[]).map((s,i) => {
            const title = (s.sectionNumber?s.sectionNumber+'  ':'')+s.title;
            return new Paragraph({
                spacing:{line:LINE,before:40,after:40},
                tabStops:[{type:TabStopType.RIGHT,position:9026,leader:'dot'}],
                children:[tr(title,{size:SZS}),new TextRun({text:'\t'+(i+3),font:FONT,size:SZS})]
            });
        });
        const toc = [
            h1('Table of Contents'),
            ...tocItems,
            new Paragraph({children:[new PageBreak()]})
        ];

        // Main content
        const mainContent = [];
        for (const sec of (docSections||[])) {
            const titleLower = (sec.title||'').toLowerCase();
            if (titleLower.includes('references') || titleLower.includes('bibliography')) {
                mainContent.push(new Paragraph({children:[new PageBreak()]}));
            }

            mainContent.push(h1((sec.sectionNumber?sec.sectionNumber+'  ':'')+sec.title));

            // Process content
            const paras = (sec.content||'').split(/\n\n+/).filter(p=>p.trim());
            for (const para of paras) {
                const t = para.trim();
                if (!t) continue;
                // Equation detection
                const eqMatch = t.match(/\[EQUATION\](.*?)\[\/EQUATION\]/s);
                if (eqMatch) {
                    mainContent.push(
                        mkP([tr(eqMatch[1].trim(),{size:SZ,bold:true,color:'1F3864'})],{center:true,before:40,after:40})
                    );
                    continue;
                }
                // Subsection heading
                if (t.match(/^\d+\.\d+\s+\w/) && t.length < 120) {
                    mainContent.push(h2(t));
                } else {
                    mainContent.push(mkP(t));
                }
            }

            // Inject matching charts
            const chartMatches = getChartsForSection(sec.title||'', sec.content||'');
            for (const [name, wCm, hCm, caption] of chartMatches) {
                mainContent.push(...embedChart(name, wCm, hCm, caption));
            }

            mainContent.push(blk());
        }

        // Word count declaration
        mainContent.push(
            new Paragraph({
                border:{top:{style:BorderStyle.SINGLE,size:6,color:'CCCCCC',space:4}},
                spacing:{line:LINE,before:80,after:60},
                children:[tr(`Word Count: ${totalWordCount||0} words (Target: ${targetWordCount||0}). Excludes cover page, TOC, tables, figures, and references.`,{size:SZS,italic:true,color:'666666'})]
            })
        );

        const doc = new Document({
            styles:{
                default:{document:{run:{font:FONT,size:SZ}}},
                paragraphStyles:[
                    {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
                        run:{size:SZH1,bold:true,font:FONT,color:'1F3864'},
                        paragraph:{spacing:{line:LINE,before:280,after:140},outlineLevel:0}},
                    {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
                        run:{size:SZH2,bold:true,font:FONT,color:'2C5282'},
                        paragraph:{spacing:{line:LINE,before:200,after:100},outlineLevel:1}},
                ]
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
            // Return base64 — works on Render ephemeral filesystem
            const b64 = buf.toString('base64');
            res.json({
                stdout: 'SUCCESS',
                stderr: '',
                returncode: 0,
                docxBase64: b64,
                filename: `${(workableTask?.document_type||'Report').replace(/\s+/g,'_')}_completed.docx`
            });
            // Cleanup tmp files
            try { fs.unlinkSync(outputPath); } catch(e) {}
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
        if (fs.existsSync(workDir)) {
            execSync(`rm -rf ${workDir}`, { encoding: 'utf8' });
        }
    } catch(e) {}
    res.json({ stdout: 'CLEANUP_OK', stderr: '', returncode: 0 });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`n8n microservice listening on port ${PORT}`);
});
