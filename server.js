const express = require('express');
const cors = require('cors');
const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

app.get('/', (req, res) => res.json({ status: 'ok', service: 'academic-report-v7' }));
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ── /extract-text ─────────────────────────────────────────────────────────────
app.post('/extract-text', (req, res) => {
    const { fileExt, base64Data, executionId } = req.body;
    if (!base64Data) return res.json({ stdout: '', stderr: 'No data', returncode: 1 });
    const inputPath = `/tmp/brief_${executionId}.${fileExt}`;
    try {
        fs.writeFileSync(inputPath, Buffer.from(base64Data, 'base64'));
        let stdout = '';
        if (fileExt === 'pdf') {
            stdout = execSync(`pdftotext -layout "${inputPath}" -`, { encoding: 'utf8', maxBuffer: 50*1024*1024 });
        } else if (fileExt === 'docx') {
            const py = `/tmp/ext_${executionId}.py`;
            fs.writeFileSync(py, `
from docx import Document
doc = Document('${inputPath}')
lines = [p.text for p in doc.paragraphs if p.text.strip()]
for t in doc.tables:
    for r in t.rows:
        for c in r.cells:
            if c.text.strip(): lines.append(c.text.strip())
print('\\n'.join(lines))
`);
            stdout = execSync(`python3 ${py}`, { encoding: 'utf8', maxBuffer: 50*1024*1024 });
            try { fs.unlinkSync(py); } catch(e) {}
        } else if (fileExt === 'pptx') {
            const py = `/tmp/ext_${executionId}.py`;
            fs.writeFileSync(py, `
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
            stdout = execSync(`python3 ${py}`, { encoding: 'utf8', maxBuffer: 50*1024*1024 });
            try { fs.unlinkSync(py); } catch(e) {}
        } else {
            stdout = Buffer.from(base64Data, 'base64').toString('utf8');
        }
        try { fs.unlinkSync(inputPath); } catch(e) {}
        res.json({ stdout: (stdout||'').trim(), stderr: '', returncode: 0 });
    } catch(e) {
        try { fs.unlinkSync(inputPath); } catch(e2) {}
        res.json({ stdout: '', stderr: e.message||String(e), returncode: 1 });
    }
});

app.post('/deps', (req, res) => res.json({ stdout: 'OK', stderr: '', returncode: 0 }));

// ── /generate-charts — FULLY DYNAMIC ──────────────────────────────────────────
app.post('/generate-charts', (req, res) => {
    const { executionId, docSections, workableTask, fullDocumentText, chartData } = req.body;
    const workDir = `/tmp/charts_${executionId}`;
    fs.mkdirSync(workDir, { recursive: true });

    // Write all data for the python script
    fs.writeFileSync(`${workDir}/sections.json`, JSON.stringify(docSections || []));
    fs.writeFileSync(`${workDir}/task.json`, JSON.stringify(workableTask || {}));
    fs.writeFileSync(`${workDir}/text.json`, JSON.stringify({ text: (fullDocumentText || '').substring(0, 5000) }));
    fs.writeFileSync(`${workDir}/chart_data.json`, JSON.stringify(chartData || { charts: [] }));

    const pyScript = `${workDir}/script.py`;
    fs.writeFileSync(pyScript, `
import sys, os, json, base64, re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

work_dir = "${workDir}"
os.makedirs(work_dir, exist_ok=True)

try:
    with open(f"{work_dir}/chart_data.json") as f: chart_data = json.load(f)
    with open(f"{work_dir}/task.json") as f: task = json.load(f)
    with open(f"{work_dir}/sections.json") as f: sections = json.load(f)
    with open(f"{work_dir}/text.json") as f: full_text = json.load(f).get("text","")
except:
    chart_data = {"charts":[]}; task = {}; sections = []; full_text = ""

# ── Colour palette ────────────────────────────────────────────────────────
NAVY='#1F3864'; BLUE='#2E86AB'; GREEN='#27AE60'; ORANGE='#E67E22'
RED='#C0392B'; PURPLE='#8E44AD'; TEAL='#16A085'; DARK='#2C3E50'; BG='#F8F9FA'

charts_generated = []
case_name = task.get('case_study','') if isinstance(task,dict) else ''

# Also do keyword-based fallback detection from content
all_text = (str(task)+str(sections)+full_text).lower()

def has_kw(*kws):
    return any(k in all_text for k in kws)

# Get AI-provided chart specs
ai_charts = chart_data.get('charts',[]) if isinstance(chart_data,dict) else []
ai_types = {c.get('type'):c for c in ai_charts if c.get('relevant')}

# Also detect from content as fallback
if not ai_types:
    if has_kw('swot','strengths','weaknesses'):
        ai_types['swot'] = {'type':'swot','relevant':True,'title':f'SWOT Analysis','data':{
            'strengths':['Strong market position','Operational capabilities','Financial stability','Brand equity'],
            'weaknesses':['Resource constraints','Limited reach','Complexity','Cost pressures'],
            'opportunities':['Market expansion','Technology adoption','Partnerships','Growth markets'],
            'threats':['Competition','Regulation','Economic uncertainty','Disruption']
        }}
    if has_kw('pestel','pestle'):
        ai_types['pestel'] = {'type':'pestel','relevant':True,'title':'PESTEL Analysis','data':{
            'political':['Policy changes','Trade regulations','Stability'],
            'economic':['Market growth','Inflation','Spending'],
            'social':['Demographics','Culture','Behaviour'],
            'technological':['Digital disruption','Innovation','Automation'],
            'environmental':['Climate','Sustainability','Carbon'],
            'legal':['Compliance','IP','Standards']
        }}
    if has_kw('stakeholder'):
        ai_types['stakeholder'] = {'type':'stakeholder','relevant':True,'title':'Stakeholder Map','data':{
            'center_label':case_name or 'PROJECT','high_influence':['Client/Sponsor','Board/Executive'],
            'medium_influence':['Regulators','Investors','Partners'],
            'low_influence':['Community','Media','End Users']
        }}
    if has_kw('market screen','market select','segmentation','internationalisation','market entry'):
        ai_types['market_screening'] = {'type':'market_screening','relevant':True,'title':'Market Screening Matrix','data':{
            'markets':['Market A','Market B','Market C','Market D','Market E'],
            'criteria':{'Market Size':[8,7,6,9,5],'Growth':[7,8,7,6,9],'Gap':[6,7,8,5,7],'Barriers':[7,6,7,8,6],'Fit':[8,7,6,7,8]}
        }}
    if has_kw("porter's five","porters five","five forces"):
        ai_types['porters'] = {'type':'porters','relevant':True,'title':"Porter's Five Forces",'data':{
            'rivalry':{'level':'HIGH','factors':['Intense competition','Price wars']},
            'new_entrants':{'level':'MEDIUM','factors':['Moderate barriers','Capital needed']},
            'substitutes':{'level':'MEDIUM','factors':['Alternatives exist','Switching costs low']},
            'supplier_power':{'level':'LOW','factors':['Many suppliers','Low differentiation']},
            'buyer_power':{'level':'HIGH','factors':['Price sensitive','Low switching cost']}
        }}
    if has_kw('balanced scorecard','bsc'):
        ai_types['bsc_distribution'] = {'type':'bsc_distribution','relevant':True,'title':'BSC Distribution','data':{
            'financial':8,'customer':7,'internal_process':9,'learning_growth':6
        }}
    if has_kw('cost overrun','cost escalat','budget overrun'):
        ai_types['cost_timeline'] = {'type':'cost_timeline','relevant':True,'title':'Cost Escalation Timeline','data':{
            'years':[2019,2020,2021,2022,2023],'costs':[100,150,200,280,350],'currency':'$','unit':'millions',
            'events':{'2019':'Initial','2023':'Final'},'original_estimate':100
        }}
    if has_kw('risk matrix','risk register','risk assess') and has_kw('project','pmbok','prince'):
        ai_types['risk_matrix'] = {'type':'risk_matrix','relevant':True,'title':'Risk Matrix','data':{
            'risks':[{'name':'Schedule delay','likelihood':4,'impact':5},{'name':'Budget overrun','likelihood':3,'impact':4},
                     {'name':'Scope creep','likelihood':3,'impact':3},{'name':'Technical failure','likelihood':2,'impact':4},
                     {'name':'Force majeure','likelihood':1,'impact':5}]
        }}
    if has_kw('resource loading','gantt','early start','late start','labourer'):
        ai_types['resource_loading'] = {'type':'resource_loading','relevant':True,'title':'Resource Loading','data':{
            'days':list(range(1,25)),'resources':[4,4,4,4,5,5,5,6,6,6,5,5,5,5,5,4,4,4,3,3,3,2,2,1],
            'constraint':6,'unit':'labourers'
        }}


# ── RENDER EACH CHART ─────────────────────────────────────────────────────
for ctype, spec in ai_types.items():
    data = spec.get('data',{})
    title = spec.get('title', ctype.replace('_',' ').title())
    if case_name and case_name not in title:
        title = f"{title} — {case_name}"

    try:
        # ── SWOT ──────────────────────────────────────────────────────────
        if ctype == 'swot':
            swot = {
                'Strengths': data.get('strengths',['Item 1','Item 2','Item 3','Item 4']),
                'Weaknesses': data.get('weaknesses',['Item 1','Item 2','Item 3','Item 4']),
                'Opportunities': data.get('opportunities',['Item 1','Item 2','Item 3','Item 4']),
                'Threats': data.get('threats',['Item 1','Item 2','Item 3','Item 4'])
            }
            fig, axes = plt.subplots(2, 2, figsize=(14, 10))
            fig.patch.set_facecolor(BG)
            cm = {'Strengths':GREEN,'Weaknesses':RED,'Opportunities':BLUE,'Threats':ORANGE}
            am = {'Strengths':axes[0,0],'Weaknesses':axes[0,1],'Opportunities':axes[1,0],'Threats':axes[1,1]}
            for k, items in swot.items():
                c=cm[k]; ax=am[k]
                ax.set_facecolor(c+'10'); ax.set_xlim(0,1); ax.set_ylim(0,1)
                ax.set_xticks([]); ax.set_yticks([])
                for sp in ax.spines.values(): sp.set_edgecolor(c); sp.set_linewidth(2.5)
                ax.text(0.5,0.95,k.upper(),ha='center',va='top',fontsize=13,fontweight='bold',color=c,transform=ax.transAxes)
                ax.axhline(y=0.89,color=c,linewidth=2,alpha=0.5)
                spacing = min(0.14, 0.82/max(len(items),1))
                for i,item in enumerate(items):
                    y = 0.82 - i*spacing
                    if y < 0.02: break
                    ax.text(0.06,y,f'\\u2022 {item}',ha='left',va='top',fontsize=8.5,color=DARK,transform=ax.transAxes)
            plt.suptitle(title, fontsize=12, fontweight='bold', y=0.98, color=NAVY)
            plt.tight_layout(rect=[0,0.02,1,0.96])
            p = f"{work_dir}/swot.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor=BG); plt.close()
            charts_generated.append(('swot',p))

        # ── PESTEL ────────────────────────────────────────────────────────
        elif ctype == 'pestel':
            labels = ['Political','Economic','Social','Technological','Environmental','Legal']
            keys = ['political','economic','social','technological','environmental','legal']
            colors = [BLUE,GREEN,ORANGE,PURPLE,TEAL,RED]
            fig, ax = plt.subplots(figsize=(14, 9))
            ax.set_xlim(-0.5,13.5); ax.set_ylim(-0.5,8.5); ax.axis('off'); fig.patch.set_facecolor('#FDFEFE')
            for idx,(label,key,color) in enumerate(zip(labels,keys,colors)):
                x = 1.0 + idx * 2.2
                items = data.get(key, ['Factor 1','Factor 2','Factor 3'])
                ax.add_patch(mpatches.FancyBboxPatch((x-0.95,0.2),2.0,7.8,boxstyle='round,pad=0.2',
                             facecolor=color+'12',edgecolor=color,linewidth=2))
                ax.text(x+0.05,7.6,label[0],ha='center',va='top',fontsize=18,fontweight='bold',color=color)
                ax.text(x+0.05,6.9,label,ha='center',va='top',fontsize=9,fontweight='bold',color=color)
                ax.axhline(y=6.5,xmin=(x-0.7)/14,xmax=(x+1.1)/14,color=color,alpha=0.5,lw=1.5)
                for i,item in enumerate(items[:4]):
                    ax.text(x+0.05,6.0-i*1.35,item,ha='center',va='top',fontsize=7.5,color=DARK,linespacing=1.3)
            ax.set_title(title,fontsize=12,fontweight='bold',pad=12,color=NAVY)
            plt.tight_layout()
            p = f"{work_dir}/pestel.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('pestel',p))

        # ── STAKEHOLDER MAP ───────────────────────────────────────────────
        elif ctype == 'stakeholder':
            fig, ax = plt.subplots(figsize=(13, 9))
            ax.set_xlim(0,10); ax.set_ylim(0,8.5); ax.axis('off'); fig.patch.set_facecolor('#FDFEFE')
            for r,c,a in [(3.8,NAVY,0.05),(2.7,BLUE,0.06),(1.5,TEAL,0.08)]:
                ax.add_patch(plt.Circle((5,4.2),r,color=c,alpha=a,zorder=1,linestyle='--'))
            cl = data.get('center_label',case_name or 'CORE')
            ax.text(5,4.2,cl.upper(),ha='center',va='center',fontsize=11,fontweight='bold',color='white',zorder=5,
                    bbox=dict(boxstyle='round,pad=0.5',facecolor=NAVY,edgecolor='white',linewidth=2))
            hi = data.get('high_influence',['Key Stakeholder 1','Key Stakeholder 2'])
            mi = data.get('medium_influence',['Stakeholder 3','Stakeholder 4'])
            lo = data.get('low_influence',['Stakeholder 5','Stakeholder 6'])
            positions = [(5,7.6),(5,0.8),(1.2,6.2),(8.8,6.2),(0.8,4.0),(9.2,4.0),(1.5,1.8),(8.5,1.8)]
            all_sh = [(s,RED) for s in hi] + [(s,ORANGE) for s in mi] + [(s,TEAL) for s in lo]
            for i,(name,color) in enumerate(all_sh[:8]):
                px,py = positions[i] if i < len(positions) else (1+i,1)
                ax.annotate('',xy=(5,4.2),xytext=(px,py),arrowprops=dict(arrowstyle='-|>',color=color,alpha=0.35,lw=1.5))
                ax.text(px,py,name,ha='center',va='center',fontsize=8,
                        bbox=dict(boxstyle='round,pad=0.35',facecolor=color+'15',edgecolor=color,linewidth=1.5))
            ax.set_title(title,fontsize=12,fontweight='bold',pad=10,color=NAVY)
            plt.tight_layout()
            p = f"{work_dir}/stakeholder.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('stakeholder',p))

        # ── MARKET SCREENING ──────────────────────────────────────────────
        elif ctype == 'market_screening':
            markets = data.get('markets',['A','B','C','D','E'])
            criteria = data.get('criteria',{'Size':[8,7,6,9,5],'Growth':[7,8,7,6,9]})
            fig, ax = plt.subplots(figsize=(14, 8))
            fig.patch.set_facecolor('#FDFEFE')
            x = np.arange(len(markets))
            w = min(0.15, 0.8/max(len(criteria),1))
            bar_colors = [BLUE,GREEN,ORANGE,PURPLE,RED,TEAL,'#F39C12','#1ABC9C']
            for i,(crit,vals) in enumerate(criteria.items()):
                vals = vals[:len(markets)]  # safety
                while len(vals) < len(markets): vals.append(5)
                ax.bar(x+i*w, vals, w, label=crit, color=bar_colors[i%len(bar_colors)], alpha=0.85, edgecolor='white')
            totals = [sum(criteria[c][j] if j<len(criteria[c]) else 0 for c in criteria) for j in range(len(markets))]
            ax2 = ax.twinx()
            mid = w*len(criteria)/2
            ax2.plot(x+mid, totals, 'D-', color=NAVY, lw=2.5, ms=9, label='Total', zorder=5, mfc='white', mew=2)
            ax2.set_ylabel('Total Score', fontsize=10, color=NAVY, fontweight='bold')
            ax.set_xlabel('Markets', fontsize=10, fontweight='bold')
            ax.set_ylabel('Score (1-10)', fontsize=10, fontweight='bold')
            ax.set_title(title, fontsize=12, fontweight='bold', pad=12, color=NAVY)
            ax.set_xticks(x+mid); ax.set_xticklabels(markets, rotation=25, ha='right', fontsize=9)
            ax.set_ylim(0,11); ax.legend(loc='upper left',fontsize=8)
            ax.grid(axis='y',alpha=0.2,linestyle='--'); ax.spines['top'].set_visible(False)
            plt.tight_layout()
            p = f"{work_dir}/market_screening.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('market_screening',p))

        # ── PORTER'S FIVE FORCES ──────────────────────────────────────────
        elif ctype == 'porters':
            fig, ax = plt.subplots(figsize=(13, 10))
            ax.set_xlim(0,10); ax.set_ylim(0,10); ax.axis('off'); fig.patch.set_facecolor('#FDFEFE')
            force_map = {
                'rivalry':        (5,5,  NAVY,  'Competitive\\nRivalry'),
                'new_entrants':   (5,8.5,GREEN, 'Threat of\\nNew Entrants'),
                'substitutes':    (5,1.5,ORANGE,'Threat of\\nSubstitutes'),
                'supplier_power': (1,5,  BLUE,  'Bargaining Power\\nof Suppliers'),
                'buyer_power':    (9,5,  RED,   'Bargaining Power\\nof Buyers'),
            }
            for key,(px,py,color,label) in force_map.items():
                fd = data.get(key,{})
                level = fd.get('level','MEDIUM') if isinstance(fd,dict) else 'MEDIUM'
                factors = fd.get('factors',['Factor 1','Factor 2']) if isinstance(fd,dict) else ['Factor 1']
                if (px,py)!=(5,5):
                    ax.annotate('',xy=(5,5),xytext=(px,py),arrowprops=dict(arrowstyle='-|>',color=color,lw=2,alpha=0.4))
                bw,bh = 2.8,2.2
                ax.add_patch(mpatches.FancyBboxPatch((px-bw/2,py-bh/2),bw,bh,boxstyle='round,pad=0.2',
                             facecolor=color+'15',edgecolor=color,linewidth=2))
                ax.text(px,py+0.5,label,ha='center',va='center',fontsize=9,fontweight='bold',color=color)
                desc = f"{level}\\n" + "\\n".join(factors[:2])
                ax.text(px,py-0.3,desc,ha='center',va='center',fontsize=7,color=DARK,linespacing=1.3)
            ax.set_title(title,fontsize=12,fontweight='bold',pad=14,color=NAVY)
            plt.tight_layout()
            p = f"{work_dir}/porters.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('porters',p))

        # ── BSC DISTRIBUTION ──────────────────────────────────────────────
        elif ctype == 'bsc_distribution':
            fig, ax = plt.subplots(figsize=(11, 7)); fig.patch.set_facecolor('#FDFEFE')
            persp = ['Financial','Customer','Internal\\nProcess','Learning &\\nGrowth']
            counts = [data.get('financial',8),data.get('customer',7),data.get('internal_process',9),data.get('learning_growth',6)]
            cs = [RED,BLUE,PURPLE,GREEN]
            bars = ax.bar(persp,counts,color=cs,edgecolor='white',linewidth=2,width=0.55,alpha=0.9)
            for b,c in zip(bars,counts):
                ax.text(b.get_x()+b.get_width()/2,b.get_height()+0.2,str(c),ha='center',va='bottom',fontsize=13,fontweight='bold',color=NAVY)
            ax.set_ylabel('Performance Failures',fontsize=11,fontweight='bold',color=DARK)
            ax.set_title(title,fontsize=12,fontweight='bold',pad=14,color=NAVY)
            ax.set_ylim(0,max(counts)+3)
            ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            ax.set_facecolor('#FAFAFA'); ax.grid(axis='y',alpha=0.15,linestyle='--')
            plt.tight_layout()
            p = f"{work_dir}/bsc_dist.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('bsc_dist',p))

        # ── COST TIMELINE ─────────────────────────────────────────────────
        elif ctype == 'cost_timeline':
            years = data.get('years',[2019,2020,2021,2022,2023])
            costs = data.get('costs',[100,150,200,280,350])
            currency = data.get('currency','$')
            unit = data.get('unit','millions')
            events = data.get('events',{})
            orig = data.get('original_estimate',costs[0] if costs else 100)
            fig, ax = plt.subplots(figsize=(13, 7)); fig.patch.set_facecolor('#FDFEFE')
            ax.fill_between(years,costs,alpha=0.08,color=RED)
            ax.plot(years,costs,'o-',color=RED,lw=2.5,ms=9,mfc='white',mew=2.5)
            for yr,label in events.items():
                yr = int(yr) if isinstance(yr,str) else yr
                if yr in years:
                    idx = years.index(yr)
                    ax.annotate(label,xy=(yr,costs[idx]),xytext=(yr,costs[idx]+max(costs)*0.08),
                                ha='center',fontsize=8.5,color=NAVY,fontweight='bold',
                                arrowprops=dict(arrowstyle='->',color=NAVY,lw=1.2),
                                bbox=dict(boxstyle='round,pad=0.25',facecolor='#EBF5FB',edgecolor=NAVY,lw=0.8))
            ax.axhline(y=orig,color=GREEN,ls='--',lw=2,alpha=0.7,label=f'Original ({currency}{orig}{unit[0]})')
            ax.set_xlabel('Year',fontsize=11,fontweight='bold',color=DARK)
            ax.set_ylabel(f'Cost ({currency} {unit})',fontsize=11,fontweight='bold',color=DARK)
            ax.set_title(title,fontsize=12,fontweight='bold',pad=14,color=NAVY)
            ax.legend(fontsize=9); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            ax.set_facecolor('#FAFAFA'); ax.grid(axis='y',alpha=0.15,linestyle='--')
            plt.tight_layout()
            p = f"{work_dir}/cost_timeline.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('cost_timeline',p))

        # ── RISK MATRIX ───────────────────────────────────────────────────
        elif ctype == 'risk_matrix':
            fig, ax = plt.subplots(figsize=(10, 8)); fig.patch.set_facecolor('#FDFEFE')
            ax.set_xlim(0,5); ax.set_ylim(0,5)
            rm = [['#92D050','#92D050','#FFFF00','#FF0000','#FF0000'],
                  ['#92D050','#FFFF00','#FFFF00','#FF0000','#FF0000'],
                  ['#92D050','#FFFF00','#FFFF00','#FFA500','#FF0000'],
                  ['#92D050','#92D050','#FFFF00','#FFA500','#FFA500'],
                  ['#92D050','#92D050','#92D050','#FFFF00','#FFFF00']]
            for i in range(5):
                for j in range(5):
                    ax.add_patch(mpatches.Rectangle((j,4-i),1,1,facecolor=rm[i][j],edgecolor='white',lw=2.5,alpha=0.75))
                    ax.text(j+0.5,4-i+0.5,str((i+1)*(j+1)),ha='center',va='center',fontsize=12,fontweight='bold',color=NAVY)
            risks = data.get('risks',[])
            for r in risks[:6]:
                lk = min(max(r.get('likelihood',3),1),5) - 0.5
                im = min(max(r.get('impact',3),1),5) - 0.5
                ax.plot(lk,im,'o',ms=14,color=NAVY,zorder=5,mec='white',mew=1.5)
                ax.text(lk,im-0.35,r.get('name','Risk'),fontsize=7,color=NAVY,ha='center',va='top',fontweight='bold')
            ax.set_xlabel('Likelihood →',fontsize=11,fontweight='bold')
            ax.set_ylabel('Impact →',fontsize=11,fontweight='bold')
            ax.set_xticks([0.5,1.5,2.5,3.5,4.5]); ax.set_xticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
            ax.set_yticks([0.5,1.5,2.5,3.5,4.5]); ax.set_yticklabels(['Very Low','Low','Medium','High','Very High'],fontsize=8)
            ax.set_title(title,fontsize=12,fontweight='bold',pad=14,color=NAVY)
            plt.tight_layout()
            p = f"{work_dir}/risk_matrix.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('risk_matrix',p))

        # ── RESOURCE LOADING ──────────────────────────────────────────────
        elif ctype == 'resource_loading':
            days = data.get('days',list(range(1,25)))
            resources = data.get('resources',[4,4,4,4,5,5,5,6,6,6,5,5,5,5,5,4,4,4,3,3,3,2,2,1])
            constraint = data.get('constraint',6)
            unit = data.get('unit','labourers')
            fig, ax = plt.subplots(figsize=(14, 6)); fig.patch.set_facecolor('#FDFEFE')
            cols = [RED if r > constraint else BLUE for r in resources]
            ax.bar(days[:len(resources)],resources,color=cols,edgecolor='white',lw=0.8,width=0.8,alpha=0.85)
            ax.axhline(y=constraint,color=RED,ls='--',lw=2,label=f'Max: {constraint} {unit}/day')
            ax.set_xlabel('Day',fontsize=11,fontweight='bold')
            ax.set_ylabel(unit.capitalize(),fontsize=11,fontweight='bold')
            ax.set_title(title,fontsize=12,fontweight='bold',pad=14,color=NAVY)
            ax.legend(fontsize=9); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
            ax.set_facecolor('#FAFAFA'); ax.grid(axis='y',alpha=0.15,linestyle='--')
            plt.tight_layout()
            p = f"{work_dir}/resource_loading.png"
            plt.savefig(p,dpi=180,bbox_inches='tight',facecolor='#FDFEFE'); plt.close()
            charts_generated.append(('resource_loading',p))

    except Exception as e:
        print(f"Chart {ctype} failed: {e}")

manifest = {}
for name,path in charts_generated:
    if os.path.exists(path):
        with open(path,'rb') as f:
            manifest[name] = {'path':path,'b64':base64.b64encode(f.read()).decode()}

with open(f"{work_dir}/manifest.json",'w') as f:
    json.dump({'charts':manifest,'count':len(manifest)},f)

print(f"CHARTS_DONE:{len(manifest)}")
for name in manifest: print(f"CHART:{name}")
`);

    try {
        const stdout = execSync(`python3 ${pyScript}`, {
            encoding: 'utf8', maxBuffer: 100*1024*1024, timeout: 120000
        });
        res.json({ stdout: stdout.trim(), stderr: '', returncode: 0 });
    } catch(e) {
        res.json({ stdout: e.stdout ? e.stdout.trim() : '', stderr: e.message||'', returncode: 1 });
    }
});

// ── /export-docx ──────────────────────────────────────────────────────────────
app.post('/export-docx', (req, res) => {
    const {
        executionId, studentName, studentId, programme, university,
        submissionDate, workableTask, totalWordCount, targetWordCount, docSections,
        fontName, fontSize, lineSpacing, marginsCm
    } = req.body;

    const FONT     = (fontName && fontName.trim()) || 'Arial';
    const FSIZE_HX = Math.round((parseFloat(fontSize) || 12) * 2);
    const LSPACING = Math.round((parseFloat(lineSpacing) || 1.5) * 240);
    const MARGIN   = Math.round((parseFloat(marginsCm) || 2.54) * 567);
    const SZ_H1    = FSIZE_HX + 8;
    const SZ_H2    = FSIZE_HX + 4;
    const SZ_SM    = FSIZE_HX - 4;
    const SZ_XS    = FSIZE_HX - 6;
    const CONTENT_W = 11906 - (MARGIN * 2);

    const workDir = `/tmp/charts_${executionId}`;
    const {
        Document, Packer, Paragraph, TextRun, ImageRun,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
        WidthType, PageNumber, PageBreak,
        Table, TableRow, TableCell, ShadingType
    } = require('docx');

    let charts = {};
    try {
        if (fs.existsSync(`${workDir}/manifest.json`)) {
            charts = JSON.parse(fs.readFileSync(`${workDir}/manifest.json`, 'utf8')).charts || {};
        }
    } catch(e) {}

    const tr = (t, o={}) => new TextRun({
        text: String(t||''), font: FONT,
        size: o.size||FSIZE_HX, bold: o.bold||false,
        italics: o.italic||false, color: o.color||'000000'
    });

    const blk = (b=0,a=0) => new Paragraph({ spacing:{line:LSPACING,before:b,after:a}, children:[tr('')] });

    function mkP(runs, o={}) {
        const align = o.center?AlignmentType.CENTER:o.right?AlignmentType.RIGHT:o.left?AlignmentType.LEFT:AlignmentType.JUSTIFIED;
        return new Paragraph({
            alignment: align,
            spacing: { line:LSPACING, before:o.before||0, after:o.after!==undefined?o.after:160 },
            indent: o.hanging ? {left:720,hanging:720} : undefined,
            children: Array.isArray(runs) ? runs : [tr(runs, o)]
        });
    }

    const h1 = t => new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: {line:LSPACING,before:360,after:200},
        children: [tr(t,{bold:true,size:SZ_H1,color:'1F3864'})]
    });
    const h2 = t => new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: {line:LSPACING,before:260,after:140},
        children: [tr(t,{bold:true,size:SZ_H2,color:'2C5282'})]
    });

    function embedChart(name, wCm, hCm, caption) {
        if (!charts[name]) return [];
        try {
            const buf = Buffer.from(charts[name].b64, 'base64');
            const wPx = Math.round(wCm * 360000 / 9525);
            const hPx = Math.round(hCm * 360000 / 9525);
            return [
                blk(120,0),
                new Paragraph({
                    alignment: AlignmentType.CENTER, spacing:{line:LSPACING,before:0,after:0},
                    children: [new ImageRun({data:Uint8Array.from(buf),transformation:{width:wPx,height:hPx},type:'png'})]
                }),
                mkP([tr(caption,{size:SZ_SM,bold:true,color:'333333'})],{center:true,before:60,after:20}),
                mkP([tr("Source: Author's own analysis.",{size:SZ_XS,italic:true,color:'666666'})],{left:true,before:0,after:200})
            ];
        } catch(e) { return []; }
    }

    const usedCharts = new Set();
    function getChartsForSection(title, content) {
        const text = (title+' '+content).toLowerCase();
        const injected = [];
        const candidates = [
            ['cost_timeline','cost',14,7,'Figure: Cost Escalation Timeline'],
            ['bsc_dist','bsc',11,7,'Figure: BSC Failure Distribution'],
            ['risk_matrix','risk',10,8,'Figure: Risk Assessment Matrix'],
            ['stakeholder','stakeholder',13,9,'Figure: Stakeholder Map'],
            ['swot','swot',14,10,'Figure: SWOT Analysis'],
            ['pestel','pestel',14,9,'Figure: PESTEL Analysis'],
            ['porters','porter',13,10,"Figure: Porter's Five Forces"],
            ['porters','five forces',13,10,"Figure: Porter's Five Forces"],
            ['market_screening','segmentation',14,8,'Figure: Market Screening Matrix'],
            ['market_screening','market',14,8,'Figure: Market Screening Matrix'],
            ['market_screening','screening',14,8,'Figure: Market Screening Matrix'],
            ['resource_loading','resource',14,6,'Figure: Resource Loading — Early Start']
        ];
        for (const [name,kw,w,h,cap] of candidates) {
            if (charts[name] && !usedCharts.has(name) && text.includes(kw)) {
                injected.push([name,w,h,cap]); usedCharts.add(name);
            }
        }
        return injected;
    }

    function renderInlineTable(lines) {
        const rows = lines.filter(l=>l.includes('|'));
        if (rows.length<2) return null;
        const parsed = rows.map(r=>r.split('|').map(c=>c.trim()).filter(c=>c))
            .filter(r=>r.length>0 && !r.every(c=>/^[-:\s]+$/.test(c)));
        if (!parsed.length) return null;
        const colCount = Math.max(...parsed.map(r=>r.length));
        const colW = Math.floor(CONTENT_W / colCount);
        const bdr = {style:BorderStyle.SINGLE,size:1,color:'BDC3C7'};
        const borders = {top:bdr,bottom:bdr,left:bdr,right:bdr};
        return new Table({
            width:{size:CONTENT_W,type:WidthType.DXA},
            columnWidths: Array(colCount).fill(colW),
            rows: parsed.map((row,ri) => new TableRow({
                children: Array.from({length:colCount},(_,ci) => {
                    const cell = row[ci]||'';
                    const isH = ri===0;
                    return new TableCell({
                        width:{size:colW,type:WidthType.DXA},
                        margins:{top:60,bottom:60,left:100,right:100},
                        shading:{fill:isH?'1F3864':(ri%2===0?'F2F6FC':'FFFFFF'),type:ShadingType.CLEAR},
                        borders,
                        children:[new Paragraph({
                            spacing:{line:240,before:0,after:0},
                            alignment:isH?AlignmentType.CENTER:AlignmentType.LEFT,
                            children:[tr(cell,{bold:isH,color:isH?'FFFFFF':'2C3E50',size:SZ_SM})]
                        })]
                    });
                })
            }))
        });
    }

    function parseContent(raw) {
        const els = [];
        for (const para of (raw||'').split(/\n\n+/)) {
            const t = para.trim();
            if (!t) continue;
            if (/^Equation\s+\d+\s*[—–\-]/i.test(t) || t.startsWith('[EQUATION]')) {
                els.push(new Paragraph({
                    alignment:AlignmentType.CENTER, spacing:{line:LSPACING,before:160,after:160},
                    border:{top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC',space:4},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC',space:4}},
                    children:[tr(t.replace(/^\[EQUATION\]/,'').replace(/\[\/EQUATION\]$/,'').trim(),{bold:true,size:FSIZE_HX,color:'1F3864'})]
                })); continue;
            }
            if (t.includes('|') && t.split('\n').filter(l=>l.includes('|')).length>=2) {
                const tbl = renderInlineTable(t.split('\n'));
                if (tbl) { els.push(tbl); els.push(blk(60,60)); continue; }
            }
            if (t.startsWith('## ')) { els.push(h2(t.replace(/^##\s*/,''))); continue; }
            if (/^\d+\.\d+[\s]+\w/.test(t) && t.length<120) { els.push(h2(t)); continue; }
            if (/^Source:/i.test(t)) { els.push(mkP([tr(t,{size:SZ_XS,italic:true,color:'666666'})],{left:true,before:0,after:120})); continue; }
            if (/^(Table|Figure)\s+\d+/i.test(t) && t.length<200) { els.push(mkP([tr(t,{size:SZ_SM,bold:true,color:'333333'})],{left:true,before:80,after:40})); continue; }
            els.push(mkP(t));
        }
        return els;
    }

    try {
        const sv = v => (v && v!=='N/A' && v!=='Student' && v!=='Not Available') ? v : '';
        const isAnon = !sv(studentName);
        const displayName = isAnon ? 'Anonymous Submission' : sv(studentName);
        const docType = workableTask?.document_type || 'Individual Report';
        const caseStudy = workableTask?.case_study || '';
        const subjectArea = workableTask?.subject_area || programme || '';
        const acadLevel = workableTask?.academic_level || 'Level 7';
        const modName = workableTask?.module_name || subjectArea;
        const modLead = workableTask?.module_lead || '';
        const schoolDept = workableTask?.school_department || '';
        const refStyle = workableTask?.referencing_style || 'Harvard';

        // ── COVER PAGE ────────────────────────────────────────────────
        const cover = [
            blk(), blk(),
            mkP([tr((university||'UNIVERSITY').toUpperCase(),{bold:true,size:SZ_H1+6,color:'1F3864'})],{center:true,after:40}),
            ...(schoolDept ? [mkP([tr(schoolDept,{size:FSIZE_HX,color:'444444'})],{center:true,after:200})] : [blk(0,160)]),
            blk(), blk(),
            mkP([tr(docType.toUpperCase(),{bold:true,size:SZ_H1+10,color:'1F3864'})],{center:true,after:60}),
            ...(caseStudy ? [mkP([tr(caseStudy,{bold:true,size:SZ_H1+2,color:'2C5282'})],{center:true,after:80})] : []),
            blk(), blk(), blk(), blk(),
            mkP([tr(`Module: ${modName}  |  ${acadLevel}`,{size:FSIZE_HX-2,color:'333333'})],{center:true,after:30}),
            ...(modLead ? [mkP([tr(`Module Lead: ${modLead}`,{size:FSIZE_HX-2,color:'333333'})],{center:true,after:30})] : []),
            mkP([tr(`Submission: ${sv(submissionDate)||''}  |  ${displayName}`,{size:FSIZE_HX-2,color:'333333'})],{center:true,after:30}),
            ...(sv(studentId) ? [mkP([tr(`Student ID: ${studentId}`,{size:FSIZE_HX-2,color:'333333'})],{center:true,after:30})] : []),
            blk(),
            mkP([tr(`Word Count: ${totalWordCount?totalWordCount.toLocaleString():'0'} (excluding cover, TOC, tables, figures, and references)`,
                {size:FSIZE_HX-2,bold:true,color:'333333'})],{center:true,after:40}),
            blk(),
            mkP([tr(`${refStyle} Referencing  |  ${FONT} ${FSIZE_HX/2}pt  |  ${parseFloat(lineSpacing)||1.5} Line Spacing  |  ${parseFloat(marginsCm)||2.54}cm margins  |  A4`,
                {size:SZ_XS,italic:true,color:'AAAAAA'})],{center:true,after:0}),
            new Paragraph({ children: [new PageBreak()] })
        ];

        // ── TOC ───────────────────────────────────────────────────────
        let TSType, TSLeader;
        try { ({TabStopType:TSType, TabStopLeader:TSLeader} = require('docx')); } catch(e) {}
        const tocItems = (docSections||[]).map(s => {
            const label = (s.sectionNumber?s.sectionNumber+'  ':'') + (s.title||'');
            return new Paragraph({
                spacing:{line:LSPACING,before:40,after:40},
                tabStops: TSType ? [{type:TSType.RIGHT,position:Math.round(CONTENT_W*0.95),leader:TSLeader?TSLeader.DOT:3}] : [],
                children: [tr(label,{size:FSIZE_HX}), new TextRun({text:'\t',font:FONT,size:FSIZE_HX})]
            });
        });

        // ── MAIN CONTENT ──────────────────────────────────────────────
        const main = [];
        for (const sec of (docSections||[])) {
            const tl = (sec.title||'').toLowerCase();
            if (tl.includes('references')||tl.includes('bibliography'))
                main.push(new Paragraph({children:[new PageBreak()]}));
            main.push(h1((sec.sectionNumber?sec.sectionNumber+'  ':'')+sec.title));
            if (tl.includes('references')||tl.includes('bibliography')) {
                for (const ref of (sec.content||'').split(/\n+/).filter(l=>l.trim()))
                    main.push(mkP([tr(ref.trim(),{size:FSIZE_HX})],{hanging:true,before:0,after:120}));
            } else {
                main.push(...parseContent(sec.content||''));
            }
            for (const [n,w,h,c] of getChartsForSection(sec.title||'',sec.content||''))
                main.push(...embedChart(n,w,h,c));
            main.push(blk());
        }

        const doc = new Document({
            styles: {
                default:{document:{run:{font:FONT,size:FSIZE_HX}}},
                paragraphStyles: [
                    {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
                     run:{size:SZ_H1,bold:true,font:FONT,color:'1F3864'},
                     paragraph:{spacing:{before:360,after:200},outlineLevel:0}},
                    {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
                     run:{size:SZ_H2,bold:true,font:FONT,color:'2C5282'},
                     paragraph:{spacing:{before:260,after:140},outlineLevel:1}}
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size:{width:11906,height:16838},
                        margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}
                    }
                },
                headers:{default:new Header({children:[new Paragraph({
                    alignment:AlignmentType.RIGHT,
                    border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'1F3864',space:4}},
                    spacing:{before:0,after:80},
                    children:[tr(`${sv(university)||''}  |  ${docType}`,{size:SZ_XS,italic:true,color:'666666'})]
                })]})},
                footers:{default:new Footer({children:[new Paragraph({
                    alignment:AlignmentType.CENTER,
                    border:{top:{style:BorderStyle.SINGLE,size:4,color:'1F3864',space:4}},
                    spacing:{before:80,after:0},
                    children:[
                        tr('Page ',{size:SZ_XS,color:'888888'}),
                        new TextRun({children:[PageNumber.CURRENT],font:FONT,size:SZ_XS,color:'888888'}),
                        tr(' of ',{size:SZ_XS,color:'888888'}),
                        new TextRun({children:[PageNumber.TOTAL_PAGES],font:FONT,size:SZ_XS,color:'888888'})
                    ]
                })]})},
                children: [...cover, h1('Table of Contents'), ...tocItems, new Paragraph({children:[new PageBreak()]}), ...main]
            }]
        });

        Packer.toBuffer(doc).then(buf => {
            const fn = docType.replace(/[^a-zA-Z0-9_\s]/g,'').replace(/\s+/g,'_');
            res.json({stdout:'SUCCESS',stderr:'',returncode:0,docxBase64:buf.toString('base64'),filename:`${fn}_completed.docx`});
            try { fs.rmSync(workDir,{recursive:true,force:true}); } catch(e) {}
        }).catch(err => res.json({stdout:'',stderr:err.message,returncode:1}));

    } catch(e) {
        res.json({stdout:'',stderr:e.message,returncode:1});
    }
});

app.post('/cleanup', (req, res) => {
    try { const d=`/tmp/charts_${req.body.executionId}`; if(fs.existsSync(d)) fs.rmSync(d,{recursive:true,force:true}); } catch(e) {}
    res.json({stdout:'OK',stderr:'',returncode:0});
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Academic Report Service v7 on port ${PORT}`));
